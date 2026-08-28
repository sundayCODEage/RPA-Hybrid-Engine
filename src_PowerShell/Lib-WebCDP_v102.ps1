# ------------------------------------------------------------------------------
# CDP (Chrome DevTools Protocol) 通信モジュール
# WebSocket経由による高速・低レイテンシのDOM操作の実現
# ------------------------------------------------------------------------------

# グローバル変数の定義
$global:CdpWebSocket = $null
$global:CdpMessageId = 0
$global:CdpEndpoint  = ""

# --- CDPセッションの確立 --- [// 内部関数 //]
function Connect-CdpSession {
    param ([int]$Port = 9222)

    $func = $MyInvocation.MyCommand.Name
    Write-DebugLog -Message "[$func] 開始: CDPセッションの確立 (Port: $Port)" -Level Info

    # 起動直後のCOM例外防止（ActiveTabId未設定時のスキップ）
    if (-not $global:ActiveTabId -or -not $global:Tabs.ContainsKey($global:ActiveTabId)) {
        Write-DebugLog -Message "[$func] ActiveTabId 未設定のため CDP 接続をスキップ" -Level Warning
        return
    }

    # アクティブタブ情報の取得
    $activeTab    = $global:Tabs[$global:ActiveTabId]
    $activeUrl    = if ($activeTab.Url) { $activeTab.Url } else { "" }
    $activeTarget = $activeTab.TargetId

    Write-DebugLog -Message "[$func] アクティブURL = '$activeUrl'" -Level Info
    if ($activeTarget) {
        Write-DebugLog -Message "[$func] アクティブTargetId = $activeTarget" -Level Info
    } else {
        Write-DebugLog -Message "[$func] TargetId 未設定 → URL マッチにフォールバック" -Level Info
    }

    # CDP/jsonからのpageターゲット探索
    $jsonUrl = "http://127.0.0.1:$Port/json"
    $pageTarget = $null

    for ($i = 1; $i -le 10; $i++) {
        try {
            $pages = Invoke-RestMethod -Uri $jsonUrl -UseBasicParsing -ErrorAction Stop

            if ($activeTarget) {
                # 1. TargetIdの優先
                $pageTarget = $pages | Where-Object {
                    $_.id -eq $activeTarget -and
                    -not [string]::IsNullOrEmpty($_.webSocketDebuggerUrl)
                } | Select-Object -First 1
            } 
            
            # 2. URLによる柔軟一致 (TargetIdがない、またはマッチしなかった場合)
            if (-not $pageTarget -and -not [string]::IsNullOrWhiteSpace($activeUrl)) {
                $normActive = $activeUrl.TrimEnd('/')
                $pageTarget = $pages | Where-Object {
                    $_.type -eq 'page' -and
                    ($_.url.TrimEnd('/') -eq $normActive -or $_.url -like "$normActive*" -or $normActive -like "$($_.url.TrimEnd('/'))*") -and
                    -not [string]::IsNullOrEmpty($_.webSocketDebuggerUrl)
                } | Select-Object -First 1
            }

            # 3. フェイルセーフ (単一Pageターゲットの自動選択/強制取得)
            if (-not $pageTarget) {
                # WebSocketが有効な page タイプのターゲットをすべて取得
                $availablePages = $pages | Where-Object { 
                    $_.type -eq 'page' -and -not [string]::IsNullOrEmpty($_.webSocketDebuggerUrl) 
                }
                
                if ($availablePages -and $availablePages.Count -ge 1) {
                    # 複数見つかっても、とりあえず最初の有効なページに強制接続してクラッシュを防ぐ
                    $pageTarget = $availablePages[0]
                    Write-DebugLog -Message "[$func] 警告: URL/TargetIdの一致に失敗しましたが、利用可能なPageターゲットへ強制接続しました (Target: $($pageTarget.url))" -Level Warning
                }
            }

            if ($pageTarget) { break }
        } catch {}

        Start-Sleep -Milliseconds 500
    }

    if (-not $pageTarget) {
        throw (New-EngineException -Func $func -Type "CDPエラー" -Message "アクティブタブに一致するCDPターゲットが見つかりませんでした")
    }

    # WebSocketDebuggerUrlの採用
    $global:CdpEndpoint = $pageTarget.webSocketDebuggerUrl
    Write-DebugLog -Message "[$func] 情報: WebSocket URL取得 ($($global:CdpEndpoint))" -Level Info

    try {
        # 既存WebSocketの安全破棄
        if ($global:CdpWebSocket -and $global:CdpWebSocket.State -ne 'Closed') {
            try { $global:CdpWebSocket.Abort() } catch {}
            try { $global:CdpWebSocket.Dispose() } catch {}
            Start-Sleep -Milliseconds 100
        }

        # 新規WebSocketの接続
        $global:CdpWebSocket = New-Object System.Net.WebSockets.ClientWebSocket
        $uri = [System.Uri]::new($global:CdpEndpoint)
        $cts = [System.Threading.CancellationTokenSource]::new()

        $task = $global:CdpWebSocket.ConnectAsync($uri, $cts.Token)
        while (-not $task.IsCompleted) {
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 10
        }

        if ($global:CdpWebSocket.State -ne 'Open') {
            throw (New-EngineException -Func $func -Type "CDPエラー" -Message "WebSocketの接続がOpen状態になりません")
        }

        $global:CdpMessageId = 0
        $global:ConnectedCdpTabId = $global:ActiveTabId
        Write-DebugLog -Message "[$func] 情報: CDP接続完了 (Target: $($pageTarget.title))" -Level Info

    } catch {
        $global:CdpWebSocket = $null
        throw (New-EngineException -Func $func -Type "CDPエラー" -Message "WebSocketの接続に失敗しました" -Details $_.Exception.Message)
    }
}

# --- CDPコマンドの送受信 ---
function Invoke-CdpCommand {
    param (
        [Parameter(Mandatory = $true)][string]$Method,
        [hashtable]$Params = @{},
        [int]$MaxRetries = 3,
        [int]$TimeoutSecPerTry = 4
    )

    $func = $MyInvocation.MyCommand.Name
    $attempt = 0

    while ($attempt -lt $MaxRetries) {
        $attempt++

        # WebSocketの接続状態確認および自動再接続
        if ($null -eq $global:CdpWebSocket -or
            $global:CdpWebSocket.State -ne [System.Net.WebSockets.WebSocketState]::Open -or
            $global:ConnectedCdpTabId -ne $global:ActiveTabId) {

            if ($null -eq $global:CdpWebSocket) {
                Write-DebugLog -Message "[$func] 情報: 初回CDPセッション接続要求" -Level Info
            } elseif ($global:ConnectedCdpTabId -ne $global:ActiveTabId) {
                Write-DebugLog -Message "[$func] 情報: アクティブタブ変更検知 → CDP再接続 ($global:ActiveTabId)" -Level Info
            } else {
                Write-DebugLog -Message "[$func] 警告: WebSocket切断検知 → 自動再接続" -Level Warning
            }

            try {
                Connect-CdpSession -Port $global:CdpPort
            } catch {
                throw (New-EngineException -Func $func -Type "CDPエラー" -Message "WebSocketの接続または再接続に失敗しました" -Details $_.Exception.Message)
            }
        }

        # ----------------------------------------------------------
        # JSON-RPCメッセージの構築
        <# CDPの仕様に従い、一意のID(スレッドセーフで採番)とメソッドを紐付けたJSONを生成 #>
        $msgId = [System.Threading.Interlocked]::Increment([ref]$global:CdpMessageId)
        $payload = @{ id = $msgId; method = $Method; params = $Params }
        $jsonString = $payload | ConvertTo-Json -Depth 5 -Compress

        $sendBytes  = [System.Text.Encoding]::UTF8.GetBytes($jsonString)
        $sendBuffer = [System.ArraySegment[byte]]::new($sendBytes)

        $cts = [System.Threading.CancellationTokenSource]::new()
        $timeoutTime = (Get-Date).AddSeconds($TimeoutSecPerTry)

        try {
            # ----------------------------------------------------------
            # 送信処理
            <# 非同期(SendAsync)で送信しつつ、完了待機中はWinFormsのDoEventsを回すことで
               ブラウザ画面（UIスレッド）のフリーズを防止する #>
            $sendTask = $global:CdpWebSocket.SendAsync(
                $sendBuffer,
                [System.Net.WebSockets.WebSocketMessageType]::Text,
                $true,
                $cts.Token
            )

            while (-not $sendTask.IsCompleted) {
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 5
                if ((Get-Date) -gt $timeoutTime) {
                    throw [System.TimeoutException]::new("CDP_TIMEOUT")
                }
            }

            # ----------------------------------------------------------
            # 受信処理
            <# WebSocketはデータ長が大きい場合（DOM全体取得など）にパケットを分割して
               送信してくるため、EndOfMessageがTrueになるまでバッファを結合し続ける #>
            $receiveBuffer  = [byte[]]::new(16384)
            $receiveSegment = [System.ArraySegment[byte]]::new($receiveBuffer)

            while ($true) {
                $responseString = ""

                while ($true) {
                    $receiveTask = $global:CdpWebSocket.ReceiveAsync($receiveSegment, $cts.Token)

                    # 受信完了までのUIフリーズ防止およびタイムアウト監視
                    while (-not $receiveTask.IsCompleted) {
                        [System.Windows.Forms.Application]::DoEvents()
                        Start-Sleep -Milliseconds 5
                        if ((Get-Date) -gt $timeoutTime) {
                            throw [System.TimeoutException]::new("CDP_TIMEOUT")
                        }
                    }

                    if ($receiveTask.IsFaulted -or $receiveTask.IsCanceled) {
                        throw [System.Exception]::new("TASK_ERROR")
                    }

                    $result = $receiveTask.Result

                    if ($result.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) {
                        throw [System.Exception]::new("WEBSOCKET_CLOSED")
                    }

                    # 受信したバイト列を文字列として結合
                    $responseString += [System.Text.Encoding]::UTF8.GetString(
                        $receiveBuffer, 0, $result.Count
                    )

                    # 分割送信の終端に達したら受信ループを抜ける
                    if ($result.EndOfMessage) { break }
                }

                # ----------------------------------------------------------
                # JSONのパース処理とID突合
                <# 非同期通信の性質上、過去にタイムアウトした別の応答が遅れて
                   届く可能性があるため、送信時の msgId と完全に一致する応答のみを採用する #>
                try {
                    $responseObj = $responseString | ConvertFrom-Json -ErrorAction Stop
                } catch {
                    continue
                }

                # idが一致した応答の返却
                if ($null -ne $responseObj.id -and $responseObj.id -eq $msgId) {
                    if ($null -ne $responseObj.error) {
                        throw (New-EngineException -Func $func -Type "CDPエラー" -Message $($responseObj.error.message) -Details "Code: $($responseObj.error.code)")
                    }
                    return $responseObj
                }
            }
        }
        catch {
            $errMsg = $_.Exception.Message

            if ($errMsg -eq "CDP_TIMEOUT") {
                $cts.Cancel()

                if ($attempt -lt $MaxRetries) {
                    Write-DebugLog -Message "[$func] 警告: 応答タイムアウト → リトライ ($attempt/$MaxRetries)" -Level Warning
                    continue
                } else {
                    throw (New-EngineException -Func $func -Type "Timeout" -Message "CDPコマンドの応答がタイムアウトしました" -Details "全 $($MaxRetries) 回リトライ失敗")
                }
            } else {
                throw (New-EngineException -Func $func -Type "CDPエラー" -Message "CDPコマンドの送受信中に予期せぬ例外が発生しました" -Details $_.Exception.Message)
            }
        }
    }
}

# --- CDP経由でのJavaScript実行 ---
function Invoke-CdpScript {
    param ([Parameter(Mandatory = $true)][string]$Js)

    $func = $MyInvocation.MyCommand.Name
    $jsCode = $Js

    # JSの安全なラップ
    <# 任意のJSコードを即時実行関数(IIFE)で包み、例外クラッシュの防止と、
       PowerShell側への確実な型伝達(JSON化)を保証する #>
    $wrappedJs = @"
(function() {
    try {
        var result = (function() { $jsCode })();

        // JavaScriptの undefined をそのまま返すとJSON.stringifyで
        // キーごと消失してしまうため、意図的に null に置換してデータ欠落を防ぐ
        return JSON.stringify({
            status: "success",
            data: result !== undefined ? result : null
        });
    } catch (e) {
        // JS内で発生したエラーをキャッチし、PowerShell側(New-EngineException)へ
        // 詳細なスタックトレースを引き渡すための規格化
        return JSON.stringify({
            status: "error",
            message: e && e.message ? e.message : String(e),
            stack: e && e.stack ? e.stack : "",
            name: e && e.name ? e.name : ""
        });
    }
})();
"@

    $params = @{
        expression    = $wrappedJs
        returnByValue = $true
        awaitPromise  = $true
    }

    $res = Invoke-CdpCommand -Method "Runtime.evaluate" -Params $params

    # JS例外の判定
    if ($null -ne $res.result.exceptionDetails) {
        $ex = $res.result.exceptionDetails
        throw (New-EngineException -Func $func -Type "JSエラー" -Message "JavaScript実行時例外が発生しました" -Details $ex.exception.description)
    }

    $rawResult = $res.result.result.value

    if ([string]::IsNullOrWhiteSpace($rawResult) -or $rawResult -eq "null") {
        throw (New-EngineException -Func $func -Type "内部エラー" -Message "CDPスクリプト実行の戻り値が空またはnullです")
    }

    # 文字列JSONのアンエスケープ
    if ($rawResult -match '^\s*"(.*)"\s*$') {
        $inner = $matches[1]
        try {
            $rawResult = [System.Text.RegularExpressions.Regex]::Unescape($inner)
        } catch {
            $rawResult = $inner
        }
    }

    # JSONのデコード処理
    try {
        $response = $rawResult | ConvertFrom-Json -ErrorAction Stop
    } catch {
        throw (New-EngineException -Func $func -Type "内部エラー" -Message "CDPからの応答(JSON)のデコードに失敗しました")
    }

    if ($response.status -ne "success") {
        throw (New-EngineException -Func $func -Type "JSエラー" -Message "$($response.name): $($response.message)" -Details $response.stack)
    }

    return $response.data
}

# --- ネイティブクリックの発火 ---
function Invoke-CdpNativeClick {
    param ([Parameter(Mandatory = $true)][string]$Selector)

    $func = $MyInvocation.MyCommand.Name
    $selectorEscaped = $Selector.Replace("'", "\'")

    # 要素の中央座標の取得
    $js = @"
    const el = document.querySelector('$selectorEscaped');
    if (!el) {
        throw new Error("Element not found: $selectorEscaped");
    }
    // 要素を画面の中央にスクロールして確実に見える状態にする
    el.scrollIntoView({ block: 'center', inline: 'center' });
    // 要素の矩形情報（BoundingBox）を取得
    const rect = el.getBoundingClientRect();
    // 中央座標を計算して PowerShell 側へ返却する
    return { 
        x: rect.left + (rect.width / 2), 
        y: rect.top + (rect.height / 2) 
    };
"@

    $res = Invoke-CdpScript -Js $js

    $x = [Math]::Round($res.x)
    $y = [Math]::Round($res.y)

    # マウスダウンおよびマウスアップの発火
    Invoke-CdpCommand -Method "Input.dispatchMouseEvent" -Params @{
        type       = "mousePressed"
        button     = "left"
        clickCount = 1
        x          = $x
        y          = $y
    } | Out-Null

    Invoke-CdpCommand -Method "Input.dispatchMouseEvent" -Params @{
        type       = "mouseReleased"
        button     = "left"
        clickCount = 1
        x          = $x
        y          = $y
    } | Out-Null
}

# --- ネイティブテキストの入力 ---
function Set-CdpNativeTextInput {
    param (
        [Parameter(Mandatory = $true)][string]$Selector,
        [Parameter(Mandatory = $true)][string]$Value
    )

    $func = $MyInvocation.MyCommand.Name
    $selectorEscaped = $Selector.Replace("'", "\'")

    # 対象要素へのフォーカスおよび値のクリア
    $js = @"
    const el = document.querySelector('$selectorEscaped');
    if (!el) {
        throw new Error("Element not found: $selectorEscaped");
    }
    // ネイティブ入力のために要素へフォーカスを当てる
    el.focus();
    // 既存の入力値をクリアする
    el.value = '';
    return true;
"@

    Invoke-CdpScript -Js $js | Out-Null

    # ネイティブ文字入力の実行
    Invoke-CdpCommand -Method "Input.insertText" -Params @{ text = $Value } | Out-Null
}
