# ------------------------------------------------------------------------------
# 自立型 WebView2 RPAエンジン（Core）
# コマンド受信および各モジュールへのルーティング
# advice by AI (2026/05...)
# ------------------------------------------------------------------------------

<#
汎用RPAエンジン 命名規則およびコーディングルール
• 実行環境: PowerShell 5.1
【1. 変数およびパラメータの基本命名規則】
• ローカル変数: camelCase（キャメルケース）
   ※ ローカル変数の解釈拡大:
      関数内の変数だけでなく、「モジュール読み込み（ドットソース展開）時の初期化処理で
      のみ使用し、以降の処理で状態として参照・操作しない作業用変数」もローカル変数と
      同義として扱い、camelCase を適用する。
   ※ パラメータ名（PascalCase）をローカルで加工・再利用する際は、必ず camelCase の別名とする。
      [NG]: $XPathEscaped = Normalize-XPath $XPath
      [OK]: $xpathEscaped = Normalize-XPath $XPath

• パラメータおよびグローバル変数: PascalCase（パスカルケース）
• 定数・システム設定値（例外ルール）:
   $global:CONFIG のように、定数・設定情報として機能する変数は大文字を許容する。

【2. 時間・タイムアウト系の命名規則（厳格化）】
• 時間や間隔を指定する変数は、必ず「単位（Sec/Ms）」を接尾辞として明記する。
   ※ Timeout（単位不明）等の曖昧な命名は使用禁止。
   - 秒単位の例: TimeoutSec, WaitTimeSec
   - ミリ秒単位の例: TimeoutMs, PollIntervalMs, InitialDelayMs, WaitMs
   ※ これらも「基本命名規則」に従い、パラメータなら $TimeoutSec、ローカルなら $timeoutSec となる。

【3. 異言語間（VBA / PowerShell / JavaScript）連携の命名規則】
• [VBA → PowerShell] 真偽値（Boolean）のコマンドライン引数渡し:
   プロセス間通信において、真偽値を文字列型（"true"/"false"）で受け取る起動引数（param）には、
   内部で保持する Boolean 変数と区別するため、意図的に Flg という接尾辞を付与する。
   - 例: [string]$IsDebugModeFlg

• [PowerShell → JavaScript] 埋め込み用変数の統一:
   PowerShellからJSのヒアドキュメント（@""@）に展開・埋め込むための加工済み変数は、
   元のパラメータが PascalCase であっても、必ず camelCase に変換して埋め込む。
   - 例: $XPath (パラメータ) → $xpathEscaped (JS埋め込み用ローカル変数)
   - 例: $FrameSelector → $frameSelectorEscaped

• [JavaScript 内部コード] 変数名の独立と統一:
   JSコード内で宣言する変数（var/let/const）は、JavaScriptの標準規約に従い
   camelCase を徹底する。PowerShell側の PascalCase 名をJS内に持ち込まない。
   - [NG]: const XPath = '$xpathEscaped';
   - [OK]: const xpath = '$xpathEscaped';
#>

<#
  ■ 外部関数 (External / Public API)
    VBA側の `rpaEngine.RunAction` から直接 `Command` として呼び出されることを想定した公開インターフェース
    (例: Invoke-WebClick, Set-WebTextInput, Enable-SilentDownload 等)
    ===========================================================================
  ■ 内部関数 (Internal Functions) 】
    
    [ログ・例外ハンドリング]
    - Write-DebugLog        : 実行ログ、デバッグ情報のコンソール出力およびファイル書き込み
    - New-EngineException   : 例外情報の標準フォーマット化（VBAのParseRpaErrorと対をなす）
    
    [待機・同期ユーティリティ]
    - Wait-Condition        : 任意の条件スクリプトブロックが $true になるまでの汎用ポーリング待機
    
    [文字列・セレクタ解析]
    - Normalize-XPath       : HTML上の改行や空白による不一致を防ぐためのXPath自動補正
    
    [コア通信・セッション制御]
    - Get-ActiveWebView     : 現在操作対象のアクティブなWebView2インスタンスを取得
    - Connect-CdpSession    : CDP (Chrome DevTools Protocol) のWebSocket接続を確立・維持
    - Invoke-WebScript      : JS実行の最上位窓口 (Native/CDPのルーティングと型統一ラッパーを適用)
        ├─ Invoke-CdpScript           : (内部) CDP経由でのJS評価
        └─ Invoke-WebView2NativeScript: (内部) Native(ExecuteScriptAsync)でのJS評価   
   
    [ウィンドウ・OSレベル制御]
    - utilFindInFrames      : (JS内部関数) 多段iframeを透過探索する再帰ロジック
    - deepQuerySelector     : (JS内部関数) Shadow DOMを貫通探索する再帰ロジック
    ===========================================================================
#>

# 起動引数の受け取り
param (
    [string]$AppSessionId = "AUTO_$(Get-Date -Format 'yyyyMMddHHmm')",
    [int]$CdpPort = 0,
    [string]$IsDebugModeFlg = "true"
)

# 起動引数のグローバル保持
$global:AppSessionId = $AppSessionId
$global:CdpPort = $CdpPort
try {
    $global:IsDebugMode = [bool]::Parse($IsDebugModeFlg)
} catch {
    $global:IsDebugMode = $true
}

#✘ Add-Type -AssemblyName System.Windows.Forms
# .NETのUIアセンブリを最優先でロードし、DPI(画面拡大率)の基準を強制統一する
Add-Type -AssemblyName System.Windows.Forms, WindowsBase, UIAutomationClient, UIAutomationTypes

try {
    # 画面サイズを取得する前にWPFのシステムを目覚めさせ、DPIを真の物理サイズに固定する
    $dummyPoint = New-Object System.Windows.Point(0,0)
    [System.Windows.Automation.AutomationElement]::FromPoint($dummyPoint) | Out-Null
} catch {}

# プライマリモニターの作業領域（タスクバーを除外した有効領域）を取得
$screenWidth  = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height

# 共通設定値の初期化
$global:CONFIG = @{
    MaxLogGenerations = 5
    DefaultTimeoutSec = 10
    BrowserWidth      = $screenWidth
    BrowserHeight     = $screenHeight
    EnableHighlight   = $true
}

# 通信モードに応じたポーリング間隔の設定
if ($global:CdpPort -gt 0) {
    # 高速ポーリング（CDPモード）
    $global:CONFIG.PollIntervalMs = 50
} else {
    # UIスレッド負荷軽減（ネイティブモード）
    $global:CONFIG.PollIntervalMs = 200
}

# システム用グローバル変数の初期化
$global:BrowserForm = $null
$global:WebViewCtrl = $null

# 多段iframeおよびShadow DOM探索用JS共通ユーティリティの定義
$global:ENGINE_JS_UTILS = @"
// --- Shadow DOM 貫通検索用関数 ---
function deepQuerySelector(selector, root) {
    root = root || document;
    var res = root.querySelector(selector);
    if (res) return res;
    var els = root.querySelectorAll('*');
    for (var i = 0; i < els.length; i++) {
        if (els[i].shadowRoot) {
            res = deepQuerySelector(selector, els[i].shadowRoot);
            if (res) return res; // 見つかったら即座に返す
        }
    }
    return null;
}

// --- iframe透過探索用関数 ---
function utilFindInFrames(win, checkFn) {
    try {
        var res = checkFn(win);
        if (res !== null && res !== false && res !== undefined) return res;
    } catch(e) {}
    for (var i = 0; i < win.frames.length; i++) {
        try {
            var found = utilFindInFrames(win.frames[i], checkFn);
            if (found !== null && found !== false && found !== undefined) return found;
        } catch(e) {}
    }
    return null;
}
"@

# タブ管理構造体の初期化
$global:Tabs = @{}
$global:ActiveTabId = $null

# ------------------------------------------------------------------------------

# 実行フォルダパスの取得
$global:ScriptDirectory = $PSScriptRoot
if ([string]::IsNullOrEmpty($global:ScriptDirectory)) {
    $global:ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

# ログフォルダの世代管理
$global:ParentLogDirectory = Join-Path $global:ScriptDirectory "Logs"
if (Test-Path $global:ParentLogDirectory) {
    $oldLogFolders = Get-ChildItem -Path $global:ParentLogDirectory -Directory | Sort-Object CreationTime
    if ($oldLogFolders.Count -ge $global:CONFIG.MaxLogGenerations) {
        $deleteCount = $oldLogFolders.Count - $global:CONFIG.MaxLogGenerations + 1
        $oldLogFolders | Select-Object -First $deleteCount | ForEach-Object {
            try { Remove-Item $_.FullName -Recurse -Force -ErrorAction Stop } catch {}
        }
    }
}

# 本セッション用ログフォルダの作成
$global:LogDir = Join-Path $global:ParentLogDirectory $AppSessionId
if (-not (Test-Path $global:LogDir)) { 
    New-Item -ItemType Directory -Path $global:LogDir -Force | Out-Null
}

# コンソール文字コードのShift-JIS強制
try {
    $sjis = [System.Text.Encoding]::GetEncoding("Shift-JIS")
    [Console]::OutputEncoding = $sjis
    [Console]::InputEncoding  = $sjis
} catch {}

# --- ログ出力の実行 --- [// 内部関数 //]
function Write-DebugLog {
    param (
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Fatal')][string]$Level = 'Info'
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $formattedMessage = "<$timestamp> <$Level> $Message"

    # コンソールへの出力
    try { [Console]::WriteLine("[LOG] $formattedMessage") } catch {}

    # ログファイルへの出力
    if ($null -ne $global:LogDir -and (Test-Path $global:LogDir)) {
        $logFilePath = Join-Path $global:LogDir "Engine_SystemLog.txt"
        try {
            $formattedMessage | Out-File -FilePath $logFilePath -Append -Encoding UTF8 -ErrorAction Stop
        } catch {
            # 書き込み失敗時のイベントログへの退避
            try {
                $sourceName = "RPA_WebView2_Engine"
                if (-not [System.Diagnostics.EventLog]::SourceExists($sourceName)) {
                    [System.Diagnostics.EventLog]::CreateEventSource($sourceName, "Application")
                }
                Write-EventLog -LogName Application -Source $sourceName -EntryType Warning -EventId 1000 -Message "ログファイル書き込み失敗: $formattedMessage"
            } catch {}
        }
    }
}

# --- 共通例外エラー生成の実行 --- [// 内部関数 //]
function New-EngineException {
    param(
        [Parameter(Mandatory=$true)][string]$Func,
        [Parameter(Mandatory=$true)]
        [ValidateSet("Timeout","未発見","JSエラー","CDPエラー","ネイティブエラー","初期化エラー","ファイルエラー","UIAエラー","引数エラー","内部エラー")]
        [string]$Type,
        [Parameter(Mandatory=$true)][string]$Message,
        [string]$Details
    )

    # エラーログの多行跨ぎ防止および1行フォーマットへの正規化（ローカル変数へ格納）
    $messageCleaned = $Message -replace "`r`n|`n|`r", " "
    
    if ([string]::IsNullOrWhiteSpace($Details)) {
        return "[$Func] [$Type]: $messageCleaned"
    } else {
        $detailsCleaned = $Details -replace "`r`n|`n|`r", " | "
        return "[$Func] [$Type]: $messageCleaned ($detailsCleaned)"
    }
}

# 起動情報ログの出力
Write-DebugLog -Message "[System] 情報: Browser/ Width-Height ($screenWidth) - ($screenHeight)" -Level Info
Write-DebugLog -Message "[System] 情報: 開発モードスイッチ ($global:IsDebugMode)" -Level Info
Write-DebugLog -Message "[System] 情報: 通信モードスイッチ ($global:CdpPort)" -Level Info

# ------------------------------------------------------------------------------

# 常時ロード対象モジュールの定義
$alwaysLoadLibs = @(
    "Lib-WebAction_v206.ps1",
    "Lib-WebXPath_v102.ps1",
    "Lib-WebDebug_v102.ps1",

    "Lib-DesktopUIA_v105.ps1",
    "Lib-DevTools_v101.ps1",
    "Lib-WebSafeAction_v101.ps1",

    "Lib-WebView2_Init_v101.ps1",
    "Lib-WebView2_Native_v101.ps1"
)

foreach ($libName in $alwaysLoadLibs) {
    $libPath = Join-Path $global:ScriptDirectory $libName
    if (Test-Path $libPath) {
        try {
            . $libPath
            Write-DebugLog -Message "[System] 成功: モジュールをロード ($libName)" -Level Success
        } catch {
            Write-DebugLog -Message "[System] 致命的エラー: $libName のロードに失敗しました ($($_.Exception.Message))" -Level Fatal
            Exit
        }
    } else {
        Write-DebugLog -Message "[System] 致命的エラー: 必須モジュール $libName が見つかりません。起動を中止します。" -Level Fatal
        Exit
    }
}

# CDPモード専用モジュールのロード
if ($global:CdpPort -gt 0) {
    $cdpLibName = "Lib-WebCDP_v102.ps1"
    $cdpLibPath = Join-Path $global:ScriptDirectory $cdpLibName

    if (Test-Path $cdpLibPath) {
        . $cdpLibPath
        Write-DebugLog -Message "[System] 成功: モジュールをロード ($cdpLibName)" -Level Success
        Start-Sleep -Milliseconds 500

        # CDPポート状態の確認
        try {
            $tcpConn = Get-NetTCPConnection -LocalPort $global:CdpPort -ErrorAction Stop | Select-Object -First 1
            Write-DebugLog -Message "[System] 情報: CDPポート ($global:CdpPort) の状態 - $($tcpConn.State) ($($tcpConn.LocalAddress))" -Level Info
        } catch {
            Write-DebugLog -Message "[System] 致命的エラー: $cdpLibName のロードに失敗しました ($($_.Exception.Message))" -Level Fatal
            Exit
        }
    } else {
        Write-DebugLog -Message "[System] 致命的エラー: CDPモジュール $cdpLibName が見つかりません。起動を中止します。" -Level Fatal
        Exit
    }
}

# ------------------------------------------------------------------------------
# タブ管理（WebView2 インスタンス管理）
# ------------------------------------------------------------------------------

# --- アクティブなWebView2インスタンスの取得 --- [// 内部関数 //]
function Get-ActiveWebView {
    # Tabs登録済みWebViewの返却
    if ($null -ne $global:ActiveTabId -and $global:Tabs.ContainsKey($global:ActiveTabId)) {
        return $global:Tabs[$global:ActiveTabId].WebView
    }
    # Tabs未登録時の後方互換対応
    return $global:WebViewCtrl
}

# --- アクティブタブの設定 ---
function Set-ActiveTab {
    param ([string]$TabId)

    $func = $MyInvocation.MyCommand.Name

    if ($global:Tabs.ContainsKey($TabId)) {
        $global:ActiveTabId = $TabId
        $global:Tabs[$TabId].WebView.BringToFront()
        return "ActiveTab = $TabId"
    }
    throw (New-EngineException -Func $func -Type "引数エラー" -Message "指定されたタブIDが見つかりません" -Details $TabId)
}

# --- タブ一覧の取得 ---（開発デバッグ用）
function List-Tabs {
    $result = @()

    foreach ($tabId in $global:Tabs.Keys) {
        $tab = $global:Tabs[$tabId]
        $webview = $tab.WebView

        # Tabs連想配列URLの優先取得（CDP同期用）
        $url = ""
        if ($tab.ContainsKey("Url") -and $tab.Url) {
            $url = $tab.Url
        } else {
            try { $url = $webview.Source.ToString() } catch {}
        }

        # WebView2からのタイトル取得
        $title = ""
        try { $title = $webview.CoreWebView2.DocumentTitle } catch {}

        $result += [PSCustomObject]@{
            TabId    = $tabId
            Url      = $url
            Title    = $title
            IsActive = ($tabId -eq $global:ActiveTabId)
        }
    }

    Write-DebugLog -Message "[List-Tabs] タブ数: $($result.Count)" -Level Info
    # パイプ (|) を使わず、-InputObject に直接配列を渡す
    $resultJson = ConvertTo-Json -InputObject @($result) -Depth 2 -Compress
    Write-DebugLog -Message "$resultJson" -Level Info
    return $resultJson
}

# --- タブの切替 ---
function Switch-Tab {
    param ([Parameter(Mandatory=$true)][string]$TabId)

    $func = $MyInvocation.MyCommand.Name

    if (-not $global:Tabs.ContainsKey($TabId)) {
        throw (New-EngineException -Func $func -Type "引数エラー" -Message "指定されたタブIDが見つかりません" -Details $TabId)
    }

    # アクティブタブの更新
    $global:ActiveTabId = $TabId

    # WebView2の前面表示
    $webview = $global:Tabs[$TabId].WebView
    try {
        $webview.BringToFront()
        $global:WebViewCtrl = $webview
    } catch {}

    if ($global:CdpPort -gt 0) {
        try {
            # Connect-CdpSession -Port $global:CdpPort
        } catch {
            Write-DebugLog -Message "[Switch-Tab] CDP再接続失敗: $($_.Exception.Message)" -Level Warning
        }
    }

    return "Switched to Tab: $TabId"
}

# --- タイトル部分一致によるタブ切替 ---
function Switch-TabByTitle {
    param ([Parameter(Mandatory=$true)][string]$TitleSubstring)
    
    $func = $MyInvocation.MyCommand.Name
    
    foreach ($tabId in $global:Tabs.Keys) {
        $title = ""
        try { 
            # WebView2コントロールからタイトルを取得
            $title = $global:Tabs[$tabId].WebView.CoreWebView2.DocumentTitle
        } catch {}
        
        if ($title -like "*$TitleSubstring*") {
            Write-DebugLog -Message "[$func] 情報: タブ切り替え ($tabId - $title)" -Level Info
            return Switch-Tab -TabId $tabId
        }
    }
    
    # 既存の例外フォーマットに統一
    throw (New-EngineException -Func $func -Type "未発見" -Message "指定されたタイトルを含むタブが見つかりません" -Details $TitleSubstring)
}

# ------------------------------------------------------------------------------
# 汎用ロジックおよび JS 実行ルーター
# ------------------------------------------------------------------------------

# --- タイムアウト付き汎用待機の実行 --- [// 内部関数 //]
function Wait-Condition {
    param (
        [Parameter(Mandatory=$true)][scriptblock]$ConditionBlock,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec,
        [int]$PollIntervalMs = $global:CONFIG.PollIntervalMs,
        [string]$TimeoutMessage = "待機処理がタイムアウトしました"
    )

    $func = $MyInvocation.MyCommand.Name

    # UIフリーズを防止した条件成立待機
    $endTime = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $endTime) {
        if (& $ConditionBlock) { return $true }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds $PollIntervalMs
    }

    throw (New-EngineException -Func $func -Type "Timeout" -Message $TimeoutMessage)
}

# --- JavaScript実行のルーティング --- [// 内部関数 //]
function Invoke-WebScript {
    param (
        [Parameter(Mandatory=$true)][string]$Js,
        [int]$Retries = 3
    )

    # アクティブWebViewの取得
    $webview = Get-ActiveWebView

    # CDP通信による最優先実行
    if ($null -ne $global:CdpWebSocket -and $global:CdpWebSocket.State -eq 'Open') {
        try {
            return Invoke-CdpScript -Js $Js
        } catch {
            Write-DebugLog -Message "[Engine] 警告: CDP実行失敗、ネイティブ実行へフォールバック ($($_.Exception.Message))" -Level Warning
        }
    }

    # ネイティブ実行へのフォールバック
    return Invoke-WebView2NativeScript -Js $Js -Retries $Retries
}

# ------------------------------------------------------------------------------

# --- エンジン設定の動的変更 ---
function Set-EngineConfig {
    param ([string]$EnableHighlight)

    $func = $MyInvocation.MyCommand.Name

    # EnableHighlightのboolean変換および更新
    if (-not [string]::IsNullOrEmpty($EnableHighlight)) {
        $global:CONFIG.EnableHighlight = [System.Convert]::ToBoolean($EnableHighlight)
#●     Write-DebugLog -Message "[$func] 設定変更: EnableHighlight = $($global:CONFIG.EnableHighlight)" -Level Info
    }

    return "Config Updated"
}

# ------------------------------------------------------------------------------
# 通信ループおよびコマンド待機
# ------------------------------------------------------------------------------

# 非同期コマンド受付キューの生成
$cmdQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

# 非同期入力監視用Runspaceの作成
$runspace = [runspacefactory]::CreateRunspace()
$runspace.Open()

# 標準入力監視スレッドの開始
$psThread = [powershell]::Create().AddScript({
    param ($Queue)

    # 親プロセスからの標準入力の常時監視
    while ($true) {
        $line = [Console]::ReadLine()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        # JSONコマンドのキュー投入
        $Queue.Enqueue($line)
        Start-Sleep -Milliseconds 50
    }
}).AddArgument($cmdQueue)

$psThread.Runspace = $runspace
$psThread.BeginInvoke() | Out-Null

# 親プロセスの特定
try {
    $global:ParentPid = (Get-CimInstance Win32_Process -Filter "ProcessId = $PID").ParentProcessId
    $parentName = (Get-Process -Id $global:ParentPid -ErrorAction SilentlyContinue).Name
    Write-DebugLog -Message "[System] 情報: 親プロセス監視開始 (PID: $global:ParentPid, Name: $parentName)" -Level Info
} catch {
    $global:ParentPid = 0
    Write-DebugLog -Message "[System] 警告: 親プロセスの特定失敗" -Level Warning
}

$lastWatchdogTime = Get-Date

# 外部プロセスへのREADY通知
try { [Console]::WriteLine("READY") } catch {}
Write-DebugLog -Message "[System] 成功: エンジン待機状態" -Level Success

# ------------------------------------------------------------------------------
# メインイベントループ
# ------------------------------------------------------------------------------
try {
    while ($global:BrowserForm.Visible) {
        # UIフリーズ防止のためのWinFormsイベント処理
        [System.Windows.Forms.Application]::DoEvents()

        $json = $null

        # ゾンビ化を防止する親プロセスの生存監視
        if ($global:ParentPid -gt 0 -and ((Get-Date) - $lastWatchdogTime).TotalSeconds -gt 2) {
            $lastWatchdogTime = Get-Date

            if (-not (Get-Process -Id $global:ParentPid -ErrorAction SilentlyContinue)) {
                Write-DebugLog -Message "[System] 警告: 親プロセスの消失検知による道連れ終了の実行" -Level Warning
                break
            }
        }

        # JSONコマンドの取り出しおよび実行
        if ($cmdQueue.TryDequeue([ref]$json)) {
            try {
                # JSONからオブジェクトへの変換
                $requestObj = $json | ConvertFrom-Json -ErrorAction Stop
                $method = $requestObj.Command

                # QuitまたはExitコマンドによる即時終了
                if ($method -in @("Quit", "QUIT", "Exit")) { break }

                # パラメータ付きデバッグログの出力
                $paramStr = if ($requestObj.Parameters) { $requestObj.Parameters | ConvertTo-Json -Compress } else { "なし" }
                if ($global:IsDebugMode) {
                    Write-DebugLog -Message "[Engine] 実行: $method | Params: $paramStr" -Level Info
                } else {
                    Write-DebugLog -Message "[Engine] 実行: $method" -Level Info
                }

                # パラメータ辞書の構築
                $parameters = @{}
                if ($null -ne $requestObj.Parameters) {
                    foreach ($prop in $requestObj.Parameters.psobject.Properties) {
                        $parameters[$prop.Name] = $prop.Value
                    }
                }

                # コマンドの動的実行
                $res = & $method @parameters | Out-String

                # 実行結果の出力
                if (-not [string]::IsNullOrWhiteSpace($res)) {
                    [Console]::WriteLine("[RESULT]$($res.Trim())")
                }
                [Console]::WriteLine("[SUCCESS]")

            } catch {
                # エラー結果の出力
                $errMsg = $_.Exception.Message -replace "`r`n", " " -replace "`n", " "
                [Console]::WriteLine("[ERROR]$errMsg")
                Write-DebugLog -Message "[Engine] コマンド実行エラー: $errMsg" -Level Error
            }
        }
        Start-Sleep -Milliseconds 20
    }

} finally {
    # ------------------------------------------------------------------------------
    # 安全な終了処理（クリーンアップ）
    # ------------------------------------------------------------------------------
    Write-DebugLog -Message "[System] 情報: エンジンシャットダウン" -Level Info

    try { $runspace.Close() } catch {}
    try { if ($null -ne $global:BrowserForm) { $global:BrowserForm.Dispose() } } catch {}

    Start-Sleep -Milliseconds 300

    # ゾンビ化防止のための自プロセス強制終了
    Stop-Process -Id $PID -Force
}
