# ------------------------------------------------------------------------------
# デバッグおよび証跡ファイルエクスポートモジュール
# 画面状態のスナップショット保存およびDOM/iframe/Window情報の解析等
# ------------------------------------------------------------------------------

# ==============================================================================
# [1] 画面状態・スナップショットの保存
# ==============================================================================

# --- 画面全体（多段iframe含む）のHTML保存 ---
function Export-WebHtml {
    param ([string]$Prefix = "WebHtml")

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $filePath  = Join-Path $global:LogDir "${Prefix}_${timestamp}.html"

    try {
        $js = @"
        function dump(win, depth) {
            let url = "UNKNOWN";
            // [CORS対策] クロスドメインのiframeにアクセスすると例外が発生するため保護
            try { url = win.location.href; } catch(e) { url = "CROSS_ORIGIN_DENIED"; }
            
            // 取得したURLをHTMLの先頭にコメントとして明記（デバッグ時の追跡用）
            let html = "\n<!-- [FRAME URL: " + url + "] -->\n";
            
            try {
                html += win.document.documentElement.outerHTML + "\n";
            } catch(e) {
                html += "<!-- DOM Access Denied -->\n";
            }
            
            for (let i = 0; i < win.frames.length; i++) {
                try { html += dump(win.frames[i], depth + 1); } catch(e) {}
            }
            return html;
        }
        return dump(window, 0);
"@

        $html = Invoke-WebScript -Js $js
        $html | Out-File -FilePath $filePath -Encoding UTF8 -Force
        return "[OK] Export-WebHtml: $filePath"

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" `
            -Message "HTMLファイルの保存に失敗しました" -Details $_.Exception.Message)
    }
}

# --- 画面スクリーンショット（CDP/WebView2）のPNG保存 ---
function Export-WebScreenshot {
    param ([string]$Prefix = "WebScreenshot")

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $guid      = [guid]::NewGuid().ToString()
    $fileName  = "${Prefix}_${timestamp}_${guid}.png"
    $filePath  = Join-Path $global:LogDir $fileName

    try {
        $useNative = $false

        # CDPモードでの撮影試行
        if ($global:CdpPort -gt 0) {
            try {
                # リトライとタイムアウトを短めに設定してCDP撮影を試みる
                $res = Invoke-CdpCommand -Method "Page.captureScreenshot" -Params @{ format = "png" } -MaxRetries 1 -TimeoutSecPerTry 3
                $bytes = [Convert]::FromBase64String($res.result.data)
                [IO.File]::WriteAllBytes($filePath, $bytes)
            } catch {
                Write-DebugLog -Message "[$func] 警告: CDP撮影失敗。ネイティブAPI(CapturePreview)へフォールバックします ($($_.Exception.Message))" -Level Warning
                $useNative = $true
            }
        } else {
            $useNative = $true
        }

        # ネイティブAPIでの撮影（CDP無効時、またはCDP失敗時）
        if ($useNative) {
            $stream = [IO.MemoryStream]::new()
            $webview = Get-ActiveWebView
            
            $task = $webview.CoreWebView2.CapturePreviewAsync(
                [Microsoft.Web.WebView2.Core.CoreWebView2CapturePreviewImageFormat]::Png,
                $stream
            )

            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $timeoutSec = 10  # タイムアウトまでの許容時間
            
            while (-not $task.IsCompleted) {
                if ($sw.Elapsed.TotalSeconds -gt $timeoutSec) {
                    throw (New-EngineException -Func $func -Type "Timeout" `
                        -Message "ネイティブスクショ取得がタイムアウトしました" -Details "${timeoutSec}秒経過")
                }
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 10
            }

            if ($task.IsFaulted) {
                throw (New-EngineException -Func $func -Type "ネイティブエラー" `
                    -Message "CapturePreviewAsyncAPI失敗" -Details $task.Exception.InnerException.Message)
            }

            # ファイル書き込み中の例外発生時も確実にメモリストリームを解放する
            $fileStream = [IO.File]::Create($filePath)
            try {
                $stream.Seek(0, [IO.SeekOrigin]::Begin) | Out-Null
                $stream.CopyTo($fileStream)
            } finally {
                if ($null -ne $fileStream) { $fileStream.Dispose() }
                if ($null -ne $stream) { $stream.Dispose() }
            }
        }
        return "[OK] Export-WebScreenshot: $filePath"

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" `
            -Message "スクリーンショットの保存処理中に例外が発生しました" -Details $_.Exception.Message)
    }
}

# ==============================================================================
# [2] DOM・要素・フレーム構造の解析とエクスポート
# ==============================================================================

# --- 指定テーブルのCSV保存およびVBA連携用データの返却 ---
function Export-WebTableToCsv {
    param (
        [Parameter(Mandatory = $true)][string]$Selector,
        [string]$FileName = "WebTable.csv"
    )

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    $selectorEscaped = $Selector.Replace("'", "\'")

    try {
        $js = @"
            $global:ENGINE_JS_UTILS
            try {
                var table = utilFindInFrames(window, function(win) {
                    try { return win.document.querySelector('$selectorEscaped') || null; }
                    catch(e) { return null; }
                });

                if (!table) return 'NOT_FOUND';

                var rows = Array.from(table.rows);
                if (rows.length === 0) return 'EMPTY_TABLE';

                // CSV出力用（人間が読む用）
                var csvData = rows.map(row => {
                    return Array.from(row.cells).map(cell => {
                        var text = (cell.innerText || '').trim().replace(/\r?\n|\r/g, ' ');
                        return '"' + text.replace(/"/g, '""') + '"';
                    }).join(',');
                }).join('\r\n');

                // VBA連携用（データ抽出用）
                var vbaData = rows.map(row => {
                    return Array.from(row.cells).map(cell => {
                        // 1. テキストがあればそれを返す
                        var text = (cell.innerText || '').trim().replace(/\r?\n|\r/g, ' ');
                        if (text !== '') return text;

                        // 2. テキストが空なら画像ファイル名を返す（汎用処理）
                        var img = cell.querySelector('img');
                        if (img && img.src) {
                            var parts = img.src.split('/');
                            return parts[parts.length - 1]; // 例: "IP10B030.png"
                        }
                        return '';
                    }).join('<T>');
                }).join('<R>');

                return { csv: csvData, vba: vbaData };
            } catch(e) {
                return 'JS_EXCEPTION: ' + e.message;
            }
"@

        $result = Invoke-WebScript -Js $js

        if ($result -eq 'NOT_FOUND')    { throw (New-EngineException -Func $func -Type "未発見" -Message "指定されたテーブルが見つかりません" -Details $Selector) }
        if ($result -eq 'EMPTY_TABLE')  { throw (New-EngineException -Func $func -Type "未発見" -Message "テーブル内に行データが存在しません" -Details $Selector) }
        if ($result -match '^JS_EXCEPTION:') { throw (New-EngineException -Func $func -Type "JSエラー" -Message "テーブル解析中にJavaScript例外が発生しました" -Details $result) }

        if ($FileName -ne "") {
            $savePath = Join-Path $global:LogDir $FileName
            $result.csv | Out-File -FilePath $savePath -Encoding UTF8 -Force
        }

        # [注意] 他のExport系関数とは異なり、VBA側でメモリ上で直接データとして活用(配列化等)することを想定（生データを返却する）
        return $result.vba

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" `
            -Message "テーブルのCSV保存処理中に例外が発生しました" -Details $_.Exception.Message)
    }
}

# --- 操作可能要素（input/button/a等）の抽出およびCSV保存 ---
function Export-WebElementsToCsv {
    param (
        [string]$FileName = "WebElements.csv",
        [int]$Limit = 1000
    )

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    $js = @"
        function dump(win, arr) {
            let els = win.document.querySelectorAll('*');
            for (let el of els) {
                let tag = el.tagName.toLowerCase();
                if (['a','button','input','select','textarea'].includes(tag) || el.id) {
                    let style = win.getComputedStyle(el);
                    // styleがnullになるエッジケースを考慮し、デフォルト値を空文字にする
                    let displayVal = style ? style.display : 'none';
                    let visibilityVal = style ? style.visibility : 'hidden';
                    
                    arr.push({
                        TagName: tag,
                        Type: el.type || '',
                        Name: el.name || '',
                        ID: el.id || '',
                        Value: (el.value || '').substring(0,50),
                        Text: (el.innerText || '').substring(0,50),
                        OuterHTML: (el.outerHTML || '').substring(0,200),
                        IsDisplayed: (displayVal !== 'none' && visibilityVal !== 'hidden')
                    });
                }
            }
            for (let i = 0; i < win.frames.length; i++) {
                try { dump(win.frames[i], arr); } catch(e) {}
            }
            return arr;
        }
        return dump(window, []);
"@

    try {
        $elements = Invoke-WebScript -Js $js

        if ($null -eq $elements -or $elements.Count -eq 0) {
            Write-DebugLog -Message "[$func] 情報: 要素なし" -Level Info
            return
        }

        $results = @()
        $idx = 1

        foreach ($el in $elements) {
            $results += [PSCustomObject]@{
                No          = $idx
                TagName     = $el.TagName
                Type        = $el.Type
                Name        = $el.Name
                ID          = $el.ID
                Value       = $el.Value
                Text        = $el.Text
                OuterHTML   = $el.OuterHTML
                IsDisplayed = $el.IsDisplayed
            }
            $idx++
            if ($idx -gt $Limit) { break }
        }

        $filePath = Join-Path $global:LogDir $FileName
        # データの蓄積、CSV出力（パイプラインでの利用）
        $results | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -Force
        return "[OK] Export-WebElementsToCsv: $filePath"

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" `
            -Message "Web要素のCSVエクスポートに失敗しました" -Details $_.Exception.Message)
    }
}

# --- iframe/frame階層構造のツリー形式CSV保存 ---
function Export-WebFrameTreeToCsv {
    param ([string]$FileName = "WebFrameTree.csv")

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    $js = @"
        function getFrameTree(win, depth, path) {
            var frames = win.document.querySelectorAll('iframe, frame');
            var tree = [];

            for (var i = 0; i < frames.length; i++) {
                var f = frames[i];
                var currentPath = path === '' ? i.toString() : path + '.' + i;

                var info = {
                    Depth: depth,
                    Path: currentPath,
                    Index: i,
                    Tag: f.tagName.toLowerCase(),
                    Id: f.id || '',
                    Name: f.name || '',
                    Src: f.src || ''
                };

                try {
                    var hasAccess = f.contentWindow && f.contentWindow.document;
                    if (hasAccess) {
                        var children = getFrameTree(f.contentWindow, depth + 1, currentPath);
                        tree.push(info);
                        tree = tree.concat(children);
                    } else {
                        info.Note = "Cross-Origin (No Access)";
                        tree.push(info);
                    }
                } catch(e) {
                    info.Note = "Access Denied";
                    tree.push(info);
                }
            }
            return tree;
        }
        return JSON.stringify(getFrameTree(window, 0, ''));
"@

    try {
        $json = Invoke-WebScript -Js $js

        if ([string]::IsNullOrWhiteSpace($json) -or $json -eq "[]") {
            Write-DebugLog -Message "[$func] 情報: iframeなし" -Level Info
            return
        }

        $frames = $json | ConvertFrom-Json

        $results = foreach ($f in $frames) {
            $indent = "  " * $f.Depth
            $prefix = if ($f.Depth -eq 0) { "■" } else { "└─" }

            [PSCustomObject]@{
                VisualHierarchy = "$indent$prefix $($f.Tag)"
                IndexPath       = $f.Path
                Id              = $f.Id
                Name            = $f.Name
                Depth           = $f.Depth
                Tag             = $f.Tag
                Src             = $f.Src
                Note            = $f.Note
            }
        }

        $filePath = Join-Path $global:LogDir $FileName
        # データの蓄積、CSV出力（パイプラインでの利用）
        $results | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -Force
        return "[OK] Export-WebFrameTreeToCsv: $filePath"

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" `
            -Message "フレームツリーの解析またはCSV保存に失敗しました" -Details $_.Exception.Message)
    }
}

# ==============================================================================
# [3] OSウィンドウレベルの解析・保存
# ==============================================================================

# --- RPAブラウザウィンドウ全体（枠含む）のPNG保存 ---
function Export-WindowScreenshot {
    param ([string]$Prefix = "WindowScreenshot")

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    try {
        Add-Type -AssemblyName System.Drawing

        if ($null -eq $global:BrowserForm) {
            throw (New-EngineException -Func $func -Type "未発見" `
                -Message "撮影対象のRPAブラウザウィンドウが存在しません")
        }

        $bounds = $global:BrowserForm.Bounds

        # --- DPIスケーリング(125%など)の補正計算 ---
        $tmpGraphics = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
        $dpiX = $tmpGraphics.DpiX
        $tmpGraphics.Dispose()
        $scale = $dpiX / 96.0  # 100%時は1.0、125%時は1.25になる

        # 物理ピクセル座標・サイズへの変換
        $phyX = [int][Math]::Round($bounds.X * $scale)
        $phyY = [int][Math]::Round($bounds.Y * $scale)
        $phyW = [int][Math]::Round($bounds.Width * $scale)
        $phyH = [int][Math]::Round($bounds.Height * $scale)

        # 補正後の物理サイズでBitmap作成
        $bmp      = New-Object System.Drawing.Bitmap $phyW, $phyH
        $graphics = [System.Drawing.Graphics]::FromImage($bmp)

        # OSレベルの画面キャプチャ (物理ピクセルで指定)
        $graphics.CopyFromScreen(
            $phyX, $phyY,
            0, 0,
            $bmp.Size,
            [System.Drawing.CopyPixelOperation]::SourceCopy
        )

        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $guid      = [guid]::NewGuid().ToString()
        $fileName  = "${Prefix}_${timestamp}_${guid}.png"
        $filePath = Join-Path $global:LogDir $fileName

        $bmp.Save($filePath, [System.Drawing.Imaging.ImageFormat]::Png)

        $graphics.Dispose()
        $bmp.Dispose()

        return "[OK] Export-WindowScreenshot: $filePath"

    } catch {
        throw (New-EngineException -Func $func -Type "内部エラー" `
            -Message "ウィンドウ全体スクリーンショットの撮影または保存に失敗しました" -Details $_.Exception.Message)
    }
}

# --- 実行中ウィンドウ一覧のCSV保存 ---
function Export-WindowHierarchyToCsv {
    param ([string]$FileName = "WindowHierarchy.csv")

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    try {
        $processes = Get-Process | Where-Object { $_.MainWindowHandle -ne 0 }

        $results = foreach ($p in $processes) {
            [PSCustomObject]@{
                ProcessName = $p.ProcessName
                Handle      = $p.MainWindowHandle
                Title       = $p.MainWindowTitle
                ID          = $p.Id
            }
        }

        $filePath = Join-Path $global:LogDir $FileName
        # データの蓄積、CSV出力（パイプラインでの利用）
        $results | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -Force
        return "[OK] Export-WindowHierarchyToCsv: $filePath"

    } catch {
        throw (New-EngineException -Func $func -Type "内部エラー" `
            -Message "OSウィンドウ一覧の取得またはCSV保存に失敗しました" -Details $_.Exception.Message)
    }
}

# ==============================================================================
# [4] 汎用デバッグユーティリティ
# ==============================================================================

# --- 任意のデバッグ文字列のテキストファイル追記 ---
function Write-DebugTextFile {
    param (
        [Parameter(Mandatory = $true)][string]$Text,
        [string]$FileName = "DebugMemo.txt"
    )

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    $filePath = Join-Path $global:LogDir $FileName

    try {
        $Text | Out-File -FilePath $filePath -Append -Encoding UTF8 -Force
        return "[OK] Write-DebugTextFile: $filePath"

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" `
            -Message "デバッグメモのファイル書き込みに失敗しました" -Details $_.Exception.Message)
    }
}

# ==============================================================================
# [Web] DOM Snapshot エクスポート( JSON保存 )
#  - 人間が読めるフラット構造に変換して出力
# ==============================================================================
function Export-WebDomSnapshot {
    param([string]$Prefix = "DOMSnapshot")

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    try {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $filePath  = Join-Path $global:LogDir "${Prefix}_${timestamp}.json"

        # 必要要素の厳選 (デザイン系を排除し、表示状態に特化)
        $params = @{
            computedStyles = @("display", "visibility", "width", "height")
        }
        $cdpResponse = Invoke-CdpCommand -Method "DOMSnapshot.captureSnapshot" -Params $params

        # 応答から .result 本体を取得
        if (-not $cdpResponse -or -not $cdpResponse.result) {
            throw (New-EngineException -Func $func -Type "CDPエラー" `
                -Message "DOM Snapshot の取得結果が空です" -Details "Invoke-CdpCommand の戻り値が null / 空でした")
        }

        # Convert-DomSnapshotFlat を適用して人間が読める形式へ変換
        $flatNodes = Convert-DomSnapshotFlat -domResult $cdpResponse.result
        # JSON化して保存 (Depthを適正化し、見やすくインデント整形)意図的に -Compress を外す
        $json = $flatNodes | ConvertTo-Json -Depth 5
        
        try {
            Set-Content -Path $filePath -Value $json -Encoding UTF8
        }
        catch {
            throw (New-EngineException -Func $func -Type "ファイルエラー" `
                -Message "DOM Snapshot の保存に失敗しました" -Details $_.Exception.Message)
        }
        return "[OK] Export-WebDomSnapshot: $filePath"

    } catch {
        # すでに New-EngineException で生成されたエラーならそのままスロー
        if ($_.Exception.Message -match "^\[.*?\] \[.*?\]:") { throw $_ }
        
        throw (New-EngineException -Func $func -Type "内部エラー" `
            -Message "DOM Snapshot の取得処理中に例外が発生しました" -Details $_.Exception.Message)
    }
}

# --- CDPのフラットなDOM配列をオブジェクトリストに展開 ---
    <# CDPの仕様（圧縮された1次元配列）と、PowerShell 5.1の処理性能の制約を克服するため、
       辞書データとノード・属性・スタイルを高速に再結合して扱いやすい構造へ変換。 #>
function Convert-DomSnapshotFlat {
    param(
        [Parameter(Mandatory=$true)]
        [object]$DomResult,

        # リクエスト時に指定したスタイルの順番を配列で定義
        [string[]]$StyleNames = @("display", "visibility", "width", "height")
    )

    $list = New-Object System.Collections.Generic.List[object]
    $strings = $DomResult.strings 

    # 子フレームを含むすべてのドキュメントをループする
    foreach ($doc in $DomResult.documents) {
        $nodes = $doc.nodes
        $layout = $doc.layout

        $count = 0
        if ($nodes.parentIndex) { $count = $nodes.parentIndex.Count }
        elseif ($nodes.nodeName) { $count = $nodes.nodeName.Count }
        
        # ----------------------------------------------------------
        # [1] DOM構造の展開
        for ($i = 0; $i -lt $count; $i++) {

            # 1要素分の器（テンプレート）を作成
            $node = [ordered]@{
                index          = $i
                parentIndex    = if ($nodes.parentIndex) { $nodes.parentIndex[$i] } else { $null }
                nodeType       = if ($nodes.nodeType) { $nodes.nodeType[$i] } else { $null }
                backendNodeId  = if ($nodes.backendNodeId) { $nodes.backendNodeId[$i] } else { $null }
                childIndexes   = if ($nodes.childNodeIndexes) { $nodes.childNodeIndexes[$i] } else { $null }
                attributes     = @{}
                computedStyles = @{}
                tagName        = ""
                id             = ""
                name           = ""
                placeholder    = ""
                value          = ""
            }

            # --- タグ名 (nodeName) の復元 ---
            # nodeName配列には直接の文字列ではなく、$strings 配列への「インデックス番号」が入っている。
            if ($nodes.nodeName -and $null -ne $nodes.nodeName[$i]) {
                $nameIndex = $nodes.nodeName[$i]
                if ($nameIndex -ge 0 -and $nameIndex -lt $strings.Count) {
                    $node.tagName = $strings[$nameIndex].ToLower()
                }
            }

            # --- 属性 (attributes) の復元 ---
            # attributes配列は [名前のインデックス, 値のインデックス, 名前のインデックス, 値のインデックス...] 
            # という連続した1次元配列として格納されているため、2つずつ(Step 2)取り出して復元する。
            if ($nodes.attributes -and $nodes.attributes[$i]) {
                $attrArray = $nodes.attributes[$i]
                for ($a = 0; $a -lt $attrArray.Count; $a += 2) {
                    $nameIndex  = $attrArray[$a]
                    $valueIndex = $attrArray[$a + 1]

                    if ($null -ne $nameIndex -and $null -ne $valueIndex) {
                        # $strings 配列から実際の文字列を引き当てる
                        $attrName  = $strings[$nameIndex]
                        $attrValue = $strings[$valueIndex]

                        if ($attrName) {
                            $node.attributes[$attrName] = $attrValue

                            # RPAの要素特定の要となる主要属性は、アクセスしやすいようルート階層にもコピー
                            if ($attrName -eq "id") { $node.id = $attrValue }
                            if ($attrName -eq "name") { $node.name = $attrValue }
                            if ($attrName -eq "placeholder") { $node.placeholder = $attrValue }
                            if ($attrName -eq "value") { $node.value = $attrValue }
                        }
                    }
                }
            }
            $list.Add($node)
        }

        # ----------------------------------------------------------
        # [2] レイアウト(スタイル)情報の紐づけ

        # レイアウトツリーが存在し、スタイル配列がある場合のみ実行
        if ($layout -and $layout.nodeIndex -and $layout.styles) {

            # layout.nodeIndex には、画面に描画されている要素のインデックス番号が入っている
            for ($L = 0; $L -lt $layout.nodeIndex.Count; $L++) {
                $nodeIdx = $layout.nodeIndex[$L]
                
                # list内の該当ノードを取得 (フレーム毎にオフセットを考慮)
                # ※ $list は累積されているため、現在のフレームの開始位置からのインデックスを計算
                $absoluteIdx = ($list.Count - $count) + $nodeIdx
                
                if ($absoluteIdx -ge 0 -and $absoluteIdx -lt $list.Count) {
                    $targetNode = $list[$absoluteIdx]
                    $styleValues = $layout.styles[$L]
                    
                    # リクエストした StyleNames の順序に合わせて文字列辞書から抽出
                    for ($s = 0; $s -lt $styleValues.Count -and $s -lt $StyleNames.Count; $s++) {
                        $strIdx = $styleValues[$s]
                        if ($strIdx -ge 0 -and $strIdx -lt $strings.Count) {
                            $targetNode.computedStyles[$StyleNames[$s]] = $strings[$strIdx]
                        }
                    }
                }
            }
        }
    }
    return $list
}

# --- DOMSnapshotとBoxModel(座標)を統合して取得 ---
    <# 画面上の全要素に対して座標取得(DOM.getBoxModel)を行うと、数千回のCDP通信が発生しフリーズの原因となる。
       まず一括でDOM構造を取得(DOMSnapshot)し、自動化の操作対象となり得るインタラクティブ要素(input, button等)に
       絞って座標を取得し、高速化と正確性を両立させる。 #>
function Get-DomSnapshotWithBoxModel {

    $func = $MyInvocation.MyCommand.Name

    # 1. 画面全体のDOMツリーと基本スタイルを一撃で取得（この時点では座標x,yは含まれない）
    $cdpResponse = Invoke-CdpCommand -Method "DOMSnapshot.captureSnapshot" -Params @{
        computedStyles = @("display","visibility","width","height")
    }
    $domResult = $cdpResponse.result
    if (-not $domResult) {
        Write-DebugLog -Message "[$func] エラー: DOMSnapshotの取得結果が空です" -Level Error
        return $null
    }

    # 2. 圧縮されたCDPの1次元配列を、扱いやすいオブジェクトリスト（連想配列）に展開
    $nodes = Convert-DomSnapshotFlat -DomResult $domResult

    # 3. 展開した全ノードの中から、RPA操作のターゲットになり得る要素だけを抽出して座標を付与
    foreach ($node in $nodes) {
        if (-not $node.tagName) { continue }
        # パフォーマンス最適化：操作対象外のタグ(divやspan等)はスキップ
        if ($node.tagName -notin @("input","button","a","select","textarea")) { continue }
        # ターゲット要素の固有ID（backendNodeId）がない場合はスキップ
        if (-not $node.backendNodeId) { continue }

        try {
            # 絞り込んだ要素に対してのみ、CDP経由で物理座標（BoxModel）を要求する
            $boxResponse = Invoke-CdpCommand -Method "DOM.getBoxModel" -Params @{
                backendNodeId = $node.backendNodeId
            }
            $model = $boxResponse.result.model
            # 取得した座標情報(content配列の先頭2つがx,y)をノード情報にマージ(合体)する           
            if ($model -and $model.content) {
                $node.boundingBox = @{
                    x      = $model.content[0]
                    y      = $model.content[1]
                    width  = $model.width
                    height = $model.height
                }
            }
        } catch {
            # [無言の理由]: 画面に描画されない要素に対するCDPエラーを無視してログスパムを防止
        }
    }

    $bbCount = ($nodes | Where-Object { $_.boundingBox }).Count
    if ($global:IsDebugMode) {
        Write-DebugLog -Message "[$func] 情報: DOMSnapshot 抽出ノード数: $($nodes.Count) / boundingBox 取得数: $bbCount" -Level Info
    }
    return $nodes
}

# ==============================================================================
# [Web] 画面レイアウト情報エクスポート( JSON保存 )
#  - Page.getLayoutMetrics を用いて viewport / content / DPI
# ==============================================================================
function Export-WebLayoutDump {
    param([string]$Prefix = "LayoutDump")

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    try {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $filePath  = Join-Path $global:LogDir "${Prefix}_${timestamp}.json"

        # --- CDP呼び出し ---
        $result = Invoke-CdpCommand -Method "Page.getLayoutMetrics" -Params @{}

        if (-not $result) {
            throw (New-EngineException -Func $func -Type "CDPエラー" `
                -Message "LayoutMetrics の取得結果が空です" -Details "Invoke-CdpCommand の戻り値が null / 空でした")
        }

        # JSON化
        $json = $result | ConvertTo-Json -Depth 16
        # ファイル保存
        try {
            Set-Content -Path $filePath -Value $json -Encoding UTF8
        }
        catch {
            throw (New-EngineException -Func $func -Type "ファイルエラー" `
                -Message "LayoutDump の保存に失敗しました" -Details $_.Exception.Message)
        }
        return "[OK] Export-WebLayoutDump: $filePath"

    } catch {
        # すでに New-EngineException で生成されたエラーならそのままスロー
        if ($_.Exception.Message -match "^\[.*?\] \[.*?\]:") { throw $_ }
        
        throw (New-EngineException -Func $func -Type "内部エラー" `
            -Message "DOM Snapshot の取得処理中に例外が発生しました" -Details $_.Exception.Message)
    }
}
