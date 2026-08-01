# ------------------------------------------------------------------------------
# デバッグおよび証跡ファイルエクスポートモジュール
# 画面状態のスナップショット保存およびDOM/iframe/Window情報の解析
# ------------------------------------------------------------------------------

# /// ブラウザ画面の取得および保存 ///

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
    // 【追加】URL取得自体を保護（クロスオリジンブロック対策）
    try { url = win.location.href; } catch(e) { url = "CROSS_ORIGIN_DENIED"; }
    
    let html = "\n";
    try {
        html += win.document.documentElement.outerHTML + "\n";
    } catch(e) {
        html += "\n";
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

        return $filePath
    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" -Message "HTMLファイルの保存に失敗しました" -Details $_.Exception.Message)
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
            while (-not $task.IsCompleted) {
                if ($sw.Elapsed.TotalSeconds -gt 10) {
                    throw (New-EngineException -Func $func -Type "Timeout" -Message "ネイティブスクショ取得がタイムアウトしました")
                }
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 10
            }

            if ($task.IsFaulted) {
                throw (New-EngineException -Func $func -Type "ネイティブエラー" -Message "CapturePreviewAsyncAPI失敗" -Details $task.Exception.InnerException.Message)
            }

            $fileStream = [IO.File]::Create($filePath)
            $stream.Seek(0, [IO.SeekOrigin]::Begin) | Out-Null
            $stream.CopyTo($fileStream)
            $fileStream.Dispose()
            $stream.Dispose()
        }

        return $filePath

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" -Message "スクリーンショットの保存処理中に例外が発生しました" -Details $_.Exception.Message)
    }
}

# /// Web要素およびフレームの解析 ///

# --- 指定テーブルのCSV保存およびVBA連携用データの返却 ---
function Export-WebTableToCsv {
    param (
        [Parameter(Mandatory = $true)][string]$Selector,
        [string]$FileName = "WebTable.csv"
    )

    $func = $MyInvocation.MyCommand.Name
    $selectorEscaped = $Selector.Replace("'", "\'")

    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

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
        return $result.vba

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" -Message "テーブルのCSV保存処理中に例外が発生しました" -Details $_.Exception.Message)
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
                    arr.push({
                        TagName: tag,
                        Type: el.type || '',
                        Name: el.name || '',
                        ID: el.id || '',
                        Value: (el.value || '').substring(0,50),
                        Text: (el.innerText || '').substring(0,50),
                        OuterHTML: (el.outerHTML || '').substring(0,200),
                        IsDisplayed: (style.display !== 'none' && style.visibility !== 'hidden')
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

        $csv = @()
        $idx = 1

        foreach ($el in $elements) {
            $csv += [PSCustomObject]@{
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
        $csv | Export-Csv -Path $filePath -Encoding UTF8 -NoTypeInformation -Force

        return $filePath

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" -Message "Web要素のCSVエクスポートに失敗しました" -Details $_.Exception.Message)
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

    $filePath = Join-Path $global:LogDir $FileName

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

        $results | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -Force

        return $filePath

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" -Message "フレームツリーの解析またはCSV保存に失敗しました" -Details $_.Exception.Message)
    }
}

# /// ウィンドウおよびOSレベルの取得 ///

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
            throw (New-EngineException -Func $func -Type "未発見" -Message "撮影対象のRPAブラウザウィンドウが存在しません")
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

        return $filePath

    } catch {
        throw (New-EngineException -Func $func -Type "内部エラー" -Message "ウィンドウ全体スクリーンショットの撮影または保存に失敗しました" -Details $_.Exception.Message)
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

        $savePath = Join-Path $global:LogDir $FileName
        $results | Export-Csv -Path $savePath -NoTypeInformation -Encoding UTF8 -Force

        return $savePath

    } catch {
        throw (New-EngineException -Func $func -Type "内部エラー" -Message "OSウィンドウ一覧の取得またはCSV保存に失敗しました" -Details $_.Exception.Message)
    }
}

# /// 汎用デバッグ情報の保存 ///

# --- 任意のデバッグ文字列のテキストファイル追記 ---
function Write-DebugTextFile {
    param (
        [Parameter(Mandatory = $true)][string]$Text,
        [string]$FileName = "DebugMemo.txt"
    )

    $func = $MyInvocation.MyCommand.Name

    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) { return }

    $filePath = Join-Path $global:LogDir $FileName

    try {
        $Text | Out-File -FilePath $filePath -Append -Encoding UTF8 -Force

        return $filePath

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" -Message "デバッグメモのファイル書き込みに失敗しました" -Details $_.Exception.Message)
    }
}
