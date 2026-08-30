# ------------------------------------------------------------------------------
# XPath 専用 Web 自動化操作モジュール
# document.evaluateを利用した動的要素の特定および操作
# ------------------------------------------------------------------------------

# --- XPathの正規化 --- [// 内部関数 //]
function Normalize-XPath {
    param([string]$XPath)

    # normalize-space包含時のスキップ
    if ($XPath -match "normalize-space") {
        return $XPath
    }

    # //*[text()='xxx'] から contains(normalize-space(.), 'xxx') への変換
    $pattern = "\[[^\]]*text\(\)\s*=\s*'([^']+)'\]"
    if ($XPath -match $pattern) {
        $value = $Matches[1]
        return ($XPath -replace $pattern, "[contains(normalize-space(.), '$value')]")
    }

    # //*[.='xxx'] から contains(normalize-space(.), 'xxx') への変換
    $pattern2 = "\[[^\]]*\.\s*=\s*'([^']+)'\]"
    if ($XPath -match $pattern2) {
        $value = $Matches[1]
        return ($XPath -replace $pattern2, "[contains(normalize-space(.), '$value')]")
    }

    return $XPath
}

# --- 指定XPath要素の出現および可視化待機 ---
function Wait-WebXPathElement {
    param (
        [Parameter(Mandatory = $true)][string]$XPath,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name

    # XPathの正規化およびJS用エスケープ
    $xpathEscaped = Normalize-XPath $XPath
    $xpathEscaped = $xpathEscaped.Replace("'", "\'")

    # 条件ブロックの定義
    $condition = {
        try {
            $js = @"
                $global:ENGINE_JS_UTILS
                var xpath = '$xpathEscaped';

                var found = utilFindInFrames(window, function(win) {
                    try {
                        var el = win.document.evaluate(xpath, win.document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                        if (el) {
                            var style = win.getComputedStyle(el);
                            var rect  = el.getBoundingClientRect();
                            var visible = (style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0' && rect.width > 0 && rect.height > 0);
                            return visible ? 'visible' : null;
                        }
                    } catch(e) {}
                    return null;
                });
                return found || 'not_found';
"@

            $res = Invoke-WebScript -Js $js

            if ($res -eq "visible") {
                # 任意ハイライト処理の実行
                if ($global:CONFIG.EnableHighlight) {
                    $jsHighlight = @"
                        $global:ENGINE_JS_UTILS
                        var xpath = '$xpathEscaped';

                        var el = utilFindInFrames(window, function(win) {
                            try { return win.document.evaluate(xpath, win.document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue || null; }
                            catch(e) { return null; }
                        });

                        if (el) {
                            var oldOutline = el.style.outline;
                            el.style.outline = '2px solid red';
                            setTimeout(() => { el.style.outline = oldOutline; }, 800);
                        }
"@

                    Invoke-WebScript -Js $jsHighlight | Out-Null
                }
                return $true
            }
        } catch {}
        return $false
    }

    $errMsg = "[$func] タイムアウト: XPath要素が見つからない ($XPath)"
    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null

    return $true
}

# --- 指定XPath要素の消滅または非表示待機 ---
function Wait-WebXPathElementDisappear {
    param (
        [Parameter(Mandatory = $true)][string]$XPath,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name

    $xpathEscaped = Normalize-XPath $XPath
    $xpathEscaped = $xpathEscaped.Replace("'", "\'")

    $js = @"
        $global:ENGINE_JS_UTILS
        var xpath = '$xpathEscaped';
        
        var found = utilFindInFrames(window, function(win) {
            try {
                var el = win.document.evaluate(xpath, win.document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (el) {
                    var style = win.getComputedStyle(el);
                    var rect = el.getBoundingClientRect();
                    // display, visibility, opacity のいずれかで非表示になれば false 扱いとする
                    var isVisible = (style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0' && rect.width > 0 && rect.height > 0);
                    if (isVisible) return true; // まだ表示されている
                }
            } catch(e) {}
            return null;
        });
        return found === true;
"@

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $isGone = $false

    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        # JS実行による要素存在の真偽値取得
        $exists = Invoke-WebView2NativeScript -Js $js -Retries 1

        # False（未発見・非表示）による消滅判定
        if ($exists -eq $false -or $exists -match "false") {
            $isGone = $true
            break
        }
        Start-Sleep -Milliseconds 500
    }

    if (-not $isGone) {
        throw (New-EngineException -Func $func -Type "Timeout" -Message "指定された時間が経過しても要素が消滅・非表示になりませんでした" -Details "${TimeoutSec}秒経過 ($XPath)")
    }
}

# --- 指定XPath要素のクリック実行（物理クリック・フォールバック対応） ---
function Invoke-WebXPathClick {
    param (
        [Parameter(Mandatory = $true)][string]$XPath,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec,
        [string]$OuterHtml = "",
        [string]$BoundingBox = ""
    )

    $func = $MyInvocation.MyCommand.Name

    $xpathEscaped = Normalize-XPath $XPath
    $xpathEscaped = $xpathEscaped.Replace("'", "\'")

    # 要素出現の待機
    Wait-WebXPathElement -XPath $XPath -TimeoutSec $TimeoutSec | Out-Null

    # クリック処理の実行 (JS)
    $js = @"
        $global:ENGINE_JS_UTILS
        var xpath = '$xpathEscaped';
        var el = utilFindInFrames(window, function(win) {
            try {
                return win.document.evaluate(xpath, win.document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue || null;
            } catch(e) { return null; }
        });

        if (el) {
            el.scrollIntoView({block:'center', inline:'center'});
            el.dispatchEvent(new MouseEvent('mouseover', {bubbles:true}));
            el.dispatchEvent(new MouseEvent('mousemove', {bubbles:true}));
            el.dispatchEvent(new MouseEvent('mousedown', {bubbles:true}));
            el.dispatchEvent(new MouseEvent('mouseup', {bubbles:true}));
            el.focus();
            el.click();
            return true;
        }
        return false;
"@

    $res = Invoke-WebScript -Js $js   
    # --- JSクリック失敗時、または座標情報がある場合 ---
    if (-not $res) {
        Write-DebugLog -Message "[$func] 警告: JSによるXPath要素のクリックが失敗しました。" -Level Warn
    }

    if (-not [string]::IsNullOrWhiteSpace($BoundingBox)) {
        try {
            $bbox = $BoundingBox | ConvertFrom-Json
            if ($bbox.width -gt 0 -and $bbox.height -gt 0) {
                # ターゲットの中心座標を計算 (画面絶対座標)
                $targetX = [int]($bbox.x + ($bbox.width / 2))
                $targetY = [int]($bbox.y + ($bbox.height / 2))
                
                Write-DebugLog -Message "[$func] 情報: BoundingBoxを使用し、座標(X:$targetX, Y:$targetY)へ物理クリックを試行します。" -Level Info
                if (Get-Command -Name "Invoke-DesktopCoordinateClick" -ErrorAction SilentlyContinue) {
                    Invoke-DesktopCoordinateClick -X $targetX -Y $targetY
                    return $true
                } else {
                    Write-DebugLog -Message "[$func] 警告: 物理クリック関数が見つかりません。Lib-DesktopUIAとの連携を確認してください。" -Level Warn
                }
            }
        } catch {
            Write-DebugLog -Message "[$func] エラー: BoundingBoxの解析または物理クリック中に例外が発生しました。" -Level Error
        }
    }

    if (-not $res) {
        throw (New-EngineException -Func $func -Type "未発見" -Message "XPath要素のクリック実行に失敗しました" -Details $XPath)
    }
}

# --- 指定XPath要素へのテキスト入力 ---
function Set-WebXPathTextInput {
    param (
        [Parameter(Mandatory = $true)][string]$XPath,
        [Parameter(Mandatory = $true)][string]$Value,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name

    $xpathEscaped = Normalize-XPath $XPath
    $xpathEscaped = $xpathEscaped.Replace("'", "\'")

    # 入力値のエスケープ処理
    $safeValue = $Value.Replace("\", "\\").Replace("'", "\'")

    Wait-WebXPathElement -XPath $XPath -TimeoutSec $TimeoutSec | Out-Null

    $js = @"
        $global:ENGINE_JS_UTILS
        var xpath = '$xpathEscaped';
        var el = utilFindInFrames(window, function(win) {
            try {
                return win.document.evaluate(xpath, win.document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue || null;
            } catch(e) { return null; }
        });

        if (el) {
            el.focus();
            el.value = '$safeValue';
            el.dispatchEvent(new Event('input',  { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
        }
        return false;
"@

    $res = Invoke-WebScript -Js $js
    if (-not $res) {
        throw (New-EngineException -Func $func -Type "未発見" -Message "XPath要素へのテキスト入力に失敗しました" -Details $XPath)
    }
}

# --- 指定XPath要素のテキスト取得 ---
function Get-WebXPathText {
    param (
        [Parameter(Mandatory = $true)][string]$XPath,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name

    $xpathEscaped = Normalize-XPath $XPath
    $xpathEscaped = $xpathEscaped.Replace("'", "\'")

    Wait-WebXPathElement -XPath $XPath -TimeoutSec $TimeoutSec | Out-Null

    $js = @"
        $global:ENGINE_JS_UTILS
        var xpath = '$xpathEscaped';
        var el = utilFindInFrames(window, function(win) {
            try {
                return win.document.evaluate(xpath, win.document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue || null;
            } catch(e) { return null; }
        });

        if (!el) throw new Error("XPath element not found: " + xpath);

        const tag = el.tagName ? el.tagName.toLowerCase() : '';
        if (tag === 'input' || tag === 'textarea' || tag === 'select') {
            return el.value;
        } else {
            return el.innerText || el.textContent;
        }
"@

    return Invoke-WebScript -Js $js
}
