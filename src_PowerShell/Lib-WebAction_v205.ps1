# ------------------------------------------------------------------------------
# Web標準操作モジュール
# DOMベースの要素探索、待機、クリック、入力などの共通アクション
# ------------------------------------------------------------------------------

# --- 指定URLへのナビゲーション実行 ---
function Invoke-WebNavigation {
    param ([Parameter(Mandatory=$true)][string]$Url)

    $func = $MyInvocation.MyCommand.Name

    try {
        # アクティブWebViewに対するNavigateの実行
        $activeWv = Get-ActiveWebView
        $activeWv.CoreWebView2.Navigate($Url)
        return "Navigating to $Url"
    } catch {
        throw (New-EngineException -Func $func -Type "ネイティブエラー" -Message "指定されたURLへのナビゲーションに失敗しました" -Details $_.Exception.Message)
    }
}

# --- ページ読み込み完了（iframe含む完全ロード）の待機 ---
function Wait-WebPageLoad {
    param ([int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec)

    $func = $MyInvocation.MyCommand.Name

    # readyState=completeおよび全iframeのDOM完成の再帰チェック
    $condition = {
        try {
            $js = @"
                function isLoaded(win) {
                    try {
                        if (win.document.readyState !== 'complete') return false;

                        // iframe の DOM がまだ構築中のケースに対応
                        let frames = win.frames;
                        for (let i = 0; i < frames.length; i++) {
                            try {
                                if (!isLoaded(frames[i])) return false;
                            } catch(e) {
                                // アクセス不可(CORS)は無視
                            }
                        }
                        return true;
                    } catch(e) {
                        return false;
                    }
                }
                return isLoaded(window);
"@

            # CDPモード時におけるネイティブ実行の強制
            $res = Invoke-WebView2NativeScript -Js $js

            if ($res -eq $true) {
                # DOM安定化のための追加待機
                Start-Sleep -Milliseconds 300
                return $true
            }
        } catch {}
        return $false
    }

    $errMsg = "[$func] タイムアウト: ページまたは iframe のロード未完了"
    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null

    return $true
}

# --- 画面全体の読み込みステータス（complete）待機 ---
function Wait-WebDocumentReady {
    param ([int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec)

    $func = $MyInvocation.MyCommand.Name

    $js = @"
        $global:ENGINE_JS_UTILS
        var res = utilFindInFrames(window, function(win) {
            try {
                return (win.document.readyState === 'complete');
            } catch(e) {
                return null;
            }
        });
        return res === true;
"@

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $isReady = $false

    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        # 修正後のラッパー構造を介して安全にBoolean（$true/$false）を受け取る
        $state = Invoke-WebScript -Js $js -Retries 1
        
        if ($state -eq $true -or $state -eq "true") {
            $isReady = $true
            break
        }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 300
    }

    if (-not $isReady) {
#✘        Write-DebugLog -Message "[$func] 警告: Document ReadyState がすべてのフレームで complete になりませんでした" -Level Warning
        throw (New-EngineException -Func $func -Type "Timeout" -Message "Document ReadyState が制限時間内に complete になりませんでした" -Details "Timeout: ${TimeoutSec}s")
    }
}

# --- URLの部分一致待機 ---
function Wait-WebUrlContains {
    param (
        [Parameter(Mandatory=$true)][string]$Substring,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name

    # URLへの指定文字列包含待機
    $condition = {
        $webview = Get-ActiveWebView
        if ($webview -and $webview.Source) {
            $url = $webview.Source.ToString()
            if ($url -like "*$Substring*") { return $true }
        }
        return $false
    }

    $errMsg = "[$func] タイムアウト: URL不一致 ($Substring)"
    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -PollIntervalMs 300 -TimeoutMessage $errMsg | Out-Null

    return "URL matched: $(Get-WebUrl)"
}

# --- ページタイトルの部分一致待機 ---
function Wait-WebTitleContains {
    param (
        [Parameter(Mandatory=$true)][string]$Substring,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name

    # タイトルへの指定文字列包含待機
    $condition = {
        $webview = Get-ActiveWebView
        if ($webview) {
            $title = $webview.CoreWebView2.DocumentTitle
            if ($title -and $title.Trim().IndexOf($Substring, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                return $true
            }
        }
        return $false
    }

    $errMsg = "[$func] タイムアウト: タイトル不一致 ($Substring)"
    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -PollIntervalMs 300 -TimeoutMessage $errMsg | Out-Null

    return "Title matched: $(Get-WebTitle)"
}

# --- 指定要素の出現および可視化待機 ---
function Wait-WebElement {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    $selectorEscaped = $Selector.Replace("'", "\'")

    # DOM出現およびdisplay/visibility/opacity/rectによる可視判定
    $condition = {
        try {
            $js = @"
                $global:ENGINE_JS_UTILS
                var selector = '$selectorEscaped';
                
                var found = utilFindInFrames(window, function(win) {
                    try {
// 変更                 var el = win.document.querySelector(selector); Shadow DOM対応へ（mode: 'open'）
                        var el = deepQuerySelector(selector, win.document);
                        if (el) {
                            var style = win.getComputedStyle(el);
                            var rect = el.getBoundingClientRect();
                            // 要素が存在し、かつ可視状態なら true を返す
                            return (style.display !== 'none' 
                                    && style.visibility !== 'hidden'
                                    && style.opacity !== '0'
                                    && rect.width > 0 && rect.height > 0);
                        }
                    } catch(e) {}
                    return null;
                });
                return found === true;
"@

            $res = Invoke-WebScript -Js $js
            if ($res) { return $true }
            
        } catch {}
        return $false
    }

    $errMsg = "[$func] タイムアウト: 要素の未出現または非表示 ($Selector)"
    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null

    return $true
}

# --- iframe内要素の出現および可視化待機 ---
function Wait-WebElementInFrame {
    param (
        [Parameter(Mandatory=$true)][string]$FrameSelector,
        [Parameter(Mandatory=$true)][string]$ElementSelector,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    $frameEscaped = $FrameSelector.Replace("'", "\'")
    $elementEscaped = $ElementSelector.Replace("'", "\'")

    # iframe.contentDocumentを用いた直接探索
    $condition = {
        try {
            $js = @"
                var frm = document.querySelector('$frameEscaped');
                if (!frm || !frm.contentDocument) return 'frame_not_found';
                var el = frm.contentDocument.querySelector('$elementEscaped');
                if (!el) return 'not_found';
                var style = frm.contentWindow.getComputedStyle(el);
                var rect = el.getBoundingClientRect();
                var isVisible = (style.display !== 'none'
                                 && style.visibility !== 'hidden'
                                 && style.opacity !== '0'
                                 && rect.width > 0 && rect.height > 0);
                return isVisible ? 'visible' : 'hidden';
"@

            $res = Invoke-WebScript -Js $js
            if ($res -eq "visible") { return $true }
        } catch {}
        return $false
    }

    $errMsg = "[$func] タイムアウト: iframe内要素の未出現 ($ElementSelector)"
    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null

    return $true
}

# --- iframe内要素のクリック実行 ---
function Invoke-WebClickInFrame {
    param (
        [Parameter(Mandatory=$true)][string]$FrameSelector,
        [Parameter(Mandatory=$true)][string]$ElementSelector,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    $frameEscaped = $FrameSelector.Replace("'", "\'")
    $elementEscaped = $ElementSelector.Replace("'", "\'")

    # iframe内でのscrollIntoViewおよびclickの実行
    Wait-WebElementInFrame -FrameSelector $FrameSelector -ElementSelector $ElementSelector -TimeoutSec $TimeoutSec | Out-Null

    $js = @"
        var frm = document.querySelector('$frameEscaped');
        if (frm && frm.contentDocument) {
            var el = frm.contentDocument.querySelector('$elementEscaped');
            if (el) {
                el.scrollIntoView({block: 'center', inline: 'center'});
                el.click();
                return true;
            }
        }
        return false;
"@

    $res = Invoke-WebScript -Js $js
    if (-not $res) { throw (New-EngineException -Func $func -Type "未発見" -Message "iframe内でのクリック実行に失敗しました" -Details $ElementSelector) }
}

# --- 指定要素の非表示またはDOM削除待機 ---
function Wait-WebElementInvisible {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    $selectorEscaped = $Selector.Replace("'", "\'")

    # 多段iframe対応と、display/visibility/opacity/rectによる判定に統一
    $condition = {
        $js = @"
            $global:ENGINE_JS_UTILS
            var selector = '$selectorEscaped';
            var found = utilFindInFrames(window, function(win) {
                try {
// 変更             var el = win.document.querySelector(selector); Shadow DOM対応へ（mode: 'open'）
                    var el = deepQuerySelector(selector, win.document);
                    if (el) {
                        var style = win.getComputedStyle(el);
                        var rect = el.getBoundingClientRect();
                        var isVisible = (style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0' && rect.width > 0 && rect.height > 0);
                        if (isVisible) return true; // まだ表示されている
                    }
                } catch(e) {}
                return null;
            });
            // どこにも無いか、あっても非表示なら 'hidden' を返す
            return found === true ? 'visible' : 'hidden';
"@

        $res = Invoke-WebScript -Js $js
        if ($res -eq "hidden") { return $true }
        return $false
    }

    $errMsg = "[$func] タイムアウト: 要素が非表示になりません ($Selector)"
    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null

    return $true
}

# --- 画面のローディングマスク解除（非表示）待機 ---（汎用）
function Wait-WebScreenUnlock {
    param ([int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec)

    $func = $MyInvocation.MyCommand.Name

    $js = @"
        function checkMask(doc, win) {
            const divs = doc.querySelectorAll('div');
            for (let d of divs) {
                const s = win.getComputedStyle(d);
                // s.zIndexが "auto" や空文字だった場合に "0" としてパースさせる
                const z = parseInt(s.zIndex || "0", 10);
                // 汎用判定：z-indexが100以上で、画面の80%以上を覆っており、透明でない
                if (!isNaN(z) && z >= 100 && (s.position === 'fixed' || s.position === 'absolute') &&
                    d.offsetWidth >= win.innerWidth * 0.8 && d.offsetHeight >= win.innerHeight * 0.8 &&
                    s.display !== 'none' && s.visibility !== 'hidden' && s.opacity !== '0') {
                    return true;
                }
            }
            
            // iframe / frame の中も再帰的に探査する
            const frames = doc.querySelectorAll('iframe, frame');
            for (let f of frames) {
                try {
                    if (f.contentDocument && f.contentWindow) {
                        if (checkMask(f.contentDocument, f.contentWindow)) return true;
                    }
                } catch(e) { 
                    // クロスドメイン(CORS)のセキュリティエラーは安全に無視
                }
            }
            return false;
        }
        return checkMask(document, window);
"@

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $wasBlocked = $false
    $initialGraceMs = 500 # ボタン押下後のマスク描画猶予時間(ms)

    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        try {
            $rawResult = Invoke-WebScript -Js $js

            # 配列返却時の末尾抽出
            if ($rawResult -is [array] -and $rawResult.Count -gt 0) {
                $isBlocked = $rawResult[-1]
            } else {
                $isBlocked = $rawResult
            }

            # 文字列・Bool値の判定ロジック
            $blockedBool = ($isBlocked -eq $true -or $isBlocked -eq "true")

            if ($blockedBool) {
                # マスク出現を検知
                $wasBlocked = $true
            }
            else {
                # マスク非表示時：一度でもマスクを検知した、または描画猶予時間を過ぎていれば完了判定
                if ($wasBlocked -or $sw.ElapsedMilliseconds -gt $initialGraceMs) {
                    if ($wasBlocked) {
                        Write-DebugLog -Message "[$func] 情報: マスク解除を確認しました" -Level Info
                    }
                    return "Screen Unlocked"
                }
            }
        } catch {
            # DOMアクセスエラー等は無視してリトライ
        }

        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 200
    }

#✘    Write-DebugLog -Message "[$func] 警告: 指定時間(${TimeoutSec}秒)内にマスクが消えませんでした。誤検知またはシステム遅延の可能性があります。" -Level Warning
    throw (New-EngineException -Func $func -Type "Timeout" -Message "画面のマスク解除がタイムアウトしました。誤検知またはシステム遅延の可能性があります。" -Details "Timeout: ${TimeoutSec}s")
    return "Timeout"
}

# --- 指定要素のクリック実行 ---
function Invoke-WebClick {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    $selectorEscaped = $Selector.Replace("'", "\'")

    # 要素出現の待機
    Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec | Out-Null

    # 多段iframeの再帰探索
    $js = @"
        $global:ENGINE_JS_UTILS
        var selector = '$selectorEscaped';
        var found = utilFindInFrames(window, function(win) {
            try {
// 変更         var el = win.document.querySelector('$selectorEscaped'); Shadow DOM対応へ（mode: 'open'）
                var el = deepQuerySelector(selector, win.document);
                if (el) {
                    el.scrollIntoView({block: 'center', inline: 'center'});
                    el.click();
                    return true;
                }
            } catch(e) {}
            return null;
        });
        return found === true;
"@

    $res = Invoke-WebScript -Js $js
    if (-not $res) {
        throw (New-EngineException -Func $func -Type "未発見" -Message "クリック実行に失敗しました" -Details $Selector)
    }
}

# --- テキストボックスへの値入力 ---
function Set-WebTextInput {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [Parameter(Mandatory=$true)][string]$Value,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    $selectorEscaped = $Selector.Replace("'", "\'")
    $valueEscaped = $Value.Replace("'", "\'").Replace("\", "\\")

    # 要素出現の待機
    Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec | Out-Null

    # value設定およびinput/changeイベントの発火
    $js = @"
        $global:ENGINE_JS_UTILS
        var selector = '$selectorEscaped';
        var val = '$valueEscaped';
        var found = utilFindInFrames(window, function(win) {
            try {
// 変更         var el = win.document.querySelector(selector); Shadow DOM対応へ（mode: 'open'）
                var el = deepQuerySelector(selector, win.document);
                if (el) {
                    el.value = val;
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    return true;
                }
            } catch(e) {}
            return null;
        });
        return found === true;
"@

    $res = Invoke-WebScript -Js $js
    if (-not $res) {
        throw (New-EngineException -Func $func -Type "未発見" -Message "入力に失敗しました" -Details $Selector)
    }
}

# --- ドロップダウンリストの指定値選択 ---
function Select-WebDropdown {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [Parameter(Mandatory=$true)][string]$Value,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    $selectorEscaped = $Selector.Replace("'", "\'")
    $valueEscaped = $Value.Replace("'", "\'")

    # 要素出現の待機
    Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec | Out-Null

    # value設定およびchangeイベントの発火
    $js = @"
        $global:ENGINE_JS_UTILS
        var selector = '$selectorEscaped';
        var val = '$valueEscaped';
        var found = utilFindInFrames(window, function(win) {
            try {
// 変更         var el = win.document.querySelector(selector); Shadow DOM対応へ（mode: 'open'）
                var el = deepQuerySelector(selector, win.document);
                if (el) {
                    el.value = val;
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    return true;
                }
            } catch(e) {}
            return null;
        });
        return found === true;
"@

    $res = Invoke-WebScript -Js $js
    if (-not $res) {
        throw (New-EngineException -Func $func -Type "未発見" -Message "ドロップダウンリストの選択に失敗しました" -Details $Selector)
    }
}

# --- チェックボックスの状態設定 ---
function Set-WebCheckbox {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [Parameter(Mandatory=$true)][bool]$State = $true,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    $selectorEscaped = $Selector.Replace("'", "\'")
    $stateStr = $State.ToString().ToLower()

    # 要素出現の待機
    Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec | Out-Null

    # 状態差分が存在する場合のみのclickおよびchangeイベント発火
    $js = @"
        $global:ENGINE_JS_UTILS
        var selector = '$selectorEscaped';
        var targetState = $stateStr; // true または false (JSのBooleanとして直接評価される)
        var found = utilFindInFrames(window, function(win) {
            try {
// 変更         var el = win.document.querySelector(selector); Shadow DOM対応へ（mode: 'open'）
                var el = deepQuerySelector(selector, win.document);
                if (el) {
                    if (el.checked !== targetState) {
                        el.click();
                        el.checked = targetState;
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                    return true;
                }
            } catch(e) {}
            return null;
        });
        return found === true; 
"@

    $res = Invoke-WebScript -Js $js
    if (-not $res) {
        throw (New-EngineException -Func $func -Type "未発見" -Message "チェックボックスの状態変更に失敗しました" -Details $Selector)
    }
}

# --- 指定要素のテキスト取得 ---
function Get-WebText {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    $selectorEscaped = $Selector.Replace("'", "\'")

    # 要素出現の待機
    Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec | Out-Null

    # innerTextまたはvalueの返却
    $js = @"
        $global:ENGINE_JS_UTILS
        var selector = '$selectorEscaped';
        var text = utilFindInFrames(window, function(win) {
            try {
// 変更         var el = win.document.querySelector(selector); Shadow DOM対応へ（mode: 'open'）
                var el = deepQuerySelector(selector, win.document);
                // 要素があれば String、無ければ null を返す
                if (el) { return el.innerText || el.value || ''; }
            } catch(e) {}
            return null;
        });
        return text !== null ? text : '';
"@

    return Invoke-WebScript -Js $js
}

# --- 指定要素の存在確認（例外を投げず True/False の文字列を返す） ---
function Test-WebElement {
    param (
        [string]$Selector = "",
        [string]$XPath = "",
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name

    if ([string]::IsNullOrEmpty($Selector) -and [string]::IsNullOrEmpty($XPath)) {
        throw (New-EngineException -Func $func -Type "引数エラー" -Message "Selector または XPath のいずれかを指定してください。")
    }

    try {
        if (-not [string]::IsNullOrEmpty($Selector)) {
            Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec | Out-Null
        } else {
            Wait-WebXPathElement -XPath $XPath -TimeoutSec $TimeoutSec | Out-Null
        }
        return "True"
    } catch {
        # タイムアウト等、要素が見つからなかった場合は例外を握りつぶして False を返す
        return "False"
    }
}

# --- Web要素の属性値(Attribute)を取得 ---
function Get-WebAttribute {
    param (
        [string]$Selector = "",
        [string]$XPath = "",
        [string]$Attribute = "",
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name

    # 1. パラメータの必須チェック
    if ([string]::IsNullOrEmpty($Attribute)) {
        throw (New-EngineException -Func $func -Type "引数エラー" -Message "Attribute パラメータは必須です。")
    }
    if ([string]::IsNullOrEmpty($Selector) -and [string]::IsNullOrEmpty($XPath)) {
        throw (New-EngineException -Func $func -Type "引数エラー" -Message "Selector または XPath のいずれかを指定してください。")
    }

#✘ Write-DebugLog -Message "[$func] 実行: 属性取得 ($Attribute)" -Level Info

    try {
        $js = ""
        # 2. Selector または XPath に応じた待機処理とJS組み立て
        if (-not [string]::IsNullOrEmpty($Selector)) {
            # 待機関数が返す余分な "True" を $null で破棄して戻り値への混入を防ぐ
            $null = Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec
            
            # JS組み立て (シングルクォートをエスケープ)
            $safeSelector = $Selector.Replace("'", "\'")
            $js = "return document.querySelector('$safeSelector').getAttribute('$Attribute');"
        }
        elseif (-not [string]::IsNullOrEmpty($XPath)) {
            # 待機関数が返す余分な "True" を $null で破棄して戻り値への混入を防ぐ
            $null = Wait-WebXPathElement -XPath $XPath -TimeoutSec $TimeoutSec
            
            # JS組み立て (シングルクォートをエスケープ)
            $safeXPath = $XPath.Replace("'", "\'")
            $js = "var node = document.evaluate('$safeXPath', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; return node ? node.getAttribute('$Attribute') : null;"
        }

        # 3. JSの実行 (内部関数を利用してブラウザに命令)
        $result = Invoke-WebView2NativeScript -Js $js
        
        # 取得結果が null の場合は空文字を返す
        if ($result -eq "null" -or $null -eq $result) {
            return ""
        }
        
        return $result
    }
    catch {
        # すでに New-EngineException で生成されたフォーマット済みエラー([関数名] [エラー種別]:)なら、そのまま上へ投げる
        if ($_.Exception.Message -match "^\[.*?\] \[.*?\]:") {
            throw $_
        }
        
        # 本当に予期せぬエラー（純粋なJS構文エラーなど）の場合のみラップする
        throw (New-EngineException -Func $func -Type "JSエラー" -Message "属性値の取得中にエラーが発生しました。" -Details $_.Exception.Message)
    }
}

# --- 現在のページURLの取得 ---
function Get-WebUrl {
    $webview = Get-ActiveWebView
    if ($webview -and $webview.Source) {
        return $webview.Source.ToString()
    }
    return $null
}

# --- 現在のページタイトルの取得 ---
function Get-WebTitle {
    try { return Invoke-WebScript -Js "return document.title;" }
    catch { return $null }
}

# --- 指定ディレクトリ・指定ファイル名への完全サイレントダウンロードを有効化 ---
function Enable-SilentDownload {
    param (
        [Parameter(Mandatory = $true)][string]$DownloadDirectory,
        [string]$FileName = "" 
    )

    $func = $MyInvocation.MyCommand.Name

    try {
        if (-not (Test-Path $DownloadDirectory)) {
            New-Item -ItemType Directory -Path $DownloadDirectory -Force | Out-Null
        }

        # 保存先とファイル名をグローバル変数にセット（毎回の書き換えに対応）
        $global:TargetDownloadDirectory = $DownloadDirectory
        $global:TargetDownloadFileName = $FileName
        
        $webview = Get-ActiveWebView
        if ($null -eq $webview -or $null -eq $webview.CoreWebView2) {
            throw "WebView2インスタンスが取得できません"
        }

        if ($global:SilentDownloadRegistered) {
            Write-DebugLog -Message "[$func] ダウンロード先を更新: Dir=$DownloadDirectory, File=$FileName" -Level Info
            return
        }

        # new() を使わず、PowerShellの「型キャスト」を使って安全にイベントを登録
        $handler = [System.EventHandler[Microsoft.Web.WebView2.Core.CoreWebView2DownloadStartingEventArgs]] {
            param($sender, $e)

            # バックグラウンドスレッドを守る
            try {
                $e.Handled = $true

                # VBAからファイル名が指定されていればそれを使用し、無ければ元ファイル名を使用
                if ([string]::IsNullOrEmpty($global:TargetDownloadFileName)) {
                    $nameToSave = [System.IO.Path]::GetFileName($e.ResultFilePath)
                } else {
                    $nameToSave = $global:TargetDownloadFileName
                }

                # 最終的な保存先フルパス
                $e.ResultFilePath = Join-Path $global:TargetDownloadDirectory $nameToSave
            } catch {
                Write-DebugLog -Message "[Download] 致命的エラー: サイレント保存中に例外発生 ($($_.Exception.Message))" -Level Error
            }
        }

        # 作成した安全なハンドラをセット
        $webview.CoreWebView2.add_DownloadStarting($handler)
        $global:SilentDownloadRegistered = $true

        Write-DebugLog -Message "[$func] サイレントダウンロードを有効化しました。" -Level Info
    } catch {
        throw (New-EngineException -Func $func -Type "初期化エラー" -Message "ダウンロード動作の設定に失敗しました" -Details $_.Exception.Message)
    }
}

# --- ダウンロード（ファイル書き込み）の完了を待機 ---
function Wait-FileDownload {
    param (
        [Parameter(Mandatory = $true)][string]$FilePath,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )
    
    $func = $MyInvocation.MyCommand.Name

    $waitSw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($waitSw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        
        # WebView2のイベント（ダウンロード開始等）をブロックさせないための息継ぎ処理
        [System.Windows.Forms.Application]::DoEvents()

        # 1. ファイルが作成されているか確認
        if (Test-Path $FilePath) {
            try {
                # 2. Chrome/Edge特有の .crdownload（一時ファイル）が存在しないか確認
                $tempFile = $FilePath + ".crdownload"
                if (-not (Test-Path $tempFile)) {
                    # 3. ファイルを排他モードで開けるか（書き込みロックが解除されたか）テスト
                    $stream = [System.IO.File]::Open($FilePath, 'Open', 'Read', 'None')
                    $stream.Close()
                    
#●                 Write-DebugLog -Message "[$func] 情報: ファイル保存完了 ($FilePath)" -Level Info
                    return $FilePath
                }
            } catch {
                # ロック中のため待機継続
            }
        }
        Start-Sleep -Milliseconds 200
    }
    
    throw (New-EngineException -Func $func -Type "Timeout" -Message "ファイルの保存完了確認がタイムアウトしました" -Details $FilePath)
}

# --- 埋め込みPDF(embed)の絶対URLを安全に取得 ---
function Get-WebEmbedPdfUrl {
    $func = $MyInvocation.MyCommand.Name

    $js = @"
        $global:ENGINE_JS_UTILS
        var url = utilFindInFrames(window, function(win) {
            try {
                // embed要素を探索（より厳密に type 属性をチェック）
                var el = win.document.querySelector('embed[type="application/pdf"], embed[src*=".pdf"]');
                if (!el) return null;
                
                var rawSrc = el.getAttribute('src') || el.src;
                
                // Chromiumベースのブラウザで直リンクのPDFを開くと src="about:blank" になる仕様への対応
                if (!rawSrc || rawSrc === 'about:blank') {
                    return win.location.href; // ページ(フレーム)のURL自体をPDFのURLとする
                }
                
                try {
                    // 絶対URLに変換して返す
                    return new URL(rawSrc, win.location.href).href;
                } catch(e) {
                    return rawSrc;
                }
            } catch(e) {
                return null;
            }
        });
        return url || '';
"@

    $res = Invoke-WebScript -Js $js

    if ([string]::IsNullOrWhiteSpace($res)) {
        throw (New-EngineException -Func $func -Type "未発見" -Message "PDFのembed要素、またはURL(src)が見つかりません")
    }

    return $res
}

# --- Fetch API を利用した裏側でのサイレントダウンロード発火 (file:// 完全互換版) ---
function Invoke-WebFetchDownload {
    param ([string]$TargetUrl = "")

    $func = $MyInvocation.MyCommand.Name

    # 1. TargetUrl が指定されていない場合は現在のページURLを使用
    $fetchUrl = $TargetUrl

    # 2. file:// プロトコルの場合、画面遷移を防ぐため Data URI (data:application/pdf;base64,...) へ変換
    if ($fetchUrl.StartsWith("file://", [System.StringComparison]::OrdinalIgnoreCase)) {
        # file:/// C:/... のパスを抽出
        $localPath = [System.Uri]::UnescapeDataString(($fetchUrl -replace '^file:///', ''))
        $localPath = $localPath -replace '/', '\'
        
        if (Test-Path $localPath) {
            $bytes = [System.IO.File]::ReadAllBytes($localPath)
            $base64 = [System.Convert]::ToBase64String($bytes)
            $fetchUrl = "data:application/pdf;base64,$base64"
        }
    }

    $targetUrlEscaped = $fetchUrl.Replace("'", "\'")

    # 3. Fetch API コアロジックを実行
    $js = @"
        try {
            var target = '$targetUrlEscaped';
            if (!target) target = window.location.href;

            fetch(target)
                .then(r => r.blob())
                .then(b => {
                    var url = window.URL.createObjectURL(b);
                    var a = document.createElement('a');
                    a.href = url;
                    a.download = 'auto_download_temp'; // 実際のファイル名はエンジン側で上書き
                    document.body.appendChild(a);
                    a.click();
                    
                    // メモリ解放とDOMのお掃除
                    setTimeout(function() {
                        window.URL.revokeObjectURL(url);
                        if (a.parentNode) a.parentNode.removeChild(a);
                    }, 1000);
                });
            return true;
        } catch(e) {
            return false;
        }
"@

    $res = Invoke-WebScript -Js $js
    
    if (-not $res) {
        throw (New-EngineException -Func $func -Type "JSエラー" -Message "Fetch APIを利用したダウンロードトリガーの発火に失敗しました" -Details $TargetUrl)
    }

    return "Fetch Download Triggered"
}
