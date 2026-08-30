# ==============================================================================
# UIAccessibility (UIA) API と WebView2 (DOM) を連携させた要素ピッカー
# ==============================================================================

# ------------------------------------------------------------------------------
# [DOM逆引きコア] UIA情報をもとにブラウザ内から最適なCSSセレクタを算出 (位置ベース優先)
function Get-WebSelectorFromUia {
    param(
        [string]$UiaName,
        [string]$UiaType,
        [string]$UiaId,
        [string]$UiaHelpText,
        [System.Windows.Automation.AutomationElement]$UiaElement
    )

    $func = $MyInvocation.MyCommand.Name

    # 1. UIA (OSレベル) の絶対座標を計算
    $rect = $UiaElement.Current.BoundingRectangle
    $screenX = $rect.Left + ($rect.Width / 2.0)
    $screenY = $rect.Top  + ($rect.Height / 2.0)

    # 2. OS絶対座標 -> ブラウザ相対座標 (Client座標) への変換
    $clientX = $screenX
    $clientY = $screenY
    if ($null -ne $global:WebViewCtrl) {
        try {
            if (-not ('System.Drawing.Point' -as [type])) { Add-Type -AssemblyName System.Drawing }
            $pt = New-Object System.Drawing.Point([int]$screenX, [int]$screenY)
            $clientPt = $global:WebViewCtrl.PointToClient($pt)
            $clientX = $clientPt.X
            $clientY = $clientPt.Y
        } catch {}
    }
    Write-DebugLog -Message "[UiaMatch] Screen(X=$screenX, Y=$screenY) -> Client(X=$clientX, Y=$clientY)" -Level Info

    # 3. DPR(画面拡大率)とスクロール量の取得、およびUIA座標の補正
    $dpr = 1.0
    $scrollX = 0
    $scrollY = 0
    try {
        $jsMetrics = "return JSON.stringify({ dpr: window.devicePixelRatio || 1, scrollX: window.scrollX || 0, scrollY: window.scrollY || 0 });"
        $metricsRes = Invoke-WebScript -Js $jsMetrics
        if (-not [string]::IsNullOrWhiteSpace($metricsRes)) {
            $metrics = $metricsRes | ConvertFrom-Json
            if ($metrics.dpr -gt 0) { $dpr = $metrics.dpr }
            $scrollX = $metrics.scrollX
            $scrollY = $metrics.scrollY
        }
    } catch {}

    # DOMとの比較用座標（CSS論理ピクセル化 ＋ スクロール加算による絶対位置化）
    $uiaCenterX = ($clientX / $dpr) + $scrollX
    $uiaCenterY = ($clientY / $dpr) + $scrollY
    Write-DebugLog -Message "[UiaMatch] Dpr=$dpr, Scroll(X=$scrollX, Y=$scrollY), Target(X=$uiaCenterX, Y=$uiaCenterY)" -Level Info

    # 4. 事前取得した DOMSnapshot の展開
    $domNodes = $global:DomNodes
    if (-not $domNodes -or $domNodes.Count -eq 0) {
        Write-DebugLog -Message "[$func] 警告: 比較用のDOMSnapshotが存在しません" -Level Warn
        return $null
    }

    # 5. 位置ベース UIA → DOM マッチング（最優先ルート）
    $bestNode = $null
    $bestScore = [double]::NegativeInfinity

    foreach ($node in $domNodes) {
        if (-not $node.boundingBox) { continue }

        $bb = $node.boundingBox
        $domCenterX = $bb.x + ($bb.width / 2.0)
        $domCenterY = $bb.y + ($bb.height / 2.0)

        # ユークリッド距離による近接度計算
        $dx = $domCenterX - $uiaCenterX
        $dy = $domCenterY - $uiaCenterY
        $distance = [math]::Sqrt($dx * $dx + $dy * $dy)
        $score = 1000.0 / (1.0 + $distance)
        
        # RPAの操作対象になりやすいタグを優遇加点
        if ($node.tagName -eq "input") { $score += 100 }
        if ($node.tagName -eq "button") { $score += 50 }
        
        # placeholder と UIA.HelpText が一致する場合はさらに強力に加点
        if ($node.placeholder -and $UiaHelpText -and $node.placeholder -eq $UiaHelpText) {
            $score += 200
        }

        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestNode = $node
        }
    }

    # 6. 位置マッチング成功時のセレクタ生成
    if ($bestNode) {
        Write-DebugLog -Message "[UiaMatch] 有力候補: Tag=$($bestNode.tagName), Id=$($bestNode.id), Name=$($bestNode.name) (Score: $([math]::Round($bestScore, 2)))" -Level Info
            
        if ($bestNode.selector) { return $bestNode.selector }
        
        # Route Aと同等の堅牢なセレクタ (IDとNameの複合) を生成
        $tagName = if ($bestNode.tagName) { $bestNode.tagName } else { "*" }
        if ($bestNode.id -and $bestNode.name) {
            return "$tagName#$($bestNode.id)[name=`"$($bestNode.name)`"]"
        }
        if ($bestNode.name)     { return ($tagName + '[name="' + $bestNode.name + '"]') }
        if ($bestNode.id)       { return ('#' + $bestNode.id) }
    }

    # 7. フォールバック：レーベンシュタイン距離による類似テキスト検索
    Write-DebugLog -Message "[$func] 情報: 有効な候補が見つからず、Fuzzy検索へフォールバックします" -Level Info
    $cssSelector = Get-WebSelectorFromUiaFuzzy -UiaName $UiaName -UiaType $UiaType -UiaId $UiaId -UiaHelpText $UiaHelpText
    if ($cssSelector) { return $cssSelector }

    return $null
}

# ------------------------------------------------------------------------------
# --- [フォールバック解析] レーベンシュタイン距離を用いた曖昧 (Fuzzy) 検索 ---
function Get-WebSelectorFromUiaFuzzy {
    param(
        [string]$UiaName,
        [string]$UiaType,
        [string]$UiaId,
        [string]$UiaHelpText
    )

    $domNodes = $global:DomNodes
    if (-not $domNodes -or $domNodes.Count -eq 0) { return $null }

    # レーベンシュタイン距離の高速計算 (1次元配列を使い回してメモリ割り当てを抑制)
    function Get-LevenshteinDistanceFast {
        param([string]$s, [string]$t)

        if ([string]::IsNullOrEmpty($s)) { return $t.Length }
        if ([string]::IsNullOrEmpty($t)) { return $s.Length }

        $v0 = New-Object int[] ($t.Length + 1)
        $v1 = New-Object int[] ($t.Length + 1)

        for ($i = 0; $i -le $t.Length; $i++) { $v0[$i] = $i }

        for ($i = 0; $i -lt $s.Length; $i++) {
            $v1[0] = $i + 1
            for ($j = 0; $j -lt $t.Length; $j++) {
                $cost = if ($s[$i] -eq $t[$j]) { 0 } else { 1 }
                $v1[$j + 1] = [Math]::Min([Math]::Min($v1[$j] + 1, $v0[$j + 1] + 1), ($v0[$j] + $cost))
            }
            # 次の行の計算に向けて配列状態をコピー
            for ($j = 0; $j -le $t.Length; $j++) { $v0[$j] = $v1[$j] }
        }
        return $v1[$t.Length]
    }

    $bestNode = $null
    $bestScore = [double]::NegativeInfinity
    
    # UIA側の情報から比較用のベース文字列を生成
    $uiaText = ("$UiaName $UiaHelpText $UiaId $UiaType").Trim()
    if ([string]::IsNullOrWhiteSpace($uiaText)) { return $null }
    
    # 探索対象を操作可能な要素に絞り込み、計算コストを削減
    $interactiveTags = @("input", "button", "a", "select", "textarea", "label", "span", "div", "img")

    foreach ($node in $domNodes) {
        if (-not $node.tagName -or $node.tagName -notin $interactiveTags) { continue }

        $domText = ""
        if ($node.placeholder) { $domText += $node.placeholder + " " }
        if ($node.name)        { $domText += $node.name + " " }
        if ($node.id)          { $domText += $node.id + " " }
        if ($node.tagName)     { $domText += $node.tagName + " " }

        $domText = $domText.Trim()
        if (-not $domText) { continue }

        # レーベンシュタイン距離の計算
        $dist = Get-LevenshteinDistanceFast -s $uiaText -t $domText
        $score = 1000.0 / (1.0 + $dist)
        if ($node.tagName -eq "input") { $score += 50 }

        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestNode = $node
        }
    }

    if ($bestNode) {
        if ($bestNode.selector) { return $bestNode.selector }
        if ($bestNode.name)     { return $bestNode.tagName + '[name="' + $bestNode.name + '"]' }
        if ($bestNode.id)       { return '#' + $bestNode.id }
    }
    return $null
}

# ------------------------------------------------------------------------------
# --- [デバッグ/診断用] セレクタ解決能力のテスト出力 (本番コード生成には影響しない) ---
function Test-WebSelectorResolution {
    param([System.Windows.Automation.AutomationElement]$UiaElement)

    # UIA 情報
    $uiaName     = $UiaElement.Current.Name
    $uiaType     = $UiaElement.Current.ControlType.ProgrammaticName
    $uiaId       = $UiaElement.Current.AutomationId
    $uiaHelpText = $UiaElement.Current.HelpText

    Write-Host "[UIA] Name=$uiaName, Type=$uiaType, Id=$uiaId, HelpText=$uiaHelpText" -ForegroundColor Yellow

    # [TEST:1] 位置ベースマッチング
    $selectorPos = Get-WebSelectorFromUia -UiaName $uiaName -UiaType $uiaType -UiaId $uiaId -UiaHelpText $uiaHelpText -UiaElement $UiaElement
    $logMsg = "[TEST:1] 位置ベースUIA→DOM/"
    if ($selectorPos) { Write-Host "$logMsg [OK] → Selector=$selectorPos" -ForegroundColor Green }
    else { Write-Host "$logMsg [NG] → Fuzzy へフォールバック" -ForegroundColor Red }

    # [TEST:2] 曖昧(Fuzzy)検索マッチング
    $selectorFuzzy = Get-WebSelectorFromUiaFuzzy -UiaName $uiaName -UiaType $uiaType -UiaId $uiaId -UiaHelpText $uiaHelpText
    $logMsg = "[TEST:2] Fuzzy(曖昧検索).../"
    if ($selectorFuzzy) { Write-Host "$logMsg [OK] → Selector=$selectorFuzzy" -ForegroundColor Green }
    else { Write-Host "$logMsg [NG] Fuzzy でも一致なし" -ForegroundColor Red }

    # [TEST:3] 最終判定結果
    $logMsg = "[TEST:3] 最終判定/"
    if ($selectorPos) {
        Write-Host "$logMsg [RESULT] 位置マッチング → $selectorPos" -ForegroundColor Cyan
        return $selectorPos
    } elseif ($selectorFuzzy) {
        Write-Host "$logMsg [RESULT] Fuzzy(曖昧検索)→ $selectorFuzzy" -ForegroundColor Cyan
        return $selectorFuzzy
    } else {
        Write-Host $logMsg "[RESULT] UIA フォールバックへ移行" -ForegroundColor Yellow
        return $null
    }
}

# ==============================================================================
# --- [VBA エントリーポイント] 要素ピッカーの実行要求窓口 ---
# ==============================================================================
function Invoke-UiaRecordStep {
    param (
        [int]$WaitSec = 3,
        [bool]$UsePointReverse = $false
    )

    $func = $MyInvocation.MyCommand.Name

    Write-DebugLog -Message "================================================================================" -Level Info
    # 最新のDOMスナップショットと座標情報を取得
    $global:DomNodes = Get-DomSnapshotWithBoxModel

    try {
        # WinForms(マウスポインタ座標取得)
        if (-not ('System.Windows.Forms.Cursor' -as [type])) {
            Add-Type -AssemblyName System.Windows.Forms
        }
        
        # ユーザーが対象要素にマウスを合わせるための待機（VBA側のカウントダウンと同期）
        Start-Sleep -Seconds $WaitSec
        
        # [MAIN]: 要素特定オーケストレーターへ処理を委譲
        $generatedCode = Process-RecordAction -IsGetNameMode $false -UsePointReverse $UsePointReverse

        if ([string]::IsNullOrWhiteSpace($generatedCode)) {
            throw (New-EngineException -Func $func -Type "未発見" -Message "対象要素の取得、またはコードの生成に失敗しました")
        }
        return $generatedCode

    } catch {
        # VBA側に、On Error Resume Next がある前提で、throw
        throw (New-EngineException -Func $func -Type "UIAエラー" -Message "UIA要素の記録処理中にエラーが発生しました" -Details $_.Exception.Message)
    }
}

# ------------------------------------------------------------------------------
# --- [共通JSコア] 物理座標から要素を逆引きするShadow DOM貫通ロジック ---
function Get-CorePointReverseJs {
    param(
        [int]$ClientX,
        [int]$ClientY
    )

    return @"
    try {
        let targetX = $ClientX;
        let targetY = $ClientY;
        const dpr = window.devicePixelRatio || 1;

        // 物理ピクセルからCSSピクセル(論理座標)への変換
        targetX = targetX / dpr;
        targetY = targetY / dpr;

        // 座標から最深部の操作対象要素を取得する再帰関数
        function deepElementFromPoint(x, y, rootNode) {
            // ヘルパ: shadowRoot 内を安全に elementFromPoint で探索（再帰）
            function probeShadow(root, px, py) {
                try {
                    if (!root) return null;
                    if (root.elementFromPoint) {
                        const el = root.elementFromPoint(px, py);
                        if (el) return el;
                    }
                    // fallback: querySelector で input/button を探す
                    const candidates = root.querySelectorAll ? root.querySelectorAll('input,button,textarea,select,a,[role="search"],[role="textbox"]') : [];
                    if (candidates && candidates.length) return candidates[0];
                } catch (e) {}
                return null;
            }

            // --------------------------------
            // 1) elementsFromPoint で重なり要素列を取得し、上から順に掘る
            let stacked = [];
            try {
                if (rootNode.elementsFromPoint) {
                    stacked = rootNode.elementsFromPoint(x, y);
                } else if (document.elementsFromPoint) {
                    stacked = document.elementsFromPoint(x, y);
                } else if (document.msElementsFromPoint) {
                    stacked = document.msElementsFromPoint(x, y);
                } else {
                    const el = (rootNode.elementFromPoint || document.elementFromPoint).call(rootNode || document, x, y);
                    if (el) stacked = [el];
                }
            } catch (e) {
                stacked = [];
            }

            // --------------------------------
            // 2) stacked の各要素について、shadowRoot / slot / 内部 input を探索
            for (let i = 0; i < stacked.length; i++) {
                let el = stacked[i];
                if (!el) continue;

                // 隠しinputは無視する
                if (el.tagName === 'INPUT' && el.type === 'hidden') continue;

                // Aタグ(ハイパーリンク)も操作対象として認識する
                if (/^(INPUT|BUTTON|TEXTAREA|SELECT|A)$/.test(el.tagName)) return el;

                try {
                    const inner = el.querySelector && el.querySelector('input,textarea,select,button,a,[role="textbox"],[role="search"]');
                    if (inner) return inner;
                } catch (e) {}

                try {
                    if (el.tagName === 'SLOT') {
                        const assigned = el.assignedElements ? el.assignedElements() : [];
                        if (assigned && assigned.length) {
                            for (let a of assigned) {
                                const found = deepElementFromPoint(x, y, a.ownerDocument || document);
                                if (found) return found;
                            }
                        }
                    }
                } catch (e) {}

                try {
                    if (el.shadowRoot) {
                        const sEl = probeShadow(el.shadowRoot, x, y);
                        if (sEl) {
                            if (sEl.shadowRoot) {
                                const deep = deepElementFromPoint(x, y, sEl.shadowRoot);
                                if (deep) return deep;
                            }
                            return sEl;
                        }

                        try {
                            const inner = el.shadowRoot.elementFromPoint ? el.shadowRoot.elementFromPoint(x, y) : null;
                            if (inner && inner !== el) {
                                if (inner.shadowRoot) {
                                    const deep = deepElementFromPoint(x, y, inner.shadowRoot);
                                    if (deep) return deep;
                                }
                                const innerInput = inner.querySelector && inner.querySelector('input,textarea,select,button,a,[role="textbox"],[role="search"]');
                                if (innerInput) return innerInput;
                                return inner;
                            }
                        } catch (e) {}
                    }
                } catch (e) {}

                try {
                    // IFRAME だけでなく FRAME にも対応する
                    if (el.tagName === 'IFRAME' || el.tagName === 'FRAME') {
                        const rect = el.getBoundingClientRect();
                        const iframeX = x - rect.left;
                        const iframeY = y - rect.top;
                        const doc = el.contentDocument;
                        if (doc) {
                            const inner = deepElementFromPoint(iframeX, iframeY, doc);
                            if (inner) return inner;
                        }
                    }
                } catch (e) {}
            }

            // --------------------------------
            // 3) stacked で見つからなければ、最後に試す（互換性およびLABEL補正）
            try {
                const fallback = (rootNode.elementFromPoint || document.elementFromPoint).call(rootNode || document, x, y);
                if (fallback) {
                    if (fallback.tagName === 'LABEL') {
                        const forId = fallback.getAttribute && fallback.getAttribute('for');
                        if (forId) {
                            const linked = document.getElementById(forId);
                            if (linked) return linked;
                        }
                        const innerInput = fallback.querySelector && fallback.querySelector('input,textarea,select');
                        if (innerInput) return innerInput;
                    }
                    return fallback;
                }
            } catch (e) {}

            return null;
        }

        // CSSセレクタおよびXPathの自動生成
        function buildCssSelector(el) {
            if (!el) return '';
            const cssEsc = (typeof CSS !== 'undefined' && CSS.escape) ? CSS.escape : s => s;
            if (el.id) {
                const idSel = '#' + cssEsc(el.id);
                try { if (document.querySelectorAll(idSel).length === 1) return idSel; } catch(e) {}
                if (el.name) return el.tagName.toLowerCase() + idSel + '[name="' + cssEsc(el.name) + '"]';
            }
            if (el.name) return el.tagName.toLowerCase() + '[name="' + cssEsc(el.name) + '"]';
            if (el.className && typeof el.className === 'string') {
                const classes = el.className.trim().split(/\s+/).map(cls => cssEsc(cls)).join('.');
                if (classes) return el.tagName.toLowerCase() + '.' + classes;
            }
            return '';
        }

        // 対象要素の絶対XPathを生成する関数 (属性(id, name, class)から一意のCSSセレクタが構築できない場合)
        function buildXPath(el) {
            if (!el) return '';
            let xpath = '';
            let node = el;
            while (node && node.nodeType === 1) {
                const siblings = node.parentNode ? node.parentNode.children : [];
                let count = 0, index = 1;
                for (let i = 0; i < siblings.length; i++) {
                    if (siblings[i].tagName === node.tagName) count++;
                    if (siblings[i] === node) { index = count; break; }
                }
                xpath = '/' + node.tagName.toLowerCase() + '[' + index + ']' + xpath;
                node = node.parentNode;
            }
            return xpath;
        }

        const targetElement = deepElementFromPoint(targetX, targetY, document);

        const log = {
            dpr: dpr,
            targetX: targetX,
            targetY: targetY,
            elementTag: targetElement ? targetElement.tagName : null,
            elementId: targetElement ? targetElement.id : null,
            elementName: targetElement ? targetElement.name : null
        };

        if (!targetElement) {
            return JSON.stringify({ error: 'elementFromPoint=null', log: log });
        }

        // 最適な要素指定パス(セレクタ)の評価と決定
        /*
         * 1. 優先的にCSSセレクタの生成を試行し、id/name/classを含む有効なパスか判定
         * 2. 条件を満たさない（または生成不可の）場合はXPathを最終手段として採用
         */
        const cssResult = buildCssSelector(targetElement);
        const finalSelector = (cssResult && (cssResult.indexOf('#') !== -1 || cssResult.indexOf('[name=') !== -1 || cssResult.indexOf('.') !== -1)) 
            ? cssResult : 'XPATH:' + buildXPath(targetElement);

        const rect = targetElement.getBoundingClientRect();
        const scrollX = window.scrollX || window.pageXOffset;
        const scrollY = window.scrollY || window.pageYOffset;

        return JSON.stringify({
            selector: finalSelector,
            outerHtml: targetElement.outerHTML || '',
            boundingBox: { 
                x: Math.round(rect.left + scrollX), 
                y: Math.round(rect.top + scrollY), 
                width: Math.round(rect.width), 
                height: Math.round(rect.height) 
            },
            log: log
        });
    } catch (e) {
        return JSON.stringify({ error: e.message });
    }
"@
}

# ------------------------------------------------------------------------------
# --- [ルートA] 物理座標ベースの要素逆引きおよび座標補正 ---
function Invoke-BrowserPointReverse {
    param(
        [System.Windows.Point]$ScreenPoint,
        [string]$TargetControlType,
        [string]$CurrentValue
    )

    # 1. スクリーン座標(OS)からクライアント座標(ブラウザ内)への変換
    $clientX = $ScreenPoint.X
    $clientY = $ScreenPoint.Y
    if ($null -ne $global:WebViewCtrl) {
        try {
            if (-not ('System.Drawing.Point' -as [type])) { Add-Type -AssemblyName System.Drawing }
            $pt = New-Object System.Drawing.Point([int]$ScreenPoint.X, [int]$ScreenPoint.Y)
            $clientPt = $global:WebViewCtrl.PointToClient($pt)
            $clientX = $clientPt.X
            $clientY = $clientPt.Y
        } catch {}
    }

    Write-DebugLog -Message "[PointReverse] Screen(X=$($ScreenPoint.X), Y=$($ScreenPoint.Y)) -> Client(X=$clientX, Y=$clientY)" -Level Info

    # 2. ブラウザ内でのヒットテスト（共通JSロジック）
    $jsPoint = Get-CorePointReverseJs -ClientX $clientX -ClientY $clientY

    $resRaw = Invoke-WebScript -Js $jsPoint
    if (-not $resRaw) { return $null }

    $domData = if ($resRaw -is [string]) { try { $resRaw | ConvertFrom-Json } catch { $null } } else { $resRaw }
    
    if ($domData.log) {
        Write-DebugLog -Message "[PointReverse] Dpr=$($domData.log.dpr), Target(X=$($domData.log.targetX), Y=$($domData.log.targetY))" -Level Info
        Write-DebugLog -Message "[PointReverse] elementFromPoint => Tag=$($domData.log.elementTag), Id=$($domData.log.elementId), Name=$($domData.log.elementName)" -Level Info
    }

    if ($null -eq $domData -or $domData.error) { return $null }

    # 3. 【座標補正ロジック】
    if ($domData.boundingBox -and $global:DomNodes) {
        $matchedNode = $global:DomNodes | Where-Object { $_.selector -and $_.selector -eq $domData.selector } | Select-Object -First 1

        if (-not $matchedNode -and $domData.log) {
            $matchedNode = $global:DomNodes | Where-Object {
                $_.tagName -eq $domData.log.elementTag.ToLower() -and
                (
                    ($_.id -and $_.id -eq $domData.log.elementId) -or
                    ($_.name -and $_.name -eq $domData.log.elementName)
                )
            } | Select-Object -First 1
        }

        # CDP座標による上書き
        if ($matchedNode -and $matchedNode.boundingBox) {
            Write-DebugLog -Message "[PointReverse] 情報: CDPの正確なBoundingBoxで座標を補正(同期)します" -Level Info
            $domData.boundingBox = $matchedNode.boundingBox
        }
    }

    Write-DebugLog -Message "[PointReverse] 確定結果: Selector='$($domData.selector)' | 座標(x:$($domData.boundingBox.x), y:$($domData.boundingBox.y))" -Level Info

    return (ConvertTo-RpaCommandCode -Selector $domData.selector -ControlType $TargetControlType -CurrentValue $CurrentValue -OuterHtml $domData.outerHtml -BoundingBox $domData.boundingBox)
}

# ------------------------------------------------------------------------------
# --- [ルートB] UIAメタデータからのDOM逆引き ---
function Invoke-UiaDomResolution {
    param(
        [string]$TargetName,
        [string]$TargetControlType,
        [string]$TargetId,
        [System.Windows.Automation.AutomationElement]$TargetElement,
        [string]$CurrentValue
    )

    $cssSelector = Get-WebSelectorFromUia -UiaName $TargetName -UiaType $TargetControlType -UiaId $TargetId -UiaHelpText "" -UiaElement $TargetElement
    if ([string]::IsNullOrWhiteSpace($cssSelector)) { return $null }

    $safeSel = $cssSelector -replace "'", "\'"
    $scriptDom = @"
        try {
            const sel = '$safeSel';
            const el = sel.indexOf('XPATH:') === 0 
                ? document.evaluate(sel.substring(6), document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue 
                : document.querySelector(sel);
            if(!el) return null;
            const rect = el.getBoundingClientRect();
            return JSON.stringify({
                outerHtml: el.outerHTML || '',
                boundingBox: { x: Math.round(rect.left), y: Math.round(rect.top), width: Math.round(rect.width), height: Math.round(rect.height) }
            });
        } catch(e) { return null; }
"@

    $resultDom = Invoke-WebScript -Js $scriptDom
    $parsedInfo = if ($resultDom -is [string]) { try { $resultDom | ConvertFrom-Json } catch { $null } } else { $resultDom }

    $outHtml = if ($parsedInfo -and $parsedInfo.outerHtml) { $parsedInfo.outerHtml } else { "" }
    $bbox = if ($parsedInfo -and $parsedInfo.boundingBox) { $parsedInfo.boundingBox } else { $null }

    return (ConvertTo-RpaCommandCode -Selector $cssSelector -ControlType $TargetControlType -CurrentValue $CurrentValue -OuterHtml $outHtml -BoundingBox $bbox)
}

# ------------------------------------------------------------------------------
# --- [ルートC] デスクトップ要素(UIAモード)へのフォールバック ---
<# ブラウザ以外のネイティブWindowsアプリがターゲットになった場合に実行される #>
function Build-FallbackStepCode {
    param(
        [string]$WindowName,
        [string]$TargetName,
        [string]$ControlType,
        [string]$CurrentValue,
        [bool]$IsGetNameMode
    )

    $stepTargetName = if ($IsGetNameMode) { "*" } else { $TargetName }
    $stepCode = Build-RpaStepCode `
        -Action "Click" `
        -WindowName $WindowName `
        -TargetName $stepTargetName `
        -ControlType $ControlType `
        -Index 1 `
        -ActionValue $CurrentValue `
        -AnchorName "" `
        -AnchorDirection "None" `
        -AnchorSteps 0

    return "[RESULT]" + $stepCode
}

# ------------------------------------------------------------------------------
# --- [ヘルパー] 最終的なRPAコマンド(PowerShell文字列)の組み上げ ---
function ConvertTo-RpaCommandCode {
    param(
        [string]$Selector,
        [string]$ControlType,
        [string]$CurrentValue,
        [string]$OuterHtml,
        [object]$BoundingBox
    )

    $outHtmlSafe = if ($OuterHtml) { $OuterHtml -replace '"', "'" -replace "`r`n|`n|`r", " " } else { "" }
    $bboxStr = if ($BoundingBox) { $BoundingBox | ConvertTo-Json -Compress -Depth 5 } else { "" }
    $safeValue = if ($CurrentValue) { $CurrentValue -replace '"', '`"' } else { "" }

    if ($Selector.StartsWith("XPATH:")) {
        $xp = $Selector.Substring(6)
        if ($ControlType -eq "Edit") {
            return "[RESULT]Set-WebXPathTextInput -XPath `"$xp`" -Value `"$safeValue`" -OuterHtml `"$outHtmlSafe`" -BoundingBox `"$bboxStr`""
        } else {
            return "[RESULT]Invoke-WebXPathClick -XPath `"$xp`" -OuterHtml `"$outHtmlSafe`" -BoundingBox `"$bboxStr`""
        }
    } else {
        if ($ControlType -eq "Edit") {
            return "[RESULT]Set-WebTextInput -Selector `"$Selector`" -Value `"$safeValue`" -OuterHtml `"$outHtmlSafe`" -BoundingBox `"$bboxStr`""
        } else {
            return "[RESULT]Invoke-WebClick -Selector `"$Selector`" -OuterHtml `"$outHtmlSafe`" -BoundingBox `"$bboxStr`""
        }
    }
}

# ==============================================================================
# --- [メインオーケストレーター] Process-RecordAction ---
# ==============================================================================
function Process-RecordAction {
    param(
        [bool]$IsGetNameMode,
        [bool]$UsePointReverse = $false
    )

    $func = $MyInvocation.MyCommand.Name

    try {
        # 1. OSレベルでのUIA要素取得 (対象要素の特定)
        $elementInfo = Get-UiaTargetInfoFromCursor
        if (-not $elementInfo) { return $null }

        $targetElement     = $elementInfo.TargetElement
        $stepWindowName    = $elementInfo.WindowName
        $targetName        = $targetElement.Current.Name
        $targetId          = $targetElement.Current.AutomationId
        
        $pName = $targetElement.Current.ControlType.ProgrammaticName
        $targetControlType = if ($pName) { $pName.Replace("ControlType.", "") } else { "Unknown" }

        # 診断テスト実行 (結果出力のみ)
        Test-WebSelectorResolution -UiaElement $targetElement | Out-Null

        # 2. UIAから現在の入力値（ValuePattern）を物理的に取得 (DOM側で保持されていない値の奪取)
        $currentValue = ""
        $valuePattern = $null
        if ($targetElement.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {
            $currentValue = $valuePattern.Current.Value -replace "`r`n|`n|`r", " "
        }

        # [ルートA] 物理座標からのDOM逆引き (ブラウザ内限定)
        if ($UsePointReverse -and ($stepWindowName -match "RPA Browser" -or $stepWindowName -match "Edge" -or $stepWindowName -match "Chrome")) {
            $rect = $targetElement.Current.BoundingRectangle
            $screenX = [int]($rect.Left + ($rect.Width  / 2))
            $screenY = [int]($rect.Top  + ($rect.Height / 2))
            
            $resultCode = Invoke-BrowserPointReverse -ScreenPoint (New-Object System.Windows.Point($screenX, $screenY)) -TargetControlType $targetControlType -CurrentValue $currentValue
            if (-not [string]::IsNullOrWhiteSpace($resultCode)) { return $resultCode }
        }

        # [ルートB] UIAメタデータからのDOM逆引き
        $resultCodeB = Invoke-UiaDomResolution -TargetName $targetName -TargetControlType $targetControlType -TargetId $targetId -TargetElement $targetElement -CurrentValue $currentValue
        if (-not [string]::IsNullOrWhiteSpace($resultCodeB)) { return $resultCodeB }

        # [ルートC] デスクトップ要素へのフォールバック
        return Build-FallbackStepCode -WindowName $stepWindowName -TargetName $targetName -ControlType $targetControlType -CurrentValue $currentValue -IsGetNameMode $IsGetNameMode

    } catch {
        throw (New-EngineException -Func $func -Type "内部エラー" -Message "記録アクション生成中にエラーが発生しました" -Details $_.Exception.Message)
    }
}

# ------------------------------------------------------------------------------
# UIAユーティリティ: 座標からの要素取得・解析

# --- 巨大コンテナ(Window等)を除外し、最小単位の要素を取得 ---
function Get-UiaElementFromPoint {
    $func = $MyInvocation.MyCommand.Name

    try {
        $point = New-Object System.Windows.Point([System.Windows.Forms.Cursor]::Position.X, [System.Windows.Forms.Cursor]::Position.Y)
        $maxRetries = 15
        $waitMs = 200
        
        for ($retry = 1; $retry -le $maxRetries; $retry++) {
            $element = [System.Windows.Automation.AutomationElement]::FromPoint($point)
            
            if ($element -and 
                $element.Current.ControlType -ne [System.Windows.Automation.ControlType]::Document -and
                $element.Current.ControlType -ne [System.Windows.Automation.ControlType]::Pane -and
                $element.Current.ControlType -ne [System.Windows.Automation.ControlType]::Window -and
                $element.Current.AutomationId -ne "__next") { 
                
                return $element 
            }
            Start-Sleep -Milliseconds $waitMs
        }
        return $element
    } catch {
        throw (New-EngineException -Func $func -Type "GetElementError" -Message "座標からのUIA要素取得に失敗しました" -Details $_.Exception.Message)
    }
}

# --- 要素の親階層を遡り、所属するウィンドウを特定する ---
function Get-UiaWindowFromElement {
    param( [System.Windows.Automation.AutomationElement]$TargetElement )

    $func = $MyInvocation.MyCommand.Name

    try {
        $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
        $current = $TargetElement
        while ($current -ne $null -and $current.Current.ControlType -ne [System.Windows.Automation.ControlType]::Window) {
            $current = $walker.GetParent($current)
        }
        return $current
    } catch {
        throw (New-EngineException -Func $func -Type "GetWindowError" -Message "親ウィンドウの特定に失敗しました" -Details $_.Exception.Message)
    }
}

# --- ウィンドウタイトルを正規化 (ブラウザ特有の動的文字、ページ数等の揺れを吸収) ---
function Format-UiaWindowName {
    param( [string]$RawWindowName )

    if ([string]::IsNullOrWhiteSpace($RawWindowName)) { return "" }
    return $RawWindowName `
        -replace "\s*および他\s*\d+\s*ページ.*$", "" `
        -replace "\s*-\s*(個人|仕事|InPrivate)(?=\s*-).*$", "" `
        -replace "\s*-[^-]+$", ""
}

# --- ターゲット要素のヒューリスティック探索（直上の親要素も含めたスコアリング評価） ---
function Get-UiaTargetInfoFromCursor {
    $func = $MyInvocation.MyCommand.Name

    try {
        $targetElement = Get-UiaElementFromPoint
        if (-not $targetElement) { return $null }

        $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
        $current = $targetElement

        # Window にぶつかるまで親を遡り(最大20階層)、最もRPA操作に適した要素を評価
        $bestCandidate = $null
        $bestScore = -1
        $depth = 0

        while ($current -ne $null -and $depth -lt 20) {
            $name = $current.Current.Name
            $id   = $current.Current.AutomationId
            $type = $current.Current.ControlType.ProgrammaticName.Replace("ControlType.","")

            if ($current.Current.ControlType -eq [System.Windows.Automation.ControlType]::Window -or
                $current.Current.ControlType -eq [System.Windows.Automation.ControlType]::Document -or
                $current.Current.ControlType -eq [System.Windows.Automation.ControlType]::Pane -or
                $id -eq "__next") {
                break
            }

            $score = 0
            if (-not [string]::IsNullOrWhiteSpace($name)) { $score += 50 }
            if (-not [string]::IsNullOrWhiteSpace($id))   { $score += 40 }
            
            # 入力・操作系のコントロールタイプは優遇
            if ($type -in @("Edit","ComboBox","Button","Hyperlink","CheckBox","RadioButton")) {
                $score += 30
            }
            # 深すぎる階層は減点 (より直下の子要素を優先)
            $score -= ($depth * 2)

            if ($score -gt $bestScore) {
                $bestScore = $score
                $bestCandidate = $current
            }

            $current = $walker.GetParent($current)
            $depth++
        }

        if ($bestCandidate) {
            $targetElement = $bestCandidate
        }

        $targetWindow = Get-UiaWindowFromElement -TargetElement $targetElement
        $windowName = if ($targetWindow) { Format-UiaWindowName -RawWindowName $targetWindow.Current.Name } else { "" }

        return @{ 
            TargetElement = $targetElement; 
            TargetWindow  = $targetWindow; 
            WindowName    = $windowName 
        }
    } catch {
        throw (New-EngineException -Func $func -Type "GetTargetInfoError" -Message "UIA探索中にエラー" -Details $_.Exception.Message)
    }
}

# ------------------------------------------------------------------------------
# --- (内部専用) UIA/Web用の実行コマンド文字列を組み立てる ---
function Build-RpaStepCode {
    param(
        [string]$Action, 
        [string]$WindowName, 
        [string]$TargetName, 
        [string]$ControlType, 
        [int]$Index, 
        [string]$ActionValue, 
        [string]$AnchorName, 
        [string]$AnchorDirection, 
        [int]$AnchorSteps
    )

    $func = $MyInvocation.MyCommand.Name

    try {
        # 文字列内のダブルクォートをエスケープ (VBAやPSで安全に実行するため)
        $safeActionValue = $ActionValue -replace '"', '`"'
        $safeTargetName  = $TargetName -replace '"', '`"'
        $safeWindowName  = $WindowName -replace '"', '`"'
        $safeAnchorName  = $AnchorName -replace '"', '`"'

        if ($Action -eq "GetName") {
            return "`$extractedValue = Execute-UIRpaStep -TargetWindowName `"$safeWindowName`" -TargetName `"$safeTargetName`" -TargetType `"$ControlType`" -Index $Index -Action `"$Action`" -ActionValue `"$safeActionValue`" -AnchorName `"$safeAnchorName`" -AnchorDirection `"$AnchorDirection`" -AnchorSteps $AnchorSteps`r`nWrite-Host `"  => 取得した値: `$extractedValue`" -ForegroundColor Cyan"
        } else {
            return "Execute-UIRpaStep -TargetWindowName `"$safeWindowName`" -TargetName `"$safeTargetName`" -TargetType `"$ControlType`" -Index $Index -Action `"$Action`" -ActionValue `"$safeActionValue`" -AnchorName `"$safeAnchorName`" -AnchorDirection `"$AnchorDirection`" -AnchorSteps $AnchorSteps"
        }
    } catch {
        throw (New-EngineException -Func $func -Type "CodeBuildError" -Message "ステップコードの生成に失敗しました" -Details $_.Exception.Message)
    }
}

# ==============================================================================
# --- [テスト用] (仮想)TEST座標からの要素取得検証 ---
# ==============================================================================
function Test-WebPointReverse {
    param (
        [int]$ClientX,
        [int]$ClientY
    )

    $func = $MyInvocation.MyCommand.Name
    Write-DebugLog -Message "[$func] TEST座標: Client(X=$ClientX, Y=$ClientY)" -Level Info

    try {
        # 共通JSロジックの呼び出し
        $jsPoint = Get-CorePointReverseJs -ClientX $ClientX -ClientY $ClientY

        # 手動ピッカーと同じ安全な関数(Invoke-WebScript)を使用してJSを実行
        $resRaw = Invoke-WebScript -Js $jsPoint
        
        Write-DebugLog -Message "[$func] 解析結果: $resRaw" -Level Info
        
        return "[RESULT]$resRaw"

    } catch {
        throw (New-EngineException -Func $func -Type "テスト実行エラー" -Message "テスト用JSの評価中にエラーが発生しました" -Details $_.Exception.Message)
    }
}
