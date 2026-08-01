# ------------------------------------------------------------------------------
# Robust DOM Utilities（フェイルセーフ・安全クリック拡張）
# ------------------------------------------------------------------------------

$global:ENGINE_DEBUG_JS_UTILS = @'

// 内部トレーサー（ログ記録機構）
var domTracker = {
    logs: [],
    log: function(msg) { this.logs.push(msg); },
    clear: function() { this.logs = []; }
};

// CSSセレクタ用に特殊文字をエスケープする関数
function cssEscape(str) {
    if (typeof CSS !== 'undefined' && typeof CSS.escape === 'function') return CSS.escape(str);
    return String(str).replace(/([^\x20-\x7E]|[ !"#$%&'()*+,./:;<=>?@[\\\]^`{|}~])/g, function (ch) { return '\\' + ch; });
}

// 要素の可視性を厳密に判定する関数 (silent = true で成功時のログを抑制)
function isVisible(el, { checkOpacity = true, minOpacity = 0.01, silent = false } = {}) {
    if (!el || el.nodeType !== 1) {
        domTracker.log("isVisible: 失敗 - 要素が null または無効なノードです。");
        return false;
    }
    
    const style = window.getComputedStyle(el);
    if (!style) {
        domTracker.log("isVisible: 失敗 - ComputedStyle が取得できません。");
        return false;
    }
    
    if (style.display === 'none' || style.visibility === 'hidden' || style.visibility === 'collapse') {
        domTracker.log("isVisible: 失敗 - CSSによって非表示です (display: " + style.display + ", visibility: " + style.visibility + ")。");
        return false;
    }
    if (el.hasAttribute('hidden')) {
        domTracker.log("isVisible: 失敗 - 'hidden' 属性が設定されています。");
        return false;
    }
    if (el.getAttribute('aria-hidden') === 'true') {
        domTracker.log("isVisible: 失敗 - 'aria-hidden=true' 属性によりスクリーンリーダーから隠されています。");
        return false;
    }
    
    if (checkOpacity) {
        const opacity = parseFloat(style.opacity);
        if (!Number.isNaN(opacity) && opacity <= minOpacity) {
            domTracker.log("isVisible: 失敗 - 透明度が高すぎます (opacity: " + opacity + ")。");
            return false;
        }
    }
    
    const rects = el.getClientRects();
    if (!rects || rects.length === 0) {
        domTracker.log("isVisible: 失敗 - 要素がレイアウト領域を持っていません (client rects = 0)。");
        return false;
    }
    
    const rect = el.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) {
        domTracker.log("isVisible: 失敗 - 要素のサイズがゼロです (幅: " + rect.width + ", 高さ: " + rect.height + ")。");
        return false;
    }
    
    // silentフラグで非表示にできる
    if (!silent) {
        domTracker.log("isVisible: 要素は可視状態です (幅: " + Math.round(rect.width) + ", 高さ: " + Math.round(rect.height) + ")");
    }
    return true;
}

// XPathから複数のDOM要素をすべて取得する関数
function getElementsByXPath(xpath, root = document) {
    try {
        const result = document.evaluate(xpath, root, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
        const nodes = [];
        for (let i = 0; i < result.snapshotLength; i++) {
            const node = result.snapshotItem(i);
            if (node && node.nodeType === 1) nodes.push(node);
        }
        return nodes;
    } catch (e) {
        return [];
    }
}

// 指定された要素から、一意に特定できる強固なCSSセレクタを逆生成する関数
function generateCssSelector(el, root = document) {
    if (!el || el.nodeType !== 1) throw new Error('element must be an Element node.');
    if (el === document.documentElement) return 'html';
    if (el === document.body) return 'body';

    // 【優先度 1位】 id属性（存在すればページ内で最強の一意性を持つ）
    const id = el.getAttribute && el.getAttribute('id');
    if (id) {
        const sel = '#' + cssEscape(id);
        try { if ((root.querySelector(sel) === el)) return sel; } catch (e) {}
    }

    // 【優先度 2位】 意味を持つ特定属性（ボタンの役割やシステムが付与した名前など）
    const ATTR_CANDIDATES = ['name', 'type', 'title', 'placeholder', 'aria-label', 'role'];
    for (const attr of ATTR_CANDIDATES) {
        const val = el.getAttribute && el.getAttribute(attr);
        if (val) {
            const sel = '[' + attr + '="' + cssEscape(val) + '"]';
            // この属性だけで一意に特定できるか（1件だけヒットし、それが自身か）を確認
            try {
                const found = root.querySelectorAll(sel);
                if (found.length === 1 && found[0] === el) return sel;
            } catch (e) {}
        }
    }

    // 【優先度 3位】 カスタムデータ属性（モダンWebフレームワークがよく使う data-* 属性）
    for (const a of Array.from(el.attributes || []).filter(at => at.name.startsWith('data-'))) {
        const sel = '[' + a.name + '="' + cssEscape(a.value) + '"]';
        try {
            const found = root.querySelectorAll(sel);
            if (found.length === 1 && found[0] === el) return sel;
        } catch (e) {}
    }

    // 【優先度 4位】 DOMツリーの階層構造（タグ名やクラス、兄弟要素の順番などを組み合わせる）
    const parts = [];
    let node = el;
    const rootElement = (root === document) ? document.documentElement : root;
    
    // 要素から親へ向かって階層を遡りながらセレクタを組み立てる
    while (node && node.nodeType === 1 && node !== rootElement) {
        let part = node.tagName.toLowerCase();
        
        // クラス名があれば最大3つまで追加して絞り込みの精度を高める
        if (node.classList && node.classList.length > 0) {
            const classes = Array.from(node.classList).slice(0, 3).map(c => '.' + cssEscape(c)).join('');
            const withClass = part + classes;
            try {
                const found = root.querySelectorAll(withClass);
                // この時点で一意になれば、これ以上親を遡る必要はないので即採用
                if (found.length === 1 && found[0] === el) {
                    parts.unshift(withClass);
                    break;
                }
            } catch (e) {}
            if (classes.length < 100) part += classes;
        }
        
        // 兄弟要素の中での順番（nth-of-type）を計算する
        const parent = node.parentNode;
        if (!parent) { parts.unshift(part); break; }
        
        const tagName = node.tagName;
        let index = 0;
        for (let i = 0; i < parent.children.length; i++) {
            if (parent.children[i].tagName === tagName) index++;
            if (parent.children[i] === node) break;
        }
        // 同名のタグが複数あれば、何番目かを指定
        if (index > 0) part += ':nth-of-type(' + index + ')';
        
        parts.unshift(part);
        const candidate = parts.join(' > ');
        // 組み立てた階層セレクタで一意になったか確認
        try {
            const found = root.querySelectorAll(candidate);
            if (found.length === 1 && found[0] === el) return candidate;
        } catch (e) {}
        node = parent; // さらに親へ
    }
    
    // 【最終チェックと保険】
    const finalCandidate = parts.join(' > ');
    try {
        const found = root.querySelectorAll(finalCandidate);
        if (found.length === 1 && found[0] === el) return finalCandidate;
    } catch (e) {}
    
    // 上記のすべてでダメだった場合の最終手段（絶対パスによる強制生成）
    node = el;
    const fullParts = [];
    while (node && node.nodeType === 1) {
        let p = node.tagName.toLowerCase();
        const id = node.getAttribute && node.getAttribute('id');
        if (id) {
            p = '#' + cssEscape(id);
            fullParts.unshift(p);
            break;
        }
        if (node.classList && node.classList.length) {
            p += Array.from(node.classList).map(c => '.' + cssEscape(c)).join('');
        } else {
            const parent = node.parentNode;
            if (parent) {
                const index = Array.prototype.indexOf.call(parent.children, node) + 1;
                p += ':nth-child(' + index + ')';
            }
        }
        fullParts.unshift(p);
        node = node.parentNode;
    }
    return fullParts.join(' > ');
}

// 要素を画面内にスクロールし、可視状態を確認した上で安全にクリックする関数
function safeClick(el, { visible = true, tryScrollIntoView = true } = {}) {
    if (!el) throw new Error('要素が null または未定義です。');
    
    //● domTracker.log('safeClick: 処理を開始します...');
    try {
        if (visible) {
            domTracker.log('safeClick: スクロール前の可視性チェックを実行します。');
            // 【修正】safeClick内部からのisVisibleチェックは、ログを抑制(silent: true)する
            if (!isVisible(el, { silent: true })) throw new Error('スクロール前の時点で要素が非表示と判定されました。');
        }
        
        if (tryScrollIntoView) {
            domTracker.log('safeClick: 画面中央へのスクロール(scrollIntoView)を試行します。');
            try { 
                el.scrollIntoView({ block: 'center', inline: 'center', behavior: 'auto' }); 
            } catch (e) {
                domTracker.log('safeClick: 警告 - スクロールに失敗しました (' + e.message + ')。');
            }
        }
        
        if (visible) {
            domTracker.log('safeClick: スクロール後の最終可視性チェックを実行します。');
            if (!isVisible(el, { silent: true })) throw new Error('スクロール後に要素が他の要素の裏に隠れたか、非表示になりました。');
        }
        
        domTracker.log('safeClick: 対象要素へのネイティブクリックを実行します！');
        el.click();
        //● domTracker.log('safeClick: クリック処理が正常に完了しました。');
    } catch (err) {
        let selectorHint = null;
        try { selectorHint = generateCssSelector(el); } catch (e) { selectorHint = '(セレクタの生成不可)'; }
        domTracker.log('safeClick: エラー - ' + err.message);
        throw new Error('安全なクリックに失敗しました。理由: ' + err.message + ' / セレクタ候補: ' + selectorHint);
    }
}

window.DOMUtils = {
    tracker: domTracker,
    cssEscape: cssEscape,
    isVisible: isVisible,
    getElementsByXPath: getElementsByXPath,
    generateCssSelector: generateCssSelector,
    safeClick: safeClick
};
'@

# --- XPathから環境に強いCSSセレクタを動的生成（複数解析対応） ---
function Get-WebCssSelectorHint {
    param ([Parameter(Mandatory=$true)][string]$XPath)
    
    $func = $MyInvocation.MyCommand.Name
    $xpathEscaped = $XPath.Replace("'", "\'")

    # --- XPathから検索ワードを抽出する処理 ---
    $searchWord = $XPath
    if ($XPath -match "contains\(\.,\s*['`"]([^'`"]+)['`"]\)") {
        $searchWord = $matches[1]
    }

    $js = $global:ENGINE_DEBUG_JS_UTILS + @"
        try {
            var els = window.DOMUtils.getElementsByXPath('$xpathEscaped', document);
            if (!els || els.length === 0) return JSON.stringify({ status: 'NOT_FOUND' });

            var results = [];
            var bestSelector = null;

            for (var i = 0; i < els.length; i++) {
                var sel = '(生成不可)';
                try { sel = window.DOMUtils.generateCssSelector(els[i], document); } catch(e){}
                // 判定のみ行うため、探索時のisVisibleはログを出さない
                var isVis = window.DOMUtils.isVisible(els[i], { silent: true });

                results.push({ index: i + 1, selector: sel, visible: isVis });

                if (!bestSelector && isVis) {
                    bestSelector = sel;
                }
            }
            
            if (!bestSelector) bestSelector = results[0].selector;

            return JSON.stringify({ status: 'SUCCESS', count: els.length, data: results, bestSelector: bestSelector });
        } catch(e) {
            return JSON.stringify({ status: 'ERROR', message: e.message });
        }
"@

    $resJson = Invoke-WebScript -Js $js
    $resObj = $resJson | ConvertFrom-Json
    
    if ($resObj.status -eq 'NOT_FOUND') { throw (New-EngineException -Func $func -Type "未発見" -Message "XPathに該当する要素が見つかりません" -Details $XPath) }
    if ($resObj.status -eq 'ERROR') { throw (New-EngineException -Func $func -Type "JSエラー" -Message "セレクタの生成に失敗しました" -Details $resObj.message) }
    
    Write-DebugLog -Message "[Robust DOM] XPath検索: [$searchWord] に合致する要素を $($resObj.count) 件発見しました。" -Level Info
    foreach ($item in $resObj.data) {
        $visMark = if ($item.visible) { "[可視〇]" } else { "[非表示]" }
        Write-DebugLog -Message "  -> 候補 $($item.index): $visMark $($item.selector)" -Level Info
    }
    Write-DebugLog -Message "[Robust DOM] 最適セレクタ: $($resObj.bestSelector)" -Level Info
    
    return $resObj.bestSelector
}

# --- 対象要素の可視化を待機し、確実にスクロール＆クリック ---
function Invoke-WebSafeClick {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )
    
    $func = $MyInvocation.MyCommand.Name
    $selectorEscaped = $Selector.Replace("'", "\'")

    $state = @{ LastLogState = "" }

    $condition = {
        $js = $global:ENGINE_DEBUG_JS_UTILS + @"
            try {
                window.DOMUtils.tracker.clear();
                var els = document.querySelectorAll('$selectorEscaped');
                window.DOMUtils.tracker.log('探索開始: セレクタに合致する要素 ' + els.length + ' 件');
                
                if (els.length === 0) {
                    window.DOMUtils.tracker.log('DOM内に該当要素が存在しません。');
                } else {
                    for (var i = 0; i < els.length; i++) {
                        window.DOMUtils.tracker.log('--- [候補 ' + (i+1) + '/' + els.length + '] の検証を開始 ---');
                        // ここは検証スタート地点なので、isVisibleの成功ログ(幅・高さ)を1回だけ出させる
                        if (window.DOMUtils.isVisible(els[i])) {
                            window.DOMUtils.safeClick(els[i]);
                            return JSON.stringify({ status: 'SUCCESS', logs: window.DOMUtils.tracker.logs });
                        } else {
                            window.DOMUtils.tracker.log('[候補 ' + (i+1) + '] は条件を満たさないためスキップします。');
                        }
                    }
                }
                return JSON.stringify({ status: 'WAIT', logs: window.DOMUtils.tracker.logs });
            } catch(e) {
                return JSON.stringify({ status: 'ERROR', message: e.message, logs: window.DOMUtils.tracker.logs });
            }
"@
        $resJson = Invoke-WebScript -Js $js
        $resObj = $resJson | ConvertFrom-Json
        
        if ($null -ne $resObj) {
            if ($global:IsDebugMode -and $null -ne $resObj.logs) {
                $currentLogState = $resObj.logs -join "|"
                if ($currentLogState -ne $state.LastLogState) {
                    foreach ($logStr in $resObj.logs) {
                        Write-DebugLog -Message "[Robust DOM] $logStr" -Level Info
                    }
                    $state.LastLogState = $currentLogState
                }
            }
            
            if ($resObj.status -eq 'SUCCESS') { return $true }
            if ($resObj.status -eq 'ERROR') {
                throw (New-EngineException -Func $func -Type "JSエラー" -Message "安全なクリック処理に失敗しました" -Details $resObj.message)
            }
        }
        
        return $false
    }

    $errMsg = "[$func] タイムアウト: 要素の出現・可視化、またはクリックに失敗しました ($Selector)"
    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null
    
    return "SafeClick Completed: $Selector"
}
