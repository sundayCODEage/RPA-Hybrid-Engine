# ------------------------------------------------------------------------------
# WebView2 ネイティブ通信モジュール
# WebView2標準APIによるJavaScript実行（CDP非依存）
# ------------------------------------------------------------------------------

# --- WebView2標準機能を利用したJavaScript非同期実行と結果取得 ---
function Invoke-WebView2NativeScript {
    param (
        [Parameter(Mandatory = $true)][string]$Js,
        [int]$Retries = 3,
        [int]$InitialDelayMs = 200
    )

    $func = $MyInvocation.MyCommand.Name
    $jsCode = $Js

    # JSの安全なラップおよびJSON形式での返却
    $wrappedJs = @"
(function() {
    try {
        const result = (function() { $jsCode })();
        // undefinedをnullに置換してJSON消失を防止
        return JSON.stringify({
            status: "success",
            data: result !== undefined ? result : null
        });
    } catch (e) {
        return JSON.stringify({
            status: "error",
            message: e && e.message ? e.message : String(e),
            stack: e && e.stack ? e.stack : "",
            name: e && e.name ? e.name : ""
        });
    }
})();
"@

    $attempt = 0
    $delayMs = $InitialDelayMs

    while ($true) {
        $attempt++

        try {
            # アクティブタブのWebView取得
            $webview = Get-ActiveWebView
            if (-not $webview) {
                throw (New-EngineException -Func $func -Type "初期化エラー" -Message "アクティブなWebViewの取得に失敗しました（タブ未初期化）")
            }

            # JSの非同期実行
            $task = $webview.CoreWebView2.ExecuteScriptAsync($wrappedJs)

            while (-not $task.IsCompleted) {
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 10
            }

            if ($task.IsFaulted) {
                throw (New-EngineException -Func $func -Type "ネイティブエラー" -Message "WebView2ネイティブAPI(ExecuteScriptAsync)の実行に失敗しました" -Details $task.Exception.InnerException.Message)
            }

            $rawResult = $task.Result

            # undefinedの扱い
            if ($rawResult -eq '"undefined"' -or $rawResult -eq 'undefined') {
                return $null
            }

            # 空またはnullの場合のリトライ処理
            if ([string]::IsNullOrWhiteSpace($rawResult) -or $rawResult -eq "null") {
                if ($attempt -lt $Retries) {
                    Start-Sleep -Milliseconds $delayMs
                    $delayMs = [Math]::Min(2000, $delayMs * 2)
                    continue
                } else {
                    throw (New-EngineException -Func $func -Type "内部エラー" -Message "WebView2ネイティブAPIの戻り値が空またはnullです")
                }
            }

            # 文字列JSONのアンエスケープ
            if ($rawResult -match '^\s*"(.*)"\s*$') {
                $inner = $Matches[1]
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
                if ($attempt -lt $Retries) {
                    Start-Sleep -Milliseconds $delayMs
                    $delayMs = [Math]::Min(2000, $delayMs * 2)
                    continue
                } else {
                    throw (New-EngineException -Func $func -Type "内部エラー" -Message "WebView2からの応答(JSON)のデコードに失敗しました")
                }
            }

            # JS側のエラー処理
            if ($response.status -ne "success") {
                $stack = $response.stack -replace "`r`n", " | "
                throw (New-EngineException -Func $func -Type "JSエラー" -Message "$($response.name): $($response.message)" -Details $stack)
            }

            return $response.data

        } catch {
            if ($attempt -lt $Retries) {
                Start-Sleep -Milliseconds $delayMs
                $delayMs = [Math]::Min(2000, $delayMs * 2)
                continue
            } else {
                $errMsg = (New-EngineException -Func $func -Type "ネイティブエラー" -Message "スクリプトの実行処理中に予期せぬエラーが発生しました" -Details $_.Exception.Message)
                throw $errMsg
            }
        }
    }
}
