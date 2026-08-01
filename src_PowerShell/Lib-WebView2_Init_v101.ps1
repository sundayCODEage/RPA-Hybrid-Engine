# ------------------------------------------------------------------------------
# WebView2 初期化モジュール
# RPA用ブラウザウィンドウの生成とWebView2エンジンの起動
# ------------------------------------------------------------------------------

# ブラウザシステムの初期化開始
Write-DebugLog -Message "[System] 開始: ブラウザシステムの初期化 ..." -Level Info
$libDir = Join-Path $global:ScriptDirectory "Libs"

# 必要なWebView2拡張DLLの動的読み込み
$requiredDlls = @(
    "Microsoft.Web.WebView2.Core.dll",
    "Microsoft.Web.WebView2.WinForms.dll"
)

foreach ($dllName in $requiredDlls) {
    $dllPath = Join-Path $libDir $dllName
    if (Test-Path $dllPath) {
        try {
            $versionInfo = (Get-Item $dllPath).VersionInfo.FileVersion
            Add-Type -Path $dllPath
            Write-DebugLog -Message "[System] 成功: DLLロード ($dllName Version: $versionInfo)" -Level Success
        } catch {
            Write-DebugLog -Message "[System] 致命的エラー: DLLロード失敗 ($dllName - $($_.Exception.Message))" -Level Fatal
            exit 1
        }
    } else {
        Write-DebugLog -Message "[System] 致命的エラー: DLLが存在しません ($dllName)" -Level Fatal
        exit 1
    }
}

# デスクトップUI関連アセンブリのロード
Add-Type -AssemblyName System.Windows.Forms, System.Drawing, UIAutomationClient, UIAutomationTypes

# RPAブラウザ用ウィンドウの生成
$global:BrowserForm = New-Object System.Windows.Forms.Form -Property @{
    Text   = "RPA Browser - $global:AppSessionId"
    Width  = $global:CONFIG.BrowserWidth
    Height = $global:CONFIG.BrowserHeight
    StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
#● WindowState   = [System.Windows.Forms.FormWindowState]::Maximized
}

# WebView2コントロールの配置
$global:WebViewCtrl = New-Object Microsoft.Web.WebView2.WinForms.WebView2
$global:WebViewCtrl.Dock = [System.Windows.Forms.DockStyle]::Fill
$global:BrowserForm.Controls.Add($global:WebViewCtrl)
$global:BrowserForm.Show()

try {
    # ユーザーデータフォルダの準備
    $userDataFolder = Join-Path $global:ScriptDirectory "UserData"
    if (-not (Test-Path $userDataFolder)) {
        New-Item -ItemType Directory -Path $userDataFolder | Out-Null
    }

    # WebView2起動オプションの設定
    $options = [Microsoft.Web.WebView2.Core.CoreWebView2EnvironmentOptions]::new()

    # ディスクキャッシュを禁止・制限する共通引数
    $cacheArgs = " --disable-disk-cache --disk-cache-size=1 --media-cache-size=1"

    if ($global:CdpPort -gt 0) {
        $options.AdditionalBrowserArguments = "--disable-gpu --disable-dev-shm-usage --remote-debugging-port=$global:CdpPort" + $cacheArgs
        Write-DebugLog -Message "[System] 起動モード: CDP有効 (Port: $global:CdpPort)" -Level Info
    } else {
        $options.AdditionalBrowserArguments = "--disable-gpu --disable-dev-shm-usage" + $cacheArgs
        Write-DebugLog -Message "[System] 起動モード: 標準 (CDP無効)" -Level Info
    }

    # WebView2エンジンの非同期起動
    $envTask = [Microsoft.Web.WebView2.Core.CoreWebView2Environment]::CreateAsync(
        $null, $userDataFolder, $options
    )
    while (-not $envTask.IsCompleted) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 10
    }

    # WebView2コントロールの初期化
    $initTask = $global:WebViewCtrl.EnsureCoreWebView2Async($envTask.Result)
    while (-not $initTask.IsCompleted) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 10
    }

    Write-DebugLog -Message "[System] 成功: WebView2エンジン初期化完了" -Level Success

    # 初回タブ(P01)の登録
    $global:Tabs["P01"] = @{
        WebView = $global:WebViewCtrl
        Parent  = $null
        Url     = $global:WebViewCtrl.Source.ToString()
    }
    $global:ActiveTabId = "P01"

    # 初回タブ(P01)のURL更新イベント登録
    $global:WebViewCtrl.CoreWebView2.add_NavigationCompleted({
        param($sender, $e)
        try {
            $global:Tabs["P01"].Url = $sender.Source
            Write-DebugLog -Message "[P01] URL更新: $($sender.Source)" -Level Info
        } catch {}
    })

    # ランタイムバージョンのログ出力
    $runtimeVersion = $global:WebViewCtrl.CoreWebView2.Environment.BrowserVersionString
    Write-DebugLog -Message "[System] 情報: 接続先ランタイム (Version: $runtimeVersion)" -Level Info

    # 新規ウィンドウ要求の仮想タブ捕獲ハンドラ定義
    $global:NewWindowHandler = {
        param($sender, $e)
        Write-DebugLog -Message "[WebView2] 情報: 新規ウィンドウ要求を検知。仮想タブとして捕獲します。" -Level Info

        try {
            $deferral = $e.GetDeferral()

            $newWebView = New-Object Microsoft.Web.WebView2.WinForms.WebView2
            $newWebView.Dock = [System.Windows.Forms.DockStyle]::Fill

            $newWebView.Tag = @{
                Parent   = $global:WebViewCtrl
                Deferral = $deferral
                EventArg = $e
            }

            $dummyHandle = $newWebView.Handle

            $newWebView.add_CoreWebView2InitializationCompleted({
                param($senderInit, $argsInit)

                $state      = $senderInit.Tag
                $parentCtrl = $state.Parent
                $deferral   = $state.Deferral
                $evtArgs    = $state.EventArg

                try {
                    if ($argsInit.IsSuccess) {
                        $tabId = [guid]::NewGuid().ToString()
                        $global:Tabs[$tabId] = @{
                            WebView = $senderInit
                            Parent  = $parentCtrl
                        }
                        $global:ActiveTabId = $tabId

                        try {
                            $global:Tabs[$tabId].TargetId = $evtArgs.WindowFeatures.TargetId
                        } catch {
                            Write-DebugLog -Message "[WebView2] 警告: targetId の取得に失敗" -Level Warning
                        }

                        $global:BrowserForm.Controls.Add($senderInit)
                        $senderInit.BringToFront()

                        $evtArgs.NewWindow = $senderInit.CoreWebView2
                        $evtArgs.Handled   = $true
                        $deferral.Complete()

                        $global:WebViewCtrl = $senderInit

                        $senderInit.CoreWebView2.add_NavigationCompleted({
                            param($navSender, $navArgs)
                            try {
                                $global:Tabs[$tabId].Url = $navSender.Source
                                Write-DebugLog -Message "[$tabId] URL更新: $($navSender.Source)" -Level Info
                            } catch {}
                        })

                        $senderInit.CoreWebView2.add_NewWindowRequested($global:NewWindowHandler)
                        $senderInit.CoreWebView2.add_WindowCloseRequested({
                            param($closeSender, $closeArgs)
                            
                            $mi = [System.Windows.Forms.MethodInvoker]{
                                try {
                                    Write-DebugLog -Message "[WebView2] 情報: ポップアップの終了要求(window.close)を検知。タブ($tabId)を破棄します。" -Level Info

                                    # 透明なゴーストとして残るのを防ぐため、コントロールを削除してメモリ解放
                                    $global:BrowserForm.Controls.Remove($senderInit)
                                    $senderInit.Dispose()
                                    
                                    # 管理リストからの完全削除
                                    if ($global:Tabs.ContainsKey($tabId)) {
                                        $global:Tabs.Remove($tabId)
                                    }

                                    # アクティブタブが閉じられたら、自動的に親画面へ戻る
                                    if ($global:ActiveTabId -eq $tabId) {
                                        # 1. まず、このポップアップを開いた親ウィンドウのTabIdを探す
                                        $nextTabId = $null
                                        foreach ($key in $global:Tabs.Keys) {
                                            if ($global:Tabs[$key].WebView -eq $parentCtrl) {
                                                $nextTabId = $key
                                                break
                                            }
                                        }

                                        # 2. 親が見つかった場合はそこへ確実に戻る
                                        if ($null -ne $nextTabId) {
                                            $global:ActiveTabId = $nextTabId
                                            $global:WebViewCtrl = $parentCtrl
                                            Write-DebugLog -Message "[WebView2] 情報: ゴーストを削除し、親タブ ($nextTabId) へ自動復帰しました。" -Level Info
                                        } else {
                                            # 3. 万が一親が見つからない場合のフォールバック（ご提案のロジックを安全化）
                                            $remainingTabs = @($global:Tabs.Keys)
                                            if ($remainingTabs.Count -gt 0) {
                                                # 順序不定のため警告を出した上で移行
                                                $global:ActiveTabId = $remainingTabs[0]
                                                $global:WebViewCtrl = $global:Tabs[$global:ActiveTabId].WebView
                                                Write-DebugLog -Message "[WebView2] 警告: 親タブ不明のため、残存タブ ($($global:ActiveTabId)) へ強制移行しました。" -Level Warning
                                            } else {
                                                $global:ActiveTabId = $null
                                                $global:WebViewCtrl = $null
                                                Write-DebugLog -Message "[WebView2] 情報: 全タブが閉じられたため、認識をリセットしました。" -Level Info
                                            }
                                        }
                                    }
                                } catch {
                                    Write-DebugLog -Message "[WebView2] 警告: ゴーストタブの破棄中にエラー ($($_.Exception.Message))" -Level Warning
                                }
                            }
                            try { $global:BrowserForm.BeginInvoke($mi) | Out-Null } catch {}
                        }.GetNewClosure())

                    } else {
                        $exMsg = if ($argsInit.InitializationException) {
                            $argsInit.InitializationException.Message
                        } else {
                            "理由不明"
                        }
                        Write-DebugLog -Message "[WebView2] エラー: サブ画面の初期化に失敗 ($exMsg)" -Level Error
                        $deferral.Complete()
                    }
                } catch {
                    Write-DebugLog -Message "[WebView2] エラー: 初期化完了イベント内で例外 ($($_.Exception.Message))" -Level Error
                    try { $deferral.Complete() } catch {}
                }
            })

            $newWebView.EnsureCoreWebView2Async($sender.Environment) | Out-Null
        } catch {
            Write-DebugLog -Message "[WebView2] エラー: 新規ウィンドウ捕獲失敗 ($($_.Exception.Message))" -Level Error
        }
    }

    $global:WebViewCtrl.CoreWebView2.add_NewWindowRequested($global:NewWindowHandler)

} catch {
    Write-DebugLog -Message "[System] 致命的エラー: WebView2 初期化失敗 ($($_.Exception.Message))" -Level Fatal
    exit 1
}

# ウィンドウの最前面表示およびフォーカス確保
try {
    $hWnd = $global:BrowserForm.Handle
    [Win32Api.Win32Utils]::ShowWindow($hWnd, 9) | Out-Null
    [Win32Api.Win32Utils]::SetForegroundWindow($hWnd) | Out-Null

    $global:BrowserForm.Activate()
    $global:WebViewCtrl.Focus()

    $global:BrowserForm.TopMost = $true
    Start-Sleep -Milliseconds 80
    $global:BrowserForm.TopMost = $false

    try {
        $global:WebViewCtrl.CoreWebView2.ExecuteScriptAsync(
            "window.focus(); if(document.activeElement) { document.activeElement.focus(); }"
        ) | Out-Null
    } catch {}

    $global:WebViewCtrl.CoreWebView2.add_NavigationCompleted({
        param ($sender, $e)

        $mi = [System.Windows.Forms.MethodInvoker]{
            try {
                [Win32Api.Win32Utils]::SetForegroundWindow($global:BrowserForm.Handle) | Out-Null
                $global:BrowserForm.Activate()
                $global:WebViewCtrl.Focus()

                try {
                    $global:WebViewCtrl.CoreWebView2.ExecuteScriptAsync(
                        "window.focus(); if(document.activeElement) { document.activeElement.focus(); }"
                    ) | Out-Null
                } catch {}
            } catch {}
        }

        try { $global:BrowserForm.BeginInvoke($mi) | Out-Null } catch {}
    })

    $global:WebViewCtrl.CoreWebView2.add_ProcessFailed({
        param ($sender, $e)

        try {
            $kind = $e.ProcessFailedKind
            $reason = ""
            try { $reason = $e.Reason } catch {}

            $errMsg = "ブラウザプロセスのクラッシュ検知 (Kind: $kind, Reason: $reason)"
            Write-DebugLog -Message "[System] 致命的エラー: $errMsg" -Level Fatal

            try { [Console]::WriteLine("[ERROR] $errMsg") } catch {}

            $mi = [System.Windows.Forms.MethodInvoker]{
                try { $global:BrowserForm.Close() } catch {}
            }
            try { $global:BrowserForm.BeginInvoke($mi) | Out-Null } catch {}
        } catch {}
    })

} catch {
    Write-DebugLog -Message "[System] 警告: フォーカス確保処理で例外発生 ($($_.Exception.Message))" -Level Warning
}

# --- WebView2プロファイルのキャッシュ・Cookie・DOMストレージの削除 ---
function Clear-WebCache {
    param ([string]$Mode = "CacheOnly")

    $func = $MyInvocation.MyCommand.Name

    try {
        $webview = Get-ActiveWebView
        if ($null -eq $webview -or $null -eq $webview.CoreWebView2) {
            throw (New-EngineException -Func $func -Type "初期化エラー" -Message "WebView2が初期化されていないか、アクティブタブが存在しません")
        }

        $kinds = [Microsoft.Web.WebView2.Core.CoreWebView2BrowsingDataKinds]::Cache -bor
                 [Microsoft.Web.WebView2.Core.CoreWebView2BrowsingDataKinds]::DiskCache -bor
                 [Microsoft.Web.WebView2.Core.CoreWebView2BrowsingDataKinds]::MemoryCaches

        if ($Mode -eq "All") {
            $kinds = $kinds -bor
                     [Microsoft.Web.WebView2.Core.CoreWebView2BrowsingDataKinds]::Cookies -bor
                     [Microsoft.Web.WebView2.Core.CoreWebView2BrowsingDataKinds]::LocalStorage -bor
                     [Microsoft.Web.WebView2.Core.CoreWebView2BrowsingDataKinds]::AllDomStorage
        }

        $task = $webview.CoreWebView2.Profile.ClearBrowsingDataAsync($kinds)

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $timeoutSec = 15

        while (-not $task.IsCompleted) {
            if ($sw.Elapsed.TotalSeconds -gt $timeoutSec) {
                throw (New-EngineException -Func $func -Type "Timeout" -Message "キャッシュクリア処理がタイムアウトしました" -Details "${timeoutSec}秒経過")
            }
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 20
        }

        return "SUCCESS"

    } catch {
        $errMsg = $_.Exception.Message -replace "`r`n", " " -replace "`n", " "
        throw (New-EngineException -Func $func -Type "内部エラー" -Message "キャッシュクリア処理中に予期せぬエラーが発生しました" -Details $errMsg)
    }
}
