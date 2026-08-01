Attribute VB_Name = "Mod_Chapter1_Basics"
Option Explicit

' ==============================================================================
' 【T1】基礎・汎用コンポーネント動作テスト (マスターシナリオ)
' .. フォーム操作 / 非同期待機 / テーブル抽出 / サイレントダウンロードの検証
' ==============================================================================
Public Sub Test_Chapter1_Master()
    On Error GoTo ErrorHandler
    
    Dim sandboxPath As String
    Dim targetUrl As String
    Dim resText As String
    
    ' タイマー用変数
    Dim startTime As Currency
    Dim endTime As Currency
    Dim elapsedTimeMs As Double
    
    ' 実行時のエンジン動作設定 (要素ハイライト機能の制御: True=有効)
    rpaEngine.RunAction "Set-EngineConfig", CreateParams("EnableHighlight", True)
    
    ' サンドボックス（ローカルHTML）の絶対パスを取得 (Windowsの \ を / に変換)
    sandboxPath = Replace(ThisWorkbook.Path & "\sandbox\", "\", "/")
    
    Debug.Print "=========================================================="
    Debug.Print "--- [T1] 基礎・汎用コンポーネント動作テスト ---"
    Debug.Print "=========================================================="

    ' ==========================================================================
    ' [テスト 1/4] 01_basic_form.html (同じ画面要素で XPath ? CSSセレクタ 同等操作の比較)
    ' ==========================================================================
    Debug.Print "--- [1/4] 01_basic_form.html テスト開始 ---"
    targetUrl = "file:///" & sandboxPath & "01_basic_form.html"
    
    ' ページ遷移 ＆ DOM構築完了の待機
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    
    ' --------------------------------------------------------------------------
    ' URLやタイトルに指定文字列が含まれるまで待機 （※ 正規の文字列の一部指定可）
    rpaEngine.RunAction "Wait-WebUrlContains", CreateParams("Substring", "basic_form")
    rpaEngine.RunAction "Wait-WebTitleContains", CreateParams("Substring", "標準フォーム操作")
    
    ' 現在のURL、およびページタイトルを取得
    resText = rpaEngine.RunAction("Get-WebUrl")
    Debug.Print "  -> [URL]   : " & resText
    resText = rpaEngine.RunAction("Get-WebTitle")
    Debug.Print "  -> [Title] : " & resText

    ' --------------------------------------------------------------------------
    ' 【パターンA】 XPath 指定による一連の操作
    ' --------------------------------------------------------------------------
    Debug.Print "  --- [A] XPathによる操作を実行 ---"
    ' 1. テキスト入力
    rpaEngine.RunAction "Set-WebXPathTextInput", CreateParams("XPath", "//input[@id='txtInput']", "Value", "GEMINI 太郎")
    ' 2. ドロップダウン選択 (財政課: D02) ※CSSで代用
    rpaEngine.RunAction "Select-WebDropdown", CreateParams("Selector", "#selDept", "Value", "D02")
    ' 3. チェックボックス ON
    rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", "//input[@id='chkAgree']")
    ' 4. 反映ボタン クリック
    rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", "//button[@id='btnApply']")
    ' 5. 表示結果の読み取り
    resText = rpaEngine.RunAction("Get-WebXPathText", CreateParams("XPath", "//span[@id='lblDisplay']"))
    Debug.Print "     > [XPath結果] 画面反映テキスト: " & resText

    Sleep 2000 ' 目視確認用のウェイト
   
    ' --------------------------------------------------------------------------
    ' 【パターンB】 CSS セレクタ 指定による「全く同じ要素」の操作（上書き実行）
    ' --------------------------------------------------------------------------
    Debug.Print "  --- [B] CSSセレクタによる操作（全く同じ要素）を実行 ---"
    ' 1. テキスト待機＆上書き入力 (CSS: #txtInput)
    rpaEngine.RunAction "Wait-WebElement", CreateParams("Selector", "#txtInput")
    rpaEngine.RunAction "Set-WebTextInput", CreateParams("Selector", "#txtInput", "Value", "GEMINI 花子")
    ' 2. ドロップダウン変更 (総務課: D01) (CSS: #selDept)
    rpaEngine.RunAction "Select-WebDropdown", CreateParams("Selector", "#selDept", "Value", "D01")
    ' 3. チェックボックス OFF (CSS: #chkAgree)
    rpaEngine.RunAction "Invoke-WebClick", CreateParams("Selector", "#chkAgree")
    ' 4. 反映ボタン クリック (CSS: #btnApply)
    rpaEngine.RunAction "Invoke-WebClick", CreateParams("Selector", "#btnApply")
    ' 5. 表示結果の読み取り (CSS: #lblDisplay)
    resText = rpaEngine.RunAction("Get-WebText", CreateParams("Selector", "#lblDisplay"))
    Debug.Print "     > [CSS結果]   画面反映テキスト: " & resText

    Sleep 1000 ' 目視確認用のウェイト
    
    ' ==========================================================================
    ' 【パターンC】 JSネイティブ実行およびCDPネイティブ操作
    ' ==========================================================================
    Debug.Print "  --- [C] JSネイティブ実行およびCDPネイティブ操作 ---"
    
    ' ※ useCdpPort のスコープ外であることと、CDP無効起動時を考慮し、
    ' On Error Resume Next によるトライ＆キャッチ方式で安全に実行します。
    On Error Resume Next
    
    ' 1. CDPコマンド (ブラウザバージョンの取得)
    ' PowerShellのオブジェクトがそのまま文字列化されるため、先頭部分のみを切り出して表示します
    resText = rpaEngine.RunAction("Invoke-CdpCommand", CreateParams("Method", "Browser.getVersion"))
    If Err.Number <> 0 Then
        Debug.Print "    > Invoke-CdpCommand エラー(CDP無効のためスキップ): " & Err.Description
        Err.Clear
    Else
        Debug.Print "    > Invoke-CdpCommand / ブラウザバージョン: " & Left(Replace(resText, vbCrLf, " "), 60) & "..."

        ' 2. JSネイティブ実行 (UserAgent)
        resText = rpaEngine.RunAction("Invoke-WebView2NativeScript", CreateParams("Js", "return navigator.userAgent;"))
        Debug.Print "    > NativeScript / UserAgent: " & Left(resText, 60) & "..."
    
        ' 3. CDP経由のJS実行 (Title)
        resText = rpaEngine.RunAction("Invoke-CdpScript", CreateParams("Js", "return document.title;"))
        Debug.Print "    > Invoke-CdpScript / title: " & resText
        
        ' 4. CDPネイティブ入力・クリック
        rpaEngine.RunAction "Set-CdpNativeTextInput", CreateParams("Selector", "#txtInput", "Value", "GEMINI 三郎 (CDP)")
        
        ' ※ lblDisplay ではなく、値を反映させるために btnApply をクリックするよう修正
        rpaEngine.RunAction "Invoke-CdpNativeClick", CreateParams("Selector", "#btnApply")
        
        ' 5. 結果表示の処理（Get-WebTextを再利用して表示結果を取得）
        resText = rpaEngine.RunAction("Get-WebText", CreateParams("Selector", "#lblDisplay"))
        Debug.Print "     > [CDP結果]   画面反映テキスト: " & resText
    End If
    
    ' エラーハンドリングを標準に戻す
    On Error GoTo ErrorHandler
    
    Sleep 1500 ' 目視確認用のウェイト

    ' ==========================================================================
    ' [テスト 2/4] 02_async_load.html (非同期待機・Invisible消滅待機・Load比較)
    ' ==========================================================================
    Debug.Print "--- [2/4] 02_async_load.html テスト開始 ---"
    targetUrl = "file:///" & sandboxPath & "02_async_load.html"

    ' ==========================================================================
    ' 【ページロード待機命令の使い分けガイド】
    '
    ' 1. Wait-WebDocumentReady (高速・軽量待機)
    '    - 判定: HTMLのDOM構造の解析完了のみ (document.readyState = 'complete')
    '    - 特徴: 画像や外部リソースのロード完了を待たず即復帰 (ミリ秒単位)
    '    - 用途: DOM要素さえあれば操作できる画面、背景の重い画像や無限通信で
    '            「Wait-WebPageLoad」がタイムアウトする画面 ("無限ロードの罠" 回避用)
    '
    ' 2. Wait-WebPageLoad (標準・完全待機)
    '    - 判定: 全リソース (画像/CSS/JS/iframe) 読み込み完了 ＋ Ajax通信の沈静化
    '    - 特徴: 画面描画と裏側の通信が完全に落ち着くまで確実・安全に待機
    '    - 用途: 通常の画面遷移時の【原則デフォルト】。描画遅延による誤動作防止
    '
    ' ★ 運用ルール: 基本は「Wait-WebPageLoad」を使い、
    '                裏で通信が走り続けて止まる特殊画面のみ「Wait-WebDocumentReady」へ切り替える。
    Debug.Print "★ コード内の【ページロード待機命令の使い分けガイド】を参照！"
    ' ==========================================================================

    ' ページ遷移
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    
    Debug.Print "比較1: Wait-WebDocumentReady (DOMツリー構築完了のみを高速待機)"
    QueryPerformanceCounter startTime
    rpaEngine.RunAction "Wait-WebDocumentReady"
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print "  -> [速度検証] Wait-WebDocumentReady 所要時間 : " & Format(elapsedTimeMs, "0.00") & " ms (DOM解析完了)"
    
    ' --------------------------------------------------------------------------
    ' ページ遷移
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    
    Debug.Print "比較2: Wait-WebPageLoad (画像・全リソース・通信沈静化を含めて完全待機)"
    QueryPerformanceCounter startTime
    rpaEngine.RunAction "Wait-WebPageLoad"
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print "  -> [速度検証] Wait-WebPageLoad      所要時間 : " & Format(elapsedTimeMs, "0.00") & " ms (全リソースロード完了)"
    
    ' --------------------------------------------------------------------------
    ' --- 非同期待機処理 ---
    ' 3秒後に動的出現する処理を開始
    rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", "//button[@id='btnStartAsync']")
    
    ' 1. ローディング表示(#loadingSpinner)が消える(Invisible/Disappear)まで待機
    Debug.Print "     > ローディング表示の消滅(Invisible)を監視中..."
    rpaEngine.RunAction "Wait-WebElementInvisible", CreateParams("Selector", "#loadingSpinner", "TimeoutSec", 10)
    Debug.Print "  -> [成功] ローカルスピナーの非表示(Invisible)を確認しました"
    
    ' 2. 処理結果(#asyncResult)がDOM上に出現(Visible)するまでポーリング待機
    rpaEngine.RunAction "Wait-WebXPathElement", CreateParams("XPath", "//div[@id='asyncResult']", "TimeoutSec", 10)
    Debug.Print "  -> [成功] 非同期結果要素の出現(Visible)を確認しました"
    
    Sleep 2000 ' 目視確認用のウェイト
    
    ' ==========================================================================
    ' [テスト 3/4] 03_table_data.html (テーブル抽出 ＆ iframe内操作検証)
    ' ==========================================================================
    Debug.Print "--- [3/4] 03_table_data.html テスト開始 ---"
    targetUrl = "file:///" & sandboxPath & "03_table_data.html"
    
    ' ページ遷移 ＆ DOM構築完了の待機
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    
    Debug.Print "★ 変更: Wait-WebPageLoad -> Wait-WebDocumentReady"
    ' Chromium（Edge/Chrome）では、PCローカル上の file:// プロトコルで <iframe> を読み込んだ際、
    ' 親画面に対して全リソースロード完了イベント（Page.loadEventFired）が正しく通知されない。（ブラウザの制限）
    rpaEngine.RunAction "Wait-WebDocumentReady"
    Sleep 1000
    
    ' --- A. パターンA: メモリへ文字列として一括取得 (職員テーブル) ---
    Dim rawTableData As String
    rawTableData = rpaEngine.RunAction("Export-WebTableToCsv", CreateParams("Selector", "#sampleTable", "FileName", ""))
    
    Dim rows() As String
    Dim cols() As String
    Dim i As Long
    
    ' <R> (行区切り) と <T> (列区切り) による二次元データのパース
    rows = Split(rawTableData, "<R>")
    Debug.Print "  -> [成功:パターンA] メモリ抽出完了 (総行数: " & UBound(rows) + 1 & " 行)"
    For i = 0 To UBound(rows)
        cols = Split(rows(i), "<T>")
        Debug.Print "     [行 " & i & "] " & Join(cols, " | ")
    Next i
    
    ' --- B. パターンB: 通常のCSVファイルとして直接出力 ---
    Dim csvFileName As String
    csvFileName = "Ch1_ExportedTable.csv"
    
    QueryPerformanceCounter startTime ' 処理開始時刻の記録
    
    ' PSエンジン内部で -Force 上書き保存されるため、VBA側の事前Killは不要
    rpaEngine.RunAction "Export-WebTableToCsv", CreateParams("Selector", "#sampleTable", "FileName", csvFileName)
    
    QueryPerformanceCounter endTime ' 処理終了時刻の記録
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print "★ テーブル出力所要時間: " & Format(elapsedTimeMs, "0.00") & " ミリ秒"

    Debug.Print "  -> [成功:パターンB] CSVファイル直接出力完了 (Logsフォルダ内)"

    ' --------------------------------------------------------------------------
    ' --- C-1. インフレーム(iframe)内要素の出現待機 ＆ クリック検証 ---
    ' --------------------------------------------------------------------------
    ' file://（ローカルHTML）環境における制限あり
    ' Chromium（Edge）の厳格な Same-Origin Policy（同一生起元ポリシー）により、
    ' ローカルファイル同士の iframe クロスドメインアクセスは Origin: null としてブラウザ通信層で一律ブロックされます。

    On Error Resume Next
    
    Debug.Print "  --- file://（ローカルHTML）環境における制限( エラー!! になる。 ） ---"
    ' 1. iframe(#testFrame) 内のボタン(#btnFrameAction)出現待機
    rpaEngine.RunAction "Wait-WebElementInFrame", CreateParams( _
        "FrameSelector", "#testFrame", _
        "ElementSelector", "#btnFrameAction", _
        "TimeoutSec", 10)
    
    ' 2. iframe(#testFrame) 内のボタン(#btnFrameAction)クリック実行
    rpaEngine.RunAction "Invoke-WebClickInFrame", CreateParams( _
        "FrameSelector", "#testFrame", _
        "ElementSelector", "#btnFrameAction")
    ' エラー状態をリセット
    On Error GoTo 0

    ' --------------------------------------------------------------------------
    ' --- C-2. 公開Webサイト (https://) での iframe 内要素の操作検証 ---
    ' --------------------------------------------------------------------------
    On Error Resume Next
    
    Debug.Print "  --- 公開Webサイト(https://)での iframe 操作検証 ---"
    ' _. 世界標準の自動化テスト用 iframe サイトへアクセス
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", "https://the-internet.herokuapp.com/iframe")
    rpaEngine.RunAction "Wait-WebPageLoad"
    Debug.Print "★ 以下テストは、公開Webサイトにアクセスできた前提で検証（On Error Resume Next）"
    
    ' 1. iframe(#mce_0_ifr) 内にあるエディタ本体(body#tinymce)の出現待機
    rpaEngine.RunAction "Wait-WebElementInFrame", CreateParams( _
        "FrameSelector", "#mce_0_ifr", _
        "ElementSelector", "body#tinymce", _
        "TimeoutSec", 15)
    Debug.Print "  -> [Wait-WebElementInFrame] iframe内のエディタ要素(body#tinymce)出現を確認"
    ' 2. iframe(#mce_0_ifr) 内のエディタ本体(body#tinymce)をクリックしてフォーカス獲得
    rpaEngine.RunAction "Invoke-WebClickInFrame", CreateParams( _
        "FrameSelector", "#mce_0_ifr", _
        "ElementSelector", "body#tinymce")
    Debug.Print "  -> [Invoke-WebClickInFrame] iframe内のエディタ本体クリックを実行しました"
    ' エラー状態をリセット
    On Error GoTo 0

    Sleep 2000 ' 目視確認用のウェイト
    
    Debug.Print "--- キャッシュクリア実行 ---"
    rpaEngine.RunAction "Clear-WebCache", CreateParams("Mode", "CacheOnly")
    DoEvents
    
    ' ==========================================================================
    ' [テスト 4/4] 04_file_download.html (サイレントダウンロード)
    ' ==========================================================================
    Debug.Print "--- [4/4] 04_file_download.html テスト開始 ---"
    targetUrl = "file:///" & sandboxPath & "04_file_download.html"
    
    Dim downloadDir As String
    Dim saveFileName As String
    Dim fullSavePath As String
    
    downloadDir = ThisWorkbook.Path & "\RPA_Downloads"
    saveFileName = "Ch1_DownloadTest.csv"
    fullSavePath = downloadDir & "\" & saveFileName
    
    ' サイレントダウンロードでは同名ファイルが存在するとEdgeが「ファイル名(1).csv」と
    ' 別名保存してしまいWait-FileDownloadが正しく検出できなくなる
    If Dir(fullSavePath) <> "" Then Kill fullSavePath
    
    ' ページ遷移 ＆ DOM構築完了の待機
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    
    QueryPerformanceCounter startTime ' 処理開始時刻の記録
    
    ' 1. 保存先フォルダとファイル名を固定予約 (OSダイアログの抑止)
    rpaEngine.RunAction "Enable-SilentDownload", CreateParams("DownloadDirectory", downloadDir, "FileName", saveFileName)
    ' 2. ダウンロード発火ボタンをクリック
    rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", "//a[@id='dlLink']")
    ' 3. 一時ファイル (.crdownload) の消滅とファイルロック解除を監視待機
    rpaEngine.RunAction "Wait-FileDownload", CreateParams("FilePath", fullSavePath, "TimeoutSec", 15)
    
    QueryPerformanceCounter endTime ' 処理終了時刻の記録
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print "★ DL所要時間: " & Format(elapsedTimeMs, "0.00") & " ミリ秒"
    
    Debug.Print "  -> [成功] サイレントダウンロード完了: " & fullSavePath
    
    MsgBox "【第1章 テスト】" & vbCrLf & _
           "　コンポーネントの動作確認が完了しました。", vbInformation

    ' ==========================================================================
    ' 【終了前処理】キャッシュクリアとセッションのクリーンアップ
    ' .. （WebView2のメモリ/ディスクキャッシュを破棄し、次回起動時の状態汚染を防ぐ）
    ' ==========================================================================
    Debug.Print "--- キャッシュクリア実行 ---"
    rpaEngine.RunAction "Clear-WebCache", CreateParams("Mode", "CacheOnly")
    DoEvents
    rpaEngine.RunAction "Clear-WebCache", CreateParams("Mode", "All")

' --- 正常終了 / エラー共通のメモリ解放領域 ---
CleanUp:
    If Not rpaEngine Is Nothing Then
        rpaEngine.CloseEngine
    End If
    Set rpaEngine = Nothing
    Exit Sub

' --- 異常系のハンドリング (RPAエンジンからの構造化エラーの捕捉) ---
ErrorHandler:
    ' PowerShellエンジン内部からスローされた例外か判定
    If Err.Source = "PS_Engine" Then
        Dim rpaErr As RpaExceptionInfo
        ' エラー文字列から [関数名] [エラー種別] メッセージ を構造体にパース
        rpaErr = ParseRpaError(Err.Description)

        Select Case rpaErr.ErrorType
            Case "未発見"
                Debug.Print "【スキップ】指定されたDOM要素が見つかりません: " & rpaErr.Details
            Case "Timeout"
                Debug.Print "【待機超過】画面の応答がタイムアウトしました (" & rpaErr.FunctionName & ")"
            Case "JSエラー", "ネイティブエラー", "内部エラー"
                MsgBox "ブラウザ制御内で致命的なエラーが発生しました。" & vbCrLf & _
                       "発生関数: " & rpaErr.FunctionName & vbCrLf & _
                       "詳細内容: " & rpaErr.Message, vbCritical, "RPA内部エラー"
            Case Else
                Debug.Print "【その他エラー】" & rpaErr.RawText
        End Select
    Else
        ' VBA標準エラー
        MsgBox "VBAマクロエラー (" & Err.Number & "): " & Err.Description, vbCritical, "VBAエラー"
    End If
    
    ' エラー発生時のエビデンス（画面スクリーンショット）を自動保存
    On Error Resume Next
    rpaEngine.RunAction "Export-WebScreenshot", CreateParams("Prefix", "ErrorHandler")
    
    MsgBox "*** エラーが発生したため処理を中断します ***" & vbCrLf & Err.Description, vbCritical, "エラー終了"
    Resume CleanUp
End Sub

