Attribute VB_Name = "Mod_TestRun1_Base一般的な"
Option Explicit

' ==============================================================================
' 【高精度タイマー用 API宣言】
' .. Win32 APIを使用してミリ秒未満（マイクロ秒単位）の処理時間を正確に測定する
' ==============================================================================
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' --- モジュールレベル変数 ---
Private freq As Currency             ' CPUの動作周波数 (タイマー精度)
    ' タイマー用変数
    Dim startTime As Currency
    Dim endTime As Currency
    Dim elapsedTimeMs As Double

Private rpaEngine As Ps_Engine       ' RPA操作エンジン本体クラス

' --- API宣言 ---
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' --- エンジンパス等の定数定義 ---
Private Const PS_ENGINE_LIB As String = "\Ps_Engine_Core_v204.ps1"

' ==============================================================================
'   RPAエンジン テスト (公開サイトを利用)
' ==============================================================================
Sub Test_Run1_Base()

    Dim enginePath As String
    Dim sessionId As String
    Dim useCdpPort As Integer   ' 通信モードの設定 (0: 標準, 9222: CDP高速通信)
    
    ' CPUのタイマー周波数を取得 (高精度計測の準備)
    QueryPerformanceFrequency freq
    
    ' --- 実行環境・セッションの初期化 ---
    enginePath = ThisWorkbook.Path & PS_ENGINE_LIB
    sessionId = "SESSION_" & Format(Now, "yyyyMMdd_HHmmss")
    ' --- CDP通信モードの設定 ---
    ' 通常の運用では標準モード(0)を推奨しますが、本テストシナリオでは
    ' CDP機能(Invoke-CdpScript等)を検証するため、ポート番号(例: 9222)を指定します。
    Debug.Print " >> 通常の運用では標準モード(0)を推奨、テストではポート番号を指定します。"
    useCdpPort = 9222
    
    Set rpaEngine = New Ps_Engine
    
    ' --- PowerShellエンジンを起動 ---
    ' 第3引数: CDPポート番号 (9222)
    ' 第4引数: IsDebugModeFlg (True = 実行時のパラメータや詳細ログをSTDOUTへ出力する)
    If Not rpaEngine.StartEngine(sessionId, enginePath, useCdpPort, True) Then
        MsgBox "RPAエンジンの起動に失敗しました。", vbCritical, "起動エラー"
        Exit Sub
    End If

    ' ==========================================================================
    ' --- テストシナリオの呼び出し ---
    ' ==========================================================================
    ' Excelを強制的に最前面に持ってくる
    Call ForceFocusExcel
    
    Dim prompt As String
    Dim ans As String
    
    prompt = "パートを選んでください：" & vbCrLf & _
            " 0. ** 全パート( UIA除く )一括テスト **" & vbCrLf & _
            " ---" & vbCrLf & _
            " 1. エンジン設定とキャッシュ操作" & vbCrLf & _
            " 2. ナビゲーションと取得系" & vbCrLf & _
            "     ＋ CSSセレクタによる標準DOM操作" & vbCrLf & _
            " 4. XPathによる要素操作" & vbCrLf & _
            " 5. ドロップダウン・チェックボックス" & vbCrLf & _
            " 6. JSネイティブ実行およびCDPネイティブ操作" & vbCrLf & _
            " 7. iframeとタブ(Window)管理" & vbCrLf & _
            " 8. エクスポート・証跡保存・テーブル解析" & vbCrLf & _
            " 9. サイレントダウンロードの正常系テスト" & vbCrLf & _
            " ---" & vbCrLf & _
            "10. UIA(デスクトップ操作) の正常系テスト (メモ帳を使用)"
            
    ans = InputBox(prompt, "テストパート選択")
    
    If ans = "" Then
        MsgBox "キャンセルされました。", vbInformation
        GoTo CleanUp
    Else
        MsgBox "選択したテストパート: ( " & ans & " )", vbInformation
    End If

    ' 裏に隠れてしまったブラウザ（RPA Browser）を一番手前に呼び戻す
    On Error Resume Next
    AppActivate "RPA Browser"
    On Error GoTo ErrorHandler ' エラー処理の設定を元に戻す
    ' 画面の切り替えが完全に落ち着くまでn秒待つ
    Application.Wait Now + TimeValue("00:00:02")

    ' 実行時のエンジン動作設定 (要素ハイライト機能の制御: True=有効)
    ' .. (有効関数: Invoke-WebXPathClick、Set-WebXPathTextInput、WebXPathText)
    rpaEngine.RunAction "Set-EngineConfig", CreateParams("EnableHighlight", True)

    ' --- (公開サイト)テスト実行 ---
    Select Case ans
        Case "0"
            If Not TEST_Part01() Then GoTo CleanUp
            Sleep 2000 ' テスト画面を目視する（以下、Sleepも同様、適時調整する）
            If Not TEST_Part02() Then GoTo CleanUp
            Sleep 2000
            If Not TEST_Part03() Then GoTo CleanUp
            Sleep 2000
            If Not TEST_Part04() Then GoTo CleanUp
            Sleep 2000
            If Not TEST_Part05() Then GoTo CleanUp
            Sleep 2000
            If Not TEST_Part06(useCdpPort) Then GoTo CleanUp
            Sleep 2000
            If Not TEST_Part07() Then GoTo CleanUp
            Sleep 2000
            If Not TEST_Part08() Then GoTo CleanUp
            Sleep 2000
            If Not TEST_Part09() Then GoTo CleanUp
            ' UIA(デスクトップ操作) ため、一括も可能だが個別テストで対応する
'xx         Sleep 2000
'xx         If Not TEST_Part10() Then GoTo CleanUp
        Case "1"
            If TEST_Part01() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case "2"
            If TEST_Part02() Then
                If TEST_Part03() Then
                    MsgBox "テストが終了しました。", vbInformation
                End If
            End If
            GoTo CleanUp
        Case "4"
            If TEST_Part04() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case "5"
            If TEST_Part05() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case "6"
            If TEST_Part06(useCdpPort) Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case "7"
            If TEST_Part07() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case "8"
            If TEST_Part08() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case "9"
            If TEST_Part09() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case "10"
            If TEST_Part10() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case Else
            MsgBox "無効な入力です。", vbExclamation
            GoTo CleanUp
    End Select

    On Error GoTo ErrorHandler
    ' ==========================================================================
    Debug.Print "=== テスト完了 ==="
    MsgBox "関数の呼び出しテストが完了しました！" & vbCrLf & _
           "Logsフォルダに出力された証跡を確認してください。", vbInformation

CleanUp:
    If Not rpaEngine Is Nothing Then
        rpaEngine.CloseEngine
    End If
    Set rpaEngine = Nothing
    Application.StatusBar = False ' ステータスバーのリセット
    Exit Sub
    
' ------------------------------------------------------------------
' 異常系のハンドリング (メインルーチン用)
' ------------------------------------------------------------------
ErrorHandler:
    If Err.Source = "PS_Engine" Then
        Dim rpaErr As RpaExceptionInfo
        rpaErr = ParseRpaError(Err.Description)
        
        Debug.Print "【実行中エラー】 関数: " & rpaErr.FunctionName & " / 種別: " & rpaErr.ErrorType
        Debug.Print " メッセージ: " & rpaErr.Message
        Debug.Print " 詳細(Details): " & rpaErr.Details
        
        MsgBox "RPAエンジン実行中にエラーが発生しました。" & vbCrLf & _
               "関数: " & rpaErr.FunctionName & vbCrLf & _
               "内容: " & rpaErr.Message, vbCritical, "RPAシステムエラー"
    Else
        MsgBox "VBAマクロエラー (" & Err.Number & "): " & Err.Description, vbCritical
    End If
    
    On Error Resume Next
    rpaEngine.RunAction "Export-WebScreenshot", CreateParams("Prefix", "Test_ErrorDump")
    Resume CleanUp
End Sub

    ' ==========================================================================
    ' 共通エラーハンドラ (サブプロシージャ用)
    ' ==========================================================================
Private Sub HandleError_Type1(ByVal procedureName As String)
    If Err.Source = "PS_Engine" Then
        Dim rpaErr As RpaExceptionInfo
        rpaErr = ParseRpaError(Err.Description)
        
        Debug.Print "【発生プロシージャ】 : " & procedureName
        Debug.Print "【実行中エラー】 関数: " & rpaErr.FunctionName & " / 種別: " & rpaErr.ErrorType
        Debug.Print " メッセージ: " & rpaErr.Message
        Debug.Print " 詳細(Details): " & rpaErr.Details
        
        MsgBox "RPAエンジン実行中にエラーが発生しました。" & vbCrLf & _
               "箇所: " & procedureName & vbCrLf & _
               "関数: " & rpaErr.FunctionName & vbCrLf & _
               "内容: " & rpaErr.Message, vbCritical, "RPAシステムエラー"
    Else
        MsgBox "VBAマクロエラー (" & Err.Number & "): " & Err.Description, vbCritical
    End If

    On Error Resume Next
    rpaEngine.RunAction "Export-WebScreenshot", CreateParams("Prefix", "Test_ErrorDump_" & procedureName)
End Sub

Private Function TEST_Part01() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 1: エンジン設定とキャッシュ操作"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler

    ' 実行時のエンジン動作設定 (要素ハイライト機能の制御: True=有効)
    ' .. (有効関数: Invoke-WebXPathClick、Set-WebXPathTextInput、WebXPathText)
    rpaEngine.RunAction "Set-EngineConfig", CreateParams("EnableHighlight", True)
    ' メモ書きテスト
    rpaEngine.RunAction "Write-DebugTextFile", CreateParams("Text", "総合テスト開始", "FileName", "TestMemo.txt")
    ' まっさらな状態でテストするため、ブラウザのキャッシュ/Cookieを全消去
    rpaEngine.RunAction "Clear-WebCache", CreateParams("Mode", "All")

    TEST_Part01 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名）
    Call HandleError_Type1("TEST_Part01")
    TEST_Part01 = False
End Function

Private Function TEST_Part02() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 2: ナビゲーションと取得系"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler

    Dim result As String
    Dim targetUrl As String

    targetUrl = "https://the-internet.herokuapp.com/login"
    ' ==========================================================================
    ' --- TEST1: Wait-WebPageLoad
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    
    QueryPerformanceCounter startTime
    ' 完全なロード完了を待機
    rpaEngine.RunAction "Wait-WebPageLoad"
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print " ◆ 計測/Wait-WebPageLoad (TimeMs) : " & Format(elapsedTimeMs, "0.00")

    ' ==========================================================================
    ' --- TEST2: Wait-WebDocumentReady
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)

    QueryPerformanceCounter startTime
    ' DOMの準備完了も併せて検証
    rpaEngine.RunAction "Wait-WebDocumentReady", CreateParams("TimeoutSec", 10)
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print " ◆ 計測/Wait-WebDocumentReady (TimeMs) : " & Format(elapsedTimeMs, "0.00")

    ' ==========================================================================
    result = rpaEngine.RunAction("Get-WebUrl")
    Debug.Print " ◆ URL取得: " & result
    result = rpaEngine.RunAction("Get-WebTitle")
    Debug.Print " ◆ タイトル取得: " & result

    ' ==========================================================================
    Debug.Print "=========================================================="
    Debug.Print "(Part 2): 画面遷移のURL・タイトル監視テストを開始..."
    ' --- TEST1: URL包含待機の計測
    targetUrl = "https://the-internet.herokuapp.com/login"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    
    ' ※ 通常の画面遷移は Wait-WebPageLoad を使用しますが、SPA(Single Page Application)や
    ' .. JavaScriptによる非同期の画面更新など、描画後にURLやタイトルが遅れて書き換わるシステムを
    ' .. 想定し、期待する文字列になるまでポーリング（監視）待機します。
    ' ------------------------------------------------------------------
    ' ※ [注意] 認証エラー等により、予期せぬ別URLにリダイレクトされた場合、
    ' .. 指定文字列が出現せずタイムアウトになるため、適切な TimeoutSec の設定が重要です。

    QueryPerformanceCounter startTime
    ' 待機: URLに "login" という文字列が含まれる状態
    rpaEngine.RunAction "Wait-WebUrlContains", CreateParams("Substring", "login", "TimeoutSec", 10)
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print " ◆ 計測/Wait-WebUrlContains (TimeMs) : " & Format(elapsedTimeMs, "0.00")
    
    ' --- TEST2: タイトル包含待機の計測
    targetUrl = "https://the-internet.herokuapp.com/login"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    
    QueryPerformanceCounter startTime
    ' 待機: タイトルに "The Internet" という文字列が含まれる状態
    rpaEngine.RunAction "Wait-WebTitleContains", CreateParams("Substring", "The Internet", "TimeoutSec", 10)
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print " ◆ 計測/Wait-WebTitleContains (TimeMs) : " & Format(elapsedTimeMs, "0.00")
    
    ' ==========================================================================
    Debug.Print "=========================================================="
    Debug.Print "(Part 2): 意図的にタイムアウトさせる異常系テスト..."
    On Error Resume Next
    ' 絶対に出現しない文字列("dummy_not_found")を3秒間待たせる。（テストのため時間短縮）
    rpaEngine.RunAction "Wait-WebUrlContains", CreateParams("Substring", "dummy_not_found", "TimeoutSec", 3)
    If Err.Number <> 0 Then
        Debug.Print " >> 想定通り、タイムアウト例外を検知しました (" & Err.Description & ")"
        Err.Clear
    Else
        Err.Raise 999, "Test", "[失敗] タイムアウト例外が発生しませんでした。"
    End If
    
    Debug.Print "=========================================================="
    ' 非同期パイプ通信の仕様上、VBA側のDebug.PrintとPowerShellからの標準出力(STDOUT)が
    ' イミディエイトウィンドウ上で前後して表示される場合があります。
    Debug.Print " ◆ 非同期パイプ通信の仕様による動作（ラグ）により、表示に混乱が発生する！"

    TEST_Part02 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名）
    Call HandleError_Type1("TEST_Part02")
    TEST_Part02 = False
End Function

Private Function TEST_Part03() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 3: CSSセレクタによる標準DOM操作"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler

    Dim result As String
    
    ' 要素の出現待機 -> 文字入力
    QueryPerformanceCounter startTime
    rpaEngine.RunAction "Wait-WebElement", CreateParams("Selector", "#username")
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print " ◆ 計測/Wait-WebElement (TimeMs) : " & Format(elapsedTimeMs, "0.00") & " ここの待機は不要！"
    
    ' ==========================================================================
    ' [パフォーマンス最適化のポイント]
    ' Invoke-WebClick や Set-WebTextInput 等の主要なアクション関数は、内部で自動的に
    ' Wait-WebElement (要素出現待機) を実行する仕様になっています。
    ' そのため、VBA側で事前にWaitを呼ぶと「二重待機」の通信ロスが発生します。
    ' 操作対象が明確な場合は、待機を省略して直接アクションコマンドを発行してください。
    ' ==========================================================================
    
    rpaEngine.RunAction "Set-WebTextInput", CreateParams("Selector", "#username", "Value", "tomsmith")
    ' 画面上のテキストを取得
    result = rpaEngine.RunAction("Get-WebText", CreateParams("Selector", "h2"))
    Debug.Print " ◆ H2テキスト: " & result
    ' クリック実行 -> クリック後に要素が非表示になる(画面遷移する)まで待機
    rpaEngine.RunAction "Invoke-WebClick", CreateParams("Selector", "button[type='submit']")
    rpaEngine.RunAction "Wait-WebElementInvisible", CreateParams("Selector", "#username")

    TEST_Part03 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名）
    Call HandleError_Type1("TEST_Part03")
    TEST_Part03 = False
End Function

Private Function TEST_Part04() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 4: XPathによる要素操作"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler
    
    Dim result As String
    Dim targetUrl As String

    targetUrl = "https://the-internet.herokuapp.com/login"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    
    ' XPath指定での待機と入力
    QueryPerformanceCounter startTime
    rpaEngine.RunAction "Wait-WebXPathElement", CreateParams("XPath", "//input[@id='password']")
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print " ◆ 計測/Wait-WebXPathElement (TimeMs) : " & Format(elapsedTimeMs, "0.00") & " ここの待機は不要！"
    
    ' ==========================================================================
    ' Invoke-WebXPathClick/ Set-WebXPathTextInput/ Get-WebXPathText には、
    ' 関数内に、要素出現を待機する Wait-WebXPathElement を包括している。
    ' ==========================================================================
    
    rpaEngine.RunAction "Set-WebXPathTextInput", CreateParams("XPath", "//input[@id='password']", "Value", "SuperSecretPassword!")
    ' Value値の取得検証
    result = rpaEngine.RunAction("Get-WebXPathText", CreateParams("XPath", "//input[@id='password']"))
    Debug.Print " ◆ XPathテキスト: " & result
    
    ' --- ログイン実行と画面遷移の確実な検知
    ' 1. Submitボタンをクリックしてログイン処理（画面遷移）を発火させる
    rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", "//button[@type='submit']")
    ' 2. 画面遷移の待機（消失監視）
    ' ※ [RPAの鉄則]：画面遷移を伴うクリック直後は、「次画面の要素が出現する」か
    ' .. 「現画面の要素が消滅する」のを必ず待機します。ここでは「パスワード入力欄が完全に消え去る」
    ' .. ことを最大5秒間監視し、システム遅延による次操作へのフライングを防止します。
    rpaEngine.RunAction "Wait-WebXPathElementDisappear", CreateParams("XPath", "//input[@id='password']", "TimeoutSec", 5)

    TEST_Part04 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名）
    Call HandleError_Type1("TEST_Part04")
    TEST_Part04 = False
End Function

Private Function TEST_Part05() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 5: ドロップダウン・チェックボックス"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler
    
    Dim result As String
    Dim targetUrl As String
    
    ' ==========================================================================
    targetUrl = "https://the-internet.herokuapp.com/checkboxes"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    ' チェックボックスを強制的にONに設定
    rpaEngine.RunAction "Set-WebCheckbox", CreateParams("Selector", "form#checkboxes input[type='checkbox']:nth-child(1)", "State", True)
    
    ' ==========================================================================
    targetUrl = "https://the-internet.herokuapp.com/dropdown"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    ' セレクトボックスのValue値を指定して選択
    rpaEngine.RunAction "Select-WebDropdown", CreateParams("Selector", "#dropdown", "Value", "2")

    TEST_Part05 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名）
    Call HandleError_Type1("TEST_Part05")
    TEST_Part05 = False
End Function

Private Function TEST_Part06(useCdpPort As Integer) As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 6: JSネイティブ実行およびCDPネイティブ操作"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler

    Dim result As String
    Dim targetUrl As String

    targetUrl = "https://ja.wikipedia.org/"

    If useCdpPort > 0 Then
        ' ブラウザ内部で直接JavaScriptを実行し戻り値を受け取る
        result = rpaEngine.RunAction("Invoke-WebView2NativeScript", CreateParams("Js", "return navigator.userAgent;"))
        Debug.Print " ◆ UserAgent: " & Left(result, 100) & "..."
    
        ' --- CDP (Chrome DevTools Protocol) 経由での直接制御テスト
        ' 1. (WebSocket通信によるJavaScriptの直接評価)
        ' .. WebView2標準のネイティブ通信(ExecuteScriptAsync等)が応答しない重い業務システムでも、
        ' .. CDP(Chrome DevTools Protocol)を用いたWebSocket通信により、ブラウザの深層へ直接命令を送り値を取得する。

        result = rpaEngine.RunAction("Invoke-CdpScript", CreateParams("Js", "return document.title;"))
         ' 2. Invoke-CdpCommand (ブラウザに対する低レイテンシなJSON-RPCコマンド発行)
        ' .. DOMの枠を超え、ブラウザ本体を直接制御して、ブラウザの内部バージョン情報を取得する
        result = rpaEngine.RunAction("Invoke-CdpCommand", CreateParams("Method", "Browser.getVersion"))
        
        rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
        rpaEngine.RunAction "Wait-WebPageLoad"
        
        ' DOMイベント（UI操作）に依存せず、裏側から直接値を注入・クリックする
        rpaEngine.RunAction "Set-CdpNativeTextInput", CreateParams("Selector", "#searchInput", "Value", "RPA")
        rpaEngine.RunAction "Invoke-CdpNativeClick", CreateParams("Selector", "#searchform button")
        rpaEngine.RunAction "Wait-WebPageLoad"
    End If

    TEST_Part06 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名）
    Call HandleError_Type1("TEST_Part06")
    TEST_Part06 = False
End Function

Private Function TEST_Part07() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 7: iframeとタブ(Window)管理"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler
    
    Dim result As String
    Dim targetUrl As String

    targetUrl = "https://the-internet.herokuapp.com/iframe"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    
    ' 通常の操作では届かないiframe内の要素に対して直接アタック
    rpaEngine.RunAction "Wait-WebElementInFrame", CreateParams("FrameSelector", "#mce_0_ifr", "ElementSelector", "body#tinymce")
    rpaEngine.RunAction "Invoke-WebClickInFrame", CreateParams("FrameSelector", "#mce_0_ifr", "ElementSelector", "body#tinymce")
    ' 現在認識している全てのタブ・ウィンドウ情報を取得する
    result = rpaEngine.RunAction("List-Tabs")
    Debug.Print " ◆ タブ一覧: " & Left(result, 100) & "..."
    
    ' フォーカス制御テスト
    rpaEngine.RunAction "Set-ActiveTab", CreateParams("TabId", "P01")
    rpaEngine.RunAction "Switch-Tab", CreateParams("TabId", "P01")
    rpaEngine.RunAction "Switch-TabByTitle", CreateParams("TitleSubstring", "The Internet")

    TEST_Part07 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名）
    Call HandleError_Type1("TEST_Part07")
    TEST_Part07 = False
End Function

Private Function TEST_Part08() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 8: エクスポート・証跡保存・テーブル解析"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler

    Dim result As String
    Dim targetUrl As String

    Dim rawData As String
    Dim rows() As String
    Dim cols() As String
    Dim i As Integer

    targetUrl = "https://the-internet.herokuapp.com/tables"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    ' 要素の出現待機
    rpaEngine.RunAction "Wait-WebElement", CreateParams("Selector", "table#table1")

    QueryPerformanceCounter startTime
    ' FileNameを指定しないことで、ファイル出力を行わず文字列としてメモリ上に高速展開する
    rawData = rpaEngine.RunAction("Export-WebTableToCsv", CreateParams("Selector", "table#table1", "FileName", ""))
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print " ◆ 計測/Export-WebTableToCsv (TimeMs) : " & Format(elapsedTimeMs, "0.00")
    
    ' 受け取ったテーブル文字列（<R>と<T>区切り）をVBAの配列にパースする
    If rawData <> "" Then
        rows = Split(rawData, "<R>")
        Debug.Print "  > 配列サイズ(行数): " & UBound(rows) + 1
        For i = 0 To UBound(rows)
            cols = Split(rows(i), "<T>")
            If UBound(cols) >= 0 Then
                Debug.Print "    行" & (i + 1) & " の A列: [" & cols(0) & "]"
            End If
        Next i
    End If

    ' ==========================================================================
    ' 各種デバッグ・証跡ファイルの生成（Logsフォルダへ出力）
    rpaEngine.RunAction "Export-WebHtml", CreateParams("Prefix", "Test_WebHtml")
    rpaEngine.RunAction "Export-WebScreenshot", CreateParams("Prefix", "Test_WebScreenshot")
    rpaEngine.RunAction "Export-WebElementsToCsv", CreateParams("FileName", "Test_WebElements.csv")
    rpaEngine.RunAction "Export-WebFrameTreeToCsv", CreateParams("FileName", "Test_WebFrameTree.csv")
    rpaEngine.RunAction "Export-WindowScreenshot", CreateParams("Prefix", "Test_WindowScreenshot")
    rpaEngine.RunAction "Export-WindowHierarchyToCsv", CreateParams("FileName", "Test_WindowHierarchy.csv")

    TEST_Part08 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名）
    Call HandleError_Type1("TEST_Part08")
    TEST_Part08 = False
End Function

Private Function TEST_Part09() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 9: サイレントダウンロードの正常系テスト"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler

    Dim result As String
    Dim targetUrl As String
    Dim saveDir As String
    Dim saveFileName As String
    Dim fullSavePath As String
    
    Dim targetHref As String
    Dim ext As String

    targetUrl = "https://the-internet.herokuapp.com/download"
    saveDir = ThisWorkbook.Path & "\RPA_Downloads"
    
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"

    ' クリック前に href 属性を取得し、拡張子を切り出す
    targetHref = rpaEngine.RunAction("Get-WebAttribute", CreateParams("XPath", "//div[@class='example']/a[1]", "Attribute", "href"))
    If InStrRev(targetHref, ".") > 0 Then
        ext = Mid(targetHref, InStrRev(targetHref, ".")) ' 例: ".jpg", ".pdf" など
    Else
        ext = ".bin" ' 拡張子不明時のフォールバック
    End If
    
    ' 動的に取得した拡張子を使って保存パスを組み立てる
    saveFileName = "SilentTest_" & Format(Now, "HHmmss") & ext
    fullSavePath = saveDir & "\" & saveFileName

    QueryPerformanceCounter startTime
    ' DLダイアログを強制バイパスするサイレント設定の有効化
    rpaEngine.RunAction "Enable-SilentDownload", CreateParams("DownloadDirectory", saveDir, "FileName", saveFileName)
    ' ダウンロードリンクを発火
    rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", "//div[@class='example']/a[1]")
    ' .crdownload の消滅とファイル確定を監視
    result = rpaEngine.RunAction("Wait-FileDownload", CreateParams("FilePath", fullSavePath, "TimeoutSec", 30))
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print " ◆ 計測/Wait-FileDownload (TimeMs) : " & Format(elapsedTimeMs, "0.00")

    If Dir(fullSavePath) <> "" Then
        Debug.Print " ◆ 保存/SilentDownload (" & fullSavePath & ")"
    Else
        Err.Raise 999, "Test", "ダウンロードされたファイルが見つかりません"
    End If
    
    TEST_Part09 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名）
    Call HandleError_Type1("TEST_Part09")
    TEST_Part09 = False
End Function

Private Function TEST_Part10() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 10: UIA(デスクトップ操作) の正常系テスト (メモ帳を使用)"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler
    
    Dim result As String
    Dim saveDir As String
    Dim saveFileName As String
    Dim fullSavePath As String
    Dim ws As Object

    saveDir = ThisWorkbook.Path & "\RPA_Downloads"
    saveFileName = "UIA_SaveTest.txt"
    fullSavePath = saveDir & "\" & saveFileName

    If Dir(fullSavePath) <> "" Then Kill fullSavePath

    Set ws = CreateObject("WScript.Shell")

    ' メモ帳の起動とフォーカス確保
    Shell "notepad.exe", vbNormalFocus
    Sleep 1000 ' 起動待機
    
    On Error Resume Next
    rpaEngine.RunAction "Switch-AppWindow", CreateParams("Name", "メモ")
    If Err.Number <> 0 Then
        Err.Clear
        rpaEngine.RunAction "Switch-AppWindow", CreateParams("Name", "Notepad")
    End If
    On Error GoTo ErrorHandler
    Sleep 500
    
    ' テキスト入力して「名前を付けて保存」ショートカットを発行
    ws.SendKeys "RPA UIA Full Action Test", True
    Sleep 500
    ws.SendKeys "^s", True
    Sleep 1000
    
    ' ==========================================================================
    ' 【 Invoke-UiaAction の重要パラメータ解説 】
    ' ・Action      : 実行する操作 (Click, Input, GetText, GetValue など)
    ' ・Name        : UI要素の表示名 (例: "キャンセル")。不要な場合は "*" を指定。
    ' ・AutomationId: アプリ内部で定義された不変のID (例: メモ帳のファイル名入力欄は "1001")。
    '                 表示名(Name)が変わってもOS言語等の影響を受けないため、確実な指定方法です。
    ' ・Mode        : 操作の実行方式 ("Pattern" または "Safe")
    '    - "Pattern" (推奨): UIA内部APIを呼び出し、裏側で処理を実行します。
    '                        ※[重要] Windows 11等の仕様変更でPattern操作がOSに弾かれた場合でも、
    '                        エンジンが自動検知して「Safe」モード(物理操作)へフォールバック(救済)する
    '                        仕組みが組み込まれています。
    '    - "Safe"          : 最初から物理的なマウスクリックやキーボード入力をシミュレートします。
    ' ==========================================================================
    
    ' Name指定による取得 (キャンセルボタンのテキストを取得)
    result = rpaEngine.RunAction("Invoke-UiaAction", CreateParams("Action", "GetText", "Name", "キャンセル", "TimeoutSec", 5))
    Debug.Print " ◆ UIAボタン取得: " & result
    Sleep 500
    ' AutomationId指定 ＋ Mode="Pattern" (裏側から瞬時にファイル名をセット)
    rpaEngine.RunAction "Invoke-UiaAction", CreateParams("Action", "Input", "Name", "*", "AutomationId", "1001", "Value", "PatternInput.txt", "Mode", "Pattern", "TimeoutSec", 5)
    Sleep 500
    ' 値の取得検証
    result = rpaEngine.RunAction("Invoke-UiaAction", CreateParams("Action", "GetValue", "Name", "*", "AutomationId", "1001", "TimeoutSec", 5))
    Debug.Print " ◆ UIA入力値取得: " & result
    Sleep 500
    ' AutomationId指定 ＋ Mode="Safe" (物理キーボード入力のエミュレートで値を上書き)
    rpaEngine.RunAction "Invoke-UiaAction", CreateParams("Action", "Input", "Name", "*", "AutomationId", "1001", "Value", "SafeInput.txt", "Mode", "Safe", "TimeoutSec", 5)
    Sleep 500
    ' Name指定 ＋ Mode="Pattern" によるクリック (キャンセルボタンを押して閉じる)
    rpaEngine.RunAction "Invoke-UiaAction", CreateParams("Action", "Click", "Name", "キャンセル", "Mode", "Pattern", "TimeoutSec", 5)
    Sleep 1000
    ' 再度ダイアログを開き、保存処理へ
    rpaEngine.RunAction "Invoke-DesktopSendKeys", CreateParams("Keys", "^s", "WaitMs", 1000)
    rpaEngine.RunAction "Invoke-UiaSafeSaveAs", CreateParams("FilePath", fullSavePath, "TimeoutSec", 10)
    
    Debug.Print " !! メモ帳プロセスを強制終了します..."
    Shell "cmd.exe /c taskkill /F /IM notepad.exe /T", vbHide
    Set ws = Nothing

    TEST_Part10 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名）
    Call HandleError_Type1("TEST_Part10")
    TEST_Part10 = False
End Function

