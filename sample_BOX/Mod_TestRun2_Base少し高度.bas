Attribute VB_Name = "Mod_TestRun2_Base少し高度"
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
'   RPAエンジン テスト2 (第2部：高度なDOM/ダウンロード/待機制御)
' ==============================================================================
Sub Test_Run2_Base()

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
             "  0. ** 第2部 全パート一括テスト **" & vbCrLf & _
             " ---" & vbCrLf & _
             " 11. Web上の Shadow DOM ページ (テキスト取得)" & vbCrLf & _
             " 12. Virtual Shadow DOM での操作テスト" & vbCrLf & _
             " 13.（マスク）、スピナー・ローディング待機" & vbCrLf & _
             " 14.（PDF） UIA保存テスト（アプローチ検証）" & vbCrLf & _
             " 15. Fetch API によるバイナリのサイレントDL" & vbCrLf & _
             " ---" & vbCrLf & _
             " 21. PDFをブラウザ内蔵ビューアで（注:Adobe Acrobat等）"
             
    ans = InputBox(prompt, "第2部 テストパート選択")
    
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
    Application.Wait Now + TimeValue("00:00:01")

    ' 実行時のエンジン動作設定 (要素ハイライト機能の制御: True=有効)
    ' .. (有効関数: Invoke-WebXPathClick、Set-WebXPathTextInput、WebXPathText)
    rpaEngine.RunAction "Set-EngineConfig", CreateParams("EnableHighlight", True)
    
    ' --- テスト実行 ---
    Select Case ans
        Case "0"
            If Not TEST_Part11() Then GoTo CleanUp
            Sleep 2000 ' テスト画面を目視する（以下、Sleepも同様、適時調整する）
            If Not TEST_Part12() Then GoTo CleanUp
            Sleep 2000
            If Not TEST_Part13() Then GoTo CleanUp
            Sleep 2000
            If Not TEST_Part14() Then GoTo CleanUp
            Sleep 2000
            If Not TEST_Part15() Then GoTo CleanUp
        Case "11"
            If TEST_Part11() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case "12"
            If TEST_Part12() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case "13"
            If TEST_Part13() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case "14"
            If TEST_Part14() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case "15"
            If TEST_Part15() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case "21"
            If TEST_Part21() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case "99" ' テスト用ダミー
            If TEST_Part99() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case Else
            MsgBox "無効な入力です。", vbExclamation
            GoTo CleanUp
    End Select

    On Error GoTo ErrorHandler
    ' ==========================================================================
    Debug.Print "=== テスト完了 ==="
    MsgBox "第2部のテストが完了しました！" & vbCrLf & _
           "Logsフォルダに出力された証跡を確認してください。", vbInformation

CleanUp:
    If Not rpaEngine Is Nothing Then
        rpaEngine.CloseEngine
    End If
    Set rpaEngine = Nothing
    Application.StatusBar = False ' ステータスバーのリセット
    Exit Sub
    
ErrorHandler:
    If Err.Source = "PS_Engine" Then
        Dim rpaErr As RpaExceptionInfo
        rpaErr = ParseRpaError(Err.Description)
        
        Debug.Print "【実行中エラー】 関数: " & rpaErr.FunctionName & " / 種別: " & rpaErr.ErrorType
        Debug.Print " メッセージ: " & rpaErr.Message
        
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
        Debug.Print "【実行中エラー】 箇所: " & procedureName & " / 関数: " & rpaErr.FunctionName
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

Private Function TEST_Part11() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 11: Web上の Shadow DOM ページ (テキスト取得)"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler

    Dim result As String
    Dim targetUrl As String
    
    targetUrl = "https://the-internet.herokuapp.com/shadowdom"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"

    Debug.Print " >> 証跡ファイルを作成します..."
    rpaEngine.RunAction "Export-WebScreenshot", CreateParams("Prefix", "Test_WebScreenshot-P11")
    rpaEngine.RunAction "Export-WebHtml", CreateParams("Prefix", "Test_WebHtml-P11")
    rpaEngine.RunAction "Export-WebElementsToCsv", CreateParams("FileName", "Test_WebElements-P11.csv")
    Debug.Print " !! Shadow DOM（ページ）のため、詳細は取得できません。"
    
    ' ==========================================================================
    ' 【 HTMLソースからの要素の見つけ方 】
    ' 出力された Test_WebHtml-P11 のHTMLソースを見ると、以下のようになっています。
    '
    ' <my-paragraph>
    '   <span slot="my-text">Let's have some different text!</span>
    ' </my-paragraph>
    '
    ' 1. 生のHTMLソースには、開発者ツール(F12)で見える「#shadow-root」という文字は存在しません。
    '    (ブラウザがJavaScriptを使って、後から動的に「影の壁」を生成しているためです)
    ' 2. しかし、影の世界へ文字を送り込むための「実体」はHTML上に存在します。
    '    それが <span slot="my-text"> というタグです。
    ' 3. この slot="my-text" という属性が、「影の世界の my-text という穴に、この文字をはめ込んでね」
    '    という目印（パスポート）になっています。
    ' 4. したがって、この目印をそのまま使い "span[slot='my-text']" と指定することで、
    '    Shadow DOMに紐づく要素を正確に狙い撃ちすることができます。
    ' ==========================================================================
    
    Debug.Print "=========================================================="
    Debug.Print " !! Export-WebHtml（静的なHTML出力）ではShadow DOMの中身は映らない。"
    Debug.Print " ◆ 要素を見つける時は、必ず「F12キー（開発者ツール）」を開く。!!"

    ' ==========================================================================
    ' 【検証1】 XPathの限界（弱点）テスト
    ' ※ XPathはブラウザの仕様上、Shadow DOM( #shadow-root )の壁を絶対に越えられないため、
    '    要素が存在していても必ずエラー(未発見)になります。これを回避するためには
    '    後述の CSSセレクタ(deepQuerySelector) を使用する必要があります。
    ' ==========================================================================
    
    Debug.Print "=========================================================="
    Debug.Print " >> 【検証1】XPathでShadow DOM内部を検索します (※意図的にエラー)"
    ' 一時的にエラーを無視する
    On Error Resume Next
    
    ' ※ TimeoutSec を短め(3秒)にして、タイムアウトエラーを発生させる（＝異常系テスト）
    Debug.Print " !! Test-WebElement を検討する: 画面の状態によって処理を分岐させる (Wait-WebElement/ Wait-WebXPathElement)"
    rpaEngine.RunAction "Wait-WebXPathElement", CreateParams("XPath", "//my-paragraph//p", "TimeoutSec", 3)
    If Err.Number <> 0 Then
        Debug.Print " >> 想定通り、XPathではShadow DOMの壁に阻まれエラーになりました。"
        Debug.Print "    エラー内容: " & Err.Description
        Err.Clear ' エラーをクリアして続行
    Else
        Err.Raise 999, "Test", "(想定外！) XPathで取得できてしまいました。"
    End If
    ' エラーハンドリングを通常に戻す
    On Error GoTo ErrorHandler
    
    Debug.Print "=========================================================="
    Debug.Print " ◆ 非同期パイプ通信の仕様による動作（ラグ）により、表示に混乱が発生する！"
    
    ' ==========================================================================
    ' 【検証2】 CSSセレクタによる正しい要素取得
    ' <span slot="my-text"> は、外側からShadow DOMへ向けて投げ込まれた要素のため、
    ' CSSセレクタで属性を指定すれば正確に捕捉可能
    ' ==========================================================================
    
    Debug.Print "=========================================================="
    Debug.Print " >> 【検証2】CSSセレクタ(Get-WebText)でテキストを取得します"
    
    targetUrl = "https://the-internet.herokuapp.com/shadowdom"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"

    ' The Internet の Shadow DOM 内の spanタグ（テキスト）を取得する、Wait-WebElement のテストを兼ねる
    ' <span slot="my-text">Let's have some different text!</span>
    rpaEngine.RunAction "Wait-WebElement", CreateParams("Selector", "span[slot='my-text']")
    result = rpaEngine.RunAction("Get-WebText", CreateParams("Selector", "span[slot='my-text']"))
    Debug.Print " >> 期待するTEXT: " & "Let's have some different text!"
    Debug.Print " ◆ GET/テキスト: " & result

    TEST_Part11 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名
    Call HandleError_Type1("TEST_Part11")
    TEST_Part11 = False
End Function

Private Function TEST_Part12() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 12: Virtual Shadow DOM での操作テスト"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler

    Dim result As String
    Dim targetUrl As String
    Dim sandboxPath As String

    ' サンドボックス（ローカルHTML）の絶対パスを取得 (Windowsの \ を / に変換)
    sandboxPath = Replace(ThisWorkbook.Path & "\sandbox\", "\", "/")
    targetUrl = "file:///" & sandboxPath & "10_shadow_dom.html"
    ' ページ遷移 ＆ DOM構築完了の待機
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    
    Debug.Print " >> 証跡ファイルを作成します..."
    rpaEngine.RunAction "Export-WebScreenshot", CreateParams("Prefix", "Test_WebScreenshot-P12")
    rpaEngine.RunAction "Export-WebHtml", CreateParams("Prefix", "Test_WebHtml-P12")

    ' ==========================================================================
    ' ★ Part 11 と Part 12 の Shadow DOM の違い
    '
    ' > Part 11 (公開サイト): パターンA [スロット(差し込み)型]
    '   Shadow DOMの中に「穴(<slot>)」だけがあり、外側から文字を差し込む作り。
    '   外側に slot="my-text" という目印(パスポート)があるため、容易にアクセス可能
    '
    ' > Part 12 (本テスト): パターンB [完全カプセル(隠蔽)型]
    '   入力欄(<input>)やボタンを、最初からShadow DOMの壁の「内側」に直接作り込む設計。
    '   差し込み口(slot)が一切無いため、通常のquerySelector等では見つからない。
    '
    ' deepQuerySelector 関数が、再帰的にすべての shadowRoot を貫通し、
    ' 完全に隠蔽された要素(#shadow-input)を捕捉・操作できるかを確認する。
    ' ==========================================================================

    ' 1. テキストボックスに入力 (Shadow DOM内部へアプローチ)、Wait-WebElement のテストを兼ねる
    rpaEngine.RunAction "Wait-WebElement", CreateParams("Selector", "#shadow-input")
    rpaEngine.RunAction "Set-WebTextInput", CreateParams("Selector", "#shadow-input", "Value", "ShadowDOM テストOK")
    ' 2. ボタンをクリック
    rpaEngine.RunAction "Invoke-WebClick", CreateParams("Selector", "#shadow-btn")
    ' 3. テキストを取得
    result = rpaEngine.RunAction("Get-WebText", CreateParams("Selector", "#result"))
    Debug.Print " ◆ GET/テキスト: " & result

    TEST_Part12 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名
    Call HandleError_Type1("TEST_Part12")
    TEST_Part12 = False
End Function

Private Function TEST_Part13() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 13: （マスク）、スピナー・ローディング待機"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler

    Dim result As String
    Dim targetUrl As String
    Dim sandboxPath As String

    ' ==========================================================================
    Debug.Print "TEST1: >> Wait-WebScreenUnlock / 画面ロック解除待機..."
    ' サンドボックス（ローカルHTML）の絶対パスを取得 (Windowsの \ を / に変換)
    sandboxPath = Replace(ThisWorkbook.Path & "\sandbox\", "\", "/")
    targetUrl = "file:///" & sandboxPath & "11_overlay_mask.html"
    ' ページ遷移 ＆ DOM構築完了の待機
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    
    ' 1. 「Start」ボタンをクリック
    Debug.Print " >> Startボタンを押下..."
    rpaEngine.RunAction "Invoke-WebClick", CreateParams("Selector", "#btnRun")
    
    QueryPerformanceCounter startTime
    ' 2. PS側に待機を指示 (セレクタ指定なし)、描画猶予(GracePeriod)を持つ
    Debug.Print " >> Wait-WebScreenUnlock による検知＆待機中..."
    result = rpaEngine.RunAction("Wait-WebScreenUnlock", CreateParams("TimeoutSec", 30))
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print " ◆ 計測/マスク消失待機 (TimeMs) : " & Format(elapsedTimeMs, "0.00")
    Debug.Print " >> Psからの応答: " & result
    
    ' 3. マスク解除後の動作確認
    result = rpaEngine.RunAction("Get-WebText", CreateParams("Selector", "#resultMsg"))
    Debug.Print " >> 期待するTEXT: 処理が完了しました！"
    Debug.Print " ◆ GET/テキスト: " & result
    
    Sleep 1000 ' テスト画面を目視する（適時調整する）
    ' ==========================================================================
    Debug.Print "=========================================================="
    Debug.Print "TEST2: >> Wait-WebElementInvisible / スピナー待機..."
    targetUrl = "https://practice-automation.com/spinners/"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    ' ページ全体ではなく、DOM構築完了で高速待機
    rpaEngine.RunAction "Wait-WebDocumentReady", CreateParams("TimeoutSec", 10)
    
    QueryPerformanceCounter startTime
    ' 1. CSSセレクタ指定でローディングマスク (#spinner) の非表示化を待機
    ' .. （ページを開いた直後に自動で回っているスピナーが消えるの待つ）
    Debug.Print " >> スピナー(#spinner)の消滅を待機中..."
    rpaEngine.RunAction "Wait-WebElementInvisible", CreateParams("Selector", "#spinner", "TimeoutSec", 30)
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print " ◆ 計測/スピナー (TimeMs) : " & Format(elapsedTimeMs, "0.00")
    
    ' 2. 待機後、テキストを取得する
    result = rpaEngine.RunAction("Get-WebText", CreateParams("Selector", "h1"))
    Debug.Print " -- 期待するTEXT: Spinners"
    Debug.Print " ◆ GET/テキスト: " & result

    Sleep 1000 ' テスト画面を目視する（適時調整する）
    ' ==========================================================================
    Debug.Print "=========================================================="
    Debug.Print "TEST3: >> Wait-WebElementInvisible / ローディング待機..."
    targetUrl = "https://the-internet.herokuapp.com/dynamic_loading/1"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    
    ' 1. 「Start」ボタンをクリック
    Debug.Print " >> Startボタンを押下..."
    rpaEngine.RunAction "Invoke-WebClick", CreateParams("Selector", "#start button")
    
    QueryPerformanceCounter startTime
    ' 2. CSSセレクタ指定でローディングマスク (#loading) の非表示(Invisible)化を待機
    Debug.Print " >> Wait-WebElementInvisible によるマスク消滅待機中..."
    rpaEngine.RunAction "Wait-WebElementInvisible", CreateParams("Selector", "#loading", "TimeoutSec", 30)
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print " ◆ 計測/ローディング (TimeMs) : " & Format(elapsedTimeMs, "0.00")
    
    ' 3. 待機後、テキストを取得する（CSS指定、ハイライトさせない）
    result = rpaEngine.RunAction("Get-WebText", CreateParams("Selector", "#finish h4"))
    Debug.Print " -- 期待するTEXT: Hello World!"
    Debug.Print " ◆ GET/テキスト: " & result
       
    Sleep 1000 ' テスト画面を目視する（適時調整する）
    ' ==========================================================================
    Debug.Print "=========================================================="
    Debug.Print "TEST4: >> Wait-WebXPathElementDisappear / ローディング待機..."
    targetUrl = "https://the-internet.herokuapp.com/dynamic_loading/1"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    
    ' 1. 「Start」ボタンをクリック
    Debug.Print " >> Startボタンを押下します..."
    rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", "//div[@id='start']/button")
    
    QueryPerformanceCounter startTime
    ' 2. ローディング要素 (<div id='loading'>) が消滅するまで待機（最大30秒）
    Debug.Print " >> ローディング画面の消滅を待機中..."
    rpaEngine.RunAction "Wait-WebXPathElementDisappear", CreateParams("XPath", "//div[@id='loading']", "TimeoutSec", 30)
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print " ◆ 計測/ローディング (TimeMs) : " & Format(elapsedTimeMs, "0.00")
    
    ' 3. 待機後、テキストを取得する
    result = rpaEngine.RunAction("Get-WebXPathText", CreateParams("XPath", "//div[@id='finish']/h4"))
    Debug.Print " -- 期待するTEXT: " & "Hello World!"
    Debug.Print " ◆ GET/テキスト: " & result

    TEST_Part13 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名
    Call HandleError_Type1("TEST_Part13")
    TEST_Part13 = False
End Function

Private Function TEST_Part14() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 14: 一般サイト(PDF) UIA保存テスト（アプローチ検証）"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler
    
    Dim result As String
    Dim targetUrl As String
    Dim saveDir As String
    Dim saveFileName As String
    Dim fullSavePath As String
    Dim success As Boolean
    
    saveDir = ThisWorkbook.Path & "\RPA_Downloads"
    saveFileName = "DummyPDF_UIA_" & Format(Now, "yyyyMMdd_HHmmss") & ".pdf"
    fullSavePath = saveDir & "\" & saveFileName
    
    If Dir(saveDir, vbDirectory) = "" Then MkDir saveDir

    targetUrl = "https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    
    ' 1. PDFビューア(プラグイン)がHTML上に立ち上がるのを待機
    rpaEngine.RunAction "Wait-WebElement", CreateParams("Selector", "embed[type='application/pdf']", "TimeoutSec", 15)
    Sleep 2000 ' レンダリングとプロセス分離の完了を待つ非同期ラグ考慮
    
    success = False
    ' ==========================================================================
    Debug.Print "=========================================================="
    Debug.Print " >> [検証0] UIAでPDFツールバーのボタンを直接クリック試行..."
    ' 一時的にエラーを無視する
    On Error Resume Next
    Err.Clear
    
    ' パターンA：「保存」ボタンを探してクリック
    rpaEngine.RunAction "Invoke-UiaAction", CreateParams("Action", "Click", "Name", "保存", "TimeoutSec", 3)
    ' パターンB：「名前を付けて保存」を探してクリック
    If Err.Number <> 0 Then
         Err.Clear
         rpaEngine.RunAction "Invoke-UiaAction", CreateParams("Action", "Click", "Name", "名前を付けて保存", "TimeoutSec", 3)
    End If
    
    ' 理由は「OSとブラウザの壁」。ブラウザ内部で独自描画されたUIのため認識されない。
    Debug.Print " !! [検証0 想定通り] ボタン認識不可によるタイムアウト (エラー: " & Err.Description & ")"
    Err.Clear
    On Error GoTo ErrorHandler ' 通常のエラー監視に戻す
    
    ' ==========================================================================
    Debug.Print "=========================================================="
    Debug.Print " >> [アプローチ1] 中央クリック ＋ (Ctrl+S) での保存を試行中..."
    On Error Resume Next
    Err.Clear
    
    ' ウィンドウタイトルの前方一致等で確実にフォーカスを奪取する
    rpaEngine.RunAction "Switch-AppWindow", CreateParams("Name", "RPA Browser")
    Sleep 500
    rpaEngine.RunAction "Invoke-DesktopCenterClick"
    Sleep 500
    
    ' PDFビューアに対して物理キー「Ctrl+S (保存)」を送信する
    rpaEngine.RunAction "Invoke-DesktopSendKeys", CreateParams("Keys", "^s", "WaitMs", 1500)
    rpaEngine.RunAction "Invoke-UiaSafeSaveAs", CreateParams("FilePath", fullSavePath, "TimeoutSec", 5)
    
    If Err.Number = 0 And Dir(fullSavePath) <> "" Then
        Debug.Print " ◆ (想定外！)／[アプローチ1] ファイルが保存されました: " & fullSavePath
        success = True
    Else
        Debug.Print " !! [アプローチ1 想定通り] ダイアログを捕捉できませんでした (エラー: " & Err.Description & ")"
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    ' ==========================================================================
    If Not success Then
        Debug.Print "=========================================================="
        Debug.Print " >> [アプローチ2] 代替ショートカット・キー操作による保存を試行中..."
        On Error Resume Next
        Err.Clear
        
        rpaEngine.RunAction "Switch-AppWindow", CreateParams("Name", "RPA Browser")
        Sleep 500
        
        ' PDFビューア上で確実にキーを受け付けさせるため、Tabキーで内部フォーカスを刺激
        rpaEngine.RunAction "Invoke-DesktopSendKeys", CreateParams("Keys", "{TAB}", "WaitMs", 300)
        ' 右クリックメニュー(Shift + F10)などの代替入力をエミュレート
        rpaEngine.RunAction "Invoke-DesktopSendKeys", CreateParams("Keys", "+{F10}", "WaitMs", 1000)
        ' 名前を付けて保存 (通常 's' や 'a' など、ブラウザ環境による)
        rpaEngine.RunAction "Invoke-DesktopSendKeys", CreateParams("Keys", "s", "WaitMs", 1500)
        rpaEngine.RunAction "Invoke-UiaSafeSaveAs", CreateParams("FilePath", fullSavePath, "TimeoutSec", 5)
        
    ' ==========================================================================
    ' 【注意】ときどき長時間の待機(フリーズ)と 0x80131505 エラーが発生する
    ' ここで Invoke-UiaSafeSaveAs を実行すると、TimeoutSec=5 を指定しているにも関わらず、
    ' 約60秒間処理が停止し、最終的に「操作がタイムアウトになりました (0x80131505)」という
    ' エラーが返ることがあります。
    '
    ' Chromium系ブラウザ(Edge/Chrome)のPDFビューアは独立したプラグインで描画されており、
    ' 右クリックメニュー等の特殊なレイヤーが開いている最中に外部からUIAスキャン(FindFirst)
    ' を受けると、アクセシビリティツリーが応答を返しなくなる。（ブロックされる）
    '
    ' PowerShell側のタイマーに関わらず、Windows OS (UIA API) 側の内部タイムアウト上限
    ' （約60秒）まで強制的に待たされた末に、UIA_E_TIMEOUT (0x80131505) が発生する。
    ' ==========================================================================

        If Err.Number = 0 And Dir(fullSavePath) <> "" Then
            Debug.Print " ◆ (想定外！)／[アプローチ2] ファイルが保存されました: " & fullSavePath
            success = True
        Else
            Debug.Print " !! [アプローチ2 想定通り] 代替アプローチもブロックされました (エラー: " & Err.Description & ")"
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    End If
    
    ' ==========================================================================
    Debug.Print "=========================================================="
    If success Then
        Debug.Print " ◆ (想定外！)／ いずれかのアプローチでPDFの保存に成功しました。"
    Else
        Debug.Print " !! すべてのUIA/キーボードアプローチがPDFプラグインにブロックされました。"
    End If

    Debug.Print "=========================================================="
'''    Debug.Print " 【正攻法へのフォールバック】"
    Debug.Print " ユーザー操作を模倣するUI操作（キー送信・UIA）は、直感的ですがブラウザの"
    Debug.Print " 仕様変更やPCの解像度、処理ラグの影響を直接受けるため、不安定要素が多い。"
    Debug.Print " !! (参考) Part 15: Fetch API によるバイナリのサイレント直接DL"
    Debug.Print " !! (参考) Part 21: Get-WebEmbedPdfUrl による内部URL抽出とFetch API連携"

    ' ==========================================================================
    ' 最終クリーンアップ（開いたままの「名前を付けて保存」ダイアログ等を閉じる）
    Debug.Print "=========================================================="
    Debug.Print " >> 後処理: 残留ダイアログの破棄と画面リセットを実行します..."
    On Error Resume Next
    
    rpaEngine.RunAction "Switch-AppWindow", CreateParams("Name", "RPA Browser")
    Sleep 500
    ' ESCキーを2回送信して、OSレベルのメニューやダイアログを強制キャンセルする
    rpaEngine.RunAction "Invoke-DesktopSendKeys", CreateParams("Keys", "{ESC}", "WaitMs", 300)
    rpaEngine.RunAction "Invoke-DesktopSendKeys", CreateParams("Keys", "{ESC}", "WaitMs", 300)
    On Error GoTo ErrorHandler
    
    ' PDFの表示を終了し、画面をリセット（空白ページへ遷移）する
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", "about:blank")
    Sleep 1000 ' 画面遷移が落ち着くまで待機
    
    TEST_Part14 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名)
    Call HandleError_Type1("TEST_Part14")
    TEST_Part14 = False
End Function

Private Function TEST_Part15() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 15: Fetch API によるバイナリのサイレント直接DL"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler

    Dim result As String
    Dim targetUrl As String
    Dim saveDir As String
    Dim saveFileName As String
    Dim fullSavePath As String
    
    Dim pdfHref As String
    Dim errNum As Long
    Dim targetXPath As String
    Dim isExist As String
    
    saveDir = ThisWorkbook.Path & "\RPA_Downloads"
    saveFileName = "名前を変更_Herokuapp_Downloaded.pdf"
    fullSavePath = saveDir & "\" & saveFileName
    
    targetUrl = "https://the-internet.herokuapp.com/download"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    
    ' 1. ダウンロードの事前予約 (ダイアログ抑止 ＆ 保存先固定)
    rpaEngine.RunAction "Enable-SilentDownload", CreateParams("DownloadDirectory", saveDir, "FileName", saveFileName)

    ' ==========================================================================
    ' [なぜ Test-WebElement で事前確認するのか？]
    ' Get-WebAttribute の内部にはすでに **要素の待機処理が含まれています** が、いきなり実行して
    ' 万が一PDFリンクが存在しなかった場合、エンジンから「未発見」の例外(エラー)がスローされ、
    ' マクロ自体が強制的に ErrorHandler へ飛んで異常終了してしまいます。
    '
    ' そのため、エラーを投げずに True/False を返す Test-WebElement を事前確認として使い、
    ' 「要素が無いならスキップする」という安全な分岐(If/Else)を構築しています。
    ' (※要素が存在した場合、直後の Get-WebAttribute 内の待機は、ほぼ0秒で通過するためロスはない)
    ' ==========================================================================

    ' 2. ページ内にある「最初のPDFリンク」の存在を確認 (判定のためTimeoutを短めに設定)
    ' .. "a[href$='.pdf']" は「href属性が .pdf で終わる最初の <a> タグ」を意味します。
    Debug.Print " >> ページ内の1番目のPDFリンクを検索中..."
    isExist = rpaEngine.RunAction("Test-WebElement", CreateParams("Selector", "a[href$='.pdf']", "TimeoutSec", 3))
    Debug.Print " !! Test-WebElement: 画面の状態によって処理を分岐させる (Wait-WebElement/ Wait-WebXPathElement)"
    
    If isExist = "True" Then
        ' 3. そのPDFリンクの URL (href属性) を取得する
        ' ※ Get-WebAttribute の内部で自動待機するため、ここでは直接取得してOK
        pdfHref = rpaEngine.RunAction("Get-WebAttribute", CreateParams("Selector", "a[href$='.pdf']", "Attribute", "href"))
        Debug.Print " ◆ 取得/1番目のPDF-URL: " & pdfHref
        
        QueryPerformanceCounter startTime
        ' 4. Fetch APIで裏側からのサイレントダウンロードを発火
        ' .. （画面のリンクをクリックせず、セッション・Cookieを保持したままバイナリを裏で直接取得する）
        rpaEngine.RunAction "Invoke-WebFetchDownload", CreateParams("TargetUrl", pdfHref)
        ' 5. 一時ファイル(.crdownload)の消失とロック解除を監視待機
        rpaEngine.RunAction "Wait-FileDownload", CreateParams("FilePath", fullSavePath, "TimeoutSec", 30)
        QueryPerformanceCounter endTime
        elapsedTimeMs = (endTime - startTime) / freq * 1000
        Debug.Print " ◆ 計測/ダウンロード (TimeMs) : " & Format(elapsedTimeMs, "0.00")
        
        If Dir(fullSavePath) <> "" Then
            Debug.Print " ◆ 保存/FetchDL: " & fullSavePath
        Else
            Err.Raise 999, "Test", "ダウンロードファイルが見つかりません"
        End If
    Else
        Debug.Print " !! [スキップ] 指定したPDFリンクが存在しませんでした。"
    End If
    
    ' ==========================================================================
    ' ページ内にある「最後(またはN番目)のPDFリンク」を取得する例
    ' .. XPathの (//a[contains(@href, '.pdf')])[last()] を使用して、
    ' .. ページ内に複数あるPDFリンクのうち、一番末尾の要素を特定します。
    '
    '   1番目: (//a[contains(@href, '.pdf')])[1]
    '   2番目: (//a[contains(@href, '.pdf')])[2]
    '   最後 : (//a[contains(@href, '.pdf')])[last()]
    ' ==========================================================================
    
    Debug.Print "=========================================================="
    Debug.Print " >> ページ内の 最後 のPDFリンクを検索中..."
    targetXPath = "(//a[contains(@href, '.pdf')])[last()]"
    isExist = rpaEngine.RunAction("Test-WebElement", CreateParams("XPath", targetXPath, "TimeoutSec", 3))

    If isExist = "True" Then
        pdfHref = rpaEngine.RunAction("Get-WebAttribute", CreateParams("XPath", targetXPath, "Attribute", "href"))
        Debug.Print " ◆ 取得/last_のPDF-URL: " & pdfHref
        ' （必要に応じてここにもFetchDLの処理を追加）
    Else
        Debug.Print " !! [スキップ] 指定したPDFリンクが存在しませんでした。"
    End If

    TEST_Part15 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名
    Call HandleError_Type1("TEST_Part15")
    TEST_Part15 = False
End Function

Private Function TEST_Part21() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 21: ブラウザ標準のPDFビューア(Get-WebEmbedPdfUrlのテスト) "
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler
    
    Dim result As String
    Dim targetUrl As String
    Dim actualPdfUrl As String
    
    Dim saveDir As String
    Dim saveFileName As String
    Dim fullSavePath As String
    Dim isExist As String
    
    saveDir = ThisWorkbook.Path & "\RPA_Downloads"
    saveFileName = "W3C_Embed_Downloaded.pdf"
    fullSavePath = saveDir & "\" & saveFileName
       
    targetUrl = "https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf"
    Debug.Print " >> PDFをブラウザ内蔵ビューアで開きます..."
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    
    ' ==========================================================================
    ' 1. ブラウザ標準のPDFビューア(embed要素)が生成されるのを待機する。
    '
    ' [なぜここでSleepが必要か？]
    ' PDFデータの読み込み完了後、ブラウザ内蔵のPDFビューア(独立プロセス/PDFium等)が起動し、
    ' 画面上に <embed> 要素が描画されるまでには、非同期のタイムラグが発生する。
    ' すぐにWait-WebElementを呼ぶと、DOMに反映される前で検知漏れ（タイムアウト）に
    ' なる可能性があるため、プロセスの切り替わりと描画猶予として待機を入れています。
    ' ==========================================================================
    Sleep 2000
    
    isExist = rpaEngine.RunAction("Test-WebElement", CreateParams("Selector", "embed[type='application/pdf']", "TimeoutSec", 3))
    Debug.Print " !! Test-WebElement: 画面の状態によって処理を分岐させる (Wait-WebElement/ Wait-WebXPathElement)"
    If isExist <> "True" Then
        Debug.Print " ◆ [スキップ] 指定したPDFリンクが存在しませんでした。"
'        Exit Function
        Err.Raise 999, "Test", "指定したPDFリンクが存在しませんでした。"
    End If
    
    ' 2. Get-WebEmbedPdfUrl を実行し、embed要素が持っている本当のURLを抽出する。（PS側で about:blank を自動補正）
    actualPdfUrl = rpaEngine.RunAction("Get-WebEmbedPdfUrl")
    Debug.Print " ◆ 取得/絶対PDF-URL: " & actualPdfUrl
    
    ' PowerShell側で about:blank を吸収した。（純粋な取得失敗（空文字）のみを判定する）
    If actualPdfUrl = "" Then
        Debug.Print "=========================================================="
        Debug.Print " !! URLの抽出に失敗しました。"
        Debug.Print "    PDFのDOM構造が特殊であるか、サードパーティ製の拡張機能が"
        Debug.Print "    内蔵ビューアの動作に干渉している可能性があります。"
        Debug.Print "=========================================================="
        Err.Raise 999, "Test", "URLの抽出に失敗しました (空のURLが返却されました)"
    End If
    
    ' 3. ダウンロードの事前予約 (ダイアログ抑止 ＆ 保存先とファイル名を固定)
    rpaEngine.RunAction "Enable-SilentDownload", CreateParams("DownloadDirectory", saveDir, "FileName", saveFileName)
    ' 4. 抽出したURLを使って、Fetch APIで裏側からバイナリを直接ダウンロード
    rpaEngine.RunAction "Invoke-WebFetchDownload", CreateParams("TargetUrl", actualPdfUrl)
    ' 5. ダウンロード完了待機
    rpaEngine.RunAction "Wait-FileDownload", CreateParams("FilePath", fullSavePath, "TimeoutSec", 30)
    
    If Dir(fullSavePath) <> "" Then
        Debug.Print " ◆ 保存/FetchDL: " & fullSavePath
    Else
        Err.Raise 999, "Test", "ダウンロードファイルが見つかりません"
    End If
    
    ' ==========================================================================
    ' 後処理（画面クリア）
    Debug.Print " >> 後処理を実行します..."
    ' （直リンク・単一タブ型）、ブラウザ内蔵PDFビューアのプロセスを解放し、真っ白な状態に戻す
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", "about:blank")
    
    ' ==========================================================================

    Debug.Print " ◆ 要件に応じた後処理パターン A、B（今回は実行スキップ）"
    ' ◆ パターンA: 同じタブ内で遷移したPDF画面から元のページへ「戻る」場合
    ' ..  (※ ブラウザの履歴が存在する場合のみ有効)
'    rpaEngine.RunAction "Invoke-WebScript", CreateParams("Js", "window.history.back();")
'    rpaEngine.RunAction "Wait-WebPageLoad"
    
    ' ◆ パターンB: JavaScriptで別ウィンドウ（別タブ）として開いたPDF画面を閉じる場合
    ' .. (※ window.open()等で開いたタブのみ window.close() が有効)
'    rpaEngine.RunAction "Invoke-WebScript", CreateParams("Js", "window.close();")
'    Sleep 500 ' タブ破棄と親画面へのフォーカス復帰待ち
    ' ==========================================================================

    TEST_Part21 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名
    Call HandleError_Type1("TEST_Part21")
    TEST_Part21 = False
End Function

Private Function TEST_Part99() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 99: xxx"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler
    ' TEST

    TEST_Part99 = True
    Exit Function
ErrorHandler:
    Call HandleError_Type1("TEST_Part99")
    TEST_Part99 = False
End Function

