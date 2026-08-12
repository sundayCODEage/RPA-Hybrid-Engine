Attribute VB_Name = "Mod_RPA_Challengeデータ同梱"
Option Explicit

' ==============================================================================
' 【高精度タイマー用 API宣言】
' .. Win32 APIを使用してミリ秒未満（マイクロ秒単位）の処理時間を正確に測定する
' ==============================================================================
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' --- モジュールレベル変数 ---
Private freq As Currency             ' CPUの動作周波数 (タイマー精度)
Private rpaEngine As Ps_Engine       ' RPA操作エンジン本体クラス

' --- API宣言 ---
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' --- エンジンパス等の定数定義 ---
Private Const PS_ENGINE_LIB As String = "\Ps_Engine_Core_v204.ps1"

' ==============================================================================
' RPAチャレンジ（起動後、「Okボタン」を押して）
' ==============================================================================
Sub Test_RPAchallenge()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet

    Dim msg As String
    Dim ans As VbMsgBoxResult
    
    ' 確認メッセージの作成
    msg = "【データシートの確認】" & vbCrLf & _
          "アクティブシート（このシート）に、データ準備が［ 未だ ］でしたら、" & vbCrLf & _
          "Array保存データを、アクティブシートへ展開しますが、必要ですか？"
              
    ' アクティブシートを確認（データシートで無い）
    If ws.Cells(1, 1).Value = "First Name" And ws.Cells(1, 2).Value = "Last Name" Then
        ' データは設定済です。
    Else
        ' はい(Yes)・いいえ(No) の選択肢付きでメッセージを出す
        ans = MsgBox(msg, vbQuestion + vbYesNo, "データ自動展開の確認")
    
        ' 「はい」が押された場合のみ処理を実行
        If ans = vbYes Then
            Call DeployDataWithHeaderToActiveSheet(ws)
        End If
    End If

' ..----------------------------------------------------------------------------
    Dim enginePath As String
    Dim sessionId As String
    Dim useCdpPort As Integer   ' 通信モードの設定 (0: 標準, 9222: CDP高速通信)

    ' タイマー用変数
    Dim startTime As Currency
    Dim endTime As Currency
    Dim elapsedTimeMs As Double

    ' CPUのタイマー周波数を取得 (高精度計測の準備)
    QueryPerformanceFrequency freq
    
    ' --- 実行環境・セッションの初期化 ---
    enginePath = ThisWorkbook.Path & PS_ENGINE_LIB
    sessionId = "SESSION_" & Format(Now, "yyyyMMdd_HHmmss")
    useCdpPort = 0      ' CDP高速通信ポート (0指定で標準モード)
    
    Set rpaEngine = New Ps_Engine
    
    ' --- PowerShellエンジンを起動 ---
    ' 第3引数: CDPポート番号 (9222)
    ' 第4引数: IsDebugModeFlg (True = 実行時のパラメータや詳細ログをSTDOUTへ出力する)
    If Not rpaEngine.StartEngine(sessionId, enginePath, useCdpPort, True) Then
        MsgBox "RPAエンジンの起動に失敗しました。", vbCritical, "起動エラー"
        Exit Sub
    End If
' ..----------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    ' --- RPA Challenge サイトへ遷移 ---
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", "https://rpachallenge.com/")
    rpaEngine.RunAction "Wait-WebPageLoad"
    
    ' Excelを強制的に最前面へ
    Call ForceFocusExcel
    
    MsgBox "ページが表示されました。" & vbCrLf & _
           "準備ができたら「 OK 」を押してください。", _
           vbInformation, "/// 準備待機 ///"
           
    ' 【対策1】裏に隠れてしまったブラウザ（RPA Browser）を一番手前に呼び戻す
    On Error Resume Next
    AppActivate "RPA Browser"
    On Error GoTo ErrorHandler ' エラー処理の設定を元に戻す

    ' 【対策2】人間が操作した後の画面の切り替えが完全に落ち着くまで2秒待つ
    Application.Wait Now + TimeValue("00:00:02")
    ' --- RPAチャレンジ開始！（ハイライトをOFFにする） ---
    rpaEngine.RunAction "Set-EngineConfig", CreateParams("EnableHighlight", False)

    ' --- [Start] ボタンをクリックして計測開始 ---
    QueryPerformanceCounter startTime '（開始時間を取得）
    rpaEngine.RunAction "Invoke-WebClick", CreateParams("Selector", "button.waves-effect.col.s12.m12.l12.btn-large.uiColorButton")

    ' --- Excelのデータ行数分ループ (全10ラウンド) ---
    ' ※2行目から11行目までデータがある
    Dim r As Long
    For r = 2 To 11
        Application.StatusBar = "Round " & (r - 1) & " を入力中..."
        
        ' ① First Name
        rpaEngine.RunAction "Set-WebXPathTextInput", _
            CreateParams("XPath", "//label[normalize-space(text())='First Name']/following-sibling::input", "Value", ws.Cells(r, 1).Value)
        ' ② Last Name
        rpaEngine.RunAction "Set-WebXPathTextInput", _
            CreateParams("XPath", "//label[normalize-space(text())='Last Name']/following-sibling::input", "Value", ws.Cells(r, 2).Value)
        ' ③ Company Name
        rpaEngine.RunAction "Set-WebXPathTextInput", _
            CreateParams("XPath", "//label[normalize-space(text())='Company Name']/following-sibling::input", "Value", ws.Cells(r, 3).Value)
        ' ④ Role in Company
        rpaEngine.RunAction "Set-WebXPathTextInput", _
            CreateParams("XPath", "//label[normalize-space(text())='Role in Company']/following-sibling::input", "Value", ws.Cells(r, 4).Value)
        ' ⑤ Address
        rpaEngine.RunAction "Set-WebXPathTextInput", _
            CreateParams("XPath", "//label[normalize-space(text())='Address']/following-sibling::input", "Value", ws.Cells(r, 5).Value)
        ' ⑥ Email
        rpaEngine.RunAction "Set-WebXPathTextInput", _
            CreateParams("XPath", "//label[normalize-space(text())='Email']/following-sibling::input", "Value", ws.Cells(r, 6).Value)
        ' ⑦ Phone Number
        rpaEngine.RunAction "Set-WebXPathTextInput", _
            CreateParams("XPath", "//label[normalize-space(text())='Phone Number']/following-sibling::input", "Value", ws.Cells(r, 7).Value)

        ' --- [Submit] ボタンをクリックして次のラウンドへ ---
        ' Submitボタンはタグとtype等で特定可能
        rpaEngine.RunAction "Invoke-WebClick", CreateParams("Selector", "input[type='submit']")
        
    Next r
    
    QueryPerformanceCounter endTime '（終了時間を取得）
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print "★ RPA Challenge: " & Format(elapsedTimeMs, "0.00") & " ミリ秒" & " 若干相違があるが!?"
    
    Application.StatusBar = False
    MsgBox "RPA Challenge 完了！ブラウザのスコアを確認してください。", vbInformation

' ..----------------------------------------------------------------------------
    Debug.Print "--- キャッシュクリア実行 ---"
'''    rpaEngine.RunAction "Clear-WebCache", CreateParams("Mode", "CacheOnly")
'''    DoEvents
    rpaEngine.RunAction "Clear-WebCache", CreateParams("Mode", "All")
    
' --- メモリ解放 ---
CleanUp:
    If Not rpaEngine Is Nothing Then
        rpaEngine.CloseEngine
    End If
    Set rpaEngine = Nothing
    Exit Sub
    
' --- 異常系のハンドリング（エラー捕捉） ---
ErrorHandler:
    If Err.Source = "PS_Engine" Then
        Dim rpaErr As RpaExceptionInfo ' 例外情報を格納する構造体（エラー文字列をパース）
        rpaErr = ParseRpaError(Err.Description)

        Select Case rpaErr.ErrorType
            Case "未発見"
                Debug.Print "【スキップ】要素が見つかりません: " & rpaErr.Details
                '● Resume Next
            Case "Timeout"
                Debug.Print "【待機超過】画面の応答がありません (" & rpaErr.FunctionName & ")"
            Case "JSエラー", "ネイティブエラー", "内部エラー"
                MsgBox "ブラウザ制御内で致命的なエラーが発生しました。" & vbCrLf & _
                       "関数: " & rpaErr.FunctionName & vbCrLf & _
                       "内容: " & rpaErr.Message, vbCritical
            Case Else
                Debug.Print "【その他エラー】" & rpaErr.RawText
        End Select
    Else
        MsgBox "VBAマクロエラー (" & Err.Number & "): " & Err.Description, vbCritical
    End If
    
    On Error Resume Next
    rpaEngine.RunAction "Export-WebScreenshot", CreateParams("Prefix", "ErrorHandler")
    MsgBox "*** エラーが発生しました:" & vbCrLf & Err.Description, vbCritical
    Resume CleanUp
End Sub

' ==============================================================================
' テストデータ： コード内のデータをアクティブシート（現在のシート）へ一括で展開する。
' .. （rpachallenge サイト からデータを取得している。）
' ==============================================================================
' --- 見出しとデータをアクティブシートに展開する処理 ---
Private Sub DeployDataWithHeaderToActiveSheet(ByVal ws As Worksheet)

    ' --- シート全体のフォーマット設定 ---
    With ws.Cells.Font
        .Name = "BIZ UDゴシック"
        .Size = 11
    End With
    
    ' A列?G列（見出しの列）の幅を設定
    ws.Columns("A:G").ColumnWidth = 18
    
    ' --- シートへデータを展開する ---
    ' 1. 見出しデータを配列として定義（1次元配列）
    Dim headers As Variant
    headers = Array("First Name", "Last Name", "Company Name", "Role in Company", "Address", "Email", "Phone Number")
    
    ' 2. 画像のデータ本体（10行 × 7列）
    Dim employees As Variant
    employees = Array( _
        Array("John", "Smith", "IT Solutions", "Analyst", "98 North Road", "jsmith@itsolutions.com", "40716543298"), _
        Array("Jane", "Dorsey", "MediCare", "Medical Engineer", "11 Crown Street", "jdorsey@mc.com", "40791345621"), _
        Array("Albert", "Kipling", "Waterfront", "Accountant", "22 Guild Street", "kipling@waterfront.com", "40735416854"), _
        Array("Michael", "Robertson", "MediCare", "IT Specialist", "17 Farburn Terrace", "mrobertson@mc.com", "40733652145"), _
        Array("Doug", "Derrick", "Timepath Inc.", "Analyst", "99 Shire Oak Road", "dderrick@timepath.com", "40799885412"), _
        Array("Jessie", "Marlowe", "Aperture Inc.", "Scientist", "27 Cheshire Street", "jmarlowe@aperture.us", "40733154268"), _
        Array("Stan", "Hamm", "Sugarwell", "Advisor", "10 Dam Road", "shamm@sugarwell.org", "40712462257"), _
        Array("Michelle", "Norton", "Aperture Inc.", "Scientist", "13 White Rabbit St", "mnorton@aperture.us", "40731254562"), _
        Array("Stacy", "Shelby", "TechDev", "HR Manager", "19 Pineapple Boulev", "sshelby@techdev.com", "40741785214"), _
        Array("Lara", "Palmer", "Timepath Inc.", "Programmer", "87 Orange Street", "lpalmer@timepath.com", "40731653845") _
    )
    
    ' 書き出す位置（アクティブシートのA1セルを起点にする）
    Dim startCell As Range
    Set startCell = ws.Range("A1")
    
    ' --- 見出しの書き出し ---
    ' 1行目のA1セルから右に向かって一括で書き出します
    startCell.Resize(1, UBound(headers) + 1).Value = headers
    
    ' --- データ本体の書き出し ---
    ' 2行目（Offsetで1行下）からデータを入れていきます
    Dim r As Long, c As Long
    For r = 0 To UBound(employees)
        For c = 0 To UBound(employees(r))
            startCell.Offset(r + 1, c).Value = employees(r)(c)
        Next c
    Next r
    
    MsgBox "アクティブシート（" & ws.Name & "）へ見出しとデータを展開しました。", vbInformation
End Sub

