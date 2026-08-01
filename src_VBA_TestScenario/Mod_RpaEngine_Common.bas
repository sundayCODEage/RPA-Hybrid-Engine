Attribute VB_Name = "Mod_RpaEngine_Common"
Option Explicit

' ==============================================================================
' 【高精度タイマー用 API宣言】
' .. Win32 APIを使用してミリ秒未満（マイクロ秒単位）の処理時間を正確に測定する
' ==============================================================================
Public Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' --- モジュールレベル変数 ---
Public freq As Currency             ' CPUの動作周波数 (タイマー精度)
Public rpaEngine As Ps_Engine       ' RPA操作エンジン本体クラス

' --- API宣言 ---
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' --- エンジンパス等の定数定義 ---
Private Const PS_ENGINE_LIB As String = "\Ps_Engine_Core_v204.ps1"

' ==============================================================================
' 【メインエントリーポイント】エンジンの初期化とテストの実行
' ==============================================================================
Public Sub Test_rpaEngine()
    Dim enginePath As String
    Dim sessionId As String
    Dim useCdpPort As Integer   ' 通信モードの設定 (0: 標準, 9222: CDP高速通信)
    
    ' CPUのタイマー周波数を取得 (高精度計測の準備)
    QueryPerformanceFrequency freq
    
    ' --- 実行環境・セッションの初期化 ---
    enginePath = ThisWorkbook.Path & PS_ENGINE_LIB
    sessionId = "SESSION_" & Format(Now, "yyyyMMdd_HHmmss")
'''    useCdpPort = 0      ' CDP高速通信ポート (0指定で標準モード)
    Debug.Print "★ 通常は、標準モード( useCdpPort = 0 ) とするが、今回はテストのため"
    useCdpPort = 9222   ' CDP高速通信ポート (0指定で標準モード)
    
    Set rpaEngine = New Ps_Engine
    
    ' --- PowerShellエンジンを起動 ---
    ' 第3引数: CDPポート番号 (9222)
    ' 第4引数: IsDebugModeFlg (True = 実行時のパラメータや詳細ログをSTDOUTへ出力する)
    If Not rpaEngine.StartEngine(sessionId, enginePath, useCdpPort, True) Then
        MsgBox "RPAエンジンの起動に失敗しました。", vbCritical, "起動エラー"
        Exit Sub
    End If

    ' --------------------------------------------------------------------------
    ' --- テストシナリオの呼び出し ---
    ' --------------------------------------------------------------------------
    ' Excelを強制的に最前面に持ってくる
    Call ForceFocusExcel
    
    Dim prompt As String
    Dim ans As String
    
    prompt = "シナリオを選んでください：" & vbCrLf & _
            "1. 【T1】  基礎・汎用コンポーネント動作テスト" & vbCrLf & _
            "2. 【T2-1】(SYS)：バッチ処理監視＆ポーリング制御" & vbCrLf & _
            "3. 【T2-2】親画面⇒ポップアップ子画面 制御" & vbCrLf & _
            "4. 【T2-3】Fetch APIによるPDFサイレントダウンロード" & vbCrLf & _
            "5. 【T2-4】(Robust)デバッグ証跡出力＆曖昧テキスト解析"
    ans = InputBox(prompt, "テストシナリオ選択")
    
    If ans = "" Then
        MsgBox "キャンセルされました。"
    Else
        MsgBox "選択したテストシナリオ: ( " & ans & " )"
    End If
    
    Select Case ans
        Case 1: Call Test_Chapter1_Master
        Case 2: Call Test_Chapter2_BatchMonitoring
        Case 3: Call Test_Chapter2_PopupWindow
        Case 4: Call Test_Chapter2_PdfFetchDownload
        Case 5: Call Test_Chapter2_RobustDom
        Case Else
            ' *
    End Select

End Sub
