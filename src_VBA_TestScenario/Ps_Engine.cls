VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ps_Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' --- API宣言 ---
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long ' 日跨ぎ防止用API
' --- グローバル定数および変数定義 ---
Private Const PS_TIMEOUT_SEC    As Integer = 30
Private Const PS_READ_RETRY_MS  As Integer = 50 ' CPU負荷と通信レスポンスの最適化を目的とした待機ミリ秒数 (50ms)
Private psProcess   As Object
Private m_IsBusy    As Boolean ' 実行中フラグ

' --- プロパティ ---
Public Property Get IsBusy() As Boolean
    IsBusy = m_IsBusy
End Property

' --- インスタンス破棄時の安全な終了処理 ---
Private Sub Class_Terminate()
    On Error Resume Next
    If Not psProcess Is Nothing Then
        Me.RunAction "Quit", Nothing
        psProcess.Terminate
        Set psProcess = Nothing
        Debug.Print "[PS_Engine] Process Terminated."
    End If
End Sub

' --- PowerShellプロセスの非同期起動および初期化待機 ---
Public Function StartEngine(sessionId As String, enginePath As String, Optional cdpPort As Integer = 0, Optional isDebugMode As Boolean = False) As Boolean
    Dim shellObj As Object
    Dim cmd As String
    Dim line As String
    Dim psDebugFlag As String
    
    If isDebugMode Then psDebugFlag = "true" Else psDebugFlag = "false"
    
    On Error GoTo Catch
    Set shellObj = CreateObject("WScript.Shell")

''' cmd = "powershell.exe -STA                 -NoProfile -ExecutionPolicy Bypass -File """ & enginePath & """ -AppSessionId """ & sessionId & """ -CdpPort " & cdpPort & " -IsDebugModeFlg " & psDebugFlag
    cmd = "powershell.exe -STA -NonInteractive -NoProfile -ExecutionPolicy Bypass -File """ & enginePath & """ -AppSessionId """ & sessionId & """ -CdpPort " & cdpPort & " -IsDebugModeFlg " & psDebugFlag
    Set psProcess = shellObj.exec(cmd)
    
    Do While Not psProcess.StdOut.AtEndOfStream
        line = psProcess.StdOut.ReadLine()
        Debug.Print "PS_Init: " & line
        If InStr(line, "READY") > 0 Then
            StartEngine = True
            Exit Function
        End If
        DoEvents
    Loop

Catch:
    StartEngine = False
End Function

' --- アクションコマンドのJSONペイロード化および送信 ---
Public Function RunAction(cmdName As String, Optional dictParams As Scripting.Dictionary = Nothing) As String
    Dim jsonPayload As String
    If psProcess Is Nothing Then Exit Function
    
    m_IsBusy = True ' 実行中ロック開始
    
    jsonPayload = BuildPayload(cmdName, dictParams)
    RunAction = SendAndReceive(jsonPayload)
    
    m_IsBusy = False ' 実行中ロック解除
End Function

' --- コマンドおよびパラメータのJSONペイロード構築 ---
Private Function BuildPayload(cmdName As String, dictParams As Scripting.Dictionary) As String
    Dim root As New Scripting.Dictionary
    Dim settings As New Scripting.Dictionary
    
    root.Add "Command", cmdName
    If dictParams Is Nothing Then
        root.Add "Parameters", New Scripting.Dictionary
    Else
        root.Add "Parameters", dictParams
    End If
    
    ' VBA-JSONによる安全なJSON文字列への自動変換
    BuildPayload = JsonConverter.ConvertToJson(root)
End Function

' --- 標準入出力パイプを介したJSONの送受信およびタイムアウト監視 ---
Private Function SendAndReceive(jsonStr As String) As String
    Dim line As String
    Dim capturedResult As String
    Dim startTick As Long
    
    Debug.Print "VBA -> PS: " & jsonStr
    
    ' PowerShellへのコマンド送信
    psProcess.StdIn.WriteLine jsonStr
    
    ' 監視開始時刻の記録 (GetTickCountを使用)
    startTick = GetTickCount()
    
    ' 応答待ちループ
    Do While True
        ' 1. プロセスの生存確認（クラッシュ検知）
        If psProcess.Status <> 0 Then
            m_IsBusy = False
            Err.Raise 998, "PS_Engine", "PowerShellプロセスが予期せず終了"
            Exit Function
        End If

        ' 2. タイムアウト判定 (ミリ秒単位で比較、日跨ぎバグを回避)
        If (GetTickCount() - startTick) > (PS_TIMEOUT_SEC * 1000&) Then
            Me.CloseEngine
            m_IsBusy = False
            Err.Raise 997, "PS_Engine", "PowerShell応答タイムアウト (" & PS_TIMEOUT_SEC & "秒)"
            Exit Function
        End If

        ' 3. 標準出力の読み取りと内容解析
        If Not psProcess.StdOut.AtEndOfStream Then
            line = psProcess.StdOut.ReadLine()
            
            If Left(line, 8) = "[RESULT]" Then
                capturedResult = Mid(line, 9)
            ElseIf Left(line, 9) = "[SUCCESS]" Then
                SendAndReceive = capturedResult
                Exit Function
            ElseIf Left(line, 7) = "[ERROR]" Then
                m_IsBusy = False
                Err.Raise 999, "PS_Engine", Trim(Mid(line, 8))
                Exit Function
            Else
                Debug.Print "PS_STDOUT: " & line
            End If
        End If
        
        DoEvents
        Sleep PS_READ_RETRY_MS
    Loop
End Function

' --- PowerShellエンジンの明示的なシャットダウン処理 ---
Public Sub CloseEngine()
    On Error Resume Next
    If Not psProcess Is Nothing Then
        psProcess.StdIn.WriteLine "{""Command"":""Quit""}"
        Sleep 500
        psProcess.Terminate
        Set psProcess = Nothing
        Debug.Print "[PS_Engine] Engine Closed explicitly."
    End If
End Sub
