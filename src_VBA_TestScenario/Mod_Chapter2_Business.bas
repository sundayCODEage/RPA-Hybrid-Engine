Attribute VB_Name = "Mod_Chapter2_Business"
Option Explicit

' ==============================================================================
' 【T2-1】（業務システム）：バッチ処理監視 ＆ ポーリング制御
' .. （マスク解除待機・疑似Ajaxテーブル解析・ポーリングループの実装パターン）
' ==============================================================================

' --- (業務)システム用 定数定義 ---
Private Const c_TimeoutSec As Integer = 30
Private Const c_帳票状態 As Integer = 6         ' テーブルの「帳票」列インデックス (0始まり)
Private Const c_処理番号 As Integer = 7         ' テーブルの「処理番号」列インデックス (0始まり)
Private Const c_処理番号_Add As Long = 900000   ' 正常終了（帳票○）時に付加する判定用オフセット値

' モジュール変数 (ループ回数カウンタ)
Private loopCount As Integer

' ==============================================================================
' T2-1 バッチ処理監視 ＆ ポーリング実行
' ==============================================================================
Public Sub Test_Chapter2_BatchMonitoring()
    On Error GoTo ErrorHandler
    
    Dim sandboxPath As String
    Dim targetUrl As String
    Dim xpathStr As String
    
    Dim rawData As String
    Dim getJobId As Long
    ' 1回目と2回目の検索結果を保持する配列
    Dim arrJobId() As Long
    ReDim arrJobId(1 To 2)
    
    ' サンドボックスの絶対パスを取得
    sandboxPath = Replace(ThisWorkbook.Path & "\Sandbox\", "\", "/")
    targetUrl = "file:///" & sandboxPath & "05_business_batch.html"
    
    Debug.Print "=========================================================="
    Debug.Print "--- [T2-1] 05_business_batch.html テスト開始 ---"
    Debug.Print "=========================================================="
    
    ' 1. テスト画面への遷移と読み込み待機
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    
    ' 初期表示時のマスク（もしあれば）を解除待機
    Call IPK_MaskUnlock
    ' 2. 対象年度の入力（テスト画面のID指定入力）
    xpathStr = "//input[@id='selectedLabel_IP325210SC~nenNendo_08']"
    rpaEngine.RunAction "Set-WebXPathTextInput", CreateParams("XPath", xpathStr, "Value", "令和08年度")
    ' 3. F11: バッチ処理の起動
    Debug.Print "--- F11: バッチ処理を起動します ---"
    xpathStr = "//li[@id='FUNC11']"
    rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", xpathStr)
    
    Debug.Print "=========================================================="
    Debug.Print "★ [バッチ処理を起動しました。F4キーで状態を更新してください。] OKを押下"
    Debug.Print "=========================================================="

    ' 起動後のマスク解除を待機
    Call IPK_MaskUnlock
      
    ' --------------------------------------------------------------------------
    ' 【ポーリング監視】バッチ完了まで F4検索 を繰り返し実行するループ
    ' --------------------------------------------------------------------------
    loopCount = 0
    
    ' --- 検索 1回目：処理中の状態を確認する ---
    Debug.Print "--- 検索1回目（処理中ステータスの確認） ---"
    Do Until arrJobId(1) > 0
        Call Wait_F4検索_Loop1(rawData, getJobId)
        arrJobId(1) = getJobId
    Loop
    Debug.Print "  -> [1回目判定結果] 処理番号: " & arrJobId(1) & " (まだ処理中)"
    
    ' --- 検索 2回目：バッチが「正常終了」して900000が加算されるのを待つ ---
    Debug.Print "--- 検索2回目（バッチ完了までポーリング監視） ---"
    Do Until arrJobId(2) > 0
        Application.Wait Now + TimeValue("00:00:02") ' 2秒間隔でポーリング
        Call Wait_F4検索_Loop1(rawData, getJobId)
        
        ' 1回目の処理番号(100001) に 900000 を足した値(900001) に変化したか検証
        If arrJobId(1) < c_処理番号_Add And (arrJobId(1) + c_処理番号_Add) = getJobId Then
            ' OK: 1回目は「処理中」だったが、今回「正常終了」へ変化した
            arrJobId(2) = getJobId
        ElseIf arrJobId(1) >= c_処理番号_Add And arrJobId(1) = getJobId Then
            ' OK: 最初から「正常終了」していたケース
            arrJobId(2) = getJobId
        Else
            ' まだ処理中のためループ継続
            Debug.Print "  -> [監視中...] 処理中または完了待ちです (得られた値: " & getJobId & ")"
            getJobId = 0
        End If
    Loop
    
    ' 実際の元の処理番号（100001）を復元算出
    getJobId = arrJobId(2) - c_処理番号_Add
    
    Debug.Print "=========================================================="
    Debug.Print "★ [バッチ完了] 処理番号: " & getJobId & " のバッチ処理が正常終了しました。"
    Debug.Print "=========================================================="

    MsgBox "【T2-1 テスト】" & vbCrLf & _
           "　バッチ完了のポーリング監視とマスク解除待機が正常に動作しました。" & vbCrLf & _
           "　確定処理番号: " & getJobId, vbInformation

    ' --------------------------------------------------------------------------
    ' 終了クリーンアップ
    rpaEngine.RunAction "Clear-WebCache", CreateParams("Mode", "All")
    
' --- 正常終了 / エラー共通のメモリ解放領域 ---
CleanUp:
    If Not rpaEngine Is Nothing Then
        rpaEngine.CloseEngine
    End If
    Set rpaEngine = Nothing
    Exit Sub

' --- テスト用（簡易） ---
ErrorHandler:
    MsgBox "【エラー発生】" & vbCrLf & _
           "関数: Test_Chapter2_BatchMonitoring" & vbCrLf & _
           "内容: " & Err.Description, vbCritical, "処理中断"
End Sub

' ==============================================================================
' 【サブ関数 1】IPK_MaskUnlock：画面オーバーレイ（#opmask）の消滅待機
' .. （業務システム特有の「画面操作不可マスク」が消えるのを待機）
' ==============================================================================
Private Sub IPK_MaskUnlock()
    Dim maskXPath As String
    ' 業務システムでよく見られるマスク要素のID候補をOR指定
    maskXPath = "//*[@id='opmask' or @id='opFreezePane' or @id='mask']"
    ' 指定要素が画面上から消失(非表示)するまで自動待機
    rpaEngine.RunAction "Wait-WebXPathElementDisappear", CreateParams("XPath", maskXPath, "TimeoutSec", c_TimeoutSec)
End Sub

' ==============================================================================
' 【サブ関数 2】Wait_F4検索_Loop1：F4検索実行 ＆ テーブル状態解析
' .. （F4キー（検索ボタン）を押し、テーブルデータをCSV取得して状態をパース）
' ==============================================================================
Private Sub Wait_F4検索_Loop1(ByRef rawData As String, ByRef getJobId As Long)
    Dim rows() As String
    Dim cols() As String
    Dim xpathStr As String
    
    ' 1. F4: 検索ボタンのクリック
    xpathStr = "//li[@id='FUNC4']"
    rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", xpathStr)
    ' 2. 検索実行時の操作不可マスクの消滅を待つ
    Call IPK_MaskUnlock
    ' 3. 動的テーブルデータの取得 (FileName="" でメモリ返却)
    xpathStr = "#IP100710G-variable-body-table"
    rawData = rpaEngine.RunAction("Export-WebTableToCsv", CreateParams("Selector", xpathStr, "FileName", ""))
    ' 取得失敗またはタイムアウト判定
    loopCount = loopCount + 1
    If rawData = "" Or rawData = "NOT_FOUND" Or rawData = "EMPTY_TABLE" Or loopCount > 30 Then
        Err.Raise 999, "Wait_F4検索", "テーブルデータの取得に失敗したか、検索がタイムアウトしました (試行回数: " & loopCount & ")"
    End If
    ' 4. 行・列の分解パース処理
    rows = Split(rawData, "<R>")
    ' rows(0)はヘッダー行のため、最新データである rows(1) を解析対象とする
    cols = Split(rows(1), "<T>")
    getJobId = 0
    ' 5. データの数値判定と「正常終了（帳票○）」フラグの加算
    If IsNumeric(cols(c_処理番号)) Then
        getJobId = CLng(cols(c_処理番号))
        ' 帳票列が「○」になっている（バッチ完成）場合は 900000 を加算して完了マークとする
        If cols(c_帳票状態) = "○" Then
            getJobId = getJobId + c_処理番号_Add
        End If
    End If
    
    ' デバッグログ出力
    Debug.Print "  -> [F4検索結果] 行数: " & UBound(rows) & " | 処理番号RAW: " & cols(c_処理番号) & " | 状態: " & cols(c_帳票状態) & " -> 判定値: " & getJobId
End Sub

' ==============================================================================
' 【サブ関数 3】Click_AndWaitDialog：クリック ＆ Ajax部分更新待機
' .. （ダイアログ表示やAjax通信を伴うボタン押下時の共通化ラッパー関数）
' ==============================================================================
Private Sub Click_AndWaitDialog(ByVal xpathStr As String, Optional ByVal waitMs As Long = 500)
    ' 1. 要素のクリック
    rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", xpathStr)
    ' 2. DOMの準備状態確認 (準備完了まで最大10秒待機)
    rpaEngine.RunAction "Wait-WebDocumentReady", CreateParams("TimeoutSec", 10)
    ' 3. オペレーションマスクの解除待機
    Call IPK_MaskUnlock
    ' 4. 必要に応じた調整スリープ
    If waitMs > 0 Then Sleep waitMs
End Sub


' ==============================================================================
' 【T2-2】親画面 ⇒ ポップアップ子画面 制御
' .. （別ウィンドウ/タブの切り替え(Switch-TabByTitle)・値選択・復帰検証）
' ==============================================================================
Public Sub Test_Chapter2_PopupWindow()
    On Error GoTo ErrorHandler
    
    Dim sandboxPath As String
    Dim targetUrl As String
    Dim xpathStr As String
    Dim resCode As String
    Dim resName As String
    
    ' サンドボックスの絶対パスを取得
    sandboxPath = Replace(ThisWorkbook.Path & "\Sandbox\", "\", "/")
    targetUrl = "file:///" & sandboxPath & "06_popup_parent.html"
    
    Debug.Print "=========================================================="
    Debug.Print "--- [T2-2] 06_popup_parent.html テスト開始 ---"
    Debug.Print "=========================================================="

    ' 1. テスト画面への遷移と読み込み待機
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    
    ' 2. F2ボタンを押して子画面(ポップアップ)を起動
    Debug.Print "--- [親画面] F2: 部門検索(子画面) を呼び出します ---"
    xpathStr = "//li[@id='FUNC2']"
    rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", xpathStr)
    ' ポップアップウィンドウが描画されるのを少し待機
    Application.Wait Now + TimeValue("00:00:01")
    
    ' 3. 子画面へアクティブタブ/ウィンドウを切り替え
    ' （タイトル文字列の一部分を指定してターゲットを切替制御）
    Debug.Print "--- [タブ切替] 子画面（部門検索）へフォーカスを移動 ---"
    rpaEngine.RunAction "Switch-TabByTitle", CreateParams("TitleSubstring", "部門検索")
    
    ' --------------------------------------------------------------------------
    ' 4. 子画面での操作（絞り込み ＆ データ選択）
    ' --------------------------------------------------------------------------
    ' 4-1. 検索キーワード入力（「財政」で絞り込み）
    xpathStr = "//input[@id='searchKeyword']"
    rpaEngine.RunAction "Set-WebXPathTextInput", CreateParams("XPath", xpathStr, "Value", "財政")
    ' 4-2. 検索ボタンクリック
    xpathStr = "//button[@id='btnSearch']"
    rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", xpathStr)
    ' 4-3. 財政課 (D02) の「選択」ボタンをクリック
    ' .. (クリックするとJSで親画面に値が渡され、子画面は自動的に window.close)
    Debug.Print "--- [子画面] 財政課(D02) を選択して確定 ---"
    xpathStr = "//button[@id='btnSelect_D02']"
    rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", xpathStr)
    ' 子画面クローズと親画面への反映処理を待機
    Application.Wait Now + TimeValue("00:00:01")
    
    ' --------------------------------------------------------------------------
    ' 5. 親画面へフォーカスを復帰
    ' --------------------------------------------------------------------------
    Debug.Print "--- [タブ切替] 親画面（伝票入力）へフォーカスを復帰 ---"
    rpaEngine.RunAction "Switch-TabByTitle", CreateParams("TitleSubstring", "伝票入力")
    ' 6. 親画面へ値が正常に引き継がれたか検証）
    ' .. （JS経由で入力欄 (deptCode / deptName) の現在値を取得
    resCode = rpaEngine.RunAction("Invoke-WebScript", CreateParams("Js", "return document.getElementById('deptCode').value;"))
    resName = rpaEngine.RunAction("Invoke-WebScript", CreateParams("Js", "return document.getElementById('deptName').value;"))
    
    Debug.Print "=========================================================="
    Debug.Print "★ [反映結果確認] 部門コード: " & resCode & " | 部門名称: " & resName
    Debug.Print "=========================================================="
    
    ' 結果判定
    If resCode = "D02" And resName = "財政課" Then
        MsgBox "【T2-2 テスト】" & vbCrLf & _
               "　ポップアップ子画面の操作と親画面への値反映が正常に行われました。" & vbCrLf & _
               "　取得結果: " & resCode & " / " & resName, vbInformation
    Else
        Err.Raise 999, "Test_PopupWindow", "親画面への値引き継ぎに失敗しました (取得値: " & resCode & " / " & resName & ")"
    End If

    ' --------------------------------------------------------------------------
    ' 終了クリーンアップ
    rpaEngine.RunAction "Clear-WebCache", CreateParams("Mode", "All")
    
' --- 正常終了 / エラー共通のメモリ解放領域 ---
CleanUp:
    If Not rpaEngine Is Nothing Then
        rpaEngine.CloseEngine
    End If
    Set rpaEngine = Nothing
    Exit Sub
    
' --- テスト用（簡易） ---
ErrorHandler:
    MsgBox "【エラー発生】" & vbCrLf & _
           "関数: Test_Chapter2_PopupWindow" & vbCrLf & _
           "内容: " & Err.Description, vbCritical, "処理中断"
End Sub

' ==============================================================================
' 【T2-3】Fetch APIによるPDFサイレントダウンロード
' .. （EdgeネイティブPDFビューアのUI操作を諦め、Fetch APIで裏側から直接奪取）
' ==============================================================================
Public Sub Test_Chapter2_PdfFetchDownload()
    On Error GoTo ErrorHandler
    
    Dim sandboxPath As String
    Dim targetUrl As String
    Dim saveDir As String
    Dim saveFileName As String
    Dim fullSavePath As String
    Dim actualPdfUrl As String
    Dim getPdfUrlJs As String
    
    ' タイマー用変数
    Dim startTime As Currency
    Dim endTime As Currency
    Dim elapsedTimeMs As Double

    ' サンドボックスの絶対パスを取得
    sandboxPath = Replace(ThisWorkbook.Path & "\Sandbox\", "\", "/")
    targetUrl = "file:///" & sandboxPath & "07_pdf_embed.html"
    
    saveDir = ThisWorkbook.Path & "\RPA_Downloads"
    saveFileName = "Ch2_FetchDownloaded.pdf"
    fullSavePath = saveDir & "\" & saveFileName
    
    ' 画面プレビュー表示用のダミーPDFを Sandbox フォルダ内に配置生成
    Call CreateDummyPdfFile(ThisWorkbook.Path & "\Sandbox\sample_report.pdf")
    
    ' 過去の保存済みテストファイルがあれば削除
    If Dir(fullSavePath) <> "" Then Kill fullSavePath
    
    Debug.Print "=========================================================="
    Debug.Print "--- [T2-3] 07_pdf_embed.html (Fetch DL) テスト開始 ---"
    Debug.Print "=========================================================="
    
    ' 1. ダウンロードの事前予約 (ダイアログ抑止 ＆ 保存先固定)
    rpaEngine.RunAction "Enable-SilentDownload", CreateParams("DownloadDirectory", saveDir, "FileName", saveFileName)
    
    ' 2. PDF画面へ遷移 ＆ embed要素の出現待機
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    ' ** このテスト **では、
    ' .. type属性がないため、単純に "//*[@id='pdfarea']//embed" または "//embed" で捕捉する
    rpaEngine.RunAction "Wait-WebXPathElement", CreateParams("XPath", "//embed", "TimeoutSec", 15)

    ' 3. 埋め込みPDF(embed)の絶対URLを安全に取得する。
    actualPdfUrl = rpaEngine.RunAction("Get-WebEmbedPdfUrl")
    Debug.Print "  -> [解析成功] 抽出した絶対PDF-URL: " & actualPdfUrl
     
    QueryPerformanceCounter startTime ' 処理開始時刻の記録
    ' 4. Fetch APIで裏側からのサイレントダウンロードを発火
    ' .. （画面のボタン（UI）を押さず、セッション・Cookieを保持したままバイナリを裏で取得する）
    Debug.Print "--- Fetch API による自己ダウンロードイベントを発火中 ---"
    rpaEngine.RunAction "Invoke-WebFetchDownload", CreateParams("TargetUrl", actualPdfUrl)
    ' 5. 一時ファイル(.crdownload)の消失とロック解除を監視待機
    rpaEngine.RunAction "Wait-FileDownload", CreateParams("FilePath", fullSavePath, "TimeoutSec", 30)
    QueryPerformanceCounter endTime ' 処理終了時刻の記録
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print "★ DL所要時間: " & Format(elapsedTimeMs, "0.00") & " ミリ秒"
    
    Debug.Print "=========================================================="
    Debug.Print "★ [Fetch DL] サイレント保存が完了しました。"
    Debug.Print "　　保存先: " & fullSavePath
    Debug.Print "=========================================================="
    
    ' --------------------------------------------------------------------------
    ' 6. 後処理
    ' ** 通常処理の例（パターンA: 同じタブ内で遷移したPDF画面から「戻る」）
    ' JS経由でブラウザの履歴を1つ戻す
'   rpaEngine.RunAction "Invoke-WebScript", CreateParams("Js", "window.history.back();")
    ' 元の画面の読み込み完了を待機
'   rpaEngine.RunAction "Wait-WebPageLoad"
    ' ** 通常処理の例（パターンB: 別ウィンドウで開いたPDF画面を閉じる）
    ' JS経由で閉じる命令を発行
'   rpaEngine.RunAction "Invoke-WebScript", CreateParams("Js", "window.close();")
    ' （エンジン側でタブ破棄と親画面へのフォーカス復帰が行われるため、少し待機）
'   Sleep 1000
    ' --------------------------------------------------------------------------

    MsgBox "【T2-3 テスト】" & vbCrLf & _
           "　Fetch APIによるPDFサイレント保存が完了しました。" & vbCrLf & _
           "　保存ファイル: " & saveFileName, vbInformation

    ' --------------------------------------------------------------------------
    ' 終了クリーンアップ
    rpaEngine.RunAction "Clear-WebCache", CreateParams("Mode", "All")

' --- 正常終了 / エラー共通のメモリ解放領域 ---
CleanUp:
    If Not rpaEngine Is Nothing Then
        rpaEngine.CloseEngine
    End If
    Set rpaEngine = Nothing
    Exit Sub

' --- テスト用（簡易） ---
ErrorHandler:
    MsgBox "【エラー発生】" & vbCrLf & _
           "関数: Test_Chapter2_PdfFetchDownload" & vbCrLf & _
           "内容: " & Err.Description, vbCritical, "処理中断"
End Sub

' ==============================================================================
' 【テスト文字】 CreateDummyPdfFile
' .. （テスト画面・ダウンロードファイルで目視確認できる文字(RPA TEST SUCCESS)付きPDFを出力）
' ==============================================================================
Private Sub CreateDummyPdfFile(ByVal pdfPath As String)
    ' 既存の古い白紙ファイルがあれば削除して再生成
    If Dir(pdfPath) <> "" Then Kill pdfPath
    
    Dim pdfContent As String
    ' テキスト描画ストリーム(BT ... ET)を含む正当なPDF1.4データ
    pdfContent = "%PDF-1.4" & vbCrLf & _
                 "1 0 obj <</Type /Catalog /Pages 2 0 R>> endobj" & vbCrLf & _
                 "2 0 obj <</Type /Pages /Kids [3 0 R] /Count 1>> endobj" & vbCrLf & _
                 "3 0 obj <</Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources <</Font <</F1 4 0 R>>>> /Contents 5 0 R>> endobj" & vbCrLf & _
                 "4 0 obj <</Type /Font /Subtype /Type1 /BaseFont /Helvetica>> endobj" & vbCrLf & _
                 "5 0 obj <</Length 62>> stream" & vbCrLf & _
                 "BT" & vbCrLf & _
                 "/F1 24 Tf" & vbCrLf & _
                 "100 700 Td" & vbCrLf & _
                 "(RPA TEST SUCCESS - FETCH DL OK) Tj" & vbCrLf & _
                 "ET" & vbCrLf & _
                 "endstream" & vbCrLf & _
                 "endobj" & vbCrLf & _
                 "xref" & vbCrLf & _
                 "0 6" & vbCrLf & _
                 "0000000000 65535 f " & vbCrLf & _
                 "0000000009 00000 n " & vbCrLf & _
                 "0000000058 00000 n " & vbCrLf & _
                 "00000000115 00000 n " & vbCrLf & _
                 "00000000244 00000 n " & vbCrLf & _
                 "00000000313 00000 n " & vbCrLf & _
                 "trailer <</Size 6 /Root 1 0 R>>" & vbCrLf & _
                 "startxref" & vbCrLf & _
                 "426" & vbCrLf & _
                 "%%EOF"
    
    Dim fNo As Integer
    fNo = FreeFile
    Open pdfPath For Output As #fNo
    Print #fNo, pdfContent;
    Close #fNo
End Sub


' ==============================================================================
' 【T2-4】（Robust DOM Utilities）デバッグ証跡出力 ＆ 曖昧テキスト自動解析テスト
' ..
' ==============================================================================
Public Sub Test_Chapter2_RobustDom()
    On Error GoTo ErrorHandler
    
    Dim sandboxPath As String
    Dim targetUrl As String
    
    ' サンドボックスの絶対パスを取得
    sandboxPath = Replace(ThisWorkbook.Path & "\Sandbox\", "\", "/")
    targetUrl = "file:///" & sandboxPath & "08_robust_dom.html"
    
    Debug.Print "=========================================================="
    Debug.Print "--- [T2-4] 08_robust_dom.html テスト開始 ---"
    Debug.Print "=========================================================="
    
    ' 1. テスト画面への遷移と読み込み待機
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    
    ' 2. 手動入力されたと見立てて「ニュース」を安全クリック
    Call Test_FindAndSafeClick("ニュース")
    ' 3. 手動入力されたと見立てて「検索」を安全クリック
    Application.Wait Now + TimeValue("00:00:01")
    Call Test_FindAndSafeClick("検索")
    ' 4. 現状の画面証跡を一括エクスポート
    Call Export_DebugInfo_All("Robust DOM 自動テスト完了時の全証跡")
    
    ' --------------------------------------------------------------------------
    ' 終了クリーンアップ
    rpaEngine.RunAction "Clear-WebCache", CreateParams("Mode", "All")

' --- 正常終了 / エラー共通のメモリ解放領域 ---
CleanUp:
    If Not rpaEngine Is Nothing Then
        rpaEngine.CloseEngine
    End If
    Set rpaEngine = Nothing
    Exit Sub

' --- テスト用（簡易） ---
ErrorHandler:
    MsgBox "【エラー発生】" & vbCrLf & _
           "関数: Test_Chapter2_PdfFetchDownload" & vbCrLf & _
           "内容: " & Err.Description, vbCritical, "処理中断"

End Sub

' ==============================================================================
' 【機能①】 画面情報の全デバッグ証跡を一括エクスポート (黄色注記用コメント対応)
' ==============================================================================
Public Sub Export_DebugInfo_All(Optional ByVal userComment As String = "手動操作後の画面解析")
    On Error GoTo ErrorHandler
    
    Dim timeStamp As String
    timeStamp = Format(Now, "yyyyMMdd_HHmmss")
    
    Debug.Print "=========================================================="
    Debug.Print "--- デバッグ証跡エクスポート開始 [" & userComment & "] ---"
    Debug.Print "=========================================================="
    
    ' 1. HTMLスナップショットの保存
    rpaEngine.RunAction "Export-WebHtml", CreateParams("FileName", "WebHtml_" & timeStamp & ".html")
    ' 2. 画面のPNGスクリーンショット保存
    rpaEngine.RunAction "Export-WebScreenshot", CreateParams("FileName", "WebScreenshot_" & timeStamp & ".png")
    ' 3. 画面内操作可能要素の属性CSV化
    rpaEngine.RunAction "Export-WebElementsToCsv", CreateParams("FileName", "WebElements_" & timeStamp & ".csv")
    ' 4. ネスト構造(FrameTree)のCSV化
    rpaEngine.RunAction "Export-WebFrameTreeToCsv", CreateParams("FileName", "WebFrameTree_" & timeStamp & ".csv")
    ' 5. OS上のウィンドウ階層(WindowHierarchy)のCSV化
    rpaEngine.RunAction "Export-WindowHierarchyToCsv", CreateParams("FileName", "WindowHierarchy_" & timeStamp & ".csv")
    
    Debug.Print "  -> [完了] デバッグ・証跡ファイルを Logs フォルダへ出力しました。"
    MsgBox "画面情報の全デバッグ証跡をエクスポートしました！" & vbCrLf & _
           "　コメント: " & userComment, vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "証跡エクスポート中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' ==============================================================================
' 【機能②】 画面に見えている文字から「最適CSSセレクタ」を自動生成し、安全クリック
' ==============================================================================
Public Sub Test_FindAndSafeClick(ByVal visibleText As String)
    On Error GoTo ErrorHandler
    
    Dim fuzzyXPath As String
    Dim cssSelector As String
    
    If Trim(visibleText) = "" Then
        MsgBox "ターゲット文字を入力してください (例: ニュース, 検索)", vbExclamation
        Exit Sub
    End If
    
    Debug.Print "=========================================================="
    Debug.Print "--- [Robust DOM] テキストターゲット探索: 『" & visibleText & "』 ---"
    Debug.Print "=========================================================="
    
    ' 1. VBA側から曖昧XPathを生成 (aタグ、buttonタグ、inputタグ等を包括)
    fuzzyXPath = "//*[self::a or self::button or self::input or @role='button'][contains(., '" & visibleText & "') or @value='" & visibleText & "']"
    ' 2. フェーズ1: 内部の可視性エンジン(isVisible)で隠し要素を弾き、最適CSSセレクタを逆生成
    Debug.Print "--- フェーズ1: Get-WebCssSelectorHint による可視判定 ＆ セレクタ逆生成 ---"
    cssSelector = rpaEngine.RunAction("Get-WebCssSelectorHint", CreateParams("XPath", fuzzyXPath))
    If cssSelector = "" Or cssSelector = "NOT_FOUND" Then
        Debug.Print "  -> [失敗] 画面上で可視状態の『" & visibleText & "』要素が見つかりませんでした。"
        MsgBox "画面上に見えている『" & visibleText & "』要素を発見できませんでした。", vbExclamation, "探索失敗"
        Exit Sub
    End If
    Debug.Print "  -> ★ [解析成功] 生成された最適CSSセレクタ: " & cssSelector
    ' 3. フェーズ2: 生成された一意のセレクタを使って scrollIntoView ＆ 重なり判定付き安全クリック
    Debug.Print "--- フェーズ2: Invoke-WebSafeClick によるスクロール ＆ 重なりチェック付き安全クリック ---"
    rpaEngine.RunAction "Invoke-WebSafeClick", CreateParams("Selector", cssSelector, "TimeoutSec", 10)
    
    Debug.Print "★ [狙撃成功] 隠し要素を回避し、目的の『" & visibleText & "』を安全にクリックしました！"
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

