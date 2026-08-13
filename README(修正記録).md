_修正記録（1回目）　2026/08/13_

2026/08/02にコード掲載しました。その後、開発テストに使っていた公開サイトでのテスト用コードを再整理する中で、一部バグに気づきました。<br>
こんなレガシーRPAコードを利用される方は ほぼ皆無と思っていますが、公開してしまった反省と、私の忘却禄として整理させていただきます。（QIITAサイトでも、コード紹介をしており、お二人から“いいね”頂きました。ごめんなさい。）<br>
今回の「公開サイトでのテストコード」は、PC内にグチャグチャに作ったテストコードを、ほんと私の忘却用に再構成しているので、コードセンスは悪いですが時間はかかりました。<br>

#### 1　モジュール　について
#### (1)　Lib-WebAction_v204　の変更について ( _v205 )
・「function Wait-WebScreenUnlock { 」（画面のローディングマスク解除（非表示）待機）は、私のコード整理モレで、古いコードを貼っていた様です。<br>
・画面全体の読み込みステータス完了を待機、「function Wait-WebDocumentReady { 」を修正しました。<br>
・指定要素の存在確認(例外を投げず True/False の文字列）、「function Test-WebElement { 」を追加しました。<br>
・Web要素の属性値(Attribute)を取得する、「function Get-WebAttribute { 」を追加しました。<br>

<details>
  <summary><i>　（一度は、確認した！ 出来てたつもり。）</i></summary>

```vba
言い訳： 単体コードの再確認をミスっていました。

' --- クリック＆ダイアログ/Ajax（部分更新）待機 ---
Private Sub Click_AndWaitDialog(ByVal xpathStr As String, Optional ByVal waitMs As Long = 500)
    rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", xpathStr)
    rpaEngine.RunAction "Wait-WebDocumentReady", CreateParams("TimeoutSec", 10)
    Call xxx_MaskUnlock
    rpaEngine.RunAction "Wait-WebScreenUnlock", CreateParams("TimeoutSec", c_TimeoutSec)
    If waitMs > 0 Then Sleep waitMs
End Sub

' --- (xxxシステム)マスク解除待機 ---
Private Sub xxx_MaskUnlock()
    Dim maskXPath As String
    maskXPath = "//*[@id='opmask' or @id='opFreezePane' or @id='mask']"
    rpaEngine.RunAction "Wait-WebXPathElementDisappear", CreateParams("XPath", maskXPath, "TimeoutSec", c_TimeoutSec)
End Sub

と、xxx_MaskUnlock　をメインに使っていたため
```
</details>

#### (2)　 Lib-DesktopUIA_v103　の変更について ( _v104 )
```text
1　UIA要素検索の「動的追従」によるフリーズ防止
　 存在しない要素を探す際、OS全体（RootElement）を検索してしまい、タイムアウトまで数十秒間プロセスがフリーズする問題がありました。
2　ターゲット要素が ValuePattern を持たない場合、自動的に内部の Edit 子要素を探索して値を読み書きするロジックを追加
3　Pattern操作がサポートされていない場合、エラーで落とさずに自動的に Safeモード（物理クリック・物理キー入力）へ切り替え
など
```

#### (3)　 Ps_Engine_Core_v204　について
```text
# 常時ロード対象モジュールの定義　／　（250行付近）
$alwaysLoadLibs = @(
    "Lib-WebAction_v204.ps1",　->　** _v205 ** に　修正しました。
    "Lib-WebXPath_v101.ps1",
    "Lib-DesktopUIA_v103.ps1",　->　** _v104 ** に　修正しました。
    ､､､
```

#### 2　 sandbox/　について
* screenshot_grid(プレビューカード)　／ サンドボックス画面の一覧画面を作成した時のコード<br>
* （公開サイトでは、テストしずらい箇所を再構成しています。）<br>
・ 10_shadow_dom　／ Shadow DOM 操作テスト<br>
・ 11_overlay_mask　／ 画面オーバーレイ（マスク）待機テスト<br>

#### 3　sample_BOX/　について
* （公開サイト用テストコード）<br>
・ Mod_TestRun1_Base一般的な.bas　／ 実行: Test_Run1_Base<br>
・ Mod_TestRun2_Base少し高度.bas　／ 実行: Test_Run2_Base<br>

```text
├📊 sample_rpa_test.xlsm    # VBA ( テストシナリオ )
 ├── Ps_Engine.cls                  ( プロセス通信とAPI実行を担うRPAエンジンのコアクラス )
 ├── JsonConverter                  ( VBA-JSON ) VBAでのJSON解析
 ├── Ps_Bridge.bas                  ( JSONパース・エラー変換などのVBA側ユーティリティ )
 ├── Mod_RpaEngine_Common.bas       ( エンジンの初期化とテストの実行司令塔 )
 ├── Mod_Chapter1_Basics.bas        ( テストシナリオ：第1章 基礎操作 )
 └── Mod_Chapter2_Business.bas      ( テストシナリオ：第2章 業務システム・応用操作 )

この中に、モジュール追加してください。
```

#### 4　docs/　について
・ 総合テスト網羅・関数一覧表 （公開サイト用テストコード）

<br>

_< 記録 >_

| x1 | x2 | 登録・更新 | その他 |
| :--- | :---: | :--- | :--- |
| Mod_RPA_Challengeデータ同梱.bas | 実行: Test_RPAchallenge | 2026/08/04 登録 | RPAチャレンジ（たぶん9秒台で完走）<br>アクティブシートに、Array保存のデータを展開可 |
| :--- | :---: | 2026/08/12 更新 | Public ⇒ Private へ修正<br>sample_rpa_test.xlsm 内のモジュール格納への修正 |
| Mod_TestRun1_Base一般的な.bas | 実行: Test_Run1_Base | 2026/08/13 登録 | :--- |
| Mod_TestRun2_Base少し高度.bas | 実行: Test_Run2_Base | 2026/08/13 登録 | :--- |


