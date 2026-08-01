## 汎用RPA操作エンジン テスト仕様・合格判定基準

### 1. 位置づけ
本RPAエンジン（VBA × PowerShell × WebView2 ハイブリッド制御基盤）の機能網羅性を確認し、今後の機能追加やバージョンアップ時に既存機能が破損していないかを確認する「リグレッションテスト（回帰テスト）」ものです。

同梱の `sandbox/` フォルダ内にあるローカルHTML群と、`sample_rpa_test.xlsm` を組み合わせることで、以下の全シナリオを検証可能です。

### 2. テストシナリオ検証・合格判定基準表

#### 2.1 システム制御・ライフサイクル
| テストID | コマンド / 関数名 | パラメータ例・主な仕様 | 検証対象シナリオ | 期待される動作・合格判定基準 |
| :--- | :--- | :--- | :--- | :--- |
| SYS-01 | StartEngine | useCdpPort:=9222, IsDebugModeFlg:=True | 全テスト共通 (起動時) | PowerShellエンジンが非同期起動し、CDPポート(9222)のListen状態と親プロセス(EXCEL)監視が正常に開始されること。 |
| SYS-02 | Set-EngineConfig | EnableHighlight:=True | シナリオ開始時 | 操作対象のDOM要素が赤枠等でハイライト表示されること。 |
| SYS-03 | Clear-WebCache | Mode:="All" | シナリオ終了時 | WebView2のメモリ/ディスクキャッシュ、Cookie、セッション領域が安全にクリアされ、残留データがないこと。 |
| SYS-04 | CloseEngine | (なし) | CleanUp 処理内 | PowerShellプロセスが安全にシャットダウンされ、プロセスやポートが開放されること。 |

#### 2.2 ナビゲーション・待機ロジック
| テストID | コマンド / 関数名 | パラメータ例・主な仕様 | 検証対象シナリオ | 期待される動作・合格判定基準 |
| :--- | :--- | :--- | :--- | :--- |
| DOM-01 | Invoke-WebNavigation | Url:="file:///..." | 01_basic_form.html | 指定されたURLへのページ遷移が正常に開始されること。 |
| DOM-02 | Wait-WebPageLoad | (設定パラメータに従う) | 01_basic_form.html 等 | 全iframeを含め、DOMの `readyState=complete` になるまで同期待機されること。 |
| DOM-08 | Wait-WebDocumentReady | TimeoutSec:=15 | 03_frame_child.html | fileプロトコル制約下等において、軽量かつ高速にDOM解析完了のみを検知し復帰できること。 |
| DOM-07 | Wait-WebXPathElement | XPath:="//div[@id='asyncResult']" | 02_async_load.html | 3秒後に動的出現する非同期要素をタイムアウト前に検知して正常通過すること。 |
| ASY-01 | Wait-WebXPathElementDisappear| XPath:="//*[@id='opmask']" | 05_business_batch.html | 処理実行中に画面全体を覆うグレーアウト(`#opmask`)の消滅を検知し、消えた瞬間に次処理へ進むこと。 |

##### 2.3 標準DOM操作・要素取得
| テストID | コマンド / 関数名 | パラメータ例・主な仕様 | 検証対象シナリオ | 期待される動作・合格判定基準 |
| :--- | :--- | :--- | :--- | :--- |
| DOM-03 | Set-WebXPathTextInput | XPath:="//input...", Value:="入力値" | 01_basic_form.html | 指定したXPathの入力欄に指定文字列が正しく挿入されること。 |
| DOM-04 | Select-WebDropdown | Selector:="#selDept", Value:="D02" | 01_basic_form.html | セレクタで指定した `<select>` 要素のドロップダウン項目(指定Value)が選択されること。 |
| DOM-05 | Invoke-WebXPathClick | XPath:="//input[@id='chkAgree']" | 01_basic_form.html | チェックボックスのトグル切り替えやボタンクリックが正常に発火すること。 |
| DOM-06 | Get-WebXPathText | XPath:="//span[@id='lblDisplay']" | 01_basic_form.html | 画面上に反映された表示テキストを正確に読み取ってVBA変数に返却できること。 |
| TBL-01 | Export-WebTableToCsv | Selector:="#sampleTable" | 03_table_data.html | テーブル全行・列データが `<R>` (行) および `<T>` (列) 区切りの単一文字列としてメモリ上に抽出・パースできること。 |

#### 2.4 高度なDOM操作 (Robust DOM / CDP)
| テストID | コマンド / 関数名 | パラメータ例・主な仕様 | 検証対象シナリオ | 期待される動作・合格判定基準 |
| :--- | :--- | :--- | :--- | :--- |
| DOM-09 | Get-WebCssSelectorHint | XPath:="//button[contains(.,'検索')]" | 08_robust_dom.html | 裏側に隠れた非表示要素を弾き、画面に見えている要素(`isVisible`)だけを特定して一意のCSSセレクタを動的生成できること。 |
| DOM-10 | Invoke-WebSafeClick | Selector:="#btn-do-search" | 08_robust_dom.html | 生成したセレクタで要素を画面中央へスクロールさせ、最終可視性チェックをパスした場合のみクリックを発火させること。 |
| CDP-01 | Invoke-CdpCommand | Method:="Browser.getVersion" | 01_basic_form.html | 動的にWebSocket(9222)へ接続し、CDPプロトコル経由で直接ブラウザ情報を取得できること（Capability Check）。 |
| CDP-02 | Set-CdpNativeTextInput | Selector:="#txtInput", Value:="ABC" | 01_basic_form.html | JavaScriptを介さず、OSレベルのキーストロークエミュレーションとして確実なテキスト入力が行えること。 |

#### 2.5 ポップアップ制御・ダウンロード
| テストID | コマンド / 関数名 | パラメータ例・主な仕様 | 検証対象シナリオ | 期待される動作・合格判定基準 |
| :--- | :--- | :--- | :--- | :--- |
| TAB-01 | Switch-TabByTitle | TitleSubstring:="部門検索" | 06_popup_parent.html | ポップアップ開いた後、タイトル部一致でアクティブ操作対象を子画面へ正常に切り替えられること。 |
| SCR-01 | Invoke-WebScript | Js:="return document.getElementById..." | 06_popup_parent.html | JavaScriptを直接実行し、子画面から親画面の入力欄に引き継がれた値を取得・検証できること。 |
| DL-01 | `nable-SilentDownload | DownloadDir:="...", FileName:="..." | 04_file_download.html | OSの「名前を付けて保存」ダイアログを強制抑制し、保存先フォルダとファイル名を事前予約・固定化できること。 |
| DL-03 | Wait-FileDownload | FilePath:="...", TimeoutSec:=30 | 04_file_download.html | `.crdownload` (一時ファイル)の消失および排他ロック解除を監視し、完全なファイル生成を確認してミリ秒単位で終了すること。 |
| DL-02 | Invoke-WebFetchDownload | TargetUrl:="file://..." | 07_pdf_embed.html | UI(見た目)のボタンを押さず、Fetch APIにより裏側ネットワーク層からPDFバイナリを直接取得し、ダウンロードを発火させること。 |

#### 2.6 デバッグ・証跡エクスポート
| テストID | コマンド / 関数名 | パラメータ例・主な仕様 | 検証対象シナリオ | 期待される動作・合格判定基準 |
| :--- | :--- | :--- | :--- | :--- |
| SYS-05 | Export-WebScreenshot | FileName:="WebScreenshot_..." | エラー発生時等 | エラー発生瞬間のブラウザ画面全体が画像ファイル(PNG)として自動退避・保存されること。 |
| SYS-06 | Export-WebHtml | FileName:="WebHtml_..." | 08_robust_dom.html | 多段iframe構造を含む、画面全体のHTMLソースコードがテキストファイルとして完全出力されること。 |
| SYS-07 | Export-WebElementsToCsv | FileName:="WebElements_..." | 08_robust_dom.html | 画面内の操作可能な全要素(input/button等)の属性値や可視状態(`isVisible`)がCSV一覧として出力されること。 |
| SYS-08 | Export-WebFrameTreeToCsv | FileName:="WebFrameTree_..." | 08_robust_dom.html | ページ内に存在するiframe/frameのネスト階層構造がツリー形式でCSV化されること。 |
