# OutlookEmailOnlyAddIn（表示名除去・メールアドレス統一表示 VSTO アドイン）

クラシック Outlook for Windows（Office 2021 / バージョン 16.0 系、POP/SMTP + ローカル PST 環境）において、送信メールの宛先（To/Cc/Bcc）および受信トレイ・送信済みアイテム一覧の表示名を除去し、メールアドレスのみで統一的に表示・送信するための VSTO アドインです。
社内ローカルドメイン（初期値: `@kinki-sharyo.local`）は除外ルールにより従来の表示名を維持します。

---

## 1. 実現可否の調査結果と採用方式の理由

| 項目 | 調査・検証結果と採用方式 |
| :--- | :--- |
| a. 送信時 (`Application.ItemSend`) | 既存の解決済み `Recipient` に対して `PR_DISPLAY_NAME` (`0x3001001F`) のみを変更しても、内部の連絡先 `EntryID` リンクが残っていると SMTP 送信時に元の表示名が復活する。そのため、除外ドメイン以外の対象 `Recipient` を一旦 `Recipients.Remove` で削除し、純粋な SMTP アドレスのみで `Recipients.Add` により再追加した直後に `PR_DISPLAY_NAME` (`0x3001001F`)・`PR_TRANSMITABLE_DISPLAY_NAME` (`0x3A20001F`) をメールアドレスで上書きし、`PR_SEND_RICH_INFO` (`0x3A40000B`) を `false` に設定して `Save()` する方式を採用。 |
| b. 送信済みアイテムの一覧表示 | 一覧の「宛先」列（`PR_DISPLAY_TO` / `0x0E04001F`）および「CC」列（`PR_DISPLAY_CC` / `0x0E03001F`）は PST ストアが計算する読み取り専用プロパティであり直接 `SetProperty` は不可。送信時または一括処理時に `MailItem.Recipients` 内の各 `Recipient` の `PR_DISPLAY_NAME` をメールアドレスに書き換えて `MailItem.Save()` することで、PST ストアに `PR_DISPLAY_TO` / `PR_DISPLAY_CC` をメールアドレスのみで再計算させる方式を採用。 |
| c. 受信メールの一覧表示 | 一覧の「差出人」列が参照する `PR_SENT_REPRESENTING_NAME` (`0x0042001F`) および `PR_SENDER_NAME` (`0x0C1A001F`) はローカル PST 内の `MailItem.PropertyAccessor` で直接書き込み可能。元の表示名をカスタム名前付きプロパティ `EmailOnlyOrigSender` に退避したうえで、SMTP アドレスで上書き・保存する方式を採用。 |
| d. 新着検知とスキャン併用 | `Items.ItemAdd` は16件超の同時追加時に発火しない制約があるため、(1) `Items.ItemAdd` 監視、(2) `SyncObject.SyncEnd`（送受信完了イベント）フック、(3) アドイン起動5秒後および定期（既定5分）の `Items.Restrict` による差分スキャン（50ms ごとに最大5件ずつ処理する UI スレッドチャンクタイマー）の3段構えを採用。 |
| e. 閲覧・返信・転送時の挙動 | `PR_SENT_REPRESENTING_NAME` / `PR_SENDER_NAME` がメールアドレスに書き換わっているため、返信・全員に返信・転送を行った際の引用ヘッダーや返信宛先にもメールアドレスのみがセットされ、元の表示名が復活しない。 |
| f. 制約事項 | 連絡先（アドレス帳）に該当アドレスが登録されている場合、一覧ビュー（テーブルビュー）は確実にメールアドレスのみになるが、閲覧ウィンドウ（プレビューペイン）の上部ヘッダーだけは Outlook 2021 のスマートルックアップ機能によりローカル連絡先名が動的表示される場合がある。 |

---

## 2. アーキテクチャ概要

### クラス構成
- `ThisAddIn.cs` : アドインのライフサイクル管理、`ItemSend` イベント登録、`FolderMonitor` 起動、リボン登録
- `Configuration/AddInSettings.cs`, `SettingsManager.cs` : `%AppData%\OutlookEmailOnlyAddIn\settings.json` によるユーザー設定の永続化と変更通知
- `Logging/FileLogger.cs` : `%AppData%\OutlookEmailOnlyAddIn\logs\addin-YYYYMMDD.log` へのログ出力（Error / Info / Debug 切替）
- `Core/MapiConstants.cs` : MAPI プロパティ DASL タグおよびカスタム名前付きプロパティ定義
- `Core/DomainFilter.cs` : ローカルドメイン判定（複数指定、サブドメイン考慮、大文字小文字無視）
- `Core/MailAddressHelper.cs` : `Recipient` / `AddressEntry` / `MailItem` からの SMTP アドレス抽出、S/MIME 判定、COM オブジェクト解放
- `Core/OutgoingMailProcessor.cs` : 送信時（`ItemSend`）の連絡先グループ展開および宛先表示名除去
- `Core/IncomingMailProcessor.cs` : 受信メール・送信済みメールの表示名書き換えと処理済みマーキング
- `Core/FolderMonitor.cs` : 全 PST ストアの受信トレイ・送信済みアイテムの `ItemAdd` 監視、`SyncEnd` 監視、定期チャンク差分スキャン
- `Core/BatchProcessor.cs` : 既存メールのフォルダ単位一括処理（非同期チャンク実行・進捗通知・中断/再開）
- `UI/AddInRibbon.cs` : Outlook メインウィンドウ「ホーム」タブへのカスタムリボングループ追加
- `UI/SettingsForm.cs` : 設定画面（WinForms）
- `UI/BatchProcessForm.cs` : 既存メール一括処理画面（フォルダツリー選択・プログレスバー・中断ボタン）

---

## 3. ビルド・インストール・アンインストール手順

### 3.1 開発環境でのビルド手順
1. Visual Studio 2019 または Visual Studio 2022 に「Office/SharePoint 開発」ワークロード（Visual Studio Tools for Office: VSTO）がインストールされていることを確認します。
2. `OutlookEmailOnlyAddIn.sln` を Visual Studio で開きます。
3. ソリューション構成を `Release`、ソリューション プラットフォームを `Any CPU`（32bit/64bit Outlook 両対応）に設定します。
4. メニューの［ビルド］→［ソリューションのビルド］を実行します（開発機ではビルド時に自動的に HKCU レジストリへアドインが登録され、そのまま Outlook を起動して動作確認できます）。

### 3.2 ClickOnce による配布用インストーラーの作成とインストール手順
1. Visual Studio のソリューション エクスプローラーで `OutlookEmailOnlyAddIn` プロジェクトを右クリックし、［プロパティ］を開きます。
2. ［署名］タブを開き、［ClickOnce マニフェストに署名する］にチェックを入れ、社内証明書を選択するか［テスト証明書の作成］をクリックして署名を設定します。
3. ［発行］タブを開き、発行フォルダーの場所（例: `publish\`）を指定して［今すぐ発行］をクリックします。
4. 生成された `publish` フォルダ一式（`setup.exe`, `OutlookEmailOnlyAddIn.vsto`, `Application Files` フォルダ）を配布先 PC へコピーします。
5. 配布先 PC で Outlook を終了した状態で `setup.exe`（または `OutlookEmailOnlyAddIn.vsto`）を実行し、画面の指示に従ってインストールします。
   - ※配布先 PC に「Visual Studio 2010 Tools for Office Runtime (VSTO Runtime)」が未導入の場合は、`setup.exe` が自動的に前提条件として導入するか、Microsoft 公式の `vstor_redist.exe` を事前インストールしてください。
   - ※本アドインは `AnyCPU` でビルドされ、ユーザー単位（`HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookEmailOnlyAddIn`）に登録されるため、32bit 版 Outlook / 64bit 版 Outlook のどちらでも同一のインストーラーで動作します。

### 3.3 アンインストール手順
1. Outlook を終了します。
2. Windows の［設定］→［アプリ］→［インストールされているアプリ］（または［コントロール パネル］→［プログラムと機能］）を開きます。
3. 一覧から `OutlookEmailOnlyAddIn` を選択して［アンインストール］を実行します。
4. ユーザー設定ファイルやログファイルも完全に削除する場合は、エクスプローラーのアドレスバーに `%AppData%\OutlookEmailOnlyAddIn` と入力し、該当フォルダを削除してください。

---

## 4. テスト観点一覧

| 分類 | No. | テスト項目 | 確認内容・期待される結果 |
| :--- | :--- | :--- | :--- |
| 正常系（送信） | 1 | 外部ドメイン宛て送信（連絡先登録あり） | 連絡先に「A様 (`a@example.com`)」と登録された宛先へ To/Cc で送信した際、送信メールのヘッダーおよび送信済みアイテム一覧の宛先から「A様」が消え、`a@example.com` のみになること。 |
| 正常系（送信） | 2 | 手入力の表示名付き宛先送信 | 宛先欄に `"B様" <b@example.org>` と直接入力して送信した際、表示名が除去され `b@example.org` のみになること。 |
| ローカルドメイン除外 | 3 | `@kinki-sharyo.local` 宛ての送信・受信 | `taro@kinki-sharyo.local`（および大文字混じりの `Taro@KINKI-SHARYO.LOCAL`）は、送信時も受信トレイ一覧でも元の表示名（「社内太郎」等）がそのまま維持されること。 |
| ローカルドメイン除外 | 4 | サブドメインの除外判定 | 設定の「サブドメインも除外対象に含める」が ON のとき `user@sub.kinki-sharyo.local` の表示名が維持され、OFF に変更すると表示名が除去されること。 |
| ローカルドメイン除外 | 5 | 外部＋社内混在メールの送信 | To に `a@example.com`（外部）と `taro@kinki-sharyo.local`（社内）を同時に指定して送信した際、外部宛先のみメールアドレス化され、社内宛先は表示名が維持されること。 |
| Bcc 設定 | 6 | Bcc 対象 ON/OFF 切替 | 設定の「Bcc の宛先も表示名除去の対象にする」が ON の場合は Bcc も表示名除去され、OFF の場合は Bcc の表示名が維持されること。 |
| 連絡先グループ | 7 | 連絡先グループ（配布リスト）宛て送信 | 社外アドレスと社内アドレスを含む連絡先グループを宛先に指定して送信した際、個別メンバーに展開されたうえで社外アドレスのみ表示名が除去されること。 |
| 正常系（受信・一覧） | 8 | 新着メール受信時の自動書き換え | 外部差出人 `"Sender Name" <sender@example.com>` からの新着メール受信時、受信トレイ一覧の「差出人」列が `sender@example.com` に書き換わること。 |
| 正常系（受信・一覧） | 9 | 大量メール受信時（17件以上同時） | 17件以上のメールを一度に POP 受信した際、`ItemAdd` 未発火分も送受信完了フックおよび差分スキャンによって自動的にメールアドレス表示へ書き換わること。 |
| 返信・転送 | 10 | 書き換え済みメールへの返信・全員に返信 | 書き換え後の受信メールに対して「返信」「全員に返信」「転送」を行った際、宛先欄や引用ヘッダーに元の表示名が復活せずメールアドレスとなること。 |
| 一括処理 | 11 | フォルダ選択一括処理・中断と再開 | 「既存メール一括処理」画面で受信トレイ・送信済みアイテムを選択して実行し、途中で「中断」ボタンを押して停止後、再度「一括処理開始」を押した際に処理済みメールがスキップされ残りから再開されること。 |
| 異常系・境界条件 | 12 | S/MIME（電子署名・暗号化）メール | S/MIME 署名付き・暗号化メールは送信時も受信・一括処理時も自動でスキップされ、署名・暗号化が破損しないこと。 |
| 異常系・境界条件 | 13 | アドレス不正・取得不可メール | 差出人アドレスが空または有効な SMTP 形式でないアイテム（システム通知等）はエラーで落ちることなく元の表示名が維持されること。 |

---

## 5. 既知の制約・今後の改善案

### 既知の制約
1. 閲覧ウィンドウ（プレビューペイン）上部ヘッダーでの連絡先名表示:
   - Outlook 2021 の仕様により、ローカルの「連絡先」フォルダに該当メールアドレスが登録されている場合、メール一覧（テーブルビュー）の差出人・宛先列は確実にメールアドレスのみになりますが、閲覧ウィンドウ内のヘッダー表示部では Outlook がローカル連絡先を動的参照して登録名を表示する場合があります（送信メールの実際のインターネットヘッダーからは確実に表示名が消えています）。
2. S/MIME メールの除外:
   - 電子署名付きメールの送信後・受信後に表示名プロパティを書き換えると署名のハッシュ検証に影響する可能性があるため、安全を優先して S/MIME メールは処理対象外としています。

### 今後の改善案
1. 元の表示名へのロールバック機能:
   - 現在、受信メールの書き換え時に元の差出人表示名をカスタム MAPI プロパティ（`EmailOnlyOrigSender`）へ自動退避しています。将来的に「選択したフォルダまたはメールの差出人表示名を元の状態に戻す（復元する）」ボタンを一括処理画面に追加することが容易に可能です。
2. グループポリシー（GPO）による全社一括設定配布対応:
   - 現在はユーザーごとに `%AppData%\OutlookEmailOnlyAddIn\settings.json` を参照していますが、レジストリ（`HKLM\Software\Policies\...`）に設定値が存在する場合はそちらを優先して読み込む機能を追加することで、システム管理者による全社統一ポリシー配布に対応できます。
