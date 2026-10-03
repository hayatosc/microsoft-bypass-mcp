# 既存 Power Automate フローのソースと更新

[README](../../README.md) · [セットアップ](../../docs/setup.md) · [仕様](../../SPEC.md)

`definition.json` は、既存の `microsoft bypass flow` を同じフロー内で拡張するための、唯一の正規ソースです。HTTP トリガー、ゲートウェイキー条件、操作スイッチ、公開用の認証・接続パラメーターの慣例を保持し、固定の読み取り・下書き操作を追加しています。名称は既存の機械識別用で、組織の承認回避を推奨する意味ではありません。

公開ファイルは秘密情報を除いた Workflow Definition Language (WDL) のレビュー・手動更新資料です。**単独でインポートできるパッケージ、Dataverse ソリューション、コネクタ作成スクリプト、ライブ環境の完成品ではありません。** 新規環境への初期フロー構築や実接続は同梱せず、既存フローがない場合の準備は別途必要です。

## 14 操作と 16 ツール

| 分類 | フロー操作 |
| --- | --- |
| Outlook メール | `list_messages`, `search_messages`, `get_message`, `list_mail_folders`, `get_conversation` |
| Outlook 添付 | `list_attachments`, `get_attachment` |
| Outlook 下書き | `create_draft`, `create_reply_draft`, `add_draft_attachment` |
| OneDrive | `onedrive_search_files`, `onedrive_list_folder`, `onedrive_get_metadata`, `onedrive_get_content` |

`outlook_inspect_attachment` / `outlook_read_attachment` は `get_attachment`、`onedrive_inspect_file` / `onedrive_read_file` は `onedrive_get_content` を共有するため、MCP 側は 16 ツール（読み取り 13、下書き書き込み 3）です。メール送信・削除・汎用プロキシはありません。

## トリガー・キー・接続の配線

生成ソースの固定構造を保持します。

- トリガー名 `manual`: `type: "Request"` / `kind: "Http"`。既存の `triggerAuthenticationType: "All"` を保持。
- 入力は `{ operation, requestId, args }`。トリガー enum と `スイッチ` の 14 分岐を一致させ、各分岐で envelope・引数・UUID v4 を検証。
- スイッチ式は `@triggerBody()?['operation']`。
- `McpGatewayKey` は `SecureString`。公開の `defaultValue` は空。Worker の `POWER_AUTOMATE_GATEWAY_KEY` と同じ実値を非公開で設定。
- Worker は `POWER_AUTOMATE_URL` に POST し、`X-MCP-Gateway-Key` を送信。

トリガー条件は次のままです。

```text
@and(not(empty(parameters('McpGatewayKey'))),equals(triggerOutputs()?['headers']?['X-MCP-Gateway-Key'],parameters('McpGatewayKey')))
```

空キーまたは不一致ならコネクタ処理を許可しません。公開の空設定を、動作中の非公開設定へ上書きしないでください。URL の署名等も秘密情報です。トリガー段階の拒否が通常のツールエラー envelope になるとは限りません。

接続パラメーター `$authentication`（`SecureObject`、公開既定値 `{}`）、`$connections`（`Object`、公開既定値 `{}`）と、各アクションの `@parameters('$authentication')` を保持します。これらの空値は実際の接続を作りません。

| 用途 | API / 接続 alias | 操作 |
| --- | --- | --- |
| Outlook | `/providers/Microsoft.PowerApps/apis/shared_office365` / `shared_office365` | `HttpRequest`。固定の Graph 経路のみ |
| OneDrive | `/providers/Microsoft.PowerApps/apis/shared_onedriveforbusiness` / `shared_onedriveforbusiness` | 下表のネイティブ操作のみ |

許可された手動更新で既存接続へ結び付けます。新しい同意・権限付与・接続作成は自動化していません。ライセンス、テナントの DLP、コネクタ利用と接続所有者の権限を確認してください。

## フィクスチャからの変更範囲

[`scripts/fixtures/microsoft-bypass-flow.pre-attachments.json`](../../scripts/fixtures/microsoft-bypass-flow.pre-attachments.json) は変更しない履歴・由来のフィクスチャです。現在のフローを別に配備するためのソースではありません。

生成器は、元のメール 3 分岐をそのまま保持するのではなく、現行仕様に更新した期待値を検証します。現在の変更は次のとおりです。

1. トリガーの operation enum を固定の許可リストへ更新。
2. 既存の認証・キー条件、スイッチ、default 分岐、パラメーターを保持し、OneDrive 検索用の空の Compose 設定 2 つを追加。
3. 分岐を固定の読み取り・下書き操作として構築。
4. Outlook のスコープ・フォルダー・限定フィルター・ページング・会話取得を追加。
5. 添付のメタデータ投影、file 種別・サイズ・base64 ガードと、リクエスト内だけのバイト通信を保持。
6. OneDrive for Business のネイティブ分岐と、最大 1,000 件の指定フォルダー集約を追加。
7. 下書き作成・送信者への返信下書き・確認済み下書きへの添付追加だけを追加。

最後に actions 内の空文字リテラルの互換変換を適用します。別フロー、追加トリガー、任意 HTTP 分岐は生成しません。

## classic designer のクリップボード互換性

`build_attachment_flow.py` は全分岐の構築後に `scripts/clipboard_compat.py` を **`actions` 部分だけ**へ適用します。WDL トークンとしての空文字 `''` を `string(null)` に変換する処理で、アポストロフィの一括置換ではありません。

空でない文字列、二重アポストロフィのエスケープ、リテラル、補間境界、JSON 型、アクション構造、接続・認証・安全設定は保持します。トリガー、パラメーター、その他のトップレベル値、由来のフィクスチャは変換しません。[公式 `string` 仕様](https://learn.microsoft.com/en-us/azure/logic-apps/expression-functions-reference#string)は、`string(null)` が null ではなく空の String になることを定義しています。

**過去の確認記録:** 旧資料では、classic designer の clipboard roundtrip が空リテラルを含む 37 の WDL 値を壊し、互換候補では同じ 37 パスの 1,149 トークンを変更して対処したと記録されています。候補の未保存 roundtrip では値・型が保持され、別の 5 WDL 値は空白だけが変化し、14 分岐と接続参照が残ったという限定的な記録です。

これは過去の未保存確認で、今回の文書更新による新しい Microsoft 側検証ではありません。保存・デプロイ・コネクタ実行・メールテストを意味しません。後の局所的な添付サイズ修正は Microsoft デザイナーで roundtrip 検証されていません。ローカルの lexical / source テストと合成動作再生も Microsoft ランタイムテストではありません。

## Outlook の契約

`list_messages` は inbox / sent / all / 指定フォルダーに限定します。フィルターなしなら `receivedDateTime desc`、限定フィルターがあれば `$orderby` を省略して `InefficientFilter` を避けます。フィルターは `isRead`、`hasAttachments`、`receivedAfter`、`receivedBefore` の固定順だけです。

`get_conversation` は `/me/messages` に対するエスケープ済み `conversationId` の完全一致で、`$orderby` も、受信トレイの部分取得による代替もありません。Worker は各ページ内だけを整列し、会話 ID 不一致と ID 重複を拒否します。

Outlook のパス ID は個別に `uriComponent` で符号化します。呼び出し元の `%` は再符号化され、区切りとして解釈されません。メールの `from` 省略 / null は Worker で空の送信者に正規化しますが、不正な非 null 送信者は拒否します。詳細は [読み取りガイド](../../docs/read-tools.md) を参照してください。

## 下書きの導入時確認

[下書きガイド](../../docs/drafts.md)の承認、引数、添付の受け渡し、権限、曖昧な結果の扱いに従います。既存接続で `Mail.ReadWrite` が許可される必要があり、無断の同意・資格情報追加は行いません。POST の再試行はすべて無効です。失敗時にも保存済みの可能性があるため、Outlook の下書きと添付を確認してから判断します。

添付 POST 成功には、要求 canonical base64 と完全一致する有限の `contentBytes`、有効な添付 ID、要求した名前が必要です。Graph `size` は非負 Int32 メタデータで実ファイル長ではありません。応答には検証済み実バイト数（1 byte〜2 MiB）を `size` として返し、Worker も入力との一致を確認します。内容の欠落・不一致は `DRAFT_WRITE_AMBIGUOUS` で停止し、再送、代替 read、追加 endpoint は使いません。

読み取り添付の Graph `size` と独立した 4 MiB metadata / raw 上限は変更しません。raw 889 / metadata 1223、raw がちょうど 2 MiB で metadata がそれより大きい場合の合成テストはありますが、テナント固有のコネクタ応答をライブ検証済みとはみなしません。

## OneDrive の契約

| 分岐 | ネイティブ operation ID |
| --- | --- |
| `onedrive_search_files` | `FindFiles` |
| `onedrive_list_folder`、`folderId` 省略 | `ListRootFolder` |
| `onedrive_list_folder`、`folderId` 指定 | `ListFolderV2` |
| `onedrive_get_metadata` | `GetFileMetadata` |
| `onedrive_get_content` の事前確認 | `GetFileMetadata` |
| `onedrive_get_content` の内容 | `GetFileContent` |

`ListFolderV2` だけに `paginationPolicy.minimumItemCount: 1000` を設定します。元資料には既存デザイナーで Pagination 設定を確認したという過去の記録がありますが、現行テナントの動作保証ではありません。継続はネイティブコネクタ内で行い、フローは最後のページが超過しても最大 1,000 件を投影し、しきい値到達・継続を不完全と示します。

`ListRootFolder` は返却配列のみ。`FindFiles` は最大 100 件で、継続入力はありません。指定フォルダーのページングが検索範囲を広げるわけではありません。MCP の一覧カーソルは Worker が毎回取得し直す有限範囲の分割で、無制限の列挙・安定したスナップショットではありません。

### 検索用の固定設定

公開生成ソースには、空入力の Compose アクションを残します。

- `OneDriveSearchRootId`: 接続所有者のネイティブルートフォルダー ID。
- `OneDriveSearchMode`: 対象デザイナーで確認した公式 `FindFiles.findMode` の機械値。

どちらかが空なら、コネクタ処理前に HTTP 503 / `ONEDRIVE_SEARCH_NOT_CONFIGURED` で停止します。通常の非ソリューション cloud-flow designer で扱う設定で、新しい workflow parameter ではありません。秘密鍵や OAuth 資格情報でもなく、Worker の環境変数でもありません。

許可された手動更新で、**呼び出し元に依存しない固定の確認済み値**を設定します。コードレビューで機械値を推測したり、provider URL をルートの代わりにしたりしません。`shared_onedriveforbusiness` の接続バインドも別途必要です。

### 内容の取得

最新メタデータで ID、通常ファイル、非空、4 MiB 以下、PDF / DOCX / XLSX の拡張子と一致または汎用 MIME 型を確認してから内容を取得します。Logic Apps の `{ "$content-type", "$content" }` の `$content` は既に base64 です。base64 形式、binary MIME、復号長とネイティブサイズを検証し、Worker にだけ `{ metadata, contentBytes }` を返します。Worker も先行メタデータとの差を検証し、リクエスト内で解析します。

## 応答と実行履歴

成功 envelope:

```json
{ "ok": true, "requestId": "...", "operation": "...", "data": {} }
```

Outlook の一覧・検索・会話・フォルダーは、Graph nextLink をフロー → Worker の内部データとして含む場合があります。Worker は検証または破棄し、URL を返したり再実行したりしません。OneDrive は有限結果と truncation / 継続の存在を示すフラグだけを返します。

エラーは `INVALID_REQUEST`、`INVALID_ARGUMENTS`、`UPSTREAM_ERROR`、`INVALID_ATTACHMENT_METADATA`、`INVALID_ATTACHMENT_CONTENT`、`INVALID_ONEDRIVE_METADATA`、`INVALID_ONEDRIVE_CONTENT`、`UNSUPPORTED_ONEDRIVE_FILE_TYPE`、`ONEDRIVE_FILE_TOO_LARGE` 等の固定値です。下書きの曖昧なエラーには再試行前の下書き確認を促す固定文を使います。検索語、件名、ID、名前、URL、生コネクタ文、バイト列、base64 はエラーに含めません。

- データを扱う `OpenApiConnection` / `Select` は安全な inputs / outputs を設定。
- Request トリガーは outputs を安全にし、引数とヘッダーの履歴露出を軽減。
- `Compose` / `ParseJson` / `Response` は安全な inputs を使い、対応する履歴表示で outputs も隠す。
- control action に非対応の `secureData` 設定は追加しない。

これらは履歴表示の露出軽減で、Microsoft の保持・監査を保証しません。Worker がデータを保存しない規則とは別です。メール・ファイル内容は信頼できない外部データとして扱います。

## 再生成とオフラインチェック

`build_attachment_flow.py` がフィクスチャから構築し、`build_read_tools_flow.py` と `build_draft_tools_flow.py` で同じ定義を拡張し、最後に actions の互換変換を適用します。配備用の別フローは生成しません。

```sh
python3 scripts/build_attachment_flow.py
python3 scripts/build_attachment_flow.py --check
bun run test:flow
```

変更予定がなければ書き込みを伴う先頭コマンドは不要で、`--check` だけで一致を確認できます。生成 JSON を直接編集せず、フィクスチャも更新しません。

回帰テストはフィクスチャ checksum、元の添付分岐 hash、互換変換後の分岐、現在の決定的な生成 bytes を固定します。過去の変換前・候補 hash の検証には、テスト用コピーだけで下書き添付結果の既知 3 値を戻します。37 パス・1,149 置換の契約も維持します。

テストはオフラインのソース契約・合成データ検証です。JSON の import、フロー保存、接続作成、メール / OneDrive 呼び出し、Microsoft ランタイム確認は行いません。

## 手動更新の順序（別途許可が必要）

1. **権限・契約を確認する。** 組織の許可、DLP、HTTP トリガーと必要コネクタのライセンス、接続所有者の権限を確認します。
2. **既存フローを非公開でバックアップする。** 接続、キーの配線、トリガー URL を含むため、Git・PR・スクリーンショット・ログに入れません。これは許可された運用者の作業で、文書更新やオフラインテストの入力にはしません。
3. **同じフローを開く。** 新規フロー、Save As、トリガー置換、未承認の接続変更はしません。
4. **デザイナー対応の入力欄で分岐と enum を再現する。** 式が参照するアクション名、run-after、失敗経路、安全設定を保持します。JSON 全体をそのままインポートできるとは考えません。
5. **Outlook / OneDrive を既存の承認済み接続に配線する。** 固定 alias と operation ID を保持し、`FindFiles.findMode` とルート ID を対象テナントで確認します。
6. **キー条件と安全な履歴設定を保持する。** 公開の空パラメーターで非公開設定を上書きしません。
7. **非公開 diff をレビューしてから保存する。** 保存後に callback URL を非公開で再確認し、変わった場合は Worker の設定と整合するまで止めます。URL を公開文面に貼りません。
8. **ライブ試験が別途許可された場合だけ確認する。** 最小の読み取り、ファイル安全ゲート、秘匿化エラーを確認します。下書き試験には書き込みの承認を追加で確認し、曖昧な結果は Outlook で確認します。

対象デザイナーでアクション、安全な履歴、run-after、接続を表現できない場合はレビューのため停止します。黙って別フローや汎用 HTTP / Graph に置き換えません。Cloudflare Access / OAuth / Worker の準備は [セットアップ](../../docs/setup.md) の別工程です。

## 公式資料

- [OneDrive for Business の操作・パラメーター](https://learn.microsoft.com/en-us/connectors/onedriveforbusiness/)
- [Logic Apps のバイナリ表現](https://learn.microsoft.com/en-us/azure/logic-apps/logic-apps-content-type)
- [Graph の添付一覧](https://learn.microsoft.com/en-us/graph/api/message-list-attachments?view=graph-rest-1.0)
- [Graph の添付取得](https://learn.microsoft.com/en-us/graph/api/attachment-get?view=graph-rest-1.0)
- [WDL の式・エスケープ](https://learn.microsoft.com/en-us/azure/logic-apps/workflow-definition-language-schema#expressions)
- [安全な入出力](https://learn.microsoft.com/en-us/azure/logic-apps/set-up-security-permissions)
