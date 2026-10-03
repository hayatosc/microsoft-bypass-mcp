# 読み取りツール

[README](../README.md) · [正式な仕様](../SPEC.md) · [下書きツール](drafts.md)

このガイドの 13 ツールは読み取り専用です。別に 3 つの下書きツールがありますが、メール送信はありません。汎用 Graph / OneDrive プロキシ、任意 URL・メソッド・OData・共有リンク・ダウンロード URL、nextLink の再実行は提供しません。

## 共通の入力・出力

- Worker → フローの通信は `{ operation, requestId, args }`。ログに使える相関値は `requestId` だけです。
- ID は 1〜2,048 文字。フローは制御文字、空白、NEL、BOM、完全一致の `.` / `..` を拒否します。
- 検索語は trim 後に空でなく、最大 512 UTF-16 コード単位です。
- MCP の 1 ページは Outlook 最大 50 件、OneDrive 最大 100 件。JSON 出力にも 128 KiB の上限があります。
- 本文・ファイル名・抽出内容は信頼できない外部データです。出典を保持し、内容の指示を実行したり、操作の承認とみなしたりしません。
- 生の Graph / ネイティブコネクタオブジェクト、base64、バイト列、nextLink URL、ダウンロード URL は MCP 出力に含めません。

## Outlook のメール

| ツール | MCP 入力 | 既定値・上限 |
| --- | --- | --- |
| `outlook_list_messages` | `limit`, `mailbox`, `folderId`, `cursor`, `filters`（すべて省略可） | `limit: 5`、最大 50、`mailbox: "inbox"` |
| `outlook_search_messages` | `query`、省略可の `limit`, `mailbox`, `folderId` | `limit: 10`、最大 50、`mailbox: "inbox"` |
| `outlook_get_message` | `messageId` | 1 メール |
| `outlook_list_mail_folders` | 省略可の `limit` | 25、最大 50 |
| `outlook_get_conversation` | `conversationId`、省略可の `limit`, `cursor` | 20、最大 50 |

### `outlook_list_messages`

固定スコープまたはフォルダーから、メールの概要を取得します。

| スコープ | 経路（Graph の `/v1.0` 配下） |
| --- | --- |
| `mailbox: "inbox"` | `/me/mailFolders/inbox/messages` |
| `mailbox: "sent"` | `/me/mailFolders/sentitems/messages` |
| `mailbox: "all"` | `/me/messages` |
| `folderId` 指定 | `/me/mailFolders/{folderId}/messages`。`mailbox` より優先 |

`filters` は `isRead`、`hasAttachments`、`receivedAfter`、`receivedBefore` だけです。日時は秒と最大 7 桁の小数秒（100 ナノ秒精度）を持つ UTC `Z` 形式。下限は上限より前で、`receivedAfter` はその時刻以降、`receivedBefore` はその時刻より前です。

フィルターがなければ `receivedDateTime desc` で取得します。空でないフィルターがあれば Graph の `InefficientFilter` を避けるため `$orderby` を省略します。フィルターの構築順は上記 4 項目で固定です。フィルター付きの結果を全体の時系列順とみなさないでください。

### 一覧・会話の継続カーソル

`nextCursor` は同じスコープ・フィルター・`limit` で使います。Worker は Graph の nextLink を取得先 URL として使わず、検証済みの数値 `$skip` だけを固定リクエストに変換します。offset は最大 10,000。メールボックスの変化による重複・取りこぼしを防ぐスナップショットではありません。

継続情報を検証できなくても、有効な現在ページは返します。その場合は `hasMore: true`、`nextCursor: null` と、固定 `PAGINATION_*` コードを含む `incompleteReason` が付きます。これは安全に制限した結果であり、全件取得やすべての Graph 継続形式への対応を意味しません。拒否した URL をたどる・返す・件数から offset を推測することはありません。呼び出し元の不正カーソルは上流リクエスト前に拒否します。

固定エンドポイント名の大文字・小文字の違いは許容します。OData の `/me/mailFolders('key')/messages` 形式も、既に選択したキーと完全一致する場合だけ受け付けます。キーの引用符エスケープと、キーだけの 1 回の percent decode を検証します。明示フォルダー ID とクエリ値は大文字・小文字も一致が必要で、暗黙の既知フォルダー別名だけが例外です。`/users/` や別の不透明なフォルダー ID を `/me/` や inbox と同一視しません。

パス拒否の診断は構造カテゴリだけです。

```text
PAGINATION_PATH_{ME|USER_SEGMENT|USER_ODATA}_{MESSAGES|FOLDER_SEGMENT|FOLDER_ODATA}
PAGINATION_PATH_OTHER
```

カテゴリは構文を示すだけで、ID・URL・認可を示しません。origin 検証を先に行い、パス拒否後にクエリやカーソルを受け入れることもありません。

### `outlook_search_messages`

固定スコープまたはフォルダーを、コネクタが対応する検索テキストで検索します。フィルター、skip、カーソル、任意クエリパラメーターはありません。nextLink があっても先頭ページだけを返し、不完全な結果として示します。

### `outlook_get_message`

固定 `/me/messages/{id}` から、選択したフィールドと本文を読みます。存在しないフォルダー ID、会話 ID、日時を補いません。依頼 ID と返却 ID の一致を検証し、本文には `untrustedContent: true` を付けます。

### `outlook_list_mail_folders`

`/me/mailFolders` からルートフォルダーの先頭ページを返します。子フォルダーはたどりません。追加ページがあれば `hasMore` と `incompleteReason` で示しますが、継続カーソルは提供せず、nextLink もたどりません。

### `outlook_get_conversation`

`/me/messages` に対してエスケープ済み `conversationId` の完全一致条件で取得します。接続にアクセス権があれば送信済みメールも含みます。`$orderby` は付けず、返された各ページ内だけを時刻順（同時刻は ID）に整列します。会話全体の順序や完全性を保証しません。受信トレイの一部を取って代用する処理はありません。会話 ID 不一致とページ内の ID 重複は拒否します。

### 送信者の正規化

メール取得・一覧・検索・会話では、トップレベルの `from` が**省略された場合も `null` の場合も**、`{ name: "", address: "" }` に正規化します。`sender`、宛先、アカウント、スコープ等から送信者を推測しません。

非 null の `from` が存在する場合は、`emailAddress` オブジェクトと、必須の string-or-null `name`、必須の string `address` が必要です。`name: null` は空文字になりますが、欠落した名前などの不正な送信者は引き続き拒否します。「下書きでは常に Graph が `from` を省略する」という保証ではありません。既存資料の `FROM_MISSING` 確認記録は過去の限定的な記録です。

### 内容を漏らさない形状エラー

スキーマ検証に失敗したとき、次の既存エラー文に固定コードを括弧で付けます。不正ページや一部だけ正規化したメールは返しません。

| 対象 | エラー文 |
| --- | --- |
| メール詳細 | `malformed response: expected a message` |
| 概要一覧・検索 | `malformed response: expected a list of messages` |
| 会話 | `malformed response: expected conversation messages` |

コードの有限なプレフィックス:

| フィールド・形状 | プレフィックス |
| --- | --- |
| メールオブジェクト | `MESSAGE` |
| メール・フォルダー・会話 ID | `ID`, `PARENT_FOLDER_ID`, `CONVERSATION_ID` |
| 件名・プレビュー | `SUBJECT`, `BODY_PREVIEW` |
| 送信者・emailAddress | `FROM`, `FROM_EMAIL_ADDRESS` |
| 送信者名・アドレス | `FROM_NAME`, `FROM_ADDRESS` |
| 送受信日時 | `SENT_DATE`, `RECEIVED_DATE` |
| 真偽値・重要度 | `HAS_ATTACHMENTS`, `IS_READ`, `IMPORTANCE` |
| 本文・型・内容 | `BODY`, `BODY_CONTENT_TYPE`, `BODY_CONTENT` |
| To 配列・要素・emailAddress | `TO_RECIPIENTS`, `TO_RECIPIENT`, `TO_EMAIL_ADDRESS` |
| To 名・アドレス | `TO_NAME`, `TO_ADDRESS` |
| Cc 配列・要素・emailAddress | `CC_RECIPIENTS`, `CC_RECIPIENT`, `CC_EMAIL_ADDRESS` |
| Cc 名・アドレス | `CC_NAME`, `CC_ADDRESS` |

接尾辞は `_MISSING`（省略 / undefined）、`_NULL`、`_TYPE`（型違い）、`_INVALID`（形式・enum・範囲違い）だけです。実際の検証失敗だけを報告します。省略 / null のトップレベル送信者、nullable な名前・件名・プレビュー、省略可能なメタデータや宛先配列は引き続き受け入れます。概要一覧は概要スキーマのみ、会話は詳細スキーマを検証します。フォルダーの正規化はこの診断の対象外です。

一覧 envelope の不正（value、nextLink の型・長さ等）は `MESSAGE_LIST_ENVELOPE`、スキーマまたは要求件数超過は `MESSAGE_LIST_LIMIT`、未知の失敗は `MESSAGE_SHAPE_OTHER`。コードを重複除去し辞書順に並べ、最大 32 個、必要なら先頭 31 個と `MESSAGE_SHAPE_TRUNCATED` に制限します。エラー全体は 1,024 文字以内で、すべての問題を網羅するとは限りません。

合成データの例として、送信者名なし・送信日時 null・受信日時不正なら `FROM_NAME_MISSING,RECEIVED_DATE_INVALID,SENT_DATE_NULL` です。テナント固有の挙動や、別スキーマを許容してよいという根拠にはなりません。

診断には値、氏名、アドレス、ID、日時、件名、本文、未知キー・パス、配列添字・件数、生 Zod エラーや Graph payload を含めません。診断ログも追加しません。通信ログの `success: true` は、その後の正規化成功を意味しない点に注意してください。

## Outlook の添付

`outlook_list_attachments`、`outlook_inspect_attachment`、`outlook_read_attachment` の入力例と制約は [添付・ファイル解析](attachments.md) を参照してください。一覧はメタデータだけ、内容取得は fileAttachment だけです。item / reference 添付や外部 URL はたどらず、生 base64 はフロー → Worker 内部に限ります。

## OneDrive

接続所有者の OneDrive for Business をネイティブコネクタで扱います。共有ライブラリ、アクセスショートカット、任意 SharePoint ドライブ、共有リンクの追跡には対応しません。

### `onedrive_search_files`

MCP 入力は `query` と省略可の `limit`（既定 10、1〜100）。内部では native `FindFiles` の `query`、固定ルート、`findMode`、`maxFileCount` を使い、メタデータだけを返します。要求上限と同じ件数なら保守的に `truncated` とし、全件とは主張しません。

`OneDriveSearchMode` と `OneDriveSearchRootId` の Compose 入力は初期状態で空です。対象デザイナーで固定値を確認し、許可された手動更新で設定するまではコネクタに接続する前に失敗します。接続のバインドも別途必要です。

**検索は最大 100 件、継続入力も検索カーソルもありません。** 同じ結果を分割しても 100 件を超えた一致を取得できないためです。不完全なら検索語を絞るか、既知のフォルダーを一覧してください。

### `onedrive_list_folder`

MCP 入力は省略可の `folderId`、`limit`（既定 50、1〜100）、`cursor`。`folderId` 省略時は `ListRootFolder`、指定時は `ListFolderV2(id)` を使います。

指定フォルダーはネイティブページングのしきい値 1,000 件で集約します。最後のページがしきい値を超えてもフローは最大 1,000 件に制限します。ルートは返却配列のみで、`ListRootFolder` のネイティブ継続対応は主張しません。

`nextCursor` は同じフォルダー / ルートと `limit` で、**その有限の集約範囲の次の部分**を読むためのものです。毎回、全範囲を取得し直します。カーソルはスコープと順序付きメタデータの fingerprint に結び付き、結果が変われば継続を拒否します。そのときはカーソルなしで最初から読み直してください。カーソルにフォルダー ID、生メタデータ、上流 URL は含めません。サーバーキャッシュもありません。

- `hasMore` は、範囲内の既知の残りと上流の取得漏れの可能性を両方含みます。
- `nextCursor` は、範囲内に既知の残りがある場合だけ返します。
- 範囲末端で `hasMore: true` / `nextCursor: null` / `incompleteReason` があるなら、それ以上はこの操作で取得できません。
- ちょうど 1,000 件でも保守的に不完全とします。無制限の列挙でも安定したスナップショットでもありません。

フォルダー集約の通信 JSON 上限は 4 MiB。超過は安全に失敗し、黙ってメタデータを捨てません。MCP 結果は最大 100 件で、128 KiB の出力上限も適用します。Worker は nextLink URL をたどらず、入力にも受け付けません。

再取得は各ページで複数のネイティブ呼び出しを繰り返し、割り当てを消費し得ます。既存資料が参照する OneDrive コネクタの公表制限は 60 秒あたり 100 呼び出しですが、現行の契約・テナントの制限と集約遅延は別途確認してください。タイムアウトは失敗した読み取りであり、全件取得の証拠ではありません。可能なら小さなフォルダーで利用します。

参考: [OneDrive コネクタの操作と制限](https://learn.microsoft.com/en-us/connectors/onedriveforbusiness/) / [ネイティブページング](https://learn.microsoft.com/en-us/azure/logic-apps/logic-apps-exceed-default-page-size-with-pagination)

### `onedrive_get_metadata`

MCP 入力は `fileId`。`GetFileMetadata(id)` を使い、フローと Worker で ID を検証します。内部の応答例:

```json
{ "Id": "...", "Name": "...", "Size": 1, "MediaType": "...", "IsFolder": false, "LastModified": "...", "ETag": "..." }
```

MCP 出力は `fileId`、`name`、`size`、`contentType`、`isFolder`、`lastModifiedDateTime`、`eTag` と、`supportedFormat`、`readable`、nullable な `limitation`。省略された任意日時・版タグは null のままです。`Path`、`NameNoExt`、`DisplayName`、`FileLocator` は内部検証用で、MCP 出力に含めません。

### `onedrive_inspect_file` / `onedrive_read_file`

入力は `fileId`、read ではさらに `selection` が必要です。[添付と同じ選択形式](attachments.md#usage)を使います。両ツールは内部操作 `onedrive_get_content` を共有します。

Worker の先行メタデータ取得に加え、フローも最新メタデータを取得し、フォルダー、空ファイル、4 MiB 超、非対応拡張子・MIME 型をバイト取得前に拒否します。PDF / DOCX / XLSX のパーサーは Outlook 添付と共通です。

Logic Apps のバイナリ表現:

```json
{ "$content-type": "...", "$content": "padded standard base64" }
```

`$content` は既に base64 です。フローは形式・MIME・復号サイズ等を検証して `{ metadata, contentBytes }` を Worker にだけ返し、Worker も先行メタデータとの一致と実バイト数を検証します。

PDF は PDF.js に構造解析を任せます。「復号ストリーム 16 MiB まで」の保証はありません。raw / ページ / 選択 / 出力上限は維持しますが、タイマーは同期処理を止められず、プラットフォームの CPU・メモリ確認が必要です。[解析の上限](attachments.md#resource-limits)を参照してください。

## 運用上の注意

フローは対応アクションの安全な入出力設定を使い、Request トリガーの出力（引数・ヘッダー）も実行履歴で隠します。これは露出の軽減であり、Microsoft の保持を保証しません。エラーは ID・検索語・名前・URL・生コネクタエラー・バイト列を返さず、Worker のログ項目も増やしません。

ライブ import、接続のバインド、スモークテストには別途許可が必要です。このガイドや合成テストの成功を、ライブ接続確認とみなさないでください。
