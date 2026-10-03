# Microsoft 365 MCP — 仕様

[README](README.md) · [セットアップ](docs/setup.md) · [読み取り](docs/read-tools.md) · [下書き](docs/drafts.md)

## 1. 目的と対象

利用権限のある Microsoft 365 のメール・添付・所有ファイルを読み、明示的に依頼された Outlook 下書きを保存する、固定機能の [Model Context Protocol (MCP)](https://modelcontextprotocol.io) サーバーです。組織の許可を得た接続を個人で運用することを想定します。ライセンス、テナント、DLP、コネクタや外部 AI 利用の制約を回避するものではありません。

汎用 Microsoft Graph MCP ではありません。Worker は Graph の認証や直接通信を行わず、Power Automate が固定の Outlook Graph 操作または OneDrive for Business のネイティブ操作を実行します。メール送信は提供しません。

## 2. 構成と境界

```text
MCP クライアント
    ↓ MCP over Streamable HTTP (/mcp)
Cloudflare Access
    ↓ 許可されたリクエスト + Access JWT
Cloudflare Workers + Hono
    ↓ HTTP POST { operation, requestId, args }
Power Automate（操作の許可リスト + 固定コネクタ操作）
    ↓
Microsoft 365
```

対応付けは `MCP tool → fixed operation → Power Automate switch case → fixed read/draft action` です。任意の URL、ルート、HTTP メソッド、Graph の生リクエスト body、OData、drive ID、site ID、共有リンク、ダウンロード URL、nextLink URL を入力として受け付けません。

Worker はリクエストごとに新しい MCP サーバーを生成します。メール・ファイル・抽出結果を DB / KV / R2 / キャッシュに永続化せず、バイト列はリクエスト内だけで扱います。ユーザーが保存を依頼した下書き・添付は Outlook に残ります。Access の署名鍵キャッシュはメール・ファイルの保存とは別です。

## 3. 認証とバインディング

本番の `/mcp` は Cloudflare Access の前段保護と、Worker 内の `Cf-Access-Jwt-Assertion` 検証を両方使います。署名鍵は `TEAM_DOMAIN` の `/cdn-cgi/access/certs` から取得し、署名・issuer・audience・有効期限を検証します。失敗時は HTTP 401 / `Unauthorized` です。

| バインディング | 必須条件 | 用途 |
| --- | --- | --- |
| `POWER_AUTOMATE_URL` | 常に必要 | Power Automate HTTP トリガーの完全な URL |
| `POWER_AUTOMATE_GATEWAY_KEY` | 常に必要 | `X-MCP-Gateway-Key` として送るキー |
| `TEAM_DOMAIN` | 本番で必要 | Cloudflare Access のチーム URL |
| `POLICY_AUD` | 本番で必要 | Access アプリケーションの AUD タグ |

ローカル開発では `TEAM_DOMAIN` / `POLICY_AUD` を両方省略できます。その場合、Worker の認証検証は無効です。片方だけなら設定エラーになります。実装は本番を自動判別して必須化しないため、運用者が両方を設定する必要があります。

秘密情報・実設定をコミットしません。Access アプリ・ポリシー、外部 OAuth とクライアント登録・接続条件は別途必要で、自動構築は同梱しません。詳しくは [セットアップ](docs/setup.md) を参照してください。

## 4. Power Automate 通信

Worker はリクエストごとに UUID v4 の `requestId` を生成し、次の形式で POST します。

```json
{ "operation": "string", "requestId": "uuid-v4", "args": {} }
```

フローの `manual` トリガーは `Request` / `Http` です。既存の `triggerAuthenticationType: "All"` と次の条件を保持します。

```text
@and(not(empty(parameters('McpGatewayKey'))),equals(triggerOutputs()?['headers']?['X-MCP-Gateway-Key'],parameters('McpGatewayKey')))
```

フローの `McpGatewayKey` は `SecureString` で、公開既定値は空です。空なら処理を許可しません。非公開の実値を `POWER_AUTOMATE_GATEWAY_KEY` と一致させます。`スイッチ` は `@triggerBody()?['operation']` を参照し、各分岐で閉じた引数スキーマと UUID v4 を検証してからコネクタ処理を行います。

成功応答:

```json
{ "ok": true, "requestId": "...", "operation": "...", "data": {} }
```

Worker は `requestId` と `operation` の一致を確認します。エラーは非 2xx、または秘匿化した `{ "ok": false, "error": { "code": "..." } }` です。Graph / ネイティブコネクタの生エラー、URL、メール・ファイル ID、検索語、件名、本文、バイト列、base64、nextLink をエラーに含めません。トリガー段階の拒否はプラットフォーム管理の応答になる場合があります。

フローにデータ保存・キャッシュ・任意 URL をたどるループはありません。Worker は POST を自動再試行せず、リダイレクトも追いません。既定の通信タイムアウトは 30 秒です。

## 5. 固定ツール一覧

読み取り 13 ツール、下書き書き込み 3 ツールの計 16 ツールです。

| MCP ツール | フロー操作 | 範囲 |
| --- | --- | --- |
| `outlook_list_messages` | `list_messages` | メール概要、固定スコープ・フィルター・ページ |
| `outlook_search_messages` | `search_messages` | 検索の先頭ページのみ |
| `outlook_get_message` | `get_message` | 指定メールの本文・詳細 |
| `outlook_list_mail_folders` | `list_mail_folders` | ルートフォルダーの先頭ページ |
| `outlook_get_conversation` | `get_conversation` | `/me/messages` 内の会話 ID 完全一致 |
| `outlook_list_attachments` | `list_attachments` | 添付メタデータのみ |
| `outlook_inspect_attachment` | `get_attachment` | Worker で添付の構造を解析 |
| `outlook_read_attachment` | `get_attachment` | 添付の指定範囲を抽出 |
| `outlook_create_draft` | `create_draft` | プレーンテキストの新規下書き保存 |
| `outlook_create_reply_draft` | `create_reply_draft` | 元メールへの返信下書きを保存。全員への返信ではない |
| `outlook_add_draft_attachment` | `add_draft_attachment` | 確認済み下書きへの小さな添付追加 |
| `onedrive_search_files` | `onedrive_search_files` | 所有ファイルのネイティブ検索 |
| `onedrive_list_folder` | `onedrive_list_folder` | 有限のフォルダー集約と、その範囲内のページ |
| `onedrive_get_metadata` | `onedrive_get_metadata` | 必要なメタデータだけを返す |
| `onedrive_inspect_file` | `onedrive_get_content` | Worker で所有ファイルの構造を解析 |
| `onedrive_read_file` | `onedrive_get_content` | 所有ファイルの指定範囲を抽出 |

添付と OneDrive の inspect / read がそれぞれ操作を共有するため、フローは 14 操作です。以下の `args` は内部通信形式です。MCP 入力の `limit` / `cursor` は Worker が `top` / `skip` 等に変換します。

## 6. Outlook の読み取り操作

### `list_messages`

```json
{
  "top": 1,
  "skip": 0,
  "mailbox": "inbox | sent | all",
  "folderId": "optional string",
  "filters": {
    "isRead": true,
    "hasAttachments": false,
    "receivedAfter": "ISO datetime",
    "receivedBefore": "ISO datetime"
  }
}
```

上の文字列は選択肢・型の説明です。実際の `mailbox` は `inbox`、`sent`、`all` のいずれか 1 つです。`top` は `1..50`、`skip` は `0..10000`。`folderId` があれば `mailbox` より優先し、固定 `/me/mailFolders/{folderId}/messages` を使います。省略時の経路は次のとおりです（Graph の `/v1.0` 配下）。

- `inbox` → `/me/mailFolders/inbox/messages`
- `sent` → `/me/mailFolders/sentitems/messages`
- `all` → `/me/messages`

概要の `$select` 順序:

```text
id,subject,from,receivedDateTime,sentDateTime,parentFolderId,conversationId,hasAttachments,importance,isRead,bodyPreview
```

フィルターなしなら `$orderby=receivedDateTime desc`。空でない限定フィルターがあれば Graph の `InefficientFilter` を避けるため `$orderby` を省略します。フィルター順は `isRead`、`hasAttachments`、`receivedAfter`、`receivedBefore` で固定です。日付は秒と最大 7 桁の小数秒を持つ UTC `Z` 形式で、下限は上限より前でなければなりません。受信日時の下限は `ge`、上限は `lt`。呼び出し元の OData は受け付けず、`$search` とフィルターを混ぜません。

Graph の `@odata.nextLink` は内部データとしてのみ扱います。Worker は origin・パス・クエリを検証し、検証済みの数値 `$skip` だけをカーソルにします。フローや Worker に URL を入力して再実行する機能はありません。未対応の継続形式なら、有効な現在ページを保ち、`hasMore: true`、`nextCursor: null`、固定 `PAGINATION_*` コードを含む `incompleteReason` を返します。件数から offset を推測しません。呼び出し元の不正カーソルは取得前に拒否します。

固定エンドポイント名だけ大文字・小文字の差を許容し、明示 ID とクエリ値は完全一致が必要です。OData のフォルダーキー形式や診断の詳細は [読み取りガイド](docs/read-tools.md) を参照してください。

### `search_messages`

内部 args は `{ "query": "trimmed nonempty string, max 512", "top": 1..50, "mailbox"?: "inbox|sent|all", "folderId"?: "string" }`。

経路選択は一覧と同じです。フィルター、skip、カーソル、任意 URL・メソッド・ルート・body は受け付けません。先頭ページだけを返し、nextLink は不完全な結果の信号としてのみ扱います。

### `get_message`

内部 args は `{ "messageId": "bounded ID" }`。固定 `/me/messages/{messageId}` を使います。詳細の `$select` は、現在の生成ソースでは概要に `toRecipients,ccRecipients,body` を続けた順序です。

```text
id,subject,from,receivedDateTime,sentDateTime,parentFolderId,conversationId,hasAttachments,importance,isRead,bodyPreview,toRecipients,ccRecipients,body
```

返された ID の一致を検証します。`from` が省略または `null` なら `{ name: "", address: "" }` に正規化します。他の属性から送信者を推測しません。非 null の `from` がある場合は `emailAddress` オブジェクト、必須の string-or-null `name`、必須の string `address` を要求し、不正な送信者は拒否します。`name: null` は空文字になります。この規則は概要一覧・検索・会話にも共通です。

### `list_mail_folders`

内部 args は `{ "top": 1..50 }`。固定 `/me/mailFolders` から、次のフィールドを選択します。

```text
id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount
```

子フォルダーを再帰的に取得しません。nextLink があれば不完全と報告し、URL は返さず、たどりません。

### `get_conversation`

内部 args は `{ "conversationId": "bounded ID", "top": 1..50, "skip": 0..10000 }`。固定 `/me/messages` を使うので、接続からアクセス可能なら送信済みメールも含みます。

```text
$filter=conversationId eq '<single-quote-doubled conversationId>'
```

詳細の `$select` は `get_message` と同じです。`InefficientFilter` を避けるため **`$orderby` は付けません**。Worker は返されたページ内だけを時刻順に整列し、会話全体の順序を保証しません。受信トレイの部分取得による代替は行わず、会話 ID 不一致とページ内 ID 重複を拒否します。

### 添付

`list_attachments` / `get_attachment` は [添付ガイド](docs/attachments.md) の契約に従います。固定のメール添付経路だけを使い、reference URL や nextLink はたどりません。一覧に `contentBytes` を返さず、fileAttachment の内容取得時のみ Worker 内部への通信に使います。読み取りの Graph `size` と実バイト数は同じとはみなさず、それぞれ独立に 4 MiB まで検証します。

## 7. OneDrive のネイティブ操作

対象は接続所有者の OneDrive for Business 内の所有ファイルです。Outlook HTTP コネクタや Graph プロキシ経路は使いません。許可する操作 ID は次のとおりです。

- `FindFiles(query,id,findMode,maxFileCount)` — `maxFileCount` は `1..100`
- `GetFileMetadata(id)`
- `GetFileContent(id,inferContentType)`
- `ListFolderV2(id)`
- `ListRootFolder()`

`shared_onedriveforbusiness` を所有者の接続に手動で結び付けます。公開ソースには接続 ID や認証秘密情報はありません。検索用 Compose の `OneDriveSearchMode`（確認済み `findMode` 機械値）と `OneDriveSearchRootId`（ネイティブルート ID）は空の固定入力で生成されます。許可された設定作業で確認・入力するまでは HTTP 503 で検索を拒否します。

### メタデータ

ネイティブ応答の `Path`、`NameNoExt`、`DisplayName`、`FileLocator` は内部検証用に許容しますが、外へ返すのは以下だけです。

```ts
type OneDriveMetadata = {
  Id: string
  Name: string
  Size: number
  MediaType: string
  IsFolder: boolean
  LastModified?: string | null
  ETag?: string | null
}
```

MCP 出力では `fileId`、`name`、`size`、`contentType`、`isFolder`、`lastModifiedDateTime`、`eTag` に正規化し、`supportedFormat`、`readable`、`limitation` を加えます。

### `onedrive_search_files`

内部 args は `{ "query": "trimmed nonempty string, max 512", "top": 1..100 }`。生成フローは native `FindFiles` の配列を上限まで選択して `value` と `truncated` を返します。件数が要求上限と同じなら、保守的に取得漏れの可能性を示します。Worker は継続の存在を示す情報も不完全判定に使いますが、ネイティブ検索の継続入力は提供しません。**最大 100 件で、検索カーソルはありません**。

### `onedrive_list_folder`

内部 args は `{ "folderId"?: "bounded ID", "top": 1..1000 }`。Worker は常に固定 1,000 件の集約範囲を要求し、MCP の `limit` は `1..100` のままです。

`folderId` 省略時は `ListRootFolder`、指定時は `ListFolderV2(id)`。指定フォルダーだけ `paginationPolicy.minimumItemCount: 1000` によるネイティブページングを使います。これは最小しきい値なので最後のページが超過することがありますが、フローは最大 1,000 件に切って返します。しきい値到達やネイティブ継続があれば不完全とします。ルートには未確認のページング設定を追加しません。

Worker は 4 MiB の通信上限内で最大 1,000 件を検証し、要求した MCP 件数だけを返します。カーソルは version・operation・folder/root のスコープ・limit・順序付き集約結果の fingerprint・有限 offset に結び付きます。**継続呼び出しごとに集約範囲を取得し直し**、変更、重複 ID、不正 offset を拒否します。保存・キャッシュはありません。カーソルは認可の証明ではありません。

`nextCursor` は取得済み範囲内の残りがある場合だけ返します。取得済み範囲を読み終えても上流が不完全なら、`hasMore: true`、`nextCursor: null` になります。無制限の列挙や変化しないスナップショットではありません。Worker はネイティブ nextLink URL を返さず、受け付けず、たどりません。[利用上の詳細](docs/read-tools.md#onedrive_list_folder)も参照してください。

### `onedrive_get_metadata`

内部 args は `{ "fileId": "bounded ID" }`。`GetFileMetadata(id)` の返却 ID をフローと Worker で検証します。

### `onedrive_get_content`

内部 args は `{ "fileId": "bounded ID" }`。フローは最新メタデータで ID、`IsFolder=false`、`0 < Size <= 4 MiB`、PDF / DOCX / XLSX の拡張子と一致または汎用 MIME 型を確認してから `GetFileContent(id,inferContentType=true)` を呼びます。

Logic Apps のバイナリ表現:

```json
{ "$content-type": "...", "$content": "base64..." }
```

`$content` は既にパディング付き標準 base64 です。フローは型・文字集合・パディング・符号化長・復号後の長さとメタデータの一致を検証し、Worker にだけ `{ metadata, contentBytes }` を返します。オブジェクト全体を base64 化しません。Worker は先行メタデータ取得との一致と復号サイズを再確認します。生コネクタオブジェクト、共有リンク、アクセスショートカット、無制限のバイト列は返しません。

## 8. 共通検証と文書の制約

フロー引数は操作ごとの閉じたスキーマです。ID は `1..2048` 文字で、制御文字、空白、NEL `U+0085`、BOM `U+FEFF`、完全一致の `.` / `..` を拒否します。パス ID は `uriComponent` で個別に符号化し、入力の `%` も再符号化するため構造上の区切りになりません。

検索語は trim 後に空でなく、最大 512 UTF-16 コード単位。通常の MCP 一覧は要求件数を超えず、Outlook は最大 50 件、OneDrive は最大 100 件です。内部フォルダー集約の 1,000 件上限は別に適用します。MCP リクエスト本文は 4 MiB、各ツール結果の JSON 表現は 128 KiB までです。

PDF / DOCX / XLSX のみを範囲指定で読み、OCR は行いません。OOXML の展開バイト数・XML 深さ等のガードを維持します。PDF の構造解析は PDF.js に任せ、手書きの文法許可リストや「復号ストリーム 16 MiB 上限」は使いません。PDF.js の内部割り当ては raw / ページ / 出力上限で制限できず、同期処理をタイマーで中断できません。Cloudflare の CPU・メモリ制約の確認が必要です。全上限は [解析ガイド](docs/attachments.md#resource-limits) に記載します。

メール・ファイル名・抽出内容は信頼できない外部データです。出典を保持し、内容の指示を実行したり、ユーザー承認とみなしたりしません。

## 9. 下書き書き込み

詳細契約は [下書きガイド](docs/drafts.md) に従います。3 ツールとも write で非冪等、`readOnlyHint: false` / `idempotentHint: false` です。MCP ホストがユーザーの明示的承認を得る必要があります。注釈だけで承認は実施されません。

固定の `/me/messages` 作成・返信・添付操作のみを許可し、送信、削除、汎用 HTTP、POST 自動再試行は提供しません。下書き添付は一致する ID と `isDraft: true` を取得で確認してから 1 回だけアップロードします。同時編集・送信とのトランザクションではありません。

アップロード成功には、有効な添付 ID、要求した名前、検証済み canonical base64 と完全一致する返却 `contentBytes` が必要です。Graph `size` は独立した非負 Int32 メタデータであり、実ファイル長とみなしません。返す `size` は検証済み実バイト数（1 byte〜2 MiB）で、Worker も入力サイズとの一致を確認します。返却バイト列の欠落・不一致は `DRAFT_WRITE_AMBIGUOUS` として停止し、再送や代替取得をしません。

タイムアウト、接続断、不正な成功応答でも、下書き・添付が既に保存されている可能性があります。再試行前に Outlook を確認します。`requestId` は相関用で、冪等性キーではありません。

## 10. プライバシーと検証範囲

Worker の通信ログは次だけです。

```ts
{ type: 'power_automate_request', requestId, operation, durationMs, status, success }
```

URL、キー、検索語、件名、宛先、ID、本文、ファイル名、バイト列、base64、nextLink をログに出しません。MCP 出力には、依頼された有限の読み取り結果や下書き結果に必要な項目だけを含め、生コネクタ応答や通信資格情報を含めません。`success: true` は通信 envelope の検証成功であり、その後の内容正規化成功とは別です。

Power Automate は対応するアクションの安全な入出力設定で実行履歴の露出を減らします。Microsoft 側の保持・監査方針や、MCP ホスト側の保存を保証するものではありません。

オフラインテストは合成データで、固定操作、経路、ネイティブ操作 ID、秘匿化、引数・サイズ制約、生成結果の一致を検証します。ライブフローの import / 保存 / 実行は行いません。既存の成功記録は過去の確認に限られ、大規模な実ファイル群の検証を意味しません。ライブ実行、コネクタ接続変更、デプロイ、マージには別途許可が必要です。公開 JSON は単独でインポートできるパッケージではありません。
