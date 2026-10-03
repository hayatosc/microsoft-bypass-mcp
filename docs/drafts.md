# Outlook の下書きツール

[README](../README.md) · [正式な仕様](../SPEC.md) · [読み取りの添付サイズ](attachments.md)

接続所有者のメールボックスに下書きを保存します。メール送信・削除、任意 Graph リクエスト、他ユーザーへのなりすましは提供しません。3 ツールの注釈は `readOnlyHint: false`、`idempotentHint: false` で、既存の読み取りツールは読み取り専用のままです。

## 承認と権限

**下書き保存や添付アップロードの前に、MCP ホストがユーザーの明示的な承認を得る必要があります。** ツール注釈は機能のヒントであって承認システムではありません。作成の承認は送信の承認ではなく、確認・送信は Outlook の通常の操作で行います。

Microsoft の API 仕様では、これらの固定操作に委任された `Mail.ReadWrite` が必要です。既存 Office 365 Outlook 接続の同意やテナントポリシーが許可するとは限りません。許可された導入時に既存接続を確認し、無断で権限追加・再認証・コネクタ作成をしないでください。実装は `Mail.Send` を使う送信操作を追加しません。

## 操作の順序

1. `outlook_create_draft` に `to`、任意の `cc` / `bcc` 配列、`subject`、プレーンテキストの `body` を渡します。成功すると `draftId` と `isDraft: true` が返ります。
2. 返信なら `outlook_create_reply_draft` に元の `messageId` とプレーンテキストの `body` を渡します。これは元メールへの返信下書きで、全員への返信ではありません。返信先は Graph の `createReply` に委ねます。元メールの `replyTo` が送信者と異なる場合があるため、保存された下書きの宛先を Outlook で確認してください。
3. 必要なら、返された `draftId` と `name`、`contentType`、canonical base64 の `contentBytes` で `outlook_add_draft_attachment` を呼びます。
4. Outlook で本文・宛先・添付を確認します。送信はこのサーバーの機能外です。

添付追加は別操作なので、失敗しても先に保存した下書きは残ります。フローは対象の ID 一致と `isDraft: true` を読み取りで確認してから添付を追加し、送信済みメールには追加しません。ただし、事前確認と同時編集・送信はトランザクションではありません。アップロード中に同じ下書きを編集・送信しないでください。

## 入力の上限とファイルの受け渡し

| 入力 | 制約 |
| --- | --- |
| `to` / `cc` / `bcc` | 合計最大 50 宛先。新規下書きの `to` は 1 件以上 |
| `subject` / `body` | 最大 512 / 20,000 UTF-16 コード単位。本文はプレーンテキスト |
| 添付 | 1 回につき 1 ファイル、実バイト数 1 byte〜2 MiB |
| `name` | 1〜255 文字。パス区切り・制御文字なし |
| `contentType` | 最大 127 文字の単純な MIME 型。パラメーターなし |
| `contentBytes` | 厳密な canonical base64。data URL、remote URL、upload session、クラウド参照、ローカルパスは不可 |

**チャットにアップロードしたファイルを、この MCP サーバーが自動で読めるわけではありません。** ホストは、自身のファイル API でユーザー承認済みのファイルを実体化し、正確なバイト列を取得してサイズを確認し、そのバイト列を base64 化して渡す必要があります。チャットのダウンロード URL、ローカルパス、Library ID を `contentBytes` の代わりに使えません。

ホストに実体化機能がなければ、ユーザーに Outlook で添付してもらってください。サーバーの 2 MiB 上限内でも、大きな base64 引数がホスト側のツール呼び出し上限を超える場合があります。

成功応答は有限の ID、名前、検証済みの**実ファイル** `size` だけで、バイト列・base64・生 Graph オブジェクトは含めません。Worker とフローはデータをリクエスト内だけで扱い、依頼された下書きと添付は Outlook に残ります。実行履歴の露出軽減は、Microsoft 側の非保持の保証ではありません。

## 添付検証と `size` の意味

Microsoft の [fileAttachment](https://learn.microsoft.com/en-us/graph/api/resources/fileattachment?view=graph-rest-1.0) は `size` を Int32 の添付サイズ、`contentBytes` を base64 の内容と定義していますが、`size` と復号後の実ファイル長が同じとは保証しません。[添付 POST](https://learn.microsoft.com/en-us/graph/api/message-post-attachments?view=graph-rest-1.0) の 201 応答例には `contentBytes` が含まれています。

下書きの確認と 1 回のアップロード後、フローは次を検証します。

1. 有効な返却添付 ID と要求した名前。Graph メタデータの `size` は非負 Int32（`0..2147483647`）に制限し、2 MiB の実ファイル上限と混同しません。
2. 返却 `contentBytes` が有限の文字列で、要求と数値としての文字列長が等しく、`contains(returnedContentBytes, requestedContentBytes)` が成立すること。[Microsoft の関数仕様](https://learn.microsoft.com/en-us/azure/logic-apps/expression-functions-reference#contains)では `contains()` は大文字・小文字を区別します。空でない canonical な要求と同じ長さのため、全内容の一致と最大 2 MiB の実バイト制約を確認できます。`equals()` の文字列比較の大文字・小文字挙動には依存しません。
3. 検証済みの要求の実バイト数を応答 `size` にします。Worker も入力との一致と実バイト上限を検証します。

メタデータの overhead を推測したり、フローで復号したり、内容を MCP に返したりしません。合成データで raw 889 byte / Graph size 1223 なら返すのは `size: 889`。raw がちょうど 2 MiB でも Graph size がそれより大きいだけでは拒否しません。固定の差分式はありません。

返却バイト列が欠落、不正、過大、非 canonical、不一致なら **`DRAFT_WRITE_AMBIGUOUS`** で停止します。再アップロードや代替の読み取りは行いません。文書化されたバイト列を返さないコネクタでは手動確認が必要で、すべてのテナントが同じ応答を返すとは主張しません。**失敗を返した時点で既に添付が保存されている可能性があります。**

読み取り専用の一覧・inspect・read は Graph メタデータの `size` を保持し、inspect / read ではメタデータと実バイト数を各 4 MiB まで独立に検証します。下書きアップロードのサイズ意味論はそれらを変更しません。[添付ガイド](attachments.md)を参照してください。

## 失敗と再試行

下書き・添付の POST アクションは再試行を無効にしています。Worker も自動再試行しません。タイムアウト、接続断、不正な成功応答は、Microsoft が保存を終えた後にも起こり得ます。

**不明な結果なら Outlook の下書きフォルダーと対象の添付を先に確認してください。むやみに再実行すると重複を作ります。** `requestId` は相関用であり、冪等性キーではありません。exactly-once の保証はありません。

通信段階の失敗は `PowerAutomateAmbiguousWriteError`、成功データの不正は `DraftError` として、下書き確認の必要性を示します。フローには `DRAFT_PREFLIGHT_FAILED`、`INVALID_DRAFT_PREFLIGHT`、`DRAFT_NOT_VERIFIED` といったアップロード前の固定拒否もありますが、Worker は書き込み操作の非 2xx を保守的に曖昧な結果として扱います。

宛先、件名、本文、ファイル名、ID、バイト列をログに出しません。既存の有限な通信ログ項目だけを使い、生コネクタエラーは返しません。

## 固定の Microsoft 操作

| 操作 | 固定経路 |
| --- | --- |
| [新規下書き](https://learn.microsoft.com/en-us/graph/api/user-post-messages) | `POST /v1.0/me/messages` |
| [返信下書き](https://learn.microsoft.com/en-us/graph/api/message-createreply) | `POST /v1.0/me/messages/{encodedMessageId}/createReply` |
| [小さな添付](https://learn.microsoft.com/en-us/graph/api/message-post-attachments) | `POST /v1.0/me/messages/{encodedDraftId}/attachments`。固定の事前取得で下書きを検証した場合のみ |

[Office 365 Outlook の HTTP アクション](https://learn.microsoft.com/en-us/connectors/office365/#send-an-http-request)は `/me/messages` 系をサポートします。ただし、生成フローはレビュー・更新用の資料であり、ライブ成功の証拠ではありません。import、同意確認、デプロイ、実際の下書きスモークテストは別途許可が必要です。オフラインテストは合成メールデータだけを使います。
