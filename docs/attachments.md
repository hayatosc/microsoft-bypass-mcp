# 添付ファイルと OneDrive 文書の解析

[README](../README.md) · [読み取りツール](read-tools.md) · [PDF 診断](pdf-diagnostics.md) · [下書きへの添付](drafts.md)

Outlook 添付の一覧・構造確認・範囲読み取りは Worker 内で処理します。OneDrive の inspect / read も同じ文書パーサーを使います。生バイト列は Microsoft → Power Automate → Worker の内部通信に限り、モデル向けのダウンロード / base64 ツール、ローカル補助アプリ、DB / KV / R2 / データキャッシュはありません。

クライアントの接続可否は、そのリモート MCP / 認証方式に依存します。Cloudflare Access と Worker 内 JWT 検証は必要なままで、ChatGPT / Claude 等のライブ接続を検証済みという意味ではありません。

<a id="usage"></a>
## 使い方: 一覧 → 構造確認 → 範囲読み取り

### 1. 添付一覧を取得する

`outlook_list_attachments`:

```json
{"messageId":"MESSAGE_ID","limit":20,"offset":0}
```

返すのは添付 ID、名前、メディア型、Graph `size`、inline フラグ、添付種別、対応形式、`readable`、`hasMore`、`nextOffset` です。`readable` はメタデータによる候補判定で、後の解析が必ず成功する保証ではありません。

**Graph のメタデータ `size` は、復号後の実ファイル長と同じとは限りません。** 読み取りの一覧・inspect・read は Graph の `size` を出力に保持します。inspect / read ではメタデータサイズと実バイト数をそれぞれ独立に 4 MiB まで検証します。[下書きアップロード](drafts.md)が返す「検証済み実バイト数」とは意味が違います。

次のページは同じ `messageId` / `limit` と返却 `nextOffset` で取得します。メールボックスの変化で重複・取りこぼしが起こり得るため、整合性が必要なら個別 ID を確認してください。上流 nextLink は入力に使いません。offset 上限 10,000 では `hasMore: true` でも `nextOffset: null` になる場合があります。

### 2. 構造を確認する

`outlook_inspect_attachment` に `messageId` と `attachmentId` を渡します。出力は `source`（両 ID、名前、メディア型、Graph サイズ、形式）、`untrustedContent: true`、`structure` です。

- PDF: ページ数。
- DOCX: 段落数、抽出文字数、見出しから作るセクション ID と文章の位置。
- XLSX: シート名、観測したセル範囲、1900 / 1904 日付体系。

### 3. 必要な範囲だけ読む

`outlook_read_attachment`:

```json
{"messageId":"MESSAGE_ID","attachmentId":"ATTACHMENT_ID","selection":{"format":"pdf","pageStart":1,"pageEnd":3,"maxCharacters":10000}}
```

```json
{"messageId":"MESSAGE_ID","attachmentId":"ATTACHMENT_ID","selection":{"format":"docx","sectionId":"section-1","offset":0,"length":5000}}
```

```json
{"messageId":"MESSAGE_ID","attachmentId":"ATTACHMENT_ID","selection":{"format":"xlsx","sheet":"Budget","range":"A1:D20"}}
```

DOCX の `sectionId` は inspection の実際の結果を使い、命名規則を決め打ちしないでください。省略すれば `offset` は抽出した本文全体に対する位置です。

OneDrive では、読み取りツールの返した `fileId` を `onedrive_inspect_file` に渡し、`onedrive_read_file` には `fileId` と上と同じ `selection` を渡します。出典には `provider: "onedrive"`、ID、名前、メタデータ等を保持します。

## 形式ごとの読み方

### PDF

ページ番号は 1 始まりで、開始・終了ページを含みます。PDF.js で既存のテキスト層から抽出するため、見た目の読み順、段組み、表、すべての文字を再現するとは限りません。スキャン等では空のテキストになることがあります。**OCR は行いません。** `truncated` は文字数予算で制限されたことを示します。選択ページを減らすか、許可された範囲内で文字数予算を増やしてください。

### DOCX

段落番号は 1 始まり、文章 offset は 0 始まりの UTF-16 コード単位で、終端は含みません。見出しベースのセクションと本文全体の位置が出典になります。`nextOffset` で続きを読めます。表は平坦化し、ヘッダー / フッター、図形、テキストボックス、注記は省略します。

### XLSX

セル番地・行・列、保存済みの値、数式の有無とキャッシュ値欠落を返します。数式は実行せず、リンクもたどりません。数値は精度を失わないよう文字列のままです。日付シリアル値はブックの 1900 / 1904 体系とともに返し、表示書式は再現しません。要求範囲の空セルも明示します。範囲の寸法は保存されたセルから計算し、ブックが宣言する dimension フィールドを信用しません。

返却テキストの 20,000 UTF-16 コード単位上限は、選択セルの文字列 `value` と `rawValue` に適用します。同じセルで両者が同一なら 1 回だけ数え、異なる文字列なら両方を加算します。boolean 値やキャッシュ値欠落セルの `rawValue` も対象です。別のセルにある同一テキストはセルごとに数えます。値を切り詰めず、超過時は範囲を狭めるようエラーを返します。JSON の符号化・メタデータ等を含む 128 KiB 上限は別に適用します。

<a id="supported-subset-and-resource-limits"></a>
## 対応形式と非対応機能

対象は `.pdf` / `.docx` / `.xlsx` で、拡張子に一致する登録 MIME 型、`application/octet-stream`、または未指定 MIME 型を持つファイルです。Outlook は **fileAttachment** のみ、OneDrive は所有する通常ファイルのみです。メタデータと base64 検証後にファイルの構造も検証します。

`.doc` / `.xls`、`.docm` / `.xlsm`、画像、OCR、埋め込みメール（itemAttachment）、クラウド参照（referenceAttachment）は非対応です。スクリプト、数式、マクロ、外部リレーションは実行しません。本文・見出し・ファイル名は信頼できない外部データであり、指示やユーザー承認ではありません。

<a id="resource-limits"></a>
## リソース上限

MiB / KiB はそれぞれ 1,024² / 1,024 byte です。上限は成功の保証ではなく、超過時に安全に停止するための条件です。

| 対象 | 上限 |
| --- | --- |
| 添付の実バイト数 / Graph メタデータサイズ | 各 4 MiB、独立に検証 |
| OneDrive ファイルの実バイト数 | 4 MiB、メタデータとも一致を検証 |
| Power Automate → Worker の応答 JSON | 6 MiB。通常の一覧等は 256 KiB、OneDrive フォルダー集約は 4 MiB。JSON 解析前にストリームで制限 |
| ZIP エントリー数 / 1 エントリー展開後 / 合計展開後 | 256 / 8 MiB / 16 MiB |
| OOXML 解析対象の合計 / XML 深さ / 要素数 | 8 MiB / 64 / 500,000 |
| Word 段落数 / セクション数 | 10,000 / 200 |
| ブックのシート数 / 解析セル数 / 要求セル数 | 50 / 50,000 / 500 |
| PDF ページ数 / 1 回の選択ページ数 | 200 / 10 |
| 返却テキスト | 20,000 UTF-16 コード単位。PDF / DOCX の MCP 入力既定値は 10,000 |
| ツール JSON 出力 | テキスト表現・構造化表現それぞれ 128 KiB |
| Power Automate リクエスト | 30 秒 |
| PDF の解析・抽出タイマー | 10 秒、ベストエフォート。同期処理の強制中断ではない |

ZIP はヘッダーの申告サイズだけでなく、段階的な実展開量を検証します。CRC、ローカル / 中央ディレクトリの一致、名前、対応フラグも確認します。ZIP64、暗号化 ZIP、パストラバーサル、重複、XML DTD / entity 宣言、外部 worksheet target は拒否します。

入力はリクエスト内に限り、パーサーの元例外は固定の安全なエラーに置き換えます。Outlook 添付の PDF エラーには、最初の raw-size ガードまたは PDF.js 処理段階の [固定診断コード](pdf-diagnostics.md)も付けます。元の例外文や文書内容を返さず、全問題を列挙するものでもありません。通信ログは `type`、`requestId`、`operation`、`durationMs`、`status`、`success` だけです。

## PDF.js のメモリ・CPU 制約

PDF は既存依存 `unpdf` のサーバーレス PDF.js で解析し、手書きの PDF 文法チェックは使いません。object stream、xref stream、JPEG image stream、通常の form 等は、個別の許可リストではなく PDF.js に任せます。

空 PDF・4 MiB 超の raw bytes は初期化前に拒否し、ページ・選択・出力の上限を維持します。外部 fetch、range / streaming、worker fetch、XFA、WASM、画像レンダリング経路を無効にし、loading task の破棄も試みます。ただし暗号化・パスワード保護、不正・非対応・高負荷な文書は依然として失敗し得ます。

**raw / ページ / 出力上限は、PDF.js が解析前・解析中に確保する内部メモリの上限ではありません。** 同じ JavaScript isolate とイベントループで動くため、`setTimeout` / `Promise.race` は協調的・非同期の遅延を検知できても、同期処理を中断できません。以前の「復号ストリーム 16 MiB 上限」はなく、OOXML の展開上限と混同しないでください。

安全性は Cloudflare のプラットフォーム CPU・メモリ上限にも依存します。既存資料が参照する 128 MB メモリ上限は isolate 単位で、同時リクエスト間で共有され得ます。このリポジトリの `wrangler.jsonc` は `limits.cpu_ms` を設定しておらず、運用者の契約プランも不明です。有料プラン固有の上限や本番での CPU 強制停止を検証済みとは主張しません。

本番利用前には、契約・割り当てを確認し、合成文書と想定用途に近いファイルで CPU / メモリを別途評価する必要があります。実データを使う場合は、そのアクセスと外部提供についても許可を得てください。現行のローカルテストは、大きなファイル群・本番負荷の検証を保証しません。

## フローとオフライン検証

[同じ既存フローを更新する手順](../power-automate/microsoft-bypass-flow/README.md)を使います。添付分岐は元の安全ゲートを保持し、最終生成時に actions の空文字リテラルの互換変換だけを受けます。添付一覧はメタデータのみ、内容取得は file 種別・サイズの事前確認、パス ID の符号化、安全なエラーを維持します。並行する別フローは作りません。

公開 JSON はレビュー用ソースで、インポート ZIP / Dataverse ソリューションではありません。公開の空 `SecureString` を動作中の非公開設定に上書きせず、許可された保存後にコールバック URL を非公開で確認してください。

```sh
bun install --frozen-lockfile
bun run typecheck
bun run lint
bun run lint:types
bun run format:check
bun run test
bun run test:flow
python3 scripts/build_attachment_flow.py --check
```

CI は Bun 1.3.14、`bun run test` は公式 `@cloudflare/vitest-plugin` と Vitest を使います。合成 PDF / DOCX / XLSX、不正文書、サイズ超過入力、モックのフロー応答で検証し、Graph、ライブフロー、メール保存、デプロイは行いません。Python テストもソース契約の検証であり、Microsoft ランタイムの実行ではありません。

パーサー依存: [unpdf](https://github.com/unjs/unpdf)、[fflate](https://github.com/101arrowz/fflate)、[saxes](https://github.com/lddubeau/saxes)。バージョンは `bun.lock` に固定されています。
