# microsoft-bypass-mcp

自分に利用権限のある Microsoft 365 のメールやファイルを、Power Automate 経由で MCP クライアントから読むための、個人運用向けリモート MCP サーバーです。Outlook の下書き保存にも対応しますが、**メールは送信しません**。

組織の承認やアクセス制御を回避するためのものではありません。大学・組織の規程、外部 AI へのデータ提供の許可、Microsoft 365 / Power Automate のライセンス、テナントの DLP ポリシーやコネクタ制限を確認してから利用してください。Microsoft アカウントがあるだけで使えるとは限りません。

## しくみ

```text
MCP クライアント
    ↓ MCP over Streamable HTTP (/mcp)
Cloudflare Access（外部 OAuth 認証・アクセス制御）
    ↓ Access JWT を Worker 内でも検証
Cloudflare Workers + Hono（リクエストごとに処理）
    ↓ HTTP POST { operation, requestId, args } + X-MCP-Gateway-Key
Power Automate（固定の操作だけを許可）
    ↓ Outlook の固定 Graph 操作 / OneDrive for Business のネイティブ操作
接続所有者の Microsoft 365
```

Worker は Microsoft Graph に直接接続しません。任意の URL・HTTP メソッド・Graph クエリを渡す汎用プロキシではなく、**16 ツール（読み取り 13、下書き書き込み 3）を 14 の固定操作に対応付けます**。

## できること

| ツール | 用途 |
| --- | --- |
| `outlook_list_messages` | 受信・送信・全体・指定フォルダーのメール概要。限定フィルターとカーソル対応 |
| `outlook_search_messages` | メール検索の先頭ページのみ |
| `outlook_get_message` | 指定メールの本文・詳細 |
| `outlook_list_mail_folders` | メールのルートフォルダー一覧。取得漏れの可能性を明示 |
| `outlook_get_conversation` | 指定会話のメール。アクセス可能な送信済みメールも対象 |
| `outlook_list_attachments` | 添付ファイルのメタデータ一覧 |
| `outlook_inspect_attachment` | PDF / DOCX / XLSX 添付の構造確認 |
| `outlook_read_attachment` | 添付の指定ページ・文章範囲・セル範囲を読む |
| `onedrive_search_files` | 所有する OneDrive ファイルの検索。最大 100 件、検索カーソルなし |
| `onedrive_list_folder` | ルートまたは指定フォルダーの一覧。指定フォルダーは最大 1,000 件の集約範囲内でページング |
| `onedrive_get_metadata` | ファイル・フォルダーのメタデータ |
| `onedrive_inspect_file` | 所有する PDF / DOCX / XLSX の構造確認 |
| `onedrive_read_file` | 所有ファイルの指定範囲を読む |
| `outlook_create_draft` | 明示的な承認を受けて新規下書きを保存 |
| `outlook_create_reply_draft` | 元メールへの返信下書きを保存。全員への返信ではない |
| `outlook_add_draft_attachment` | 下書きであることを確認し、小さな添付を追加 |

ファイル読み取りは最大 4 MiB、下書きへの添付は 1 回 2 MiB までです。OCR、画像抽出、共有ライブラリやショートカットの追跡には対応しません。OneDrive の一覧カーソルは毎回取得し直す有限の集約結果を分割するもので、無制限の列挙や安定したスナップショットではありません。

## 導入の流れ

必要なのは Bun **1.3.14**、Python 3、Cloudflare アカウント、許可された Microsoft 365 / Power Automate 環境、およびリモート MCP と選択した外部 OAuth 認証に対応するクライアントです。

1. **組織の許可・契約・接続権限を確認する。** 読み取りだけでなく下書きの書き込みも対象です。
2. **このチェックアウトで依存関係とオフライン検証を用意する。**
   ```sh
   bun install --frozen-lockfile
   python3 scripts/build_attachment_flow.py --check
   bun run test
   bun run test:flow
   ```
3. **許可された既存 Power Automate フローを手動更新する。** [フローの手順](power-automate/microsoft-bypass-flow/README.md)に従い、接続、`McpGatewayKey`、操作分岐を設定します。公開 `definition.json` はレビュー用ソースで、単独でインポートできるパッケージではありません。既存フローがない場合の初期構築は同梱していません。
4. **Worker の 4 つのバインディングを設定する。** `POWER_AUTOMATE_URL`、`POWER_AUTOMATE_GATEWAY_KEY`、`TEAM_DOMAIN`、`POLICY_AUD`。URL とキーを公開しないでください。
5. **Cloudflare Access と外部 OAuth / クライアント側の接続条件を整える。** Access アプリ・ポリシーや OAuth の登録・連携は別途必要です。このリポジトリはそれらの自動構築を提供しません。
6. **各種チェックとレビューを終え、別途許可された公開・接続確認へ進む。** 本番で Access の前段保護と Worker 内の JWT 検証を両方有効にします。

設定例、ローカル起動、全チェック、公開前の確認事項は **[セットアップガイド](docs/setup.md)** にまとめています。`bun run dev` は実際のフローを呼べるため、オフラインテストとは区別してください。

## 安全に使うために

- メール本文、ファイル名、抽出した文章・セルは**信頼できない外部データ**です。内容に書かれた指示を実行したり、操作の承認とみなしたりしないでください。
- Worker はメール・ファイルを DB / KV / R2 / キャッシュへ保存しません。明示的に保存した下書き・添付は Outlook に残ります。MCP クライアントや Microsoft 側の保持方針は別途確認が必要です。
- 読み取り結果に生バイト列や base64、ダウンロード URL は返しません。添付を下書きへ渡すには、ホストが承認済みの実バイト列を用意する必要があります。
- 下書き書き込みは**非冪等**です。タイムアウトなどで結果が不明な場合は、Outlook の下書きと添付を確認してから判断し、むやみに再実行しないでください。
- フロー URL、キー、接続情報を Git・ログ・スクリーンショットに含めないでください。Power Automate の安全な入出力設定は実行履歴の露出を減らしますが、サービス側の非保持を保証しません。
- PDF は PDF.js で解析します。ページ数や出力の上限だけでは内部メモリ使用量を制限できず、本番の CPU・メモリ制約の確認が必要です。

## ドキュメント

- [SPEC.md](SPEC.md) — 正式な仕様、固定操作、通信・セキュリティ境界
- [セットアップ](docs/setup.md) — 手動作業、認証、4 バインディング、Bun / Vitest の開発手順
- [読み取りツール](docs/read-tools.md) — 入力、ページング、不完全な結果、送信者の正規化
- [下書きツール](docs/drafts.md) — 承認、添付の受け渡し、サイズ検証、再試行の注意
- [添付・ファイル解析](docs/attachments.md) — PDF / DOCX / XLSX の使い方と制限
- [PDF 診断コード](docs/pdf-diagnostics.md) — 内容を漏らさない固定エラーの見方
- [Power Automate フロー](power-automate/microsoft-bypass-flow/README.md) — 正規ソースの由来と既存フローの更新
- [テスト用フィクスチャ](scripts/fixtures/README.md) — 変更しない履歴資料

ローカルテストは合成データとモックを使います。成功しても Microsoft 側の実行、クライアント接続、実ファイルの大規模検証が済んだことにはなりません。

## ライセンス

このリポジトリのライセンスはまだ指定されていません。再利用の条件はメンテナーに確認してください。
