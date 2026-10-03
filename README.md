# microsoft-bypass-mcp

Microsoft 365のOutlookやOneDrive等のサービスとChatGPT/Claudeなどのコネクタとの連携がテナント管理者によって許可されていない場合に、Power Automate経由でバイパスすることでアクセス可能にするMCP

## しくみ

```text
MCP クライアント
    ↓ MCP over Streamable HTTP (/mcp)
Cloudflare Access（外部 OAuth 認証・アクセス制御）
    ↓ Access JWT
Cloudflare Workers + Hono（リクエストごとに処理）
    ↓ HTTP POST { operation, requestId, args } + X-MCP-Gateway-Key
Power Automate
    ↓ Graph API
Microsoft 365 Services
```

## ツールリスト

| ツール | 用途 |
| --- | --- |
| `outlook_list_messages` | Outlookのメールのリストを表示 |
| `outlook_search_messages` | Outlookのメールを検索 |
| `outlook_get_message` | Outlookのメール本文を取得 |
| `outlook_list_mail_folders` | Outlookのメールフォルダ一覧を表示 |
| `outlook_get_conversation` | Outlookの指定したメールのやり取りを一斉取得 |
| `outlook_list_attachments` | Outlookの添付ファイルを取得 |
| `outlook_inspect_attachment` | Outlookの添付されたファイルの構造を取得 |
| `outlook_read_attachment` | Outlookの添付されたファイルの中身を取得 |
| `onedrive_search_files` | OneDriveのファイルを検索 |
| `onedrive_list_folder` | OneDriveの指定されたフォルダの中身のリストを表示 |
| `onedrive_get_metadata` | OneDriveの指定されたファイル・フォルダのメタデータを取得 |
| `onedrive_inspect_file` | OneDriveの指定されたファイルの構造を取得 |
| `onedrive_read_file` | OneDriveの指定されたファイルの中身を取得 |
| `outlook_create_draft` | Outlookで下書きを作成 |
| `outlook_create_reply_draft` | Outlookの特定のメールに対して返信を下書きとして作成 |
| `outlook_add_draft_attachment` | OneDrive |

## 導入の流れ

- Bun
- Python 3
- Cloudflare アカウント
- Microsoft 365 / Power Automate

```sh
bun install --frozen-lockfile
python3 scripts/build_attachment_flow.py --check
bun run test
bun run test:flow
```

詳細は **[セットアップガイド](docs/setup.md)** 

## ライセンス

MIT
