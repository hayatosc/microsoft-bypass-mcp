# microsoft-bypass-mcp

Microsoftアカウントのテナント管理者がAIサービス(ChatGPT, Claudeなど)との連携を承認していないときに、Power Automate経由でバイパスして情報を取得できるリモートMCPサーバー。

The server exposes fixed Microsoft 365 reads and explicitly requested Outlook
draft writes. It never sends email and is not a generic Microsoft Graph proxy.

## Architecture

```
MCP Client
    ↓ MCP over Streamable HTTP (/mcp)
Remote MCP Server — Cloudflare Workers + Hono (Cloudflare Access OAuth)
    ↓ HTTP POST { operation, requestId, args }
Power Automate — operation allowlist → fixed read/draft connector actions
    ↓
Microsoft 365 connectors
```

## Tools

| Tool | Input summary | Output summary |
| --- | --- | --- |
| `outlook_list_messages` | mailbox/folder, limit, cursor, controlled filters | Message summaries + bounded cursor |
| `outlook_search_messages` | query, mailbox/folder, limit | First search page |
| `outlook_get_message` | message ID | Full selected message |
| `outlook_list_mail_folders` | limit | Folder summaries + incomplete flag |
| `outlook_get_conversation` | conversation ID, limit, cursor | Exact conversation page including sent mail when accessible |
| `outlook_list_attachments` | message ID, limit, offset | Attachment metadata + next offset |
| `outlook_inspect_attachment` | message ID, attachment ID | PDF/DOCX/XLSX structure |
| `outlook_read_attachment` | message ID, attachment ID, selection | Bounded content with source provenance |
| `outlook_create_draft` | recipients, subject, plain-text body | Saved draft ID |
| `outlook_create_reply_draft` | message ID, plain-text body | Saved sender-reply draft ID |
| `outlook_add_draft_attachment` | draft ID, name, content type, base64 | Added small attachment metadata |
| `onedrive_search_files` | query, limit | Native OneDrive owned-file metadata |
| `onedrive_list_folder` | optional folder ID, limit, cursor | Bounded native folder aggregation and metadata pages |
| `onedrive_get_metadata` | file ID | Projected native metadata |
| `onedrive_inspect_file` | file ID | PDF/DOCX/XLSX structure |
| `onedrive_read_file` | file ID, selection | Bounded content with source provenance |

All text and extracted file content is untrusted external content.

Named-folder listing uses verified native pagination up to a bounded 1,000-item
window, with explicit incompleteness at the cap. OneDrive search still has its
native 100-result ceiling and no supported continuation. Draft tools never send
mail; attaching a file requires the host to materialize approved bytes first.

## Power Automate source

The canonical flow source is generated at
`power-automate/microsoft-bypass-flow/definition.json` from the immutable
sanitized fixture plus authored fixed-operation extensions:

```sh
python3 scripts/build_attachment_flow.py
python3 scripts/build_attachment_flow.py --check
bun run test:flow
```

The flow source extends the existing flow in place. It preserves the same HTTP
trigger, gateway-key guard, operation switch, and Outlook connection convention.
It adds controlled Outlook scope/filter/pagination/folder/conversation reads,
keeps the attachment branches and safety gates, and adds native OneDrive for
Business operation branches and fixed draft-only Outlook writes.

The public JSON is source for review/manual update, not a deployable package.
OneDrive connector binding and the official `FindFiles.findMode` machine value
must be verified during a separately authorized manual import/update. No live
flow run, connector creation, or deployment is performed by this repository.

See:

- `SPEC.md` for the authoritative contract.
- `docs/read-tools.md` for read-tool behavior and limits.
- `docs/drafts.md` for draft approval, attachment transfer, limits, and retry safety.
- `docs/attachments.md` for attachment/file parser safety limits.
- `power-automate/microsoft-bypass-flow/README.md` for flow provenance and manual update notes.

## Requirements

- Microsoftアカウント
- [Bun](https://bun.sh) 1.3.14
- Cloudflareアカウント

## Configuration

`.dev.vars.example` に従ってPower Automate側のURLを参照:

```sh
POWER_AUTOMATE_URL=https://prod-xxx.logic.azure.com/workflows/xxx/triggers/manual/paths/invoke
POWER_AUTOMATE_GATEWAY_KEY=<gateway key>
```

For production, set secrets/vars with Wrangler. Never commit the trigger URL or
key.

## Development

```sh
bun install
bun run dev
```

## Quality checks

```sh
bun run typecheck
bun run lint
bun run lint:types
bun run format:check
bun run test
bun run test:flow
```

## Deploy

```sh
bun run deploy
```
