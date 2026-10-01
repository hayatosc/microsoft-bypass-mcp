# microsoft-bypass-mcp

Microsoftアカウントのテナント管理者がAIサービス(ChatGPT, Claudeなど)との連携を承認していないときに、Power Automate経由でバイパスして情報を取得できるリモートMCPサーバー

## Specification

### Architecture

```
MCP Client
    ↓  MCP over Streamable HTTP (/mcp)
Remote MCP Server — Cloudflare Workers + Hono (Cloudflare Access OAuth)
    ↓  HTTP POST { operation, requestId, args }
Power Automate — operation allowlist → fixed Graph endpoints, M365 auth
    ↓
Microsoft Graph
```

### Tools

| Tool | Input | Output |
| --- | --- | --- |
| `outlook_list_messages` | `{ limit?: number }` (default 5, 1–50) | `{ messages: MessageSummary[], hasMore: boolean }` |
| `outlook_search_messages` | `{ query: string, limit?: number }` (default 10, 1–50) | `{ messages: MessageSummary[], hasMore: boolean }` |
| `outlook_get_message` | `{ messageId: string }` | `MessageDetail` |
| `outlook_list_attachments` | `{ messageId, limit?, offset? }` | Metadata + next offset |
| `outlook_inspect_attachment` | `{ messageId, attachmentId }` | PDF pages, DOCX sections or XLSX sheets |
| `outlook_read_attachment` | `{ messageId, attachmentId, selection }` | Bounded content with source provenance |

## Power Automate source

The existing cloud flow is tracked as a [sanitized source snapshot](power-automate/microsoft-bypass-flow/README.md).
It includes the actual operation routing and Graph queries, plus redaction and
re-export instructions. It is not an importable package.

The new [attachment-reader flow source](power-automate/attachment-reader-flow/README.md)
adds attachment operations; the Worker needs that flow extension configured before
attachment tools work. Neither JSON is a directly importable package.
See [attachment support, examples, and safety limits](docs/attachments.md).
Parsing runs on the Worker for remote chat clients; no local helper is required.

## Requirements

- Microsoftアカウント
- [Bun](https://bun.sh) 1.3.14 and Node.js 24 for offline workerd tests
- Cloudflareアカウント

## Configuration

`.dev.vars.example` に従ってPower Automate側のURLを参照

```
POWER_AUTOMATE_URL=https://prod-xxx.logic.azure.com/workflows/xxx/triggers/manual/paths/invoke
POWER_AUTOMATE_GATEWAY_KEY=<gateway key>
```

For production, set these as Worker secrets/vars (`wrangler secret put`).

## Development

```sh
bun install
bun run dev          # local Worker (wrangler dev)
```

## Quality checks

```sh
bun run typecheck    # tsc --noEmit
bun run lint         # oxlint
bun run lint:types   # oxlint --type-aware
bun run format:check # oxfmt --check
bun run test         # vitest
bun run test:flow    # offline flow contract tests
bun run test:worker  # bundled production app in real workerd, synthetic documents
```

## Deploy

```sh
bun run deploy       # wrangler deploy
```
