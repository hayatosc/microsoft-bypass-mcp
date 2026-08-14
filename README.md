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

## Requirements

- Microsoftアカウント
- [Bun](https://bun.sh)
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
```

## Deploy

```sh
bun run deploy       # wrangler deploy
```
