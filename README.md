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

The existing [microsoft bypass flow source](power-automate/microsoft-bypass-flow/README.md)
is extended in place: the same HTTP trigger and operation switch retain the three
mail branches and add attachment list/get branches. There is one canonical JSON
at `power-automate/microsoft-bypass-flow/definition.json`, with an in-place update
procedure for the existing flow. It is source for review, not an importable package.
The Worker needs that existing flow updated before attachment tools work.
See [attachment support, examples, and safety limits](docs/attachments.md).
Parsing runs on the Worker for remote chat clients; no local helper is required.

Tests use the existing Vitest workflow with Cloudflare's official
[`@cloudflare/vitest-plugin`](https://developers.cloudflare.com/workers/testing/vitest-integration/)
integration, configured from `wrangler.jsonc`. Hono request tests and synthetic
attachment fixtures run through `bun run test`; there is no separate runtime-test
script or direct Miniflare dependency.

## Requirements

- Microsoftアカウント
- [Bun](https://bun.sh) 1.3.14
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
```

## Deploy

```sh
bun run deploy       # wrangler deploy
```
