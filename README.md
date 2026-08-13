# University Microsoft 365 Read-only MCP

A read-only [Model Context Protocol](https://modelcontextprotocol.io) server that
exposes a fixed surface of read tools for university Microsoft 365 resources via
a Power Automate HTTP-trigger intermediary. The current implementation covers
the Outlook mailbox (three tools); other Microsoft 365 apps (Teams, OneDrive,
etc.) are added as new features. See [SPEC.md](./SPEC.md) for the full
specification.

## Specification

### Architecture

```
MCP Client
    ↓  MCP over Streamable HTTP (/mcp)
Remote MCP Server — Cloudflare Workers + Hono (this repo)
    ↓  HTTP POST { operation, requestId, args }
Power Automate — operation allowlist → fixed Graph endpoints, M365 auth
    ↓
Microsoft Graph — Outlook mailbox
```

- **Read-only.** Exactly three Outlook tools, three operations, fixed Graph
  endpoints. The server never authenticates to Graph and never calls Graph
  directly.
- **Stateless.** A fresh `McpServer` is created per request.
- **Auth** for `/mcp` is delegated to Cloudflare Access (OAuth) in front of the
  Worker; this app performs no auth itself.

### Tools

| Tool | Input | Output |
| --- | --- | --- |
| `outlook_list_messages` | `{ limit?: number }` (default 5, 1–100) | `{ messages: MessageSummary[] }` |
| `outlook_search_messages` | `{ query: string, limit?: number }` (default 10, 1–100) | `{ messages: MessageSummary[] }` |
| `outlook_get_message` | `{ messageId: string }` | `MessageDetail` |

### Power Automate protocol

Request (POST JSON):

```jsonc
{ "operation": "list_messages | search_messages | get_message", "requestId": "<uuid>", "args": { ... } }
```

- `list_messages` → `args: { top }`
- `search_messages` → `args: { query, top }`
- `get_message` → `args: { messageId }`

Response (2xx):

```jsonc
{ "ok": true, "requestId": "<uuid>", "operation": "...", "data": { /* Graph response */ } }
```

The server normalizes `data` into the tool output schemas below.

### Message shapes

```ts
type Recipient = { name: string; address: string }

type MessageSummary = {
  id: string; subject: string; from: Recipient
  receivedDateTime: string; hasAttachments: boolean
  importance: 'low' | 'normal' | 'high'; isRead: boolean
  bodyPreview: string
}

type MessageDetail = {
  id: string; subject: string; from: Recipient
  to: Recipient[]; cc: Recipient[]
  receivedDateTime: string; hasAttachments: boolean
  importance: 'low' | 'normal' | 'high'; isRead: boolean
  body: { contentType: 'text' | 'html'; content: string }
}
```

## Requirements

- [Bun](https://bun.sh)
- A Cloudflare account with Wrangler authentication (for deployment)
- Cloudflare Access (OAuth) in front of the Worker to protect `/mcp`
- A Power Automate HTTP-trigger flow that accepts `{ operation, requestId, args }`
  and proxies Microsoft Graph read calls (see SPEC §7)

## Configuration

Set the required environment variables. For local development, copy
`.dev.vars.example` to `.dev.vars` and fill in real values:

```
POWER_AUTOMATE_URL=https://prod-xxx.logic.azure.com/workflows/xxx/triggers/manual/paths/invoke
```

For production, set these as Worker secrets/vars (`wrangler secret put`).

## Development

```sh
bun install
bun run dev          # local Worker (wrangler dev)
```

The MCP endpoint is served at `/mcp` (protected by Cloudflare Access OAuth in
front of the Worker; this app performs no auth itself). A public info page is
served at `/`.

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
