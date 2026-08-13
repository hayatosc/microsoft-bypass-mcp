# University Outlook Read-only MCP

A read-only [Model Context Protocol](https://modelcontextprotocol.io) server that
exposes a fixed surface of three tools for reading a university Outlook mailbox
via a Power Automate HTTP-trigger intermediary. See [SPEC.md](./SPEC.md) for the
full specification.

## Tools

- `outlook_list_messages` — list the most recent messages (metadata only)
- `outlook_search_messages` — search messages by free-text query (metadata only)
- `outlook_get_message` — fetch a full message (including body) by ID

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
