# university-m365-mcp

A read-only [Model Context Protocol (MCP)](https://modelcontextprotocol.io) server that lets an
LLM read university Microsoft 365 resources (Microsoft Graph) through a Power Automate
HTTP-trigger intermediary. It targets environments where a direct Microsoft Graph OAuth
integration is unavailable: Power Automate is the authentication boundary, and the MCP
server never talks to Graph directly.

See [`SPEC.md`](./SPEC.md) for the authoritative specification.

## Architecture

```
MCP Client
    ↓  MCP over Streamable HTTP (/mcp)
Remote MCP Server — Cloudflare Workers + Hono (Cloudflare Access)
    ↓  HTTP POST { operation, requestId, args }
Power Automate — operation allowlist → fixed Graph endpoints, M365 auth
    ↓
Microsoft Graph
```

The boundary is fixed: `fixed tools → fixed operations → fixed Graph endpoints`. There is no
generic Graph proxy.

## Tools

| Tool | Input | Output |
| --- | --- | --- |
| `outlook_list_messages` | `{ limit?: number }` (default 5, 1–50) | `{ messages: MessageSummary[], hasMore: boolean }` |
| `outlook_search_messages` | `{ query: string, limit?: number }` (default 10, 1–50) | `{ messages: MessageSummary[], hasMore: boolean }` |
| `outlook_get_message` | `{ messageId: string }` | `MessageDetail` |

## Authentication

The `/mcp` endpoint is protected in two layers:

1. **Cloudflare Access** (OAuth) sits in front of the Worker; only clients admitted by the
   Access policy reach it.
2. The Worker additionally validates the Access JWT via the `Cf-Access-Jwt-Assertion`
   header (defense in depth), using `TEAM_DOMAIN` and `POLICY_AUD` environment variables.

Direct `*.workers.dev` URLs are disabled (`workers_dev: false`, `preview_urls: false`) so the
Worker is reachable only through the Access-protected custom domain.

## Requirements

- A Microsoft 365 account (for the Power Automate flow)
- A Cloudflare account (for Workers + Access)
- [Bun](https://bun.sh)

## Configuration

Copy `.dev.vars.example` to `.dev.vars` and set:

```sh
POWER_AUTOMATE_URL=https://prod-xxx.logic.azure.com/workflows/xxx/triggers/manual/paths/invoke
```

For production, set these as Worker secrets/vars (`wrangler secret put`), plus:

```sh
TEAM_DOMAIN=https://<your-team-name>.cloudflareaccess.com
POLICY_AUD=<application-audience-aud-tag>
```

`TEAM_DOMAIN` and `POLICY_AUD` enable Access JWT validation; leave them unset for local
development (where no Access is in front).

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

Because `workers_dev` is disabled, deploy requires a custom domain route on the Worker.
