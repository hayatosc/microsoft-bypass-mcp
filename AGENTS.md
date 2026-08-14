# AGENTS.md

## Project

University Microsoft 365 Read-only MCP — a read-only [Model Context Protocol](https://modelcontextprotocol.io)
server (Cloudflare Workers + Hono) exposing read tools for university Microsoft
365 resources through a Power Automate HTTP-trigger intermediary. The current
implementation covers the Outlook mailbox (three tools); other Microsoft 365
apps (Teams, OneDrive, etc.) are added as new features. See
[`SPEC.md`](./SPEC.md) for the authoritative specification and
[`README.md`](./README.md) for a summary.

## Commands

```sh
bun install            # install dependencies
bun run dev            # local Worker (wrangler dev)
bun run typecheck      # tsc --noEmit
bun run lint           # oxlint
bun run lint:types     # oxlint --type-aware
bun run format         # oxfmt --write src
bun run format:check   # oxfmt --check src
bun run test           # vitest
bun run deploy         # wrangler deploy
```

Run `typecheck`, `lint`, `lint:types`, `format:check`, and `test` before
committing. Package manager is `bun`.

## Structure

```
src/
  index.ts                 # Worker entry (export default app)
  app.ts                   # Hono app: GET / info + /mcp (stateless MCP handler)
  lib/env.ts               # Bindings + validated env access (fail fast)
  lib/access-auth.ts       # Cloudflare Access JWT validation middleware
  lib/power-automate.ts    # stateless Power Automate client
  features/outlook/
    server.ts              # createOutlookMcpServer factory (3 mail tools)
    schema.ts              # zod schemas (tool I/O + message shapes)
    normalize.ts           # Graph response -> normalized shapes
```

## Hard constraints

- **Read-only.** Only the three tools (`outlook_list_messages`,
  `outlook_search_messages`, `outlook_get_message`). Never add a generic Graph
  proxy tool (`{ url, method, body }` passthrough).
- **No persistence.** Never store mail data (no DB/KV/R2/cache). Mail data must
  not outlive a request.
- **Logging hygiene.** Never log the Power Automate URL, mail bodies, queries,
  subjects, or message IDs; the client emits one structured log entry per call
  with only `type` + `requestId` + `operation` + `durationMs` + `status` + `success`.
- **Auth** is Cloudflare Access (OAuth) in front of the Worker, plus in-Worker
  validation of the Access JWT (`Cf-Access-Jwt-Assertion`). Keep both; do not add
  app-level auth back.
- `POWER_AUTOMATE_URL` lives in `.dev.vars` (gitignored); `TEAM_DOMAIN` and
  `POLICY_AUD` are set as Cloudflare Workers secrets (`wrangler secret put`).
  Never commit any of them.

## Conventions

- TypeScript ESM-first: `strict`, `verbatimModuleSyntax`, no `any` / unsafe `as`.
- zod v4 for validation. MCP SDK v2 (`@modelcontextprotocol/server`):
  `McpServer` + `registerTool` behind `createMcpHandler` (fresh server per request).
- Docs and code comments in English.
