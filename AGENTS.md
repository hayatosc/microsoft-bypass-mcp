# AGENTS.md

## Project

University Outlook Read-only MCP — a read-only [Model Context Protocol](https://modelcontextprotocol.io)
server (Cloudflare Workers + Hono) exposing three tools that read a university
Outlook mailbox through a Power Automate HTTP-trigger intermediary. See
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
  lib/power-automate.ts    # stateless Power Automate client
  features/outlook/
    server.ts              # createOutlookMcpServer factory (3 tools)
    schema.ts              # zod schemas (tool I/O + message shapes)
    normalize.ts           # Graph response -> normalized shapes
```

## Hard constraints

- **Read-only.** Only the three tools (`outlook_list_messages`,
  `outlook_search_messages`, `outlook_get_message`). Never add a generic Graph
  proxy tool (`{ url, method, body }` passthrough).
- **No persistence.** Never store mail data (no DB/KV/R2/cache). Mail data must
  not outlive a request.
- **Logging hygiene.** Never log the Power Automate URL or mail bodies; only log
  `operation` + `requestId`.
- **Auth** is delegated to Cloudflare Access (OAuth) in front of the Worker. Do
  not add app-level auth back.
- `POWER_AUTOMATE_URL` lives in `.dev.vars` (gitignored); never commit it.

## Conventions

- TypeScript ESM-first: `strict`, `verbatimModuleSyntax`, no `any` / unsafe `as`.
- zod v4 for validation. MCP SDK `McpServer` +
  `WebStandardStreamableHTTPServerTransport` (fresh server per request).
- Docs and code comments in English.
