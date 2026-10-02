# AGENTS.md

## Project

University Microsoft 365 MCP — a fixed read and draft [Model Context Protocol](https://modelcontextprotocol.io)
server (Cloudflare Workers + Hono) exposing fixed read and draft-only tools for university
Microsoft 365 resources through a Power Automate HTTP-trigger intermediary. The
current implementation covers Outlook mailbox and attachment reads, controlled
draft creation/attachments, and native OneDrive for Business owned-file reads. See [`SPEC.md`](./SPEC.md)
for the authoritative specification and [`README.md`](./README.md) for a summary.

## Commands

```sh
bun install            # install dependencies
bun run dev            # local Worker (wrangler dev)
bun run typecheck      # tsc --noEmit
bun run lint           # oxlint
bun run lint:types     # oxlint --type-aware
bun run format         # oxfmt --write src
bun run format:check   # oxfmt --check src
bun run test           # Vitest with official Cloudflare Workers integration
bun run test:flow      # offline exported-flow contract and redaction tests (Python 3)
bun run deploy         # wrangler deploy
```

Run `typecheck`, `lint`, `lint:types`, `format:check`, `test`, and `test:flow` before
committing. Package manager is `bun`.

## Structure

```
src/
  index.ts                 # Worker entry (export default app)
  app.ts                   # Hono app: GET / info + /mcp (stateless MCP handler)
  lib/env.ts               # Bindings + validated env access (fail fast)
  lib/access-auth.ts       # Cloudflare Access JWT validation middleware
  lib/power-automate.ts    # stateless Power Automate client
  features/outlook/        # Outlook schemas, normalization, attachment parsing
  features/onedrive/       # native-connector owned-file adapters and tools
  features/documents/      # shared bounded parser dispatch/source adapter
scripts/
  build_attachment_flow.py # canonical flow generator for Outlook + OneDrive ops
  build_read_tools_flow.py # read/paging augmentation; no second flow output
  build_draft_tools_flow.py # draft-only augmentation of the same canonical flow
  test_attachment_flow.py  # offline flow contract/redaction tests
  test_read_tools_flow.py  # extension routing, native gates and contract tests
power-automate/
  microsoft-bypass-flow/   # generated canonical flow source + manual update notes
```

## Hard constraints

- **Fixed surface.** The surface has 13 read tools and 3 draft-only write tools backed by 14 fixed operations,
  documented in `SPEC.md`: Outlook message/folder/conversation/attachment reads
  and OneDrive owned-file search/list/metadata/inspect/read backing operations, plus
  explicit Outlook draft creation, sender-reply draft creation, and attachment
  addition to verified drafts. Never add sending or unrelated mutation tools.
  Never add a generic Graph, Outlook, OneDrive, URL, method, body, query, or
  nextLink passthrough tool.
- **No server persistence.** Never store mail or file data in DB/KV/R2/cache.
  Transport data must not outlive a request; explicitly requested drafts and
  attachments persist in Outlook. Draft tools are non-idempotent: no blind retry.
- **Logging hygiene.** Never log the Power Automate URL, queries, subjects,
  message IDs, file IDs, body text, bytes, base64, or nextLink URLs. The Worker
  logs only `type`, `requestId`, `operation`, `durationMs`, `status`, and `success`.
- **Auth.** Cloudflare Access (OAuth) stays in front of the Worker and in-Worker
  Access JWT validation remains defense in depth. Do not add app-level auth back.
- Secrets (`POWER_AUTOMATE_URL`, `POWER_AUTOMATE_GATEWAY_KEY`, `TEAM_DOMAIN`,
  `POLICY_AUD`) must never be committed.

## Conventions

- TypeScript ESM-first: `strict`, `verbatimModuleSyntax`, no `any` / unsafe `as`.
- zod v4 for validation. MCP SDK v2 (`@modelcontextprotocol/server`):
  `McpServer` + `registerTool` behind `createMcpHandler` (fresh server per request).
- Docs and code comments in English.

## Flow-source safety

- Evolve `power-automate/microsoft-bypass-flow/definition.json` only through
  `scripts/build_attachment_flow.py`. The sanitized pre-attachment fixture in
  `scripts/fixtures/` is immutable and remains a provenance fixture, not another
  flow source.
- Preserve the existing trigger/auth/gateway shape and extend the same switch.
  Do not introduce a parallel flow, a second trigger, or connector creation.
- Keep attachment bytes and OneDrive bytes request-local. Never expose base64,
  raw Graph objects, native connector raw objects, sharing links, download URLs,
  access shortcuts, or cross-drive IDs to MCP callers.
- Enforce transport, raw-file, page/cell and output limits. Preserve OOXML expanded-byte
  and XML guards. Delegate PDF structure parsing to the existing PDF.js library;
  PDF internal resource bounds rely on the documented platform CPU/memory limits,
  not a handwritten grammar whitelist or a claimed decoded-byte cap. Keep these
  limitations explicit. Parser exceptions must not reveal document data.
- Outlook uses only fixed Graph endpoints and fixed query construction. OneDrive
  uses only native OneDrive for Business connector operation IDs; never use the
  Outlook HTTP connector for OneDrive.
- Use synthetic fixtures only. Live flow runs, deployment, OneDrive binding, and
  merging require separate authorization. See `docs/attachments.md` and
  `power-automate/microsoft-bypass-flow/README.md`.
