# University Microsoft 365 Read-only MCP — Specification

## 1. Overview

A read-only [Model Context Protocol (MCP)](https://modelcontextprotocol.io) server
that lets an LLM read university Microsoft 365 resources (Microsoft Graph)
through a Power Automate HTTP-trigger intermediary.

The server is **not** a "Microsoft Graph MCP". It is a
**"University Microsoft 365 Read-only MCP"**: a fixed, allow-listed surface of
read tools that map onto fixed operations, which Power Automate turns into
fixed Microsoft Graph calls. The server never authenticates to Graph and never
talks to Graph directly.

The current implementation covers the Outlook mailbox with exactly three tools
(§8). Other Microsoft 365 apps (Teams, OneDrive, SharePoint, etc.) are added as
new features, each following the same pattern: fixed tools -> fixed operations
-> fixed Graph endpoints.

## 2. Why Power Automate is in the path

Direct access to Microsoft Graph from the MCP server requires an OAuth
app registration, consent, and secret/token management. Power Automate already
holds the Microsoft 365 identity and can call Graph with the flow owner's
credentials. Routing read operations through Power Automate therefore:

- removes all Graph authentication from the MCP server,
- restricts the Graph surface to whatever the flow explicitly implements,
- keeps credentials out of the MCP server runtime.

## 3. Architecture and responsibility split

```
MCP Client
    判断 user intent

Remote MCP Server (this repo)
    input validation
    tool exposure
    Power Automate API calls
    response normalization
    logging / timeout

Cloudflare Access (in front of the Worker)
    authentication (OAuth / access policy)

Power Automate (external, not in this repo)
    operation allowlist
    fixed Microsoft Graph endpoints
    Microsoft 365 authentication

Microsoft Graph
    Outlook mailbox access (currently the only implemented feature)
```

The boundary is fixed and must not blur:

```
MCP -> fixed tools -> fixed operations -> Power Automate -> fixed Graph APIs
```

## 4. Transport and server lifecycle

The MCP server is exposed as a **Streamable HTTP** endpoint at `/mcp`, served by
a [Hono](https://hono.dev) app on Cloudflare Workers.

It follows the **stateless server factory** pattern from the official MCP
TypeScript SDK (v2): a fresh `McpServer` is created from a factory for **every
request**. No state survives between requests.

```ts
export function createOutlookMcpServer(client: PowerAutomateClient): McpServer {
  const server = new McpServer({ name: 'university-m365', version: '0.1.0' })
  // ...server.registerTool(...) x3...
  return server
}
```

Per request, the `/mcp` handler wraps the factory in a `createMcpHandler()`
entry and delegates via `handler.fetch(request)`. The factory runs once per
request, so each request gets a fresh server; the handler serves the modern
2026-07-28 protocol revision and, via the legacy stateless fallback, 2025-era
traffic.

The Power Automate client itself is **stateless** and is constructed per request
from the environment-derived base URL.

## 5. Endpoints and authentication

| Method | Path  | Purpose                                        |
|--------|-------|------------------------------------------------|
| GET    | `/`   | Human-readable info page (server name, tools). |
| ALL    | `/mcp`| MCP Streamable HTTP endpoint.                  |

The `/mcp` endpoint is protected in two layers:

1. **Cloudflare Access** (OAuth) sits in front of the Worker (on the
   `*.workers.dev` URL and/or a custom domain), so only clients admitted by the
   Access policy reach `/mcp`.
2. The Worker additionally validates the Access JWT carried on the
   `Cf-Access-Jwt-Assertion` header (via `TEAM_DOMAIN` + `POLICY_AUD`), as
   defense in depth against requests that reach the Worker without passing
   Access.

When `TEAM_DOMAIN`/`POLICY_AUD` are unset (local development, or before Access
is configured), JWT validation is skipped.

## 6. Environment

| Variable            | Required | Purpose                                                     |
|---------------------|----------|-------------------------------------------------------------|
| `POWER_AUTOMATE_URL`| yes      | Power Automate HTTP-trigger URL.                            |
| `TEAM_DOMAIN`       | prod     | Cloudflare Access team domain (`https://<team>…access.com`). |
| `POLICY_AUD`        | prod     | Cloudflare Access Application Audience (AUD) tag.           |

Missing required variables fail fast at startup/request time (no silent
defaults). `TEAM_DOMAIN` and `POLICY_AUD` must be set together; a partially
configured pair fails fast.

`TEAM_DOMAIN` and `POLICY_AUD` are set as **Cloudflare Workers secrets**
(`wrangler secret put`), not in `.dev.vars`. Leaving them unset skips JWT
validation, which is the intended behavior for local development (`wrangler dev`
has no Access in front). Values are:

- `TEAM_DOMAIN=https://<team>.cloudflareaccess.com`
- `POLICY_AUD=<aud-tag-from-access-dashboard>`

## 7. Power Automate protocol

The MCP server calls the Power Automate HTTP trigger with `POST` JSON. The
contract below is the **verified** shape of the real flow (reverse-engineered
from live responses).

### Request body

```jsonc
{
  "operation": "list_messages | search_messages | get_message",
  "requestId": "<uuid v4, generated per request>",
  "args": {
    // list_messages
    //   { "top": <number> }
    // search_messages
    //   { "query": "<string>", "top": <number> }
    // get_message
    //   { "messageId": "<string>" }
  }
}
```

- `requestId` is a UUID generated by the MCP server (via `crypto.randomUUID()`)
  for correlation. It is the only value the server is allowed to log.
- `top` is the Graph `$top` equivalent; the MCP tool argument is named `limit`
  and is translated to `top` here.

### Success response (HTTP 2xx)

The flow wraps the Graph response in an envelope:

```jsonc
{
  "ok": true,
  "requestId": "<uuid v4>",
  "operation": "<echo of the operation>",
  "data": { /* Graph response */ }
}
```

The MCP server validates that `requestId` and `operation` echo the request; a
mismatch is surfaced as a tool error.

`data` depends on the operation:

- `list_messages`, `search_messages` -> Graph list response
  `{ "value": [ <message summary>... ] }`. Each item carries the flow's `$select`
  fields: `id`, `subject`, `from`, `receivedDateTime`, `isRead`, `importance`,
  `hasAttachments`, `bodyPreview`. Recipients are **not** selected here.
- `get_message` -> a single Graph message resource with `$select` fields:
  `id`, `subject`, `from`, `toRecipients`, `ccRecipients`, `receivedDateTime`,
  `isRead`, `importance`, `hasAttachments`, `body`.

### Error response

Non-2xx (e.g. `400` on schema mismatch, `502` on upstream Graph failure) with a
body of the shape `{ "error": { "code", "message", ... } }`.

The MCP server is responsible for **normalizing** `data` into the tool output
schemas (§8). Non-2xx responses and malformed payloads are surfaced as tool
errors.

> Field availability is controlled entirely by the flow's Graph `$select`. If
> the flow is updated to also `$select` `toRecipients` in list/search, or
> `sentDateTime`/`conversationId` in get, the MCP server shapes in §8 must be
> extended to match.

## 8. MCP tools

Three tools. Every tool defines an `inputSchema` and an `outputSchema`, and
returns both `content` (text, for the LLM) and `structuredContent` (validated
against `outputSchema`, for programs).

### Tool result shape (all tools)

```ts
return {
  content: [{ type: 'text', text: JSON.stringify(result) }],
  structuredContent: result,
}
```

### `outlook_list_messages`

List the most recent inbox messages (metadata only).

- Input: `{ limit?: number }` — default `5`, range `1..50`.
- Output: `{ messages: MessageSummary[], hasMore: boolean }`

### `outlook_search_messages`

Search inbox messages by a free-text query.

- Input: `{ query: string, limit?: number }` — `query` required (non-empty),
  `limit` default `10`, range `1..50`.
- Output: `{ messages: MessageSummary[], hasMore: boolean }`

### `outlook_get_message`

Fetch a full message (including body) by ID.

- Input: `{ messageId: string }`
- Output: `MessageDetail`

### Normalized message shapes

```ts
type Recipient = { name: string; address: string }

type MessageSummary = {
  id: string
  subject: string
  from: Recipient
  receivedDateTime: string
  hasAttachments: boolean
  importance: 'low' | 'normal' | 'high'
  isRead: boolean
  bodyPreview: string
}

type MessageDetail = {
  id: string
  subject: string
  from: Recipient
  to: Recipient[]
  cc: Recipient[]
  receivedDateTime: string
  hasAttachments: boolean
  importance: 'low' | 'normal' | 'high'
  isRead: boolean
  body: { contentType: 'text' | 'html'; content: string }
}
```

## 9. Non-functional requirements

### No persistence

The MCP server must not persist mail content anywhere. Forbidden:

```
DB storage, KV storage, R2 storage, cache storage,
sending bodies to analytics, sending bodies to error monitoring
```

Mail data must not outlive the request that produced it.

### Logging hygiene

- The Power Automate URL must never appear in responses or logs.
- Mail bodies, queries, subjects, and message IDs must never be logged.
- Power Automate calls emit one structured log entry per call containing only
  `type`, `requestId`, `operation`, `durationMs`, `status`, and `success`.

### Timeout

Power Automate calls carry an explicit timeout (abort signal); the default is
30 seconds. On timeout the tool returns an error rather than hanging.

## 10. Out of scope (not implemented in MVP)

```
send_message, create_draft, delete_message, move_message,
mark_as_read, mark_as_unread, attachments, pagination,
mail-folder selection, sent items, calendar, contacts,
other Microsoft 365 apps (Teams, OneDrive, SharePoint, etc.),
arbitrary Graph API proxy
```

Individual tools are added later only if a concrete need arises. A generic
proxy tool — passing `{ url, method, body }` from the MCP tool to Power
Automate — must **never** be built.

## 11. Tech stack and tooling

- Runtime: Cloudflare Workers (`wrangler`)
- Web framework: Hono
- MCP: `@modelcontextprotocol/server` (v2)
- JWT validation: `jose`
- Validation/schemas: Zod v4
- Package manager: bun
- Language: TypeScript (ESM-first, `strict`)
- Typecheck: `tsc --noEmit`
- Test: Vitest
- Lint: oxlint
- Format: oxfmt

## 12. Testing

### Unit tests (mock `fetch`)

For `list_messages`, `search_messages`, `get_message`, verify:

- the correct operation is sent,
- a `requestId` is generated,
- `limit` is translated to `top`,
- response normalization,
- error handling,
- timeout.

### Integration test

Use the official **MCP Inspector** against the `/mcp` Streamable HTTP endpoint
to exercise `outlook_list_messages`, `outlook_search_messages`,
`outlook_get_message`.

### End-to-end (gated)

Tests that run only when `POWER_AUTOMATE_URL` points at the real flow. Normal
CI must never touch the real university mailbox.

## 13. Acceptance criteria

MVP is complete when all of the following hold.

### MCP

- Can connect to `/mcp` over Streamable HTTP.
- `tools/list` shows the three tools.

### list

`outlook_list_messages({ limit: 5 })` returns metadata for up to the 5 most
recent messages.

### search

`outlook_search_messages({ query: "PMDA", limit: 10 })` returns search results.

### get

An ID obtained from list/search, passed to `outlook_get_message({ messageId })`,
returns the full message body.

### Security

- The Power Automate URL never appears in responses.
- The Power Automate URL never appears in logs.
- `/mcp` is unreachable unless Cloudflare Access (OAuth) admits the client.
- Mail bodies never appear in logs.
- No generic Graph proxy exists.

## 14. Design principles

This server is "University Microsoft 365 Read-only MCP", not "Microsoft Graph
MCP". Keep the layer boundary intact: fixed tools -> fixed operations -> fixed
Graph APIs. Prefer the smallest version that works end-to-end and grow from
there.
