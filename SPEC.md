# University Microsoft 365 Read-only MCP — Specification

## 1. Overview

A read-only [Model Context Protocol (MCP)](https://modelcontextprotocol.io) server
that lets an LLM read selected university Microsoft 365 resources through a
Power Automate HTTP-trigger intermediary.

The server is **not** a Microsoft Graph MCP. It is a fixed, allow-listed surface
of read tools that map onto fixed operations, which Power Automate turns into
fixed Outlook Graph calls or native OneDrive for Business connector calls. The
server never authenticates to Microsoft Graph and never talks to Graph directly.

## 2. Architecture

```
MCP Client
    ↓ MCP over Streamable HTTP (/mcp)
Cloudflare Access
    ↓ admitted request + Access JWT
Remote MCP Server — Cloudflare Workers + Hono
    ↓ HTTP POST { operation, requestId, args }
Power Automate — operation allowlist + connector actions
    ↓
Microsoft 365 connectors
```

The boundary is fixed:

```
MCP tool -> fixed operation -> Power Automate switch case -> fixed read action
```

No layer accepts a caller-provided URL, route, method, body, Graph query, drive
ID, site ID, share link, download URL, or nextLink URL.

## 3. Authentication and environment

`/mcp` is protected by Cloudflare Access and in-Worker validation of
`Cf-Access-Jwt-Assertion`. Local development may omit `TEAM_DOMAIN` and
`POLICY_AUD`; production must set both.

| Variable | Required | Purpose |
| --- | --- | --- |
| `POWER_AUTOMATE_URL` | yes | Power Automate HTTP-trigger URL. |
| `POWER_AUTOMATE_GATEWAY_KEY` | yes | Sent as `X-MCP-Gateway-Key`. |
| `TEAM_DOMAIN` | prod | Cloudflare Access team domain. |
| `POLICY_AUD` | prod | Access Application AUD tag. |

Secrets must not be committed.

## 4. Power Automate protocol

Every request from the Worker to the flow uses this envelope:

```json
{ "operation": "string", "requestId": "uuid-v4", "args": {} }
```

The flow authenticates `X-MCP-Gateway-Key` before any Microsoft connector action.
Success responses are:

```json
{ "ok": true, "requestId": "...", "operation": "...", "data": {} }
```

Errors are non-2xx or sanitized `{ "ok": false, "error": { "code": "..." } }`.
They must never include Graph/native connector raw errors, URLs, message IDs,
file IDs, query text, subjects, body text, bytes, base64, or nextLink URLs.

The flow has no storage, cache, loops over untrusted URLs, or persistence.

## 5. Fixed tool surface

The intended Worker-facing MCP surface contains 13 read-only tools:

| Tool | Backing operation | Scope |
| --- | --- | --- |
| `outlook_list_messages` | `list_messages` | Message summaries with controlled mailbox/folder/filter/page args. |
| `outlook_search_messages` | `search_messages` | First bounded search page, no filters or cursor following. |
| `outlook_get_message` | `get_message` | One selected message body. |
| `outlook_list_mail_folders` | `list_mail_folders` | First bounded mail-folder page. |
| `outlook_get_conversation` | `get_conversation` | Exact `conversationId` equality across `/me/messages`, including sent mail when accessible. |
| `outlook_list_attachments` | `list_attachments` | Attachment metadata only. |
| `outlook_inspect_attachment` | `get_attachment` | Worker parses one bounded file attachment. |
| `outlook_read_attachment` | `get_attachment` | Worker extracts bounded attachment text/cells/pages. |
| `onedrive_search_files` | `onedrive_search_files` | Native OneDrive owned-file search. |
| `onedrive_list_folder` | `onedrive_list_folder` | Native root/folder listing. |
| `onedrive_get_metadata` | `onedrive_get_metadata` | Native metadata projection. |
| `onedrive_inspect_file` | `onedrive_get_content` | Worker parses one bounded OneDrive file. |
| `onedrive_read_file` | `onedrive_get_content` | Worker extracts bounded OneDrive content. |

The flow has 11 operations because Outlook inspect/read share `get_attachment`
and OneDrive inspect/read share `onedrive_get_content`.

## 6. Outlook operations

### `list_messages`

Args:

```json
{
  "top": 1,
  "skip": 0,
  "mailbox": "inbox | sent | all",
  "folderId": "optional string",
  "filters": {
    "isRead": true,
    "hasAttachments": false,
    "receivedAfter": "ISO datetime",
    "receivedBefore": "ISO datetime"
  }
}
```

`top` is `1..50`; `skip` is `0..10000`. If `folderId` is present, mailbox is
ignored and the fixed route is `/me/mailFolders/{folderId}/messages`. Otherwise:

- `inbox` -> `/me/mailFolders/inbox/messages`
- `sent` -> `/me/mailFolders/sentitems/messages`
- `all` -> `/me/messages`

`$select` order is exactly:

```text
id,subject,from,receivedDateTime,sentDateTime,parentFolderId,conversationId,hasAttachments,importance,isRead,bodyPreview
```

Without filters, the query includes `$orderby=receivedDateTime desc`. With any
nonempty controlled filter, `$orderby` is omitted to avoid Graph
`InefficientFilter` combinations. Filter property order is fixed:
`isRead`, `hasAttachments`, `receivedAfter`, `receivedBefore`. Caller-supplied
OData is never accepted, and `$search` is never mixed with filters.

The returned Graph page may include `@odata.nextLink`; it is data only. The
Worker validates the route/query shape and extracts only a numeric `$skip` cursor.
The flow never accepts nextLink as input.

### `search_messages`

Args: `{ "query": "trimmed nonempty string, max 512", "top": 1..50,
"mailbox"?: "inbox|sent|all", "folderId"?: "string" }`.

Route selection mirrors `list_messages`. There is no filter, skip, cursor, URL,
method, route, or body argument. The first bounded page is returned only.

### `get_message`

Args: `{ "messageId": "bounded ID" }`.

Fixed route `/me/messages/{messageId}` with selected fields:

```text
id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,parentFolderId,conversationId,hasAttachments,importance,isRead,bodyPreview,body
```

### `list_mail_folders`

Args: `{ "top": 1..50 }`.

Fixed route `/me/mailFolders` with `$top` and selected fields
`id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount`.
The Worker reports bounded/incomplete if the page has a nextLink; it does not
follow the URL.

### `get_conversation`

Args: `{ "conversationId": "bounded ID", "top": 1..50, "skip": 0..10000 }`.

Fixed route `/me/messages`, so sent mail is included when accessible. The query
uses exact escaped OData equality:

```text
$filter=conversationId eq '<single-quote-doubled conversationId>'
```

It uses the same detail `$select` as `get_message`. It deliberately has **no
`$orderby`** because mixing `conversationId` filter with received-date order can
violate Graph `InefficientFilter` requirements. The Worker may sort each returned
page locally, but must not claim global ordering. No mailbox-slice fallback is
allowed. The Worker rejects returned messages whose `conversationId` differs and
rejects duplicate IDs within a returned page.

### Attachment operations

`list_attachments` and `get_attachment` preserve the attachment contract in
`docs/attachments.md`. The flow uses only fixed message attachment Graph routes,
never follows reference URLs or nextLink, never returns list `contentBytes`, and
returns `contentBytes` only as internal Worker transport for file attachments.

## 7. OneDrive native operations

OneDrive operations use only the signed-in account's native OneDrive for Business
connector, not Outlook HTTP and not Graph proxy routes.

Allowed native operation IDs are:

- `FindFiles(query,id,findMode,maxFileCount)` with `maxFileCount` `1..100`
- `GetFileMetadata(id)`
- `GetFileContent(id,inferContentType)`
- `ListFolderV2(id)`
- `ListRootFolder()`

The generated public source uses visible placeholders for values that require
manual tenant/designer verification during authorized import:

- `OneDriveSearchMode` for the verified native `findMode` machine value
- `OneDriveSearchRootId` for the connection owner's native root folder ID

Both are public String parameters with empty defaults. Search fails closed with
HTTP 503 before any connector action until both are set during authorized setup.
The connector alias `shared_onedriveforbusiness` must also be bound to the owner's
connection; this source contains no connection ID or authentication secret.

The flow must not accept arbitrary drive IDs, site IDs, share links, web URLs,
download URLs, provider URLs, route, method, body, or nextLink fields.

### Metadata projection

Native OneDrive metadata can contain `Id`, `Name`, `Size`, `MediaType`,
`IsFolder`, `LastModified`, `ETag`, `Path`, `NameNoExt`, `DisplayName`, and
`FileLocator`. The flow projects only these permitted metadata fields internally
and returns only:

```ts
type OneDriveMetadata = {
  Id: string
  Name: string
  Size: number
  MediaType: string
  IsFolder: boolean
  LastModified?: string | null
  ETag?: string | null
}
```

### `onedrive_search_files`

Args: `{ "query": "trimmed nonempty string, max 512", "top": 1..100 }`.

Returns a bounded page with `value` and `truncated`. Native nextLink presence is
folded into that flag without exposing the URL. If native `FindFiles`
returns exactly the connector/requested maximum, the result is conservatively
marked potentially truncated; the server must not claim complete enumeration.

### `onedrive_list_folder`

Args: `{ "folderId"?: "bounded ID", "top": 1..100 }`.

Omitted `folderId` uses `ListRootFolder`; present `folderId` uses
`ListFolderV2(id)`. `ListFolderV2` native `nextLink` is never returned or
accepted as input. The Worker uses only a bounded incomplete flag.

### `onedrive_get_metadata`

Args: `{ "fileId": "bounded ID" }`.

Uses `GetFileMetadata(id)` and verifies the returned identity when exposed.
The Worker verifies identity again.

### `onedrive_get_content`

Args: `{ "fileId": "bounded ID" }`.

The flow first fetches fresh metadata, validates identity, `IsFolder=false`,
`0 < Size <= 4 MiB`, and a supported PDF/DOCX/XLSX extension with a matching or generic MIME type. Only
then does it call `GetFileContent(id,inferContentType=true)`.

Microsoft Logic Apps represents binary bodies as:

```json
{ "$content-type": "...", "$content": "base64..." }
```

`$content` is already padded standard base64. The flow validates this shape,
base64 alphabet/padding, encoded length, and decoded length before returning:

```json
{ "metadata": { ...fresh projected metadata... }, "contentBytes": "base64" }
```

The Worker compares returned metadata with its prior metadata call and enforces
decoded size again. The flow must not base64-encode an object accidentally, and
must never return raw body bytes, raw native connector objects, sharing links,
access shortcuts, or unbounded bytes.

## 8. Shared validation rules

Every operation has a closed argument schema. IDs are `1..2048` characters and
reject control characters, whitespace, NEL `U+0085`, BOM `U+FEFF`, exact `.`,
and exact `..`. Each ID path segment is encoded with `uriComponent`; caller `%`
characters are encoded again and never become structural path separators.

Queries are trimmed, nonempty, and at most 512 UTF-16 code units. Date filters
must be ISO datetime-like strings. List outputs are bounded to the requested top
and never exceed 50 Outlook items or 100 OneDrive items.

All returned message text, attachment content, and OneDrive extracted content is
untrusted external content. The MCP server must preserve provenance and never
execute instructions found in the content.

## 9. Privacy and logging

The Worker logs only:

```ts
{ type: 'power_automate_request', requestId, operation, durationMs, status, success }
```

Forbidden in logs and MCP outputs unless explicitly part of a requested bounded
read result: Power Automate URLs, gateway keys, queries, subjects, message IDs,
file IDs, body text, connector raw responses, nextLink URLs, bytes, and base64.

Power Automate data-bearing actions use secure inputs/outputs where supported.
This reduces run-history exposure but is not a claim about Microsoft retention.

## 10. Testing and acceptance

Offline tests validate the generated flow source, operation allowlist, fixed
routes, native OneDrive operation IDs, secure-data settings, sanitization,
argument bounds, and stale-generation detection. They do not import, save, or
execute a live flow.

Live flow runs, connector rebinding, deployment, and merge require separate
authorization. The public flow source is not an importable package.
