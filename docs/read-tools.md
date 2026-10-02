# Read tools

The read tools in this document are read-only. The server also has separately
documented [Outlook draft tools](drafts.md). It does
not expose a generic Microsoft Graph proxy, a generic OneDrive proxy, arbitrary
URLs, arbitrary methods, caller-supplied OData, download URLs, sharing links, or
Power Automate nextLink replay.

## Shared behavior

- Every Worker-to-flow request is `{ operation, requestId, args }`.
- `requestId` is the only caller correlation value that may appear in logs.
- IDs are bounded to 1–2048 characters and reject controls, whitespace, NEL,
  BOM, exact `.`, and exact `..`.
- Queries are trimmed, nonempty, and at most 512 UTF-16 code units.
- Outlook pages are bounded to 50 items; OneDrive pages are bounded to 100 items.
- Text, message bodies, attachment extracts, and OneDrive extracts are untrusted
  external content and include provenance in MCP output.
- Raw Graph/native connector objects, base64, bytes, nextLink URLs, and download
  URLs are not MCP output.

## Outlook mail tools

### `outlook_list_messages`

Lists message summaries from a fixed mailbox scope or folder.

Supported scope:

- `mailbox: "inbox"` -> inbox folder
- `mailbox: "sent"` -> sent items folder
- `mailbox: "all"` -> `/me/messages`
- `folderId` -> `/me/mailFolders/{folderId}/messages`; overrides `mailbox`

Supported filters are controlled only:

- `isRead`
- `hasAttachments`
- `receivedAfter`
- `receivedBefore`

Time bounds are UTC timestamps with seconds and up to seven fractional digits
(100-nanosecond precision). The lower bound must be strictly before the upper.

When no filter is present, the flow uses `receivedDateTime desc`. When any
filter is present, the flow omits `$orderby` to avoid Graph `InefficientFilter`
failures. The fixed filter order is `isRead`, `hasAttachments`, `receivedAfter`,
`receivedBefore`.

If the upstream continuation cannot be verified, the normalized current page is
still returned with `hasMore: true`, `nextCursor: null`, and an explicit
`incompleteReason` containing only a fixed `PAGINATION_*` rejection code. This is
safe degradation, not a promise that all native pagination shapes are supported.
No rejected URL is followed or echoed, and no offset is invented from item count.
Fixed Graph endpoint names allow case variation. The equivalent anchored
`/me/mailFolders('key')/messages` OData form is accepted only for the exact
already-selected key (with OData quote escaping and one key-only percent decode).
Caller-specified folder IDs and query values remain case-sensitive; only implicit
well-known folder aliases permit case variation. `/users/` routes and different
opaque folder IDs are never treated as equivalent to `/me/` or an inbox alias. Invalid caller cursors still fail before any
upstream request. Conversation pages use the same continuation behavior.

Rejected path diagnostics use only fixed structural categories:
`PAGINATION_PATH_{ME|USER_SEGMENT|USER_ODATA}_{MESSAGES|FOLDER_SEGMENT|FOLDER_ODATA}`,
or `PAGINATION_PATH_OTHER`. They expose neither mailbox/folder keys nor URLs.
These categories describe syntax, not identity or authorization: a recognized
OData key must still match the selected key, and `/users/` routes stay rejected.
Origin checks run first; path rejection still prevents query/cursor acceptance.

### `outlook_search_messages`

Searches a fixed mailbox scope or folder with connector-supported search text.
It has no filters, skip, cursor following, or arbitrary query parameters.
Search nextLink, if present, is treated only as a bounded incomplete signal.

### `outlook_get_message`

Reads one message by ID using a fixed `/me/messages/{id}` route and selected
fields only. It does not synthesize missing folder IDs, conversation IDs, or
timestamps.

### `outlook_list_mail_folders`

Lists the first bounded page of root folders from `/me/mailFolders`, selecting
only folder metadata; child folders are not traversed. If Graph reports additional pages, the Worker reports an
incomplete result; it does not expose or follow the nextLink URL.

### `outlook_get_conversation`

Reads messages across `/me/messages` with exact escaped `conversationId`
equality, so sent mail is included when the account has access. The query does
not include `$orderby`; the Worker may sort each returned page locally but must
not claim global ordering. There is no inbox-only client-side fallback.

### Sender normalization on message reads

For `outlook_get_message`, `outlook_list_messages`, `outlook_search_messages`,
and `outlook_get_conversation`, an omitted top-level `from` normalizes exactly
like the already-supported `from: null`: `{ name: "", address: "" }`. This narrow
correction is based on a verified sanitized `FROM_MISSING` rejection, not a
universal guarantee that Graph omits `from` for drafts. No sender is inferred
from `sender`, recipients, account identity, mailbox, or other properties.
A present non-null `from` must still contain an `emailAddress` object with a
required string-or-null `name` and required string `address`; a null name still
normalizes to an empty string. Malformed present senders remain rejected.

### Sanitized message-shape rejection diagnostics

`outlook_get_message` retains `malformed response: expected a message` and appends
fixed shape codes in parentheses on normalization schema failure. Message summary
lists (`outlook_list_messages` and `outlook_search_messages`) retain
`malformed response: expected a list of messages` and use the same field codes,
collapsed across entries. `outlook_get_conversation` retains
`malformed response: expected conversation messages` and uses the same bounded
helper in list mode, including detail/recipient field codes collapsed across
messages. No rejected page or partially normalized message is returned. Apart
from the omitted-sender correction above, schema acceptance, error classes,
requested page limits, the 50-item list ceiling, ID identity checks, conversation
duplicate checks, sorting, and successful output remain unchanged. Folder
normalization is not changed by this diagnostic.

The finite field-prefix allowlist is:

| Schema field or shape | Static code prefix |
| --- | --- |
| Message object | `MESSAGE` |
| Message/folder/conversation IDs | `ID`, `PARENT_FOLDER_ID`, `CONVERSATION_ID` |
| Subject and preview | `SUBJECT`, `BODY_PREVIEW` |
| Sender object and email-address object | `FROM`, `FROM_EMAIL_ADDRESS` |
| Sender name and address | `FROM_NAME`, `FROM_ADDRESS` |
| Sent/received UTC timestamps | `SENT_DATE`, `RECEIVED_DATE` |
| Booleans and importance | `HAS_ATTACHMENTS`, `IS_READ`, `IMPORTANCE` |
| Body object, content type and content | `BODY`, `BODY_CONTENT_TYPE`, `BODY_CONTENT` |
| To-recipient array, entry and email-address object | `TO_RECIPIENTS`, `TO_RECIPIENT`, `TO_EMAIL_ADDRESS` |
| To-recipient name/address | `TO_NAME`, `TO_ADDRESS` |
| Cc-recipient array, entry and email-address object | `CC_RECIPIENTS`, `CC_RECIPIENT`, `CC_EMAIL_ADDRESS` |
| Cc-recipient name/address | `CC_NAME`, `CC_ADDRESS` |

Each field prefix has only these suffixes: `_MISSING` (absent/undefined), `_NULL`,
`_TYPE` (wrong schema type), or `_INVALID` (invalid format, enum string or bound).
Only actual schema failures produce codes: omitted or null top-level senders,
nullable names/subjects/previews, optional metadata and omitted recipient arrays
remain accepted. Message summary lists validate the existing **summary** schema
only; body/to/cc detail codes are not added to summary-list validation.
Conversation entries validate the existing **detail** schema.

List envelope failures use `MESSAGE_LIST_ENVELOPE`, including invalid list/value
shapes or nextLink type/length. Exceeding the existing schema/requested limit uses
`MESSAGE_LIST_LIMIT`, never a count. Unknown paths or unrecognized failures use
`MESSAGE_SHAPE_OTHER`. Codes are deduplicated and lexically ordered, independent
of entry/recipient indices and input key order. At most 32 codes are returned;
if needed the first 31 are followed by `MESSAGE_SHAPE_TRUNCATED`. The entire
shape-rejection error is bounded to 1,024 characters. An abbreviated result is
not a complete inventory of every rejection.

For example, a **synthetic** message with a missing sender name, null sent time
and malformed received time reports
`FROM_NAME_MISSING,RECEIVED_DATE_INVALID,SENT_DATE_NULL`. These codes identify
shape discrepancies, not mailbox identities, evidence of a particular tenant's
behavior, or permission to accept a different schema. Any later normalization
change requires separately reviewed evidence and authorization.

Errors contain no input values, names, addresses, IDs, timestamps, subjects,
body content, unknown keys/paths, array indices/counts, raw Zod messages/issues,
or raw Graph payloads. There is no diagnostic logging: the transport telemetry
whitelist remains only `type`, `requestId`, `operation`, `durationMs`, `status`,
and `success`. Transport `success: true` can still precede a normalization
rejection; this meaning is unchanged. Offline tests use synthetic fixtures only.

## Outlook attachment tools

`outlook_list_attachments`, `outlook_inspect_attachment`, and
`outlook_read_attachment` are documented in `docs/attachments.md`. Important
points:

- Listing returns metadata only.
- Content reads support file attachments only.
- Reference/item attachments and external URLs are not followed.
- Parser output is bounded and includes source provenance.
- Raw base64 remains internal flow-to-Worker transport and is never MCP output.

## OneDrive tools

OneDrive tools use the signed-in account's native OneDrive for Business
connector. They are intentionally limited to owned-file/search/list/metadata and
bounded content reads.

### `onedrive_search_files`

Uses native `FindFiles` with `query`, root binding, `findMode`, and
`maxFileCount`. Results are metadata projections only. If the connector returns
exactly the requested/maximum count, the result is conservatively marked
potentially truncated; the server must not claim complete enumeration.

The `OneDriveSearchMode` and `OneDriveSearchRootId` Compose configuration actions
have fixed empty-string inputs by default. Search fails closed before connector access until both have
verified tenant/designer values. Bind the native connection separately during an
authorized manual update.

Native search exposes no supported continuation argument and returns at most 100
results. No search cursor is implemented: paging slices of those same results
would not retrieve matches beyond the connector cap. Narrow the query or browse
a known folder when a search is incomplete.

### `onedrive_list_folder`

Uses `ListRootFolder` when no `folderId` is supplied and `ListFolderV2(id)` when
one is supplied. For named folders, the native connector's supported pagination
setting follows its own continuation tokens, with a threshold of 1,000 items.
The flow projects at most 1,000 metadata records even when the final native page
overshoots that threshold. Root listing is bounded to the returned array and does
not claim supported native continuation for `ListRootFolder`.

The MCP `limit` remains 1–100 (default 50). `nextCursor` retrieves the next slice
of that fixed bounded window. Each call re-fetches the window; there is no server
cache. Cursors bind to the exact folder/root, limit, and a fingerprint of the
ordered metadata window. A changed window rejects continuation: restart without
a cursor. No cursor contains a folder ID, raw metadata, or an upstream URL.

`hasMore` includes both known records remaining in the window and possible
upstream truncation. `nextCursor` exists only for known records in the window.
At the window boundary, `hasMore: true` with `nextCursor: null` and an
`incompleteReason` means more native data may exist but cannot be retrieved by
this bounded operation. Reaching exactly 1,000 is conservatively incomplete;
this is not unlimited enumeration or a stable snapshot of a changing drive.

The Worker never follows native nextLink URLs, and callers cannot supply them.
A folder window has a 4 MiB transport ceiling; an over-limit response fails
safely rather than silently dropping metadata. Normal MCP results still contain
at most 100 records and remain subject to the shared output-size limit.

Re-fetching can repeat multiple native requests on every page and consume the
connector's request allowance. The published OneDrive connector limit is 100
calls per 60 seconds; actual tenant behavior and aggregation latency need live
verification. A timeout is a failed read, not evidence that enumeration is
complete. Use a smaller folder where possible; this design deliberately adds no
persistent cache.

Sources: [OneDrive connector operations and limits](https://learn.microsoft.com/en-us/connectors/onedriveforbusiness/)
and [native runtime pagination](https://learn.microsoft.com/en-us/azure/logic-apps/logic-apps-exceed-default-page-size-with-pagination).

### `onedrive_get_metadata`

Uses `GetFileMetadata(id)`. The internal flow-to-Worker metadata projection is:

```json
{ "Id": "...", "Name": "...", "Size": 1, "MediaType": "...", "IsFolder": false, "LastModified": "...", "ETag": "..." }
```

The MCP result normalizes these to `fileId`, `name`, `size`, `contentType`,
`isFolder`, `lastModifiedDateTime`, and `eTag`, plus `supportedFormat`, `readable`,
and a nullable `limitation`. Missing optional timestamps/version tags stay null.

Other native fields such as `Path`, `NameNoExt`, `DisplayName`, and
`FileLocator` are permitted only for internal validation/projection and are not
part of MCP output.

### `onedrive_inspect_file` and `onedrive_read_file`

Both use `onedrive_get_content` as the transport operation. The flow fetches
fresh metadata first, then rejects folders, empty files, files over 4 MiB, and
unsupported types before downloading bytes. Supported file classes are PDF,
DOCX, and XLSX with a supported extension and matching or generic MIME type.
Both OneDrive and Outlook use the same document parsers. PDF structure parsing
is delegated to PDF.js; the former 16 MiB PDF decoded-stream guarantee is removed.
Raw file, page, selection and output caps remain. Timers cannot preempt synchronous
parser work, and deployed platform CPU/memory limits still require verification.
See [attachment limits](attachments.md#supported-subset-and-resource-limits) for the full details.

Power Automate/Logic Apps binary bodies are expected as:

```json
{ "$content-type": "...", "$content": "padded standard base64" }
```

The flow treats `$content` as already-base64 data, validates it deliberately,
and returns `{ metadata, contentBytes }` only to the Worker. The Worker compares
fresh metadata to its first metadata call and enforces decoded byte limits again.

## Security notes

- The flow uses secure inputs/outputs on supported data-bearing actions.
- The Request trigger secures outputs so caller arguments/headers are hidden in
  run history. This is obfuscation, not a guarantee about Microsoft retention.
- Sanitized errors do not echo IDs, queries, subjects, names, URLs, raw connector
  errors, bytes, or base64.
- Flow run-history security settings are not a guarantee about Microsoft service
  retention; review tenant retention separately.
- Live import, connector binding, and smoke testing require separate authorization.
