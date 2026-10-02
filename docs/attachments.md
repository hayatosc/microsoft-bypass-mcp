# Attachment reading for remote chat clients

The three attachment tools run entirely in the Worker. ChatGPT/Claude clients can
list files, inspect their structure, then request bounded text or cell ranges.
Raw attachment bytes travel only from Graph through Power Automate to the Worker;
there is no model-facing download/base64 tool, local helper, DB, KV, R2 or cache.
Existing Cloudflare Access and in-Worker JWT validation remain unchanged. Actual
client connectivity still depends on that client's supported remote MCP/auth flow;
this change does not claim a live ChatGPT or Claude connection was tested.

## Examples and provenance

First call `outlook_list_attachments`:

```json
{"messageId":"MESSAGE_ID","limit":20,"offset":0}
```

It returns attachment ID, name, media type, Graph size, inline flag, attachment
kind, supported format, eligibility (`readable`), `hasMore`, and `nextOffset`.
Eligibility is based on metadata; inspect/read can still reject a malformed or
unsupported file. Use the returned offset with the same message and limit. Offset
paging can duplicate/skip items if the mailbox changes; inspect individual IDs
when consistency matters. Never use an upstream nextLink as an input. At the
10,000-offset cap `hasMore` may be true while `nextOffset` is null.

Call `outlook_inspect_attachment` with `messageId` and `attachmentId`. It returns
`source` (both IDs, name, media type, Graph size and format), `untrustedContent:
true`, and a `structure` object. Then call `outlook_read_attachment`:

```json
{"messageId":"MESSAGE_ID","attachmentId":"ATTACHMENT_ID","selection":{"format":"pdf","pageStart":1,"pageEnd":3,"maxCharacters":10000}}
```

```json
{"messageId":"MESSAGE_ID","attachmentId":"ATTACHMENT_ID","selection":{"format":"docx","sectionId":"section-1","offset":0,"length":5000}}
```

```json
{"messageId":"MESSAGE_ID","attachmentId":"ATTACHMENT_ID","selection":{"format":"xlsx","sheet":"Budget","range":"A1:D20"}}
```

Use actual DOCX section IDs returned by inspection, rather than assuming a naming
scheme. Without `sectionId`, DOCX offset is relative to the whole extracted body.

- PDF page numbers are 1-based, inclusive. Text comes from the existing text layer
  through PDF.js and may not preserve visual reading order, tables, columns or all
  glyphs. Pages can have empty text, especially scans. `truncated` means the
  requested character budget was reached; request fewer pages or a larger permitted
  budget.
- DOCX paragraph numbers are 1-based; text offsets are 0-based, end-exclusive
  UTF-16 code units in extracted main-body text. Heading-defined sections and
  global source offsets provide provenance. `nextOffset` allows continuation.
  Tables are flattened; headers/footers, drawings, text boxes and notes are omitted.
- XLSX returns individual cell addresses, row/column, raw cached values and flags
  for formulas without cached results. Formulas never execute, and links are not
  followed. Numbers remain strings to avoid precision loss. Date serials stay raw
  with the workbook's 1900/1904 date system; number formatting is not reproduced.
  Empty requested cells are explicit. Dimensions describe observed cells rather
  than trusting the workbook's advertised dimension field.

## Supported subset and resource limits

Only `.pdf`, `.docx` and `.xlsx` **fileAttachment** objects are accepted, with a
matching registered media type or `application/octet-stream`/unspecified type.
Legacy `.doc`/`.xls`, macro-enabled `.docm`/`.xlsm`, images, OCR, embedded mail
items and cloud/reference attachments are unsupported. File signatures/structure
are validated after metadata and base64 checks. No scripts, formulas, macros or
external relationships are executed. Attachment text, headings and filenames are
untrusted external data, never instructions or user authorization.

| Resource | Limit |
| --- | --- |
| Attachment raw bytes and Graph size | 4 MiB each, independently checked |
| Transport JSON | 6 MiB; list JSON 256 KiB, streamed before JSON parsing |
| ZIP entries / expanded entry / total expanded | 256 / 8 MiB / 16 MiB |
| Parsed OOXML (aggregate) | 8 MiB, depth 64, 500,000 XML elements |
| Word paragraphs / sections | 10,000 / 200 |
| Workbook sheets / parsed cells / requested cells | 50 / 50,000 / 500 |
| PDF pages / selected pages | 200 / 10 |
| Returned text | 20,000 UTF-16 characters; defaults 10,000 |
| Serialized tool JSON | 128 KiB (text and structured representations each) |
| Power Automate request | 30 seconds |

ZIP expansion is checked against actual incremental output, not just size headers;
CRC, local/central directory agreement, names and supported flags are validated.
ZIP64, encrypted archives, traversal, duplicates, XML DTD/entity declarations and
external worksheet targets are rejected. Inputs are request-local and parser
exceptions are replaced with fixed, sanitized tool errors. PDF failures additionally
include an allowlisted [diagnostic code](pdf-diagnostics.md) for the first active
raw-size guard or PDF.js parser stage, without original exception text or document
values. Diagnostics do not identify every issue in a file. Removing the local
structure whitelist allows constructs supported by PDF.js. Transport logs contain
only operation, request ID, timing, status and success.

PDF handling now uses the existing `unpdf` serverless PDF.js dependency instead of
a handwritten PDF structure validator. The Worker still rejects empty PDFs and raw
PDF bytes over 4 MiB before parser initialization, caps documents at 200 pages,
extracts at most 10 selected pages, returns at most 20,000 UTF-16 characters, uses
request-local bytes, disables external fetch/range/streaming/worker fetch/XFA/WASM,
disables image rendering paths, observes a best-effort 10 second timeout, and
always attempts to destroy the PDF.js loading task. Object streams, xref streams,
JPEG image streams, normal forms and other PDF structures are left to PDF.js rather
than individually whitelisted or rejected by local grammar checks.

This does not mean every valid PDF is supported. Encrypted/password-protected,
malformed, unsupported or resource-heavy documents can still fail in PDF.js and are
reported only through sanitized diagnostics. PDF.js runs in the same serverless
JavaScript isolate and on the same event loop as the Worker. `setTimeout` and
`Promise.race` observe cooperative/asynchronous delays only; they cannot preempt
synchronous parser work, and they do not restore the former handwritten 16 MiB
decoded-stream bound. Raw/page/output caps do not bound PDF.js internal allocation
before or during parsing. Resource safety therefore also relies on Cloudflare
platform CPU and memory ceilings; the documented 128 MB memory limit is per-isolate
and can be shared across concurrent requests. This repository does not set
`limits.cpu_ms` in `wrangler.jsonc`, the account plan is not known here, and no
paid-only Worker limit or production CPU-kill behavior is claimed by this patch.
Cloudflare plan/quota verification and representative CPU/memory checks with
synthetic and expected real-world attachments are rollout gates before separately
authorized production use.

## Flow setup and verification

Use the [canonical existing-flow definition and in-place update procedure](../power-automate/microsoft-bypass-flow/README.md).
The same HTTP trigger and operation switch retain the two attachment branches
byte-for-byte while the [read-tool expansion](read-tools.md) evolves the mail
branches and adds native OneDrive reads. Attachment metadata-only projection,
file type/size preflight, encoded path segments and sanitized errors remain intact.
There is no parallel replacement flow. The JSON is reviewable source, not an importable ZIP or
Dataverse solution. Preserve the existing private connection and gateway setup;
the public blank SecureString parameter deliberately fails closed. Verify the
callback URL privately after an authorized save rather than assuming continuity.

Offline checks use synthetic PDF/DOCX/XLSX documents, malformed files and size-limit
attacks. `bun run test` uses the official Cloudflare Workers Vitest integration
(`@cloudflare/vitest-plugin`) for Hono `app.request`, MCP tool calls, and file
parsers with mocked flow responses. They do not contact Graph, invoke a live flow,
persist mail, or deploy anything. Power Automate tests check the source contract,
including exact legacy-mail compatibility, without executing Microsoft's service.

```sh
bun install --frozen-lockfile
bun run typecheck
bun run lint
bun run lint:types
bun run format:check
bun run test
bun run test:flow
```

Parser packages: [unpdf](https://github.com/unjs/unpdf) (serverless PDF.js build),
[fflate](https://github.com/101arrowz/fflate), and
[saxes](https://github.com/lddubeau/saxes). Versions are locked in `bun.lock`.
