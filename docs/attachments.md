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
  and may not preserve visual reading order, tables, columns or all glyphs. Pages
  can have empty text, especially scans. `truncated` means the requested character
  budget was reached; request fewer pages or a larger permitted budget.
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
exceptions are replaced with fixed, sanitized tool errors. Transport logs contain
only operation, request ID, timing, status and success.

PDF support is intentionally conservative: PDF 1.0–1.7 classic cross-reference tables with
direct stream lengths, and unfiltered or Flate streams. Encrypted PDFs, object/xref
streams, indirect stream lengths or filters, other stream codecs, DecodeParms, incremental revisions, inline images, Form
XObjects and Type3 fonts are rejected before PDF.js. This means many otherwise-valid PDFs need conversion
outside this service; it does not promise universal PDF extraction. These guards
bound actual stream expansion before invoking the parser, rather than trusting the
compressed attachment size. CMap/CID indexes are restricted to 16 bits, with
65,536 aggregate mappings and width expansions; repeated page-content references
are independently bounded at 16 MiB and 10,000 references. No OCR or rendered image
extraction is provided.

The PDF parser also has a best-effort wall-time limit and cleanup. Synchronous JS
cannot be forcibly interrupted by a timer, so platform CPU/memory limits remain a
last backstop; these tests do not establish production throughput or worst-case
latency. Keep Cloudflare plan/CPU limits appropriate and load-test synthetic files
before an authorized production rollout. No Worker settings are changed here.

## Flow setup and verification

Use the [canonical existing-flow definition and in-place update procedure](../power-automate/microsoft-bypass-flow/README.md).
The same HTTP trigger and operation switch retain the existing three mail branches
and add two attachment branches, with metadata-only list projection, file type/size
preflight, encoded path segments and sanitized attachment errors. There is no
parallel replacement flow. The JSON is reviewable source, not an importable ZIP or
Dataverse solution. Preserve the existing private connection and gateway setup;
the public blank SecureString parameter deliberately fails closed. Verify the
callback URL privately after an authorized save rather than assuming continuity.

Offline checks use synthetic PDF/DOCX/XLSX documents, malformed files and size-limit
attacks. `bun run test` uses the official Cloudflare Workers Vitest integration
(`@cloudflare/vitest-plugin`) for Hono `app.request`, MCP tool calls, and file
parsers with mocked flow responses. They do not contact Graph, invoke a live flow, persist
mail, or deploy anything. Power Automate tests check the source contract, including
exact legacy-mail compatibility, without executing Microsoft's service.

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
