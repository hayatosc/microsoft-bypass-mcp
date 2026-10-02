# Fixed PDF diagnostic codes

The PDF parser preserves the generic unsupported/malformed/encrypted/limit message
and adds one fixed code. The code identifies the **first active raw-size guard or
PDF.js runtime stage**, not every issue in the file. A code does not mean PDF.js
cannot parse a different file with similar features. Diagnosing a particular file
requires observing its own diagnostic result.

`PdfParseError` provenance is kept in a private WeakMap. Public errors are rebuilt
from the issued, runtime-allowlisted code; original messages/stacks and mutable
error properties are never propagated. No document-derived values are part of a
code. No new MCP tool, output schema, logging, data storage or resource access has
been added. PDF structures are now delegated to PDF.js instead of a local
whitelist, broadening the accepted subset. Existing selector errors are unchanged.

## Code reference

| Code | First rule/stage |
| --- | --- |
| `PDF_CREATE_DOCUMENT` | PDF.js loading task creation failed |
| `PDF_EXTRACTION` | PDF.js inspection or text extraction failed |
| `PDF_INITIALIZATION` | PDF.js module initialization failed |
| `PDF_LOAD` | PDF.js document load failed |
| `PDF_PAGE_LIMIT` | Page count exceeded |
| `PDF_PASSWORD` | PDF.js reported an encrypted or password-protected document |
| `PDF_RAW_SIZE` | Empty PDF or raw byte limit exceeded |
| `PDF_TEXT_CHUNK` | PDF.js returned an invalid text chunk |
| `PDF_TIMEOUT` | PDF parse or extraction wall-time timeout |
| `PDF_UNKNOWN` | Unrecognized PDF diagnostic |

`PDF_PASSWORD` is used only when PDF.js exposes its standardized password error
name. Other encrypted or malformed documents may surface as `PDF_LOAD` if PDF.js
cannot classify them more specifically without exposing document-derived values.

## Operational use after separately authorized publication

Retry the existing `outlook_inspect_attachment` on the same message/attachment IDs.
Read the bracketed fixed code in the existing error response. If inspection passes,
use the existing bounded read tool to test extraction. Keep IDs and content private.
No original document bytes need to be exported for this first-stage diagnosis.

The handwritten PDF grammar/preflight checker has been removed. Object streams,
xref streams, JPEG image streams and other PDF structures are now delegated to the
existing unpdf/PDF.js parser with external resource access and rendering disabled.
Do not interpret the shorter diagnostic list as complete memory isolation: raw,
page, selection and output caps remain, but PDF.js internal allocation is governed
by the runtime and Cloudflare platform limits. Deployment and live verification
remain separate from local testing and are not implied by this patch.
