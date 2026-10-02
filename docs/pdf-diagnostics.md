# Fixed PDF diagnostic codes

The PDF parser preserves the generic unsupported/malformed/encrypted/limit message
and adds one fixed code. The code identifies the **first failing validation rule
or runtime stage**, not every issue in the file. Combined guards may still have
multiple listed causes. A guard code does not mean PDF.js cannot parse the file.
Diagnosing a particular file requires observing its own diagnostic result.

`PdfParseError` provenance is kept in a private WeakMap. Public errors are rebuilt
from the issued, runtime-allowlisted code; original messages/stacks and mutable
error properties are never propagated. No document-derived values are part of a
code. No new MCP tool, output schema, logging, data storage, resource access, or
parser support has been added. Existing selector errors are unchanged.

## Code reference

| Code | First rule/stage |
| --- | --- |
| `PDF_ACTIVE_CONTENT` | Active content marker unsupported |
| `PDF_ARRAY_LIMIT` | Array entry limit exceeded |
| `PDF_CMAP_ARRAY` | Array destination on a non-range CMap entry |
| `PDF_CMAP_BUDGET` | CMap count or mapping-byte budget exceeded |
| `PDF_CMAP_DESTINATION` | CMap destination outside supported range |
| `PDF_CMAP_ENTRIES` | Too many destination entries for CMap range |
| `PDF_CMAP_HEX` | Invalid or oversized CMap hexadecimal value |
| `PDF_CMAP_RANGE` | Reversed CMap range |
| `PDF_CMAP_TOKEN` | CMap entry is not a hexadecimal string |
| `PDF_CONTENT_ARRAY` | Content reference does not resolve to stream or array |
| `PDF_CONTENT_ARRAY_BUDGET` | Repeated content-array byte budget exceeded |
| `PDF_CONTENT_BUDGET` | Aggregate repeated page-content byte budget exceeded |
| `PDF_CONTENT_REFERENCE` | Missing or cyclic content reference |
| `PDF_CONTENT_REFERENCE_BUDGET` | Content-reference count or depth exceeded |
| `PDF_CREATE_DOCUMENT` | PDF.js loading task creation failed |
| `PDF_DECODE_PARAMETERS` | Non-null DecodeParms unsupported |
| `PDF_DECOMPRESSION` | Unexpected decompression failure |
| `PDF_DICTIONARY_DELIMITER` | Invalid dictionary delimiter |
| `PDF_DICTIONARY_KEY` | Dictionary key is invalid, duplicate, or forbidden |
| `PDF_DIFFERENCES_ARRAY` | Font Differences is not an array |
| `PDF_DIFFERENCES_ENTRY` | Invalid or out-of-range Differences entry |
| `PDF_DOCUMENT_PAGE_COUNT` | PDF.js returned invalid or unsupported page count |
| `PDF_ENCRYPTED` | Encryption marker unsupported |
| `PDF_EOF` | Missing terminal EOF marker |
| `PDF_EXTERNAL_ENCODING` | External stream filter or parameters unsupported |
| `PDF_EXTERNAL_STREAM` | External stream file unsupported |
| `PDF_FLATE_TRUNCATED` | Flate stream is too short |
| `PDF_FONT_CODE_RANGE` | Font character code outside supported range |
| `PDF_FONT_CYCLE` | Cyclic font-width reference |
| `PDF_FONT_NUMBER` | Font measurement is not finite numeric data |
| `PDF_FONT_REFERENCE` | Missing font-width value or reference-depth limit |
| `PDF_FORBIDDEN_NAME` | Unsupported name |
| `PDF_FORM_XOBJECT` | Form XObject unsupported |
| `PDF_HEADER` | Unsupported PDF header |
| `PDF_HEX_CHARACTER` | Invalid hexadecimal string character |
| `PDF_HEX_UNTERMINATED` | Unterminated hexadecimal string |
| `PDF_HYBRID_XREF` | Hybrid cross-reference stream unsupported |
| `PDF_INCREMENTAL` | Incremental revision marker unsupported |
| `PDF_INDIRECT_LENGTH` | Indirect stream Length unsupported |
| `PDF_INDIRECT_TYPE` | Indirect structural Type or Subtype |
| `PDF_INITIALIZATION` | PDF.js module initialization failed |
| `PDF_INLINE_IMAGE` | Inline-image token detected in decoded bytes |
| `PDF_LEXER_DELIMITER` | Unexpected delimiter |
| `PDF_LEXER_EOF` | Expected another PDF token |
| `PDF_LEXER_EXPECTED_TOKEN` | Unexpected structural token |
| `PDF_LEXER_INTEGER` | Expected a nonnegative safe integer |
| `PDF_LOAD` | PDF.js document load failed |
| `PDF_NAME_ESCAPE` | Invalid escaped name |
| `PDF_NAME_LENGTH` | Name length limit exceeded |
| `PDF_NUMBER_RANGE` | Numeric value outside supported range |
| `PDF_OBJECT_ID` | Invalid, duplicate, or nonzero-generation object |
| `PDF_OBJECT_LIMIT` | Object count exceeded |
| `PDF_OBJECT_STREAM` | Object streams unsupported |
| `PDF_OPERATION` | PDF inspection or extraction operation failed |
| `PDF_PAGE_LIMIT` | Page count exceeded |
| `PDF_PREFLIGHT` | Unexpected preflight failure |
| `PDF_RAW_SIZE` | Empty PDF or raw byte limit exceeded |
| `PDF_REFERENCE_RANGE` | Unsupported object reference or generation |
| `PDF_STARTXREF` | startxref offset or trailing-token mismatch |
| `PDF_STREAM_BOUNDS` | Stream extends beyond file bytes |
| `PDF_STREAM_DICTIONARY` | Invalid stream dictionary or buffered stream tokens |
| `PDF_STREAM_EXPANSION` | Per-stream or aggregate expanded-byte limit exceeded |
| `PDF_STREAM_FILTER` | Other unsupported stream filter value |
| `PDF_STREAM_LENGTH` | Invalid direct stream length |
| `PDF_STREAM_NEWLINE` | Missing supported newline before stream bytes |
| `PDF_STRING_DEPTH` | String nesting limit exceeded |
| `PDF_STRING_UNTERMINATED` | Unterminated literal string |
| `PDF_TEXT_CHUNK` | PDF.js returned an invalid text chunk |
| `PDF_TIMEOUT` | PDF parse or extraction wall-time timeout |
| `PDF_TOKEN_LENGTH` | Token length limit exceeded |
| `PDF_TOKEN_LIMIT` | Token count exceeded |
| `PDF_TRAILER` | Trailer is not a dictionary |
| `PDF_TRAILER_ROOT_SIZE` | Invalid trailer root reference or size |
| `PDF_TYPE3_FONT` | Type3 fonts unsupported |
| `PDF_UNKNOWN` | Unrecognized PDF diagnostic |
| `PDF_VALUE_DEPTH` | Value nesting limit exceeded |
| `PDF_VALUE_SYNTAX` | Unsupported value syntax |
| `PDF_VERTICAL_WIDTHS` | Invalid vertical-width tuple length |
| `PDF_WIDTH_ARRAY` | Font widths are not an array |
| `PDF_WIDTH_BUDGET` | Aggregate font-width expansion exceeded |
| `PDF_WIDTH_CODE_RANGE` | Width array exceeds supported character range |
| `PDF_WIDTH_RANGE` | Reversed font-width range |
| `PDF_XFA` | XFA unsupported |
| `PDF_XREF_COUNT` | Invalid cross-reference subsection/count limit |
| `PDF_XREF_COVERAGE` | Cross-reference does not index every object |
| `PDF_XREF_ENTRY` | Cross-reference generation, uniqueness, or offset mismatch |
| `PDF_XREF_STATE` | Unsupported cross-reference entry state |
| `PDF_XREF_STREAM` | Cross-reference streams unsupported |
| `PDF_FILTER_ASCII85` | ASCII85Decode stream filter or A85 alias unsupported |
| `PDF_FILTER_ASCIIHEX` | ASCIIHexDecode stream filter or AHx alias unsupported |
| `PDF_FILTER_CCITT` | CCITTFaxDecode stream filter or CCF alias unsupported |
| `PDF_FILTER_CHAIN` | Empty or multiple-entry stream filter array unsupported |
| `PDF_FILTER_CRYPT` | Crypt stream filter unsupported |
| `PDF_FILTER_DCT` | DCTDecode (JPEG) stream filter or DCT alias unsupported |
| `PDF_FILTER_FLATE_ALIAS` | Fl abbreviation unsupported (full FlateDecode required) |
| `PDF_FILTER_INDIRECT` | Indirect stream filter or filter-array member unsupported |
| `PDF_FILTER_JBIG2` | JBIG2Decode stream filter unsupported |
| `PDF_FILTER_JPX` | JPXDecode (JPEG 2000) stream filter unsupported |
| `PDF_FILTER_LZW` | LZWDecode stream filter or LZW alias unsupported |
| `PDF_FILTER_RUNLENGTH` | RunLengthDecode stream filter or RL alias unsupported |

A known filter code can describe a direct name or a one-element filter array.
Multiple/empty arrays and indirect filters use fixed shape codes. Unknown values
stay generic; no arbitrary filter names or document values are returned.

## Operational use after separately authorized publication

Retry the existing `outlook_inspect_attachment` on the same message/attachment IDs.
Read the bracketed fixed code in the existing error response. If inspection passes,
use the existing bounded read tool to test extraction. Keep IDs and content private.
No original document bytes need to be exported for this first-stage diagnosis.
Do not remove the indicated guard as a shortcut: its resource-bounding role still
needs a verified replacement. Deployment and live verification remain separate
from local testing and are not implied by this patch.
