# Authored attachment-reader flow source

`definition.json` is a **new, sanitized, review-only Workflow Definition Language
source extension**, built by `scripts/build_attachment_flow.py`. It is not an
export of a running flow, an importable legacy ZIP, a Dataverse solution, or a
live-tested deployment template. No live flow was edited, run, imported, or
deployed to produce it.

The [original exported snapshot](../microsoft-bypass-flow/README.md) remains
unchanged. Its connector convention, immutable-ID headers, inbox query scope,
message field selections, and fail-closed gateway guard informed this authored
version. The snapshot's observed quirks remain documented with that snapshot;
they are not silently rewritten as historical behavior.

## Fixed operation contract

The Worker still sends an authenticated HTTP POST to the trigger. **All six
outbound connector request actions are GETs** to fixed Microsoft Graph paths;
there is no URL, method, body, `$expand`, or resource-path passthrough.

| Operation | Required `args` | Graph GET path under `https://graph.microsoft.com/v1.0/me/` |
| --- | --- | --- |
| `list_messages` | `{ "top": 1..50 }` | `mailFolders/inbox/messages` with the original metadata selection/order |
| `search_messages` | `{ "query": "nonempty string", "top": 1..50 }` | `mailFolders/inbox/messages` with an encoded search and original selection |
| `get_message` | `{ "messageId": "..." }` | `messages/{encoded messageId}` with the original full-message selection |
| `list_attachments` | `{ "messageId": "...", "top": 1..50, "skip": 0..10000 }` | `messages/{encoded messageId}/attachments?$select=id,name,contentType,size,isInline&$top={top}&$skip={skip}` |
| `get_attachment` | `{ "messageId": "...", "attachmentId": "..." }` | Metadata preflight on `messages/{encoded messageId}/attachments/{encoded attachmentId}?$select=id,name,contentType,size,isInline`, followed by the same resource without `$select` only after acceptance |

IDs are 1–2048 characters. Whitespace/control characters and exact `.` / `..` are
rejected; each ID is independently passed through `uriComponent`. A supplied
`%` is encoded again rather than treated as a pre-encoded path. Extra request
fields and operation arguments are rejected. `requestId` must be a UUID v4,
matching the Worker's `crypto.randomUUID()` calls.

The legacy three operations retain their successful data shapes and Graph
queries. This new version deliberately tightens their request validation,
requires `top` where applicable, returns HTTP 400 for missing IDs, and adds
explicit upstream failure/timeout responses. These changes apply only to the
new source, not the preserved exported snapshot.

### Successful responses

Each success is `{ "ok": true, "requestId": "...", "operation": "...", "data": ... }`.

- `list_attachments`: `data.value` contains only `@odata.type`, `id`, `name`,
  `contentType`, `size`, and `isInline`. The Graph request excludes
  `contentBytes`; a second Select projection also strips unexpected extra
  fields from the returned metadata. `data["@odata.nextLink"]` is an opaque
  string or `""` when absent. Neither the flow nor Worker follows this URL.
  The Worker uses only a nonempty string value to derive `hasMore`, strips the URL from
  MCP output, and calculates a bounded `nextOffset` itself
- `get_attachment`: `data` contains those six metadata fields plus
  `contentBytes`. Base64 is an **internal flow-to-Worker transport field**;
  it is never an MCP tool result. The Worker decodes it in request memory,
  validates the actual file bytes, extracts bounded PDF/DOCX/XLSX content, and
  returns inspection or reading results

The metadata preflight requires a nonnegative integer `size`, a matching ID,
and the exact `#microsoft.graph.fileAttachment` type annotation. Reference,
item, and unknown attachment types return a fixed error without fetching their
content, expanding nested items, or following any cloud link. Missing type
annotations fail closed; the flow must never assume that an attachment is a
file simply because its name has a supported extension.

The declared size must be at most **4 MiB (4,194,304 bytes)** before the content
request. The full response is revalidated for file type, ID, declared size,
base64 alphabet/padding and a length divisible by four. A simple character-class
pattern avoids a nested repeated-group pattern on multi-megabyte inputs. The
padded-base64 decoded length is independently
bounded at 4 MiB, including the two-byte edge case that an encoded-length ceiling
alone would permit. Metadata size and decoded size are not assumed equal.
The Worker independently enforces its decoded-byte and response-envelope limits.
The metadata preflight cannot prevent a changed or incorrect upstream size from
causing a larger connector response to be fetched; such a response is rejected
before success. The Worker provides another bounded response-consumption layer.

### Sanitized failures

| HTTP | Code | Cause |
| --- | --- | --- |
| 400 | `INVALID_REQUEST` | Invalid envelope, request ID, or operation allowlist |
| 400 | `INVALID_ARGUMENTS` | Missing, extra, malformed, or out-of-range operation arguments |
| 400 | `INVALID_OPERATION` | Defensive switch default |
| 413 | `ATTACHMENT_TOO_LARGE` | Metadata preflight exceeds 4 MiB |
| 415 | `UNSUPPORTED_ATTACHMENT_TYPE` | Metadata identifies a reference, item, or unknown type |
| 502 | `INVALID_ATTACHMENT_METADATA` | Invalid list/preflight metadata or wrong returned ID |
| 502 | `INVALID_ATTACHMENT_CONTENT` | Invalid, changed, or oversized full content response |
| 502 | `UPSTREAM_ERROR` | Connector failure or timeout, including Graph authorization/not-found/throttling errors |

Every authored error body is fixed text of the form
`{ "ok": false, "error": { "code": "...", "message": "..." } }`. It does not echo
arguments, IDs, attachment names, URLs, Graph error text, or connector outputs.
Graph retry policy is `none`; retries are not allowed to silently extend one
read into an unbounded connector operation. Responses specify `Cache-Control:
no-store`. Malformed HTTP bodies rejected before the trigger, gateway-condition
rejections, platform failures, and transport timeouts can still produce
platform-managed behavior rather than these authored envelopes.

## Authentication, privacy, and private configuration

- The `McpGatewayKey` parameter is `SecureString` with an empty default. The
  trigger checks that it is nonempty and exactly matches `X-MCP-Gateway-Key`;
  the checked-in definition cannot accept an empty/missing key
- `$authentication` / `$connections` have empty defaults. `shared_office365` is
  a generic connector alias, not an account-specific connection identifier
- No tenant/account/resource metadata, real gateway key, callback URL,
  connection reference, or mailbox content is included in this folder
- Trigger, connector, and Select action inputs/outputs are secured for run
  history. Parse JSON and Response support Secure Inputs only; enabling that
  also obscures their outputs. Control-flow actions have no unsupported
  `secureData` settings. Every downstream data action is explicitly secured
- These settings obscure run-history data; they do **not** prove that Microsoft
  retains no data. Review the environment's history retention, access policies,
  connector behavior, and monitoring settings before a real deployment. The
  Worker's no-persistence rule remains independent of external-service retention
- Preserve Cloudflare Access and the Worker's JWT check as well as this gateway
  guard. Tenant policies and the Outlook connection's permissions still apply

## Rebuild and offline checks

From the repository root:

```sh
python3 scripts/build_attachment_flow.py
python3 scripts/build_attachment_flow.py --check
python3 -m unittest discover -s scripts -p 'test_*flow.py' -v
python3 scripts/test_export_power_automate.py
```

The generator reads the preserved snapshot, writes only this authored JSON, and
never calls Microsoft. Tests pin the snapshot checksum, operation/argument parity
with Worker types, GET-only fixed endpoints, secure-data settings, sanitized
errors, and generation reproducibility. A small local interpreter executes the
**checked-in expression/action subset** with synthetic Graph results to exercise
request rejection, run-after failure paths, encoded IDs, metadata-only projection,
preflight gates, pagination, and size boundaries. It is not a Microsoft runtime
emulator. Passing tests establishes the offline source contract, not successful
Power Automate import or live connector behavior.

Do not run `export_power_automate.py` against this authored version; that exporter
is deliberately limited to the original three-operation historical snapshot.

## Authorized setup and validation still required

This JSON cannot be selected as a legacy flow package, and must not be passed to
`pac solution pack` as if it were a solution. A later, separately authorized
private deployment must use the appropriate Power Automate designer or a real
package/solution export-and-import workflow. The generic workflow source alone
does not provide connection references, environment bindings, or an installer.
In particular, custom workflow parameters are not guaranteed to be exposed by
the package importer; the maintainer must verify how the gateway key is privately
bound in their environment without replacing the fail-closed guard.

Before enabling any real copy, the maintainer must verify:

1. The permitted Office 365 Outlook connection is bound and the private gateway
   key matches `POWER_AUTOMATE_GATEWAY_KEY`; neither the key nor callback URL is
   added to Git, PR text, screenshots, or test output
2. The target designer/runtime accepts the authored action shapes, strict Parse
   JSON schemas, regular expressions, secure-history settings, and run-after paths
3. Graph/connector `$select` omits attachment bytes, preserves `@odata.type`, and
   returns the documented metadata shape; missing/invalid metadata stays an error
4. Graph/connector `$top` / `$skip` and `@odata.nextLink` behavior supports the
   fixed-offset pagination contract; no next-link URL is ever fetched
5. With explicitly authorized synthetic fixtures, file/type/size gates and
   connector failure/timeout responses behave as expected, and no base64 or
   source/download URL appears in an MCP tool result or application log
6. The existing three message tools still work and Cloudflare Access remains in
   force; any live validation uses the minimum necessary account data

These are outstanding deployment checks, not claims of completed live testing.

## Microsoft references

- [Get attachment: types, GET paths, and response shapes](https://learn.microsoft.com/en-us/graph/api/attachment-get?view=graph-rest-1.0)
- [List message attachments](https://learn.microsoft.com/en-us/graph/api/message-list-attachments?view=graph-rest-1.0)
- [Workflow action definitions: Parse JSON, Select, and Response](https://learn.microsoft.com/en-us/azure/logic-apps/logic-apps-workflow-actions-triggers)
- [Secure inputs/outputs and supported actions](https://learn.microsoft.com/en-us/azure/logic-apps/set-up-security-permissions)
- [Non-solution flow export/import](https://learn.microsoft.com/en-us/power-automate/export-import-flow-non-solution)
