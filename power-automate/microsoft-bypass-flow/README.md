# Existing Power Automate flow: microsoft bypass flow

`definition.json` is the **single canonical source for extending the existing
microsoft bypass flow in place**. Its HTTP trigger and existing operation switch
are retained. Two attachment cases are added to that switch; no parallel flow,
new endpoint, or replacement mail implementation is introduced.

This is sanitized Workflow Definition Language source for review, **not an
importable package, Dataverse solution, or live-tested deployment template**.
No live flow was edited, run, imported, or deployed to produce this change.

## Exactly what changes

Relative to the sanitized export from 2026-10-01 (UTC):

1. Append `list_attachments` and `get_attachment` to the existing HTTP trigger's
   `operation` enum
2. Add those two cases under the existing `actions.スイッチ.cases`

Everything else in the original definition is preserved: Japanese action names,
`manual` trigger settings and gateway guard, connection/auth parameters, switch
expression/default, and the complete `list_messages`, `get_message`, and
`search_messages` branches. Their inbox scope, Graph queries, Prefer headers,
response bodies, dependencies, and retry behavior are unchanged.

The new cases validate their own envelopes and arguments. There is no global
Parse JSON gate or tightened shared argument schema in front of legacy mail.
The existing shared trigger schema can reject malformed requests before any
case runs; those platform-managed responses are not replaced by this extension.

### Preserved mail behavior and export quirks

- List/search still default omitted `top` to 20; the trigger's 1–50 bound remains
- Legacy `requestId` remains any string; only attachment cases require UUID v4
- Missing `messageId` still returns HTTP 200, `ok: false`, `MISSING_MESSAGE_ID`
- The existing switch default remains HTTP 400 `INVALID_OPERATION`
- Mail connector failures/timeouts still have no explicit error-response action
- `contentVersion` remains the exported literal `"undefined"`
- The trigger retains `triggerAuthenticationType: "All"` and its key condition

These quirks are not fixed as a side effect of attachment support. Offline tests
prove preservation, not Microsoft schema/import acceptance of the export.

## Fixed operation contract

The Worker sends one authenticated HTTP POST to this same flow's trigger. All
six outbound connector actions are GETs through the existing generic
`shared_office365` alias and `$authentication` convention. No caller-controlled
URL, method, body, `$expand`, or resource path is accepted.

| Operation | Arguments | Fixed path under `https://graph.microsoft.com/v1.0/me/` |
| --- | --- | --- |
| `list_messages` | Existing `top`, optional with default 20 | `mailFolders/inbox/messages`, original selection/order |
| `search_messages` | Existing `query`, optional `top` | `mailFolders/inbox/messages`, original encoded search/selection |
| `get_message` | Existing `messageId` check | `messages/{encoded messageId}`, original full-message selection |
| `list_attachments` | Required `messageId`, `top` 1–50, `skip` 0–10000 | `messages/{encoded messageId}/attachments?$select=id,name,contentType,size,isInline&$top={top}&$skip={skip}` |
| `get_attachment` | Required `messageId`, `attachmentId` | Metadata preflight on `messages/{encoded messageId}/attachments/{encoded attachmentId}?$select=id,name,contentType,size,isInline`, then the same resource without `$select` only after acceptance |

Within the two new cases, IDs are 1–2048 characters. Whitespace/control characters
and exact `.` / `..` are rejected; each ID is independently passed through
`uriComponent`. A supplied `%` is encoded again rather than treated as a
pre-encoded path. Extra envelope/argument fields are rejected. `requestId` must
be UUID v4, matching the Worker's `crypto.randomUUID()` calls. These restrictions
do not alter the old mail branches.

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
| 400 | `INVALID_REQUEST` | Invalid envelope/request ID inside an attachment case |
| 400 | `INVALID_ARGUMENTS` | Missing, extra, malformed, or out-of-range operation arguments |
| 413 | `ATTACHMENT_TOO_LARGE` | Metadata preflight exceeds 4 MiB |
| 415 | `UNSUPPORTED_ATTACHMENT_TYPE` | Metadata identifies a reference, item, or unknown type |
| 502 | `INVALID_ATTACHMENT_METADATA` | Invalid list/preflight metadata or wrong returned ID |
| 502 | `INVALID_ATTACHMENT_CONTENT` | Invalid, changed, or oversized full content response |
| 502 | `UPSTREAM_ERROR` | Connector failure or timeout, including Graph authorization/not-found/throttling errors |

Every new attachment error body is fixed text of the form
`{ "ok": false, "error": { "code": "...", "message": "..." } }`. It does not echo
arguments, IDs, attachment names, URLs, Graph error text, or connector outputs.
Attachment Graph retry policy is `none`; retries are not allowed to silently extend one
read into an unbounded connector operation. Attachment responses specify `Cache-Control:
no-store`. Malformed HTTP bodies rejected before the trigger, gateway-condition
rejections, platform failures, and transport timeouts can still produce
platform-managed behavior rather than these authored envelopes.

## Authentication, privacy, and provenance

The original Package (.zip) export supplied `properties.definition` from
`Microsoft.Flow/flows/<id>/definition.json`. Its unchanged sanitized version is
stored at `scripts/fixtures/microsoft-bypass-flow.pre-attachments.json` solely as
a generator input and regression fixture. It is **not another deployable flow**.
Raw exports remain outside Git.

The original export had a literal secret in the `X-MCP-Gateway-Key` guard.
Sanitization replaced it with `parameters('McpGatewayKey')`, an empty-default
`SecureString`, and a non-empty check. This fail-closed public representation is
not a change already made to the live flow. `$authentication` / `$connections`
have empty public defaults; `shared_office365` is a generic alias, not an
account-specific connection ID. No real key, callback URL, tenant identity,
connection reference, package manifest, or mailbox data is published.

All new attachment connector/Select inputs and outputs are secured in run
history. New Parse JSON and Response actions use Secure Inputs (which also
obscures their outputs); control actions have no unsupported `secureData` setting.
The existing trigger and mail actions retain their exported history settings.
Do not claim this extension hides all historical mail data or the trigger's
request envelope. Obfuscation is not proof that Microsoft retains no data;
review environment retention/access policies separately. The Worker's
request-local/no-persistence rule is independent of Microsoft's retention.

## Rebuild and offline checks

```sh
python3 scripts/build_attachment_flow.py
python3 scripts/build_attachment_flow.py --check
bun run test:flow
```

The generator reads only the frozen sanitized baseline, appends the two cases
and enum entries, and writes only the canonical `definition.json` in this folder.
Make attachment-source changes in the generator, then regenerate; direct edits
to the generated JSON will be overwritten. The generator never reads its output
as a baseline and never calls Microsoft. A checksum pins the historical fixture.
Tests remove the additions and require complete equality with that fixture,
then check legacy default-top/missing-ID/failure behavior, new attachment gates,
fixed GET routes, source redaction, dependency paths and reproducibility.

A small interpreter executes the checked-in action/expression subset with
synthetic responses **after HTTP trigger admission**. Separate checks cover the
trigger schema. This is not a Microsoft runtime emulator or a full Power Automate
import/schema validation; the preserved `contentVersion` quirk is intentional.

`scripts/export_power_automate.py` remains a sanitizer for the original
three-operation export. Its tests use the historical fixture. Do not use it to
overwrite the canonical extended definition; a future five-operation re-export
needs its own private review and sanitizer update first.

## Update the existing flow in place (separate authorization required)

Use this one update path when a live update is separately authorized:

1. Privately export/retain a backup of the **existing** microsoft bypass flow and
   record its current connection binding, gateway configuration, and trigger URL.
   Keep the backup and secrets outside Git, PRs, screenshots, and logs
2. Open **that same flow → Edit** in Power Automate. Preserve its HTTP trigger,
   Office 365 Outlook connection, and existing three mail cases. Do not create a
   new flow, use Save As, replace the trigger, or rebind the connection just to
   add attachment support
3. In the existing HTTP trigger's request schema, append only the two operation
   enum values shown above. In the existing operation switch, add the two cases
   from this folder's `definition.json`, recreating their actions and expression
   fields in the designer. Retain the new action names because expressions refer
   to them. Configure their run-after branches, retry policy, and secure-history
   settings as specified by the source. The complete JSON is a review reference,
   not a file accepted by the legacy package importer or a whole-definition paste
   instruction; use the designer's supported fields for the existing flow
4. Preserve the existing private gateway-key guard and connection binding. Do not
   paste the public empty `$connections` / `$authentication` defaults or blank
   gateway parameter over a working private configuration. The public key
   parameter is a redaction convention; do not assume the package importer
   exposes custom parameters. Any deliberate change to private key binding must
   still fail closed and match the Worker's `POWER_AUTOMATE_GATEWAY_KEY`
5. Review the diff against the private backup, then save the same flow. Privately
   compare its resulting HTTP callback URL with the Worker's existing
   `POWER_AUTOMATE_URL`. Preserving the flow and trigger is the intended update
   path, **not a guarantee of URL continuity**. If it changed, stop and reconcile
   the private Worker setting before testing; never copy the URL into public text
6. Under explicit live-test authorization, verify all three existing mail tools
   and the attachment list/inspect/read path with minimal synthetic fixtures.
   Verify the runtime accepts Parse JSON schemas, regexes, run-after paths and
   secure-history settings; connector `$select` preserves `@odata.type` while
   excluding bytes; `$top`/`$skip` pagination works; and type/size/failure gates
   never leak bytes or URLs to MCP. Preserve Cloudflare Access/JWT protection

If the target designer cannot express the specified actions/settings, stop for
review of that concrete limitation rather than silently creating a replacement
flow. An alternative package/solution/API update is a separate deployment choice,
not a capability supplied by this JSON. Live save, binding, callback continuity,
and connector behavior remain unverified in this PR. Do not run this source
through `pac solution pack` as if it were a solution.

## Microsoft references

- [Get attachment: types, GET paths, and response shapes](https://learn.microsoft.com/en-us/graph/api/attachment-get?view=graph-rest-1.0)
- [List message attachments](https://learn.microsoft.com/en-us/graph/api/message-list-attachments?view=graph-rest-1.0)
- [Workflow action definitions: Parse JSON, Select, and Response](https://learn.microsoft.com/en-us/azure/logic-apps/logic-apps-workflow-actions-triggers)
- [Secure inputs/outputs and supported actions](https://learn.microsoft.com/en-us/azure/logic-apps/set-up-security-permissions)
- [Non-solution flow export/import](https://learn.microsoft.com/en-us/power-automate/export-import-flow-non-solution)
