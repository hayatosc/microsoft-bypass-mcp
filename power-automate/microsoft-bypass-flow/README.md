# Existing Power Automate flow: microsoft bypass flow

`definition.json` is the single canonical source for extending the existing
`microsoft bypass flow` in place. It preserves the same HTTP trigger, gateway-key
condition, operation switch, and public connector/auth parameter convention while
expanding the fixed read and draft-only operation surface.

This is sanitized Workflow Definition Language source for review and authorized
manual update. It is **not** an importable package, Dataverse solution, connector
creation script, or live deployment artifact.

## Current authored scope

The flow switch contains 14 backing operations for 16 MCP tools:

- Outlook mail: `list_messages`, `search_messages`, `get_message`,
  `list_mail_folders`, `get_conversation`
- Outlook attachments: `list_attachments`, `get_attachment`
- Outlook drafts: `create_draft`, `create_reply_draft`, `add_draft_attachment`
- OneDrive native: `onedrive_search_files`, `onedrive_list_folder`,
  `onedrive_get_metadata`, `onedrive_get_content`

`outlook_inspect_attachment` and `outlook_read_attachment` share
`get_attachment`. `onedrive_inspect_file` and `onedrive_read_file` share
`onedrive_get_content`.

## Exactly what changes from the sanitized fixture

The immutable fixture at `scripts/fixtures/microsoft-bypass-flow.pre-attachments.json`
remains a provenance fixture. The generator deliberately versions the mail branch
expectations instead of requiring the original three mail branches to stay byte-for-byte
unchanged.

Relative to that fixture, the generated source:

1. Replaces the trigger operation enum with the fixed flow-operation allowlist.
2. Keeps the trigger's authentication/key guard, switch, default branch and existing
   parameters, while adding two fixed, empty-default OneDrive search Compose settings.
3. Rebuilds the switch cases as sanitized fixed read and draft-only operations.
4. Adds controlled Outlook mailbox/folder/filter/pagination/conversation support.
5. Keeps attachment metadata/content branches with request-local byte transport and
   size/type/base64 gates.
6. Adds native OneDrive for Business branches, with supported native folder
   pagination aggregated into a bounded 1,000-item metadata window.
7. Adds draft creation, sender-reply draft creation, and attachment upload only
   after a fixed lookup verifies the matching target is a draft. No send action.

No generic proxy branch is introduced.

## Classic designer clipboard compatibility

After the read/draft augmentations, `scripts/build_attachment_flow.py` applies
`scripts/clipboard_compat.py` only to the completed `actions` subtree. It parses
WDL tokens and replaces actual empty string literals (`''`) with `string(null)`;
this is not a global apostrophe replacement. Nonempty strings, doubled-apostrophe
escapes, literal text, interpolation boundaries, JSON types, action topology and
connector/auth/security settings remain exact. Trigger, parameters (including
sanitized auth placeholders), and every other top-level value are untouched.
The immutable provenance fixture is never transformed.

The classic designer's native clipboard roundtrip previously corrupted 37 WDL
leaves containing empty literals. The validated compatibility candidate changes
1,149 tokens across those 37 action-value paths, with no key changes.
Microsoft's [official `string` reference](https://learn.microsoft.com/en-us/azure/logic-apps/expression-functions-reference#string)
explicitly guarantees `string(null)` produces an empty **String**, not null.

Independent **unsaved** native classic-designer clipboard roundtrip validation
passed for the candidate: all token values and types survived in the 37 affected
leaves; five WDL leaves changed only whitespace, and all 14 cases and connector
references remained. This was not a save, deployment, connector execution or mail
test. Local tests are offline lexical/source contracts and synthetic behavior
replays, **not Microsoft runtime tests**. Any save or live test still requires
separate authorization; this source integration performs neither.

## Outlook contract notes

`list_messages` supports inbox, sent items, all messages, or a specific mail
folder. Without filters it orders by `receivedDateTime desc`. With any controlled
filter it omits `$orderby` to avoid Graph `InefficientFilter` combinations. The
only filter properties are `isRead`, `hasAttachments`, `receivedAfter`, and
`receivedBefore`, in that order.

`get_conversation` uses `/me/messages` with exact escaped `conversationId`
equality. It has no `$orderby` and no client-side mailbox-slice fallback. The
Worker may sort each returned page locally and must reject returned items with a
mismatched `conversationId` or duplicate ID.

All Outlook IDs are independently `uriComponent`-encoded. A caller-supplied `%`
is encoded again and cannot become a path separator or query delimiter.

## Draft rollout notes

See [`docs/drafts.md`](../../docs/drafts.md) for the approval, argument, attachment
transfer, permission, and ambiguous-outcome contract. The existing Outlook
connection must support `Mail.ReadWrite`; do not add consent or credentials
without authorization. All draft POST actions disable retries. An ambiguous
failure may leave a draft or attachment saved, so inspect before retrying.
Attachment POST results must include bounded `contentBytes` exactly matching the
validated canonical request, a valid attachment ID, and the exact requested name.
Graph `size` is bounded nonnegative Int32 metadata, not raw-file length; success
projects verified raw-file bytes as `size` (1 byte through 2 MiB) for the Worker's existing
input/output checks. Missing or mismatching returned bytes fail closed as
`DRAFT_WRITE_AMBIGUOUS`, with no retry, fallback read, or extra endpoint. The
read-only attachment branches and their independent 4 MiB metadata/raw limits
are unchanged. Offline synthetic coverage includes raw 889 / metadata 1223 and
exactly 2 MiB raw with larger metadata; connector response behavior remains an
authorized rollout check, not a claim of live verification.
The public source remains sanitized and must not contain live bindings or keys.

## OneDrive contract notes

OneDrive branches use only the native OneDrive for Business connector:

| Branch | Native operation ID |
| --- | --- |
| `onedrive_search_files` | `FindFiles` |
| `onedrive_list_folder` without `folderId` | `ListRootFolder` |
| `onedrive_list_folder` with `folderId` | `ListFolderV2` |
| `onedrive_get_metadata` | `GetFileMetadata` |
| `onedrive_get_content` metadata preflight | `GetFileMetadata` |
| `onedrive_get_content` bytes | `GetFileContent` |

Only `ListFolderV2` enables native `paginationPolicy.minimumItemCount: 1000`.
The operation's Pagination setting was verified in the existing designer; its
native continuation happens within the connector runtime. The flow takes at
most 1,000 records and flags the threshold/continuation as potentially incomplete.
No arbitrary nextLink HTTP replay or guessed skip token is introduced.
`ListRootFolder` remains an array-returning operation. `FindFiles` remains capped
at 100 with no supported continuation; pagination does not expand search coverage.

The generated source intentionally contains these visible placeholders:

- `OneDriveSearchMode` (Compose configuration action, empty input by default)
- `OneDriveSearchRootId` (Compose configuration action, empty input by default)

They must be resolved or confirmed in the Power Automate designer during a
separately authorized manual update. The OneDrive for Business connector
documentation identifies the operation IDs and parameters, but this offline
stage does not verify the tenant-specific designer serialization for `findMode`
or root binding. Do not invent a different value in code review, and do not use
provider URLs as a substitute.

Search returns a sanitized HTTP 503 before connector access while either binding
is empty. Bind the `shared_onedriveforbusiness` connector alias separately to the
owner's connection. Neither setting is a secret or an OAuth credential. These named Compose actions
work in the ordinary non-solution cloud-flow designer without adding workflow
parameters. Their inputs must be fixed verified values, never caller expressions.

OneDrive content reads fetch fresh metadata before content, reject folders,
reject empty files, reject files over 4 MiB, and require a supported extension
and matching (or generic) MIME type before `GetFileContent`. Logic Apps binary bodies are handled as
`{ "$content-type", "$content" }`; `$content` is already base64. The flow
validates base64 shape, binary MIME and exact decoded/native file size, then returns only `{ metadata, contentBytes }` to the
Worker for request-local parsing.

## Successful response envelopes

Every successful branch returns:

```json
{ "ok": true, "requestId": "...", "operation": "...", "data": {} }
```

Outlook list/search/conversation/folder branches may include Graph nextLink in
flow-to-Worker data only. The Worker validates or discards it and never exposes
or replays a URL. OneDrive list/search branches expose bounded truncation/native
nextLink-presence flags rather than cursor URLs.

## Sanitized failure envelopes

New authored branches use fixed errors such as `INVALID_REQUEST`,
`INVALID_ARGUMENTS`, `UPSTREAM_ERROR`, `INVALID_ATTACHMENT_METADATA`,
`INVALID_ATTACHMENT_CONTENT`, `INVALID_ONEDRIVE_METADATA`,
`INVALID_ONEDRIVE_CONTENT`, `UNSUPPORTED_ONEDRIVE_FILE_TYPE`, and
`ONEDRIVE_FILE_TOO_LARGE`. Draft errors also use a fixed ambiguous-write warning
that instructs the caller to inspect Drafts before retrying.

Errors do not echo queries, subjects, IDs, names, URLs, raw connector text,
bytes, or base64. Platform-level trigger rejection or gateway rejection can still
produce platform-managed behavior.

## Secure history and privacy

Data-bearing OpenApiConnection and Select actions use secure inputs/outputs.
The existing Request trigger now explicitly secures outputs (including request
arguments and headers), without changing its identity or authentication guard.
Compose, Parse JSON and Response actions use secure inputs, which also hides their outputs
in supported Logic Apps/Power Automate run history. Control actions do not carry
unsupported `secureData` settings.

This obfuscates action history; it is not proof of Microsoft retention behavior.
The Worker no-persistence rule remains separate.

## Rebuild and offline checks

`build_attachment_flow.py` retains the original attachment cases, calls
`build_read_tools_flow.py` and `build_draft_tools_flow.py` to augment the same
definition, then applies the actions-only clipboard compatibility pass. No second
deployable flow is generated. Regression tests pin the original attachment hashes
as provenance, the precise compatible attachment content/hashes, the immutable
fixture, and deterministic current canonical bytes. Historical pre-pass and
validated-candidate hashes remain pinned by restoring only the three localized
draft attachment-result values in a test-only copy; the size fix itself has not
been roundtrip-tested in Microsoft's designer. The clipboard pass's 37 paths
and 1,149 replacements remain unchanged.

```sh
python3 scripts/build_attachment_flow.py
python3 scripts/build_attachment_flow.py --check
bun run test:flow
```

The tests are offline source-contract tests. They do not import this JSON, save a
flow, create connectors, call a mailbox, call OneDrive, or verify Microsoft
runtime behavior.

## Manual update path (separate authorization required)

1. Privately export and back up the existing flow and its current connection,
   gateway-key binding, and trigger URL. Keep the backup and secrets out of Git,
   PRs, screenshots, and logs.
2. Open the same existing flow in Power Automate. Do not create a new flow,
   use Save As, replace the trigger, or rebind connectors without explicit review.
3. Recreate the generated switch cases and trigger operation enum through the
   designer-supported fields. Preserve action names referenced by expressions.
4. Bind Outlook actions to the existing Office 365 Outlook connection convention.
5. Bind OneDrive actions to the signed-in OneDrive for Business connection and
   verify the official `FindFiles.findMode` machine value plus root behavior in
   that tenant/designer.
6. Preserve the private gateway-key guard. Do not paste public empty parameter
   defaults over a working private configuration.
7. Save only after reviewing the private diff. If the trigger callback URL
   changes, stop and reconcile the private Worker setting; never paste the URL
   into public text.
8. Only under explicit live-test authorization, smoke test minimal Outlook and
   OneDrive reads, attachment/file gates, and sanitized failures.

If the target designer cannot express a generated action, secure-history setting,
run-after branch, or OneDrive binding, stop for review. Do not silently create a
replacement flow or substitute a generic HTTP/Graph call.

## References

- OneDrive for Business connector operation IDs and parameters: https://learn.microsoft.com/en-us/connectors/onedriveforbusiness/
- Logic Apps binary body representation: https://learn.microsoft.com/en-us/azure/logic-apps/logic-apps-content-type
- Microsoft Graph message attachments: https://learn.microsoft.com/en-us/graph/api/message-list-attachments?view=graph-rest-1.0
- Microsoft Graph get attachment: https://learn.microsoft.com/en-us/graph/api/attachment-get?view=graph-rest-1.0
- WDL expressions and escaping: https://learn.microsoft.com/en-us/azure/logic-apps/workflow-definition-language-schema#expressions
- Secure inputs/outputs: https://learn.microsoft.com/en-us/azure/logic-apps/set-up-security-permissions
