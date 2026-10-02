# Outlook draft tools

This extension writes drafts in the signed-in mailbox. It does not send email,
delete messages, expose arbitrary Graph requests, or add mailbox impersonation.
The existing read tools retain their read-only annotations. The three draft
tools advertise `readOnlyHint: false` and `idempotentHint: false`.

## Approval and access

The MCP host must obtain the user’s approval before saving a draft or uploading
an attachment. Tool annotations are capability hints, not an approval system.
Approval to create a draft does not authorize sending it. Review and send in
Outlook using the user’s normal workflow.

The fixed Graph operations require delegated `Mail.ReadWrite`, as documented by
Microsoft. Existing Office 365 Outlook connector consent and tenant policy may
or may not permit them. Verify the existing connection during an authorized
rollout; do not silently add permissions, reauthenticate, or create a connector.
No `Mail.Send` operation is introduced by this implementation.

## Workflow

1. Call `outlook_create_draft` with `to`, optional `cc` and `bcc` arrays, `subject`,
   and plain-text `body`. The result identifies the saved draft
2. Or call `outlook_create_reply_draft` with an existing `messageId` and plain-text
   `body`. Microsoft creates a sender-reply draft; this is not reply-all. Outlook
   chooses the original message’s applicable reply recipient
3. If needed, call `outlook_add_draft_attachment` for the returned `draftId`, with
   a safe filename, MIME content type, and canonical base64 `contentBytes`
4. Review the draft and its attachments in Outlook before sending

Attachment creation is deliberately a separate operation. If it fails, the
previously created draft still exists. The attachment flow first reads the target
and requires its matching ID and `isDraft: true`; it does not attach to an
already-sent message. This check is not a transaction with concurrent Outlook
edits or sending. Avoid modifying/sending the same draft during an upload.

## Bounds and file transfer

- At most 50 recipients total across `to`, `cc`, and `bcc`
- Subject at most 512 characters; plain-text body at most 20,000 characters
- One attachment per call, nonempty, at most 2 MiB decoded bytes
- Filenames at most 255 characters, without paths or control characters
- Canonical base64 only; no data URLs, upload URLs, remote fetches, upload
  sessions, cloud-file pointers, or arbitrary paths
- Responses contain bounded identifiers and metadata only, never attachment
  bytes/base64 or raw Graph objects

A chat attachment is not automatically readable by this MCP server. The host
must first materialize the user-approved file through its supported file API,
read the exact bytes locally, verify size, and base64-encode those bytes for the
typed attachment tool. Do not substitute a chat download URL, a local path, or a
Library ID for `contentBytes`. A host without that materialization capability
must ask the user to attach the file in Outlook. Large base64 tool arguments may
also exceed a host’s own tool-call limit even below the server’s limit.

The Worker and flow keep data request-local; the requested saved draft and its
attachments persist only in the user’s mailbox. The flow secures data-bearing
run-history inputs/outputs and does not add storage or caches.

## Failure and retry behavior

Draft/attachment POST actions have retries disabled. A timeout, connection loss,
or malformed success response can occur after Microsoft has written the draft
or attachment. Inspect Drafts and the target draft’s attachments before retrying;
otherwise a retry can create duplicates. Request IDs are correlation values,
not idempotency keys. The tool cannot guarantee exactly-once writes.

Do not log recipients, subjects, bodies, file names, IDs, or attachment bytes.
Only the existing bounded request telemetry fields are logged. Raw connector
errors are never returned to the MCP caller.

## Fixed Microsoft operations

- [Create a message draft](https://learn.microsoft.com/en-us/graph/api/user-post-messages):
  `POST /v1.0/me/messages`
- [Create a reply draft](https://learn.microsoft.com/en-us/graph/api/message-createreply):
  `POST /v1.0/me/messages/{encodedMessageId}/createReply`
- [Add a small attachment](https://learn.microsoft.com/en-us/graph/api/message-post-attachments):
  `POST /v1.0/me/messages/{encodedDraftId}/attachments`, only after a fixed
  message lookup verifies the target is a draft
- [Office 365 Outlook connector HTTP action](https://learn.microsoft.com/en-us/connectors/office365/#send-an-http-request)
  explicitly supports the `/me/messages` resource family

The generated flow is review/update material, not evidence of successful live
execution. Import, consent checks, deployment, and a live draft-write smoke test
require separate authorization. Offline tests use synthetic mailbox data only.
