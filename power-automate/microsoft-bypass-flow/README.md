# Power Automate source: microsoft bypass flow

`definition.json` is a sanitized, readable snapshot of the existing cloud flow,
exported through **Export → Package (.zip)** on 2026-10-01 (UTC). It comes from
`properties.definition` in the package's `Microsoft.Flow/flows/<id>/definition.json`.
The action names, expressions, request schema, Graph queries, and response bodies
are preserved from that export, except for the authentication redaction below.

This is a source-review artifact, **not an importable package or a tested deployment
template**. It is not a Dataverse solution export. No live flow was changed or run
when this snapshot was added. The raw ZIP stays outside Git.

## What the flow implements

- A Request / HTTP trigger named `manual`, accepting `operation`, `requestId`, and `args`
- Exactly `list_messages`, `search_messages`, and `get_message`
- Three GET requests through the Office 365 Outlook connector's `HttpRequest` action
- Fixed Microsoft Graph `/v1.0/me/` endpoints; list and search use the inbox
- `top` constrained to 1–50 by the trigger schema, with a flow-side default of 20
- Immutable message IDs; full-message bodies requested as plain text
- Success envelopes with `ok`, `requestId`, `operation`, and `data`

The connector alias `shared_office365` is generic and retained. It is not an
account-specific connection ID. Japanese action names are retained because other
expressions reference them by name.

## Redactions and safe configuration

The public source omits the resource ID/name, creator and tenant metadata,
modification metadata, account-specific connection references, and package
manifests/maps. No callback URL, gateway key, account identity, or mailbox data is
included.

The exported trigger used a literal secret in its `X-MCP-Gateway-Key` equality
condition. The only authentication transformation is:

1. Replace that literal with `parameters('McpGatewayKey')`
2. Add a `SecureString` parameter whose default is empty
3. Add a non-empty check so this public snapshot cannot accept an empty or missing
   gateway key

The parameter and non-empty check are **sanitization-derived**, not changes made
to the existing flow. This transformation is reversible when preparing an
explicitly authorized private deployment. It does not prove that the Power
Automate package importer exposes custom workflow parameters.

Before any deployment, privately configure the gateway key and bind a permitted
Office 365 Outlook connection. Use the same key as the Worker's
`POWER_AUTOMATE_GATEWAY_KEY`; keep it and the callback URL out of Git, PRs, logs,
and screenshots. Preserve both the gateway guard and Cloudflare Access protection.
Tenant policies and connector permissions still apply.

## Re-export and review

Export only the intended flow using **Package (.zip)**, and keep the ZIP outside
this checkout. From the repository root:

```sh
python3 scripts/export_power_automate.py /private/path/flow.zip \
  power-automate/microsoft-bypass-flow/definition.json
bun run test:flow
git diff -- power-automate/microsoft-bypass-flow/definition.json
```

The helper is deliberately specific to this flow. It rejects unexpected operation
or connector structures, unfamiliar gateway guards, retained identity markers,
and unexpected URL hosts. It does not replace manual secret review. If it fails,
inspect the input privately; never paste a raw definition or error containing
private values into an issue or PR. Review every changed expression before commit.

## Observed differences from the intended contract

These are recorded as exported, not silently corrected:

- `contentVersion` is the literal string `"undefined"`
- A missing `messageId` produces HTTP 200 with `ok: false` and
  `MISSING_MESSAGE_ID`; the MCP client rejects this as an invalid success envelope
- The unknown-operation branch returns HTTP 400 with `INVALID_OPERATION`
- Successful connector calls have response actions, but there is no explicit
  response action for upstream failure or timeout (and no configured HTTP 502
  envelope)
- The trigger uses `triggerAuthenticationType: "All"`; its explicit gateway-key
  condition is therefore important

Static tests protect the operation allowlist, GET-only calls, key redaction, and
these preserved behaviors. No import, deployment, or live mailbox integration test
has been performed.

## Import and lifecycle management

For an actual non-solution import, use a privately retained original package and
rebind its connection through Power Automate's package importer. This sanitized
JSON alone cannot be selected as a legacy package. Importing or updating a live flow
requires a separate authorized deployment step.

Microsoft recommends Dataverse solutions for ALM. Moving this flow into a
solution is a separate change; do not treat this folder as a solution package or
send it to `pac solution pack`.

- [Export and import a non-solution flow](https://learn.microsoft.com/en-us/power-automate/export-import-flow-non-solution)
- [Export a solution and its JSON workflow definitions](https://learn.microsoft.com/en-us/power-automate/export-flow-solution)
