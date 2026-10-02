"""Extend the existing flow source in place; never contacts Power Automate.

The frozen sanitized export is a compatibility fixture, not a second flow.
Only the operation allowlist and two attachment cases are added to that baseline.
"""

from copy import deepcopy
import json
from pathlib import Path
import sys
from urllib.parse import quote

ROOT = Path(__file__).resolve().parents[1]
BASELINE = ROOT / "scripts/fixtures/microsoft-bypass-flow.pre-attachments.json"
DESTINATION = ROOT / "power-automate/microsoft-bypass-flow/definition.json"
MAIL_OPERATIONS = ("list_messages", "get_message", "search_messages")
ATTACHMENT_OPERATIONS = ("list_attachments", "get_attachment")
MAX_BYTES = 4 * 1024 * 1024
MAX_BASE64 = ((MAX_BYTES + 2) // 3) * 4
BASE64_CHUNK = 8192
METADATA_FIELDS = ["@odata.type", "id", "name", "contentType", "size", "isInline"]
GRAPH_ROOT = "https://graph.microsoft.com/v1.0/me/"
# Percent signs are allowed; uriComponent encodes them again, so caller-supplied
# encoded separators never become structural path separators.
ID_SCHEMA = {
    "type": "string", "minLength": 1, "maxLength": 2048,
}
# Union of the existing JavaScript/Python whitespace interpretations, plus C0
# and DEL. In particular, reject both NEL (U+0085) and BOM (U+FEFF).
FORBIDDEN_ID_CODEPOINTS = (*range(0x20), 0x20, 0x7F, 0x85, 0xA0, 0x1680,
    *range(0x2000, 0x200B), 0x2028, 0x2029, 0x202F, 0x205F, 0x3000, 0xFEFF)
TOP_SCHEMA = {"type": "integer", "minimum": 1, "maximum": 50}
SKIP_SCHEMA = {"type": "integer", "minimum": 0, "maximum": 10000}
ARGUMENTS = {
    "list_messages": {"top": TOP_SCHEMA},
    "search_messages": {"query": {"type": "string", "minLength": 1}, "top": TOP_SCHEMA},
    "get_message": {"messageId": ID_SCHEMA},
    "list_attachments": {"messageId": ID_SCHEMA, "top": TOP_SCHEMA, "skip": SKIP_SCHEMA},
    "get_attachment": {"messageId": ID_SCHEMA, "attachmentId": ID_SCHEMA},
}


def object_schema(properties, *, strict=False):
    result = {"type": "object", "properties": deepcopy(properties), "required": list(properties)}
    if strict:
        result["additionalProperties"] = False
    return result


def secure(action):
    # ParseJson and Response support Secure Inputs only, which also obscures
    # their outputs. Control-flow actions do not support secureData.
    properties = ["inputs"] if action["type"] in {"ParseJson", "Response"} else ["inputs", "outputs"]
    action["runtimeConfiguration"] = {"secureData": {"properties": properties}}
    return action


def after(name, *statuses):
    return {name: list(statuses or ("Succeeded",))}


def response(status, body, run_after=None):
    return secure({
        "type": "Response", "kind": "Http", "runAfter": run_after or {},
        "inputs": {"statusCode": status, "headers": {"Content-Type": "application/json", "Cache-Control": "no-store"}, "body": body},
    })


def error(status, code, message, run_after=None):
    # No request IDs, Graph errors, URLs, connector outputs, or argument echoes
    # in errors; even an invalid envelope gets a fixed, non-sensitive response.
    return response(status, {"ok": False, "error": {"code": code, "message": message}}, run_after)


def success(data, run_after=None):
    return response(200, {
        "ok": True,
        "requestId": "@triggerBody()?['requestId']",
        "operation": "@triggerBody()?['operation']",
        "data": data,
    }, run_after)


def literal_key(value):
    return "@" + value if value.startswith("@") else value


def schema_literals(value):
    # WDL evaluates JSON property names and string values even inside schemas.
    # Escape literal leading @ characters while retaining their runtime meaning.
    if isinstance(value, dict):
        return {literal_key(key): schema_literals(child) for key, child in value.items()}
    if isinstance(value, list):
        return [schema_literals(child) for child in value]
    return "@" + value if isinstance(value, str) and value.startswith("@") else value


def parse(content, schema, run_after=None):
    return secure({"type": "ParseJson", "runAfter": run_after or {}, "inputs": {"content": content, "schema": schema_literals(schema)}})


def graph(uri, run_after=None):
    parameters = {
        "Uri": uri, "Method": "GET", "CustomHeader1": 'Prefer: IdType="ImmutableId"',
        "ContentType": "application/json",
    }
    return secure({
        "type": "OpenApiConnection", "runAfter": run_after or {},
        "inputs": {
            "parameters": parameters,
            "host": {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
                "connectionName": "shared_office365", "operationId": "HttpRequest",
            },
            "authentication": "@parameters('$authentication')",
            "retryPolicy": {"type": "none"},
        },
    })


def failure(actions, dependency, *, status=502, code="UPSTREAM_ERROR", message="The upstream read could not be completed."):
    actions[f"Error_{dependency}"] = error(status, code, message, after(dependency, "Failed", "TimedOut"))


def condition(expression, actions, otherwise, run_after=None):
    return {"type": "If", "expression": expression, "actions": actions, "else": {"actions": otherwise}, "runAfter": run_after or {}}


def field(action, name):
    return f"@body('{action}')?['{name}']"


def projection(action):
    return {literal_key(key): field(action, key) for key in METADATA_FIELDS}


def uuid_v4_guard(value):
    # Called only after the request schema has enforced a 36-character string.
    stripped = f"toLower(replace({value}, '-', ''))"
    remainder = stripped
    for char in "0123456789abcdef":
        remainder = f"replace({remainder}, '{char}', '')"
    checks = [f"equals(length({stripped}), 32)", f"equals({remainder}, '')"]
    checks += [f"equals(substring({value}, {offset}, 1), '-')" for offset in (8, 13, 18, 23)]
    checks += [f"equals(substring({value}, 14, 1), '4')", f"contains('89ab', toLower(substring({value}, 19, 1)))"]
    return "and(" + ", ".join(checks) + ")"


def id_guard(value):
    # Called only after the schema has bounded/typed the ID. Percent signs from
    # the caller are encoded to %25, so literal strings such as '%20' stay valid.
    encoded = f"uriComponent({value})"
    cleaned = encoded
    # UTF-8 never encodes a non-control character with byte 00..1F. These two
    # prefixes replace 32 nested calls without rejecting literal '%0'/'%1'.
    forbidden = ("%0", "%1", *(quote(chr(codepoint), safe="") for codepoint in FORBIDDEN_ID_CODEPOINTS if codepoint >= 0x20))
    for encoded_character in forbidden:
        cleaned = f"replace({cleaned}, '{encoded_character}', '')"
    return f"and(not(equals({value}, '.')), not(equals({value}, '..')), equals(length({encoded}), length({cleaned})))"


def base64_alphabet_guard(value):
    # URI-safe ASCII consists of alphanumerics plus -_.!~*'(). Permit +/= via
    # explicit encoding, reject the remaining punctuation. Apply only to small
    # chunks, never construct a multi-megabyte string expression result.
    escaped = value
    for char, encoded in (("+", "%2B"), ("/", "%2F"), ("=", "%3D")):
        escaped = f"replace({escaped}, '{char}', '{encoded}')"
    checks = [f"equals(uriComponent({value}), {escaped})"]
    checks += [f"not(contains({value}, '{char.replace(chr(39), chr(39) * 2)}'))" for char in "-_.!~*'()"]
    return "and(" + ", ".join(checks) + ")"


def request_args(operation):
    # Validate only inside the new cases. A global validator would tighten the
    # legacy mail contract (optional top, arbitrary string requestId, etc.).
    envelope = f"Validate_{operation}_request"
    envelope_schema = object_schema({
        "operation": {"type": "string", "enum": [operation]},
        "requestId": {"type": "string", "minLength": 36, "maxLength": 36},
        "args": {"type": "object"},
    }, strict=True)
    actions = {envelope: parse("@triggerBody()", envelope_schema)}
    failure(actions, envelope, status=400, code="INVALID_REQUEST", message="The request envelope is invalid.")
    request_format = f"Validate_{operation}_request_format"
    actions[request_format] = parse("@" + uuid_v4_guard(f"body('{envelope}')?['requestId']"), {"type": "boolean", "enum": [True]}, after(envelope))
    failure(actions, request_format, status=400, code="INVALID_REQUEST", message="The request envelope is invalid.")
    name = f"Validate_{operation}_args"
    actions[name] = parse(f"@body('{envelope}')?['args']", object_schema(ARGUMENTS[operation], strict=True), after(request_format))
    failure(actions, name, status=400, code="INVALID_ARGUMENTS", message="The operation arguments are invalid.")
    ids = f"Validate_{operation}_ids"
    checks = [id_guard(f"body('{name}')?['{field}']") for field in ARGUMENTS[operation] if field.endswith("Id")]
    guard = checks[0] if len(checks) == 1 else "and(" + ", ".join(checks) + ")"
    actions[ids] = parse("@" + guard, {"type": "boolean", "enum": [True]}, after(name))
    failure(actions, ids, status=400, code="INVALID_ARGUMENTS", message="The operation arguments are invalid.")
    return name, ids, actions


def metadata_schema():
    return object_schema({
        "@odata.type": {"type": "string", "minLength": 1, "maxLength": 128},
        "id": ID_SCHEMA,
        "name": {"type": "string", "maxLength": 512},
        "contentType": {"type": "string", "maxLength": 256},
        "size": {"type": "integer", "minimum": 0},
        "isInline": {"type": "boolean"},
    })


def list_case():
    operation = "list_attachments"
    validate, validated_ids, actions = request_args(operation)
    uri = (
        f"@concat('{GRAPH_ROOT}messages/', uriComponent(body('{validate}')?['messageId']), "
        "'/attachments?$select=id,name,contentType,size,isInline&$top=', "
        f"string(body('{validate}')?['top']), '&$skip=', string(body('{validate}')?['skip']))"
    )
    actions["Graph_list_attachments"] = graph(uri, after(validated_ids))
    failure(actions, "Graph_list_attachments")
    schema = object_schema({"value": {"type": "array", "maxItems": 50, "items": metadata_schema()}})
    schema["properties"]["@odata.nextLink"] = {"type": "string"}
    actions["Validate_attachment_list"] = parse("@body('Graph_list_attachments')", schema, after("Graph_list_attachments"))
    failure(actions, "Validate_attachment_list", code="INVALID_ATTACHMENT_METADATA", message="The upstream attachment metadata is invalid.")
    selected = {literal_key(key): f"@item()?['{key}']" for key in METADATA_FIELDS}
    selected["id"] = "@if(" + id_guard("item()?['id']") + ", item()?['id'], null)"
    actions["Select_attachment_metadata"] = secure({
        "type": "Select", "runAfter": after("Validate_attachment_list"),
        "inputs": {"from": "@body('Validate_attachment_list')?['value']", "select": selected},
    })
    failure(actions, "Select_attachment_metadata", code="INVALID_ATTACHMENT_METADATA", message="The upstream attachment metadata is invalid.")
    actions["Validate_projected_attachment_list"] = parse("@body('Select_attachment_metadata')", {"type": "array", "maxItems": 50, "items": metadata_schema()}, after("Select_attachment_metadata"))
    failure(actions, "Validate_projected_attachment_list", code="INVALID_ATTACHMENT_METADATA", message="The upstream attachment metadata is invalid.")
    actions["Respond_list_attachments"] = success({
        "value": "@body('Select_attachment_metadata')",
        # This is opaque data only. No action ever evaluates/fetches this URL;
        # the Worker discards it after computing hasMore and builds its own cursor.
        "@@odata.nextLink": "@coalesce(body('Validate_attachment_list')?['@odata.nextLink'], '')",
    }, after("Validate_projected_attachment_list"))
    return {"case": operation, "actions": actions}


def get_case():
    operation = "get_attachment"
    validate, validated_ids, actions = request_args(operation)
    path = (
        f"@concat('{GRAPH_ROOT}messages/', uriComponent(body('{validate}')?['messageId']), "
        f"'/attachments/', uriComponent(body('{validate}')?['attachmentId'])"
    )
    actions["Graph_attachment_metadata"] = graph(path + ", '?$select=id,name,contentType,size,isInline')", after(validated_ids))
    failure(actions, "Graph_attachment_metadata")
    actions["Validate_attachment_metadata"] = parse("@body('Graph_attachment_metadata')", metadata_schema(), after("Graph_attachment_metadata"))
    failure(actions, "Validate_attachment_metadata", code="INVALID_ATTACHMENT_METADATA", message="The upstream attachment metadata is invalid.")

    content_actions = {"Graph_attachment_content": graph(path + ")")}
    failure(content_actions, "Graph_attachment_content")
    content_schema = metadata_schema()
    content_schema["properties"]["@odata.type"]["enum"] = ["#microsoft.graph.fileAttachment"]
    content_schema["properties"]["size"]["maximum"] = MAX_BYTES
    content_schema["properties"]["contentBytes"] = {
        "type": "string", "maxLength": MAX_BASE64,
    }
    content_schema["required"].append("contentBytes")
    content_actions["Validate_attachment_content"] = parse("@body('Graph_attachment_content')", content_schema, after("Graph_attachment_content"))
    failure(content_actions, "Validate_attachment_content", code="INVALID_ATTACHMENT_CONTENT", message="The upstream attachment content is invalid or exceeds the limit.")
    content_actions["Select_attachment_content_alphabet"] = secure({
        "type": "Select", "runAfter": after("Validate_attachment_content"),
        "inputs": {"from": f"@chunk(body('Validate_attachment_content')?['contentBytes'], {BASE64_CHUNK})", "select": "@" + base64_alphabet_guard("item()")},
    })
    failure(content_actions, "Select_attachment_content_alphabet", code="INVALID_ATTACHMENT_CONTENT", message="The upstream attachment content is invalid or exceeds the limit.")
    content_actions["Validate_attachment_content_alphabet"] = parse("@body('Select_attachment_content_alphabet')", {"type": "array", "maxItems": (MAX_BASE64 + BASE64_CHUNK - 1) // BASE64_CHUNK, "items": {"type": "boolean", "enum": [True]}}, after("Select_attachment_content_alphabet"))
    failure(content_actions, "Validate_attachment_content_alphabet", code="INVALID_ATTACHMENT_CONTENT", message="The upstream attachment content is invalid or exceeds the limit.")
    # A padded base64 string's encoded-length ceiling alone allows up to two
    # extra decoded bytes. Check the decoded length without materializing bytes.
    b64 = "body('Validate_attachment_content')?['contentBytes']"
    decoded_size = f"sub(mul(div(length({b64}), 4), 3), if(endsWith({b64}, '=='), 2, if(endsWith({b64}, '='), 1, 0)))"
    accepted = projection("Validate_attachment_content")
    accepted["contentBytes"] = "@body('Validate_attachment_content')?['contentBytes']"
    content_actions["Check_content_size_and_identity"] = condition({"and": [
        {"lessOrEquals": ["@" + decoded_size, MAX_BYTES]},
        {"equals": [field("Validate_attachment_content", "id"), field(validate, "attachmentId")]},
        {"equals": [f"@mod(length({b64}), 4)", 0]},
        {"equals": [f"@or(equals(indexOf({b64}, '='), -1), equals(indexOf({b64}, '='), sub(length({b64}), if(endsWith({b64}, '=='), 2, 1))))", True]},
    ]}, {"Respond_get_attachment": success(accepted)}, {
        "Reject_attachment_content": error(502, "INVALID_ATTACHMENT_CONTENT", "The upstream attachment content is invalid or exceeds the limit.")
    }, after("Validate_attachment_content_alphabet"))

    size_gate = condition({"lessOrEquals": [field("Validate_attachment_metadata", "size"), MAX_BYTES]}, content_actions, {
        "Reject_attachment_size": error(413, "ATTACHMENT_TOO_LARGE", "The attachment exceeds the 4 MiB limit.")
    })
    type_gate = condition({"equals": [field("Validate_attachment_metadata", "@odata.type"), "#microsoft.graph.fileAttachment"]}, {
        "Check_attachment_size": size_gate,
    }, {"Reject_attachment_type": error(415, "UNSUPPORTED_ATTACHMENT_TYPE", "Only file attachments can be read.")})
    actions["Check_attachment_identity"] = condition({"equals": [field("Validate_attachment_metadata", "id"), field(validate, "attachmentId")]}, {
        "Check_attachment_type": type_gate,
    }, {"Reject_attachment_metadata": error(502, "INVALID_ATTACHMENT_METADATA", "The upstream attachment metadata is invalid.")}, after("Validate_attachment_metadata"))
    return {"case": operation, "actions": actions}


def build_definition():
    definition = json.loads(BASELINE.read_text(encoding="utf-8"))
    # Preserve the existing HTTP trigger, auth guard, connection parameters,
    # switch/default and complete mail branches, including exported quirks.
    # Attachment-only argument validation belongs in the new cases below;
    # no global schema fields or strictness are changed for legacy mail calls.
    operation_schema = definition["triggers"]["manual"]["inputs"]["schema"]["properties"]["operation"]
    operation_schema["enum"].extend(ATTACHMENT_OPERATIONS)
    cases = definition["actions"]["スイッチ"]["cases"]
    cases["list_attachments"] = list_case()
    cases["get_attachment"] = get_case()
    from build_read_tools_flow import augment_read_tools
    return augment_read_tools(definition)


def main():
    rendered = json.dumps(build_definition(), ensure_ascii=False, indent=2) + "\n"
    if "--check" in sys.argv:
        if not DESTINATION.exists() or DESTINATION.read_text(encoding="utf-8") != rendered:
            raise SystemExit("Existing flow source is stale; run scripts/build_attachment_flow.py.")
    elif len(sys.argv) == 1:
        DESTINATION.parent.mkdir(parents=True, exist_ok=True)
        DESTINATION.write_text(rendered, encoding="utf-8")
    else:
        raise SystemExit("Usage: python3 scripts/build_attachment_flow.py [--check]")


if __name__ == "__main__":
    main()
