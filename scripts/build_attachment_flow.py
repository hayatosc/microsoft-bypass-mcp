"""Build the review-only attachment flow source; never contacts Power Automate.

This is authored Workflow Definition Language, not a sanitized live export.
The original snapshot is read for its connector and gateway contract, not edited.
"""

from copy import deepcopy
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
SNAPSHOT = ROOT / "power-automate/microsoft-bypass-flow/definition.json"
DESTINATION = ROOT / "power-automate/attachment-reader-flow/definition.json"
MAX_BYTES = 4 * 1024 * 1024
MAX_BASE64 = ((MAX_BYTES + 2) // 3) * 4
METADATA_FIELDS = ["@odata.type", "id", "name", "contentType", "size", "isInline"]
GRAPH_ROOT = "https://graph.microsoft.com/v1.0/me/"
# Percent signs are allowed; uriComponent encodes them again, so caller-supplied
# encoded separators never become structural path separators.
ID_SCHEMA = {
    "type": "string", "minLength": 1, "maxLength": 2048,
    "pattern": r"^(?!\.{1,2}$)[^\s\u0000-\u001f\u007f]+$(?![\s\S])",
}
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
        "requestId": "@body('Validate_request')?['requestId']",
        "operation": "@body('Validate_request')?['operation']",
        "data": data,
    }, run_after)


def parse(content, schema, run_after=None):
    return secure({"type": "ParseJson", "runAfter": run_after or {}, "inputs": {"content": content, "schema": deepcopy(schema)}})


def graph(uri, run_after=None, *, text_body=False):
    parameters = {
        "Uri": uri, "Method": "GET", "CustomHeader1": 'Prefer: IdType="ImmutableId"',
        "ContentType": "application/json",
    }
    if text_body:
        parameters["CustomHeader2"] = 'Prefer: outlook.body-content-type="text"'
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
    return {key: field(action, key) for key in METADATA_FIELDS}


def request_args(operation):
    name = f"Validate_{operation}_args"
    actions = {name: parse("@body('Validate_request')?['args']", object_schema(ARGUMENTS[operation], strict=True))}
    failure(actions, name, status=400, code="INVALID_ARGUMENTS", message="The operation arguments are invalid.")
    return name, actions


def metadata_schema():
    return object_schema({
        "@odata.type": {"type": "string", "minLength": 1, "maxLength": 128},
        "id": ID_SCHEMA,
        "name": {"type": "string", "maxLength": 512},
        "contentType": {"type": "string", "maxLength": 256},
        "size": {"type": "integer", "minimum": 0},
        "isInline": {"type": "boolean"},
    })


def message_case(operation, snapshot):
    validate, actions = request_args(operation)
    old_case = snapshot["actions"]["スイッチ"]["cases"][operation]

    def find_request(value):
        if isinstance(value, dict):
            if value.get("type") == "OpenApiConnection":
                return value
            for child in value.values():
                found = find_request(child)
                if found:
                    return found
        return None

    # Preserve all legacy Graph selections, inbox scope, encoded IDs/search,
    # and Prefer headers. Arguments now come from explicit validated bodies.
    original = find_request(old_case)
    uri = original["inputs"]["parameters"]["Uri"]
    uri = uri.replace("triggerBody()?['args']", f"body('{validate}')")
    name = f"Graph_{operation}"
    actions[name] = graph(uri, after(validate), text_body=operation == "get_message")
    failure(actions, name)
    actions[f"Respond_{operation}"] = success(f"@body('{name}')", after(name))
    return {"case": operation, "actions": actions}


def list_case():
    operation = "list_attachments"
    validate, actions = request_args(operation)
    uri = (
        f"@concat('{GRAPH_ROOT}messages/', uriComponent(body('{validate}')?['messageId']), "
        "'/attachments?$select=id,name,contentType,size,isInline&$top=', "
        f"string(body('{validate}')?['top']), '&$skip=', string(body('{validate}')?['skip']))"
    )
    actions["Graph_list_attachments"] = graph(uri, after(validate))
    failure(actions, "Graph_list_attachments")
    schema = object_schema({"value": {"type": "array", "maxItems": 50, "items": metadata_schema()}})
    schema["properties"]["@odata.nextLink"] = {"type": "string"}
    actions["Validate_attachment_list"] = parse("@body('Graph_list_attachments')", schema, after("Graph_list_attachments"))
    failure(actions, "Validate_attachment_list", code="INVALID_ATTACHMENT_METADATA", message="The upstream attachment metadata is invalid.")
    actions["Select_attachment_metadata"] = secure({
        "type": "Select", "runAfter": after("Validate_attachment_list"),
        "inputs": {"from": "@body('Validate_attachment_list')?['value']", "select": {key: f"@item()?['{key}']" for key in METADATA_FIELDS}},
    })
    failure(actions, "Select_attachment_metadata", code="INVALID_ATTACHMENT_METADATA", message="The upstream attachment metadata is invalid.")
    actions["Respond_list_attachments"] = success({
        "value": "@body('Select_attachment_metadata')",
        # This is opaque data only. No action ever evaluates/fetches this URL;
        # the Worker discards it after computing hasMore and builds its own cursor.
        "@odata.nextLink": "@coalesce(body('Validate_attachment_list')?['@odata.nextLink'], '')",
    }, after("Select_attachment_metadata"))
    return {"case": operation, "actions": actions}


def get_case():
    operation = "get_attachment"
    validate, actions = request_args(operation)
    path = (
        f"@concat('{GRAPH_ROOT}messages/', uriComponent(body('{validate}')?['messageId']), "
        f"'/attachments/', uriComponent(body('{validate}')?['attachmentId'])"
    )
    actions["Graph_attachment_metadata"] = graph(path + ", '?$select=id,name,contentType,size,isInline')", after(validate))
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
        "pattern": r"^[A-Za-z0-9+/]*={0,2}$(?![\s\S])",
    }
    content_schema["required"].append("contentBytes")
    content_actions["Validate_attachment_content"] = parse("@body('Graph_attachment_content')", content_schema, after("Graph_attachment_content"))
    failure(content_actions, "Validate_attachment_content", code="INVALID_ATTACHMENT_CONTENT", message="The upstream attachment content is invalid or exceeds the limit.")
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
    ]}, {"Respond_get_attachment": success(accepted)}, {
        "Reject_attachment_content": error(502, "INVALID_ATTACHMENT_CONTENT", "The upstream attachment content is invalid or exceeds the limit.")
    }, after("Validate_attachment_content"))

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
    snapshot = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
    envelope_schema = object_schema({
        "operation": {"type": "string", "enum": list(ARGUMENTS)},
        "requestId": {"type": "string", "minLength": 36, "maxLength": 36, "pattern": r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-4[0-9a-fA-F]{3}-[89aAbB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}$"},
        "args": {"type": "object"},
    }, strict=True)
    manual = deepcopy(snapshot["triggers"]["manual"])
    # Runtime ParseJson (rather than an early trigger schema failure) lets bad
    # request shapes receive our fixed 400 body. The gateway remains fail-closed.
    manual["inputs"]["schema"] = {}
    secure(manual)
    actions = {"Validate_request": parse("@triggerBody()", envelope_schema)}
    failure(actions, "Validate_request", status=400, code="INVALID_REQUEST", message="The request envelope is invalid.")
    cases = {name: message_case(name, snapshot) for name in ("list_messages", "search_messages", "get_message")}
    cases["list_attachments"] = list_case()
    cases["get_attachment"] = get_case()
    actions["Route_operation"] = {
        "type": "Switch", "runAfter": after("Validate_request"),
        "expression": "@body('Validate_request')?['operation']", "cases": cases,
        "default": {"actions": {"Reject_operation": error(400, "INVALID_OPERATION", "The operation is not supported.")}},
    }
    return {
        "$schema": snapshot["$schema"], "contentVersion": "1.0.0.0",
        "parameters": deepcopy(snapshot["parameters"]), "triggers": {"manual": manual},
        "actions": actions,
        "description": "Authored read-only attachment extension source. Not a live export or an importable package.",
    }


def main():
    rendered = json.dumps(build_definition(), ensure_ascii=False, indent=2) + "\n"
    if "--check" in sys.argv:
        if not DESTINATION.exists() or DESTINATION.read_text(encoding="utf-8") != rendered:
            raise SystemExit("Attachment flow source is stale; run scripts/build_attachment_flow.py.")
    elif len(sys.argv) == 1:
        DESTINATION.parent.mkdir(parents=True, exist_ok=True)
        DESTINATION.write_text(rendered, encoding="utf-8")
    else:
        raise SystemExit("Usage: python3 scripts/build_attachment_flow.py [--check]")


if __name__ == "__main__":
    main()
