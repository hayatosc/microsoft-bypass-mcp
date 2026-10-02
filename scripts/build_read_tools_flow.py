"""Augment the canonical flow with reviewed Codex-authored read-tool cases.

This module never contacts Power Automate or writes a separate flow. The existing
build_attachment_flow generator calls it after constructing its preserved
attachment cases from the immutable fixture.

The generated source is review/update material for the existing flow. OneDrive
native connection names/placeholders must be bound manually during an authorized
Power Automate update; this script intentionally does not create connectors.
"""

from __future__ import annotations

from copy import deepcopy
import json
from urllib.parse import quote


GRAPH_ROOT = "https://graph.microsoft.com/v1.0/me/"
OUTLOOK_CONNECTION = {
    "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
    "connectionName": "shared_office365",
    "operationId": "HttpRequest",
}
ONEDRIVE_API = "/providers/Microsoft.PowerApps/apis/shared_onedriveforbusiness"
ONEDRIVE_CONNECTION = "shared_onedriveforbusiness"
# Official connector docs identify FindFiles parameters query, id, findMode and
# maxFileCount, but the live designer must verify the exact machine value used
# by that tenant before import. Keeping the placeholder visible prevents an
# invented value from silently entering a production flow.
ONEDRIVE_FIND_MODE_PLACEHOLDER = "@outputs('OneDriveSearchMode')"
# Power Automate's OneDrive root can require tenant/designer-specific binding.
# The ListRootFolder action is used for omitted folderId; this placeholder only
# documents that no arbitrary provider URL/root ID is accepted from callers.
ONEDRIVE_ROOT_BINDING_PLACEHOLDER = "@outputs('OneDriveSearchRootId')"

MAX_ATTACHMENT_BYTES = 4 * 1024 * 1024
MAX_BASE64 = ((MAX_ATTACHMENT_BYTES + 2) // 3) * 4
BASE64_CHUNK = 8192

OUTLOOK_TOP_SCHEMA = {"type": "integer", "minimum": 1, "maximum": 50}
ONEDRIVE_TOP_SCHEMA = {"type": "integer", "minimum": 1, "maximum": 100}
ONEDRIVE_FOLDER_WINDOW_SIZE = 1000
ONEDRIVE_FOLDER_TOP_SCHEMA = {"type": "integer", "minimum": 1, "maximum": ONEDRIVE_FOLDER_WINDOW_SIZE}
SKIP_SCHEMA = {"type": "integer", "minimum": 0, "maximum": 10000}
ID_SCHEMA = {"type": "string", "minLength": 1, "maxLength": 2048}
QUERY_SCHEMA = {"type": "string", "minLength": 1, "maxLength": 512}
ISO_SCHEMA = {"type": "string", "minLength": 20, "maxLength": 28}

SUMMARY_SELECT = "id,subject,from,receivedDateTime,sentDateTime,parentFolderId,conversationId,hasAttachments,importance,isRead,bodyPreview"
DETAIL_SELECT = SUMMARY_SELECT + ",toRecipients,ccRecipients,body"
FOLDER_SELECT = "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount"
ATTACHMENT_METADATA_FIELDS = ["@odata.type", "id", "name", "contentType", "size", "isInline"]
ONEDRIVE_METADATA_FIELDS = ["Id", "Name", "Size", "MediaType", "IsFolder", "LastModified", "ETag", "Path", "NameNoExt", "DisplayName", "FileLocator"]
ONEDRIVE_RETURN_FIELDS = ["Id", "Name", "Size", "MediaType", "IsFolder", "LastModified", "ETag"]

MAIL_OPERATIONS = ("list_messages", "get_message", "search_messages")
OUTLOOK_OPERATIONS = (
    "list_messages",
    "search_messages",
    "get_message",
    "list_mail_folders",
    "get_conversation",
    "list_attachments",
    "get_attachment",
)
ONEDRIVE_OPERATIONS = (
    "onedrive_search_files",
    "onedrive_list_folder",
    "onedrive_get_metadata",
    "onedrive_get_content",
)
# Worker-facing tools are 13; flow operations are 11 because inspect/read tools
# share get_attachment or get_content transport operations.
FLOW_OPERATIONS = OUTLOOK_OPERATIONS + ONEDRIVE_OPERATIONS

ARGUMENTS = {
    "list_messages": {
        "top": OUTLOOK_TOP_SCHEMA,
        "skip": SKIP_SCHEMA,
        "mailbox": {"type": "string", "enum": ["inbox", "sent", "all"]},
        "folderId": ID_SCHEMA,
        "filters": {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "isRead": {"type": "boolean"},
                "hasAttachments": {"type": "boolean"},
                "receivedAfter": ISO_SCHEMA,
                "receivedBefore": ISO_SCHEMA,
            },
        },
    },
    "search_messages": {
        "query": QUERY_SCHEMA,
        "top": OUTLOOK_TOP_SCHEMA,
        "mailbox": {"type": "string", "enum": ["inbox", "sent", "all"]},
        "folderId": ID_SCHEMA,
    },
    "get_message": {"messageId": ID_SCHEMA},
    "list_mail_folders": {"top": OUTLOOK_TOP_SCHEMA},
    "get_conversation": {"conversationId": ID_SCHEMA, "top": OUTLOOK_TOP_SCHEMA, "skip": SKIP_SCHEMA},
    "list_attachments": {"messageId": ID_SCHEMA, "top": OUTLOOK_TOP_SCHEMA, "skip": SKIP_SCHEMA},
    "get_attachment": {"messageId": ID_SCHEMA, "attachmentId": ID_SCHEMA},
    "onedrive_search_files": {"query": QUERY_SCHEMA, "top": ONEDRIVE_TOP_SCHEMA},
    "onedrive_list_folder": {"folderId": ID_SCHEMA, "top": ONEDRIVE_FOLDER_TOP_SCHEMA},
    "onedrive_get_metadata": {"fileId": ID_SCHEMA},
    "onedrive_get_content": {"fileId": ID_SCHEMA},
}
REQUIRED_ARGS = {
    "list_messages": ["top"],
    "search_messages": ["query", "top"],
    "get_message": ["messageId"],
    "list_mail_folders": ["top"],
    "get_conversation": ["conversationId", "top"],
    "list_attachments": ["messageId", "top", "skip"],
    "get_attachment": ["messageId", "attachmentId"],
    "onedrive_search_files": ["query", "top"],
    "onedrive_list_folder": ["top"],
    "onedrive_get_metadata": ["fileId"],
    "onedrive_get_content": ["fileId"],
}

FORBIDDEN_ID_CODEPOINTS = (
    *range(0x20),
    0x20,
    0x7F,
    0x85,
    0xA0,
    0x1680,
    *range(0x2000, 0x200B),
    0x2028,
    0x2029,
    0x202F,
    0x205F,
    0x3000,
    0xFEFF,
)


def object_schema(properties: dict, *, required: list[str] | None = None, strict: bool = False) -> dict:
    schema = {"type": "object", "properties": deepcopy(properties), "required": list(required if required is not None else properties)}
    if strict:
        schema["additionalProperties"] = False
    return schema


def literal_key(value: str) -> str:
    return "@" + value if value.startswith("@") else value


def schema_literals(value):
    if isinstance(value, dict):
        return {literal_key(k): schema_literals(v) for k, v in value.items()}
    if isinstance(value, list):
        return [schema_literals(v) for v in value]
    return "@" + value if isinstance(value, str) and value.startswith("@") else value


def secure(action: dict) -> dict:
    props = ["inputs"] if action["type"] in {"ParseJson", "Response", "Compose"} else ["inputs", "outputs"]
    action["runtimeConfiguration"] = {"secureData": {"properties": props}}
    return action


def after(name: str, *statuses: str) -> dict:
    return {name: list(statuses or ("Succeeded",))}


def response(status: int, body, run_after: dict | None = None) -> dict:
    return secure({
        "type": "Response",
        "kind": "Http",
        "runAfter": run_after or {},
        "inputs": {"statusCode": status, "headers": {"Content-Type": "application/json", "Cache-Control": "no-store"}, "body": body},
    })


def error(status: int, code: str, message: str, run_after: dict | None = None) -> dict:
    return response(status, {"ok": False, "error": {"code": code, "message": message}}, run_after)


def success(data, run_after: dict | None = None) -> dict:
    return response(200, {"ok": True, "requestId": "@triggerBody()?['requestId']", "operation": "@triggerBody()?['operation']", "data": data}, run_after)


def parse(name: str, content, schema: dict, run_after: dict | None = None) -> tuple[str, dict]:
    return name, secure({"type": "ParseJson", "runAfter": run_after or {}, "inputs": {"content": content, "schema": schema_literals(schema)}})


def graph(name: str, uri: str, run_after: dict | None = None, *, body_text: bool = False) -> tuple[str, dict]:
    params = {"Uri": uri, "Method": "GET", "CustomHeader1": 'Prefer: IdType="ImmutableId"', "ContentType": "application/json"}
    if body_text:
        params["CustomHeader2"] = 'Prefer: outlook.body-content-type="text"'
    return name, secure({
        "type": "OpenApiConnection",
        "runAfter": run_after or {},
        "inputs": {"parameters": params, "host": OUTLOOK_CONNECTION, "authentication": "@parameters('$authentication')", "retryPolicy": {"type": "none"}},
    })


def onedrive(name: str, operation_id: str, parameters: dict, run_after: dict | None = None) -> tuple[str, dict]:
    return name, secure({
        "type": "OpenApiConnection",
        "runAfter": run_after or {},
        "inputs": {
            "parameters": parameters,
            "host": {"apiId": ONEDRIVE_API, "connectionName": ONEDRIVE_CONNECTION, "operationId": operation_id},
            "authentication": "@parameters('$authentication')",
            "retryPolicy": {"type": "none"},
        },
    })


def failure(actions: dict, dependency: str, *, status: int = 502, code: str = "UPSTREAM_ERROR", message: str = "The upstream read could not be completed.") -> None:
    actions[f"Error_{dependency}"] = error(status, code, message, after(dependency, "Failed", "TimedOut"))


def condition(expression, actions: dict, otherwise: dict, run_after: dict | None = None) -> dict:
    return {"type": "If", "expression": expression, "actions": actions, "else": {"actions": otherwise}, "runAfter": run_after or {}}


def field(action: str, name: str) -> str:
    return f"@body('{action}')?['{name}']"


def uuid_v4_guard(value: str) -> str:
    stripped = f"toLower(replace({value}, '-', ''))"
    remainder = stripped
    for char in "0123456789abcdef":
        remainder = f"replace({remainder}, '{char}', '')"
    checks = [f"equals(length({stripped}), 32)", f"equals({remainder}, '')"]
    checks += [f"equals(substring({value}, {offset}, 1), '-')" for offset in (8, 13, 18, 23)]
    checks += [f"equals(substring({value}, 14, 1), '4')", f"contains('89ab', toLower(substring({value}, 19, 1)))"]
    return "and(" + ", ".join(checks) + ")"


def id_guard(value: str) -> str:
    encoded = f"uriComponent({value})"
    cleaned = encoded
    forbidden = ("%0", "%1", *(quote(chr(cp), safe="") for cp in FORBIDDEN_ID_CODEPOINTS if cp >= 0x20))
    for encoded_character in forbidden:
        cleaned = f"replace({cleaned}, '{encoded_character}', '')"
    return f"and(not(equals({value}, '.')), not(equals({value}, '..')), equals(length({encoded}), length({cleaned})))"


def query_guard(value: str) -> str:
    trimmed = f"trim({value})"
    return f"and(not(empty({trimmed})), lessOrEquals(length({trimmed}), 512))"


def iso_guard(value: str) -> str:
    # Parse with Microsoft's timestamp function as well as restricting the
    # alphabet. Coalescing makes absent optional bounds safe even when function
    # arguments are evaluated eagerly. The Worker contract accepts UTC Z only.
    text = f"coalesce({value}, '2000-01-01T00:00:00Z')"
    remainder = text
    for char in "0123456789-:.TZ":
        remainder = f"replace({remainder}, '{char}', '')"
    return f"and(equals({remainder}, ''), equals(substring({text}, 10, 1), 'T'), endsWith({text}, 'Z'), greaterOrEquals(ticks({text}), 0))"


def base64_alphabet_guard(value: str) -> str:
    escaped = value
    for char, encoded in (("+", "%2B"), ("/", "%2F"), ("=", "%3D")):
        escaped = f"replace({escaped}, '{char}', '{encoded}')"
    checks = [f"equals(uriComponent({value}), {escaped})"]
    checks += [f"not(contains({value}, '{char.replace(chr(39), chr(39) * 2)}'))" for char in "-_.!~*'()"]
    return "and(" + ", ".join(checks) + ")"


def request_args(operation: str) -> tuple[str, str, dict]:
    actions: dict = {}
    envelope = f"Validate_{operation}_request"
    actions[envelope] = parse(envelope, "@triggerBody()", object_schema({"operation": {"type": "string", "enum": [operation]}, "requestId": {"type": "string", "minLength": 36, "maxLength": 36}, "args": {"type": "object"}}, strict=True))[1]
    failure(actions, envelope, status=400, code="INVALID_REQUEST", message="The request envelope is invalid.")
    request_format = f"Validate_{operation}_request_format"
    actions[request_format] = parse(request_format, "@" + uuid_v4_guard(f"body('{envelope}')?['requestId']"), {"type": "boolean", "enum": [True]}, after(envelope))[1]
    failure(actions, request_format, status=400, code="INVALID_REQUEST", message="The request envelope is invalid.")
    args_name = f"Validate_{operation}_args"
    actions[args_name] = parse(args_name, f"@body('{envelope}')?['args']", object_schema(ARGUMENTS[operation], required=REQUIRED_ARGS[operation], strict=True), after(request_format))[1]
    failure(actions, args_name, status=400, code="INVALID_ARGUMENTS", message="The operation arguments are invalid.")
    checks: list[str] = []
    for key in ARGUMENTS[operation]:
        if key.endswith("Id") or key == "conversationId":
            checks.append(id_guard(f"coalesce(body('{args_name}')?['{key}'], 'unused')"))
        if key == "query":
            checks.append(query_guard(f"body('{args_name}')?['{key}']"))
    if operation == "list_messages":
        checks += [iso_guard(f"body('{args_name}')?['filters']?['receivedAfter']"), iso_guard(f"body('{args_name}')?['filters']?['receivedBefore']")]
        after_value = f"body('{args_name}')?['filters']?['receivedAfter']"
        before_value = f"body('{args_name}')?['filters']?['receivedBefore']"
        checks.append(f"or(equals({after_value}, null), equals({before_value}, null), less(ticks(coalesce({after_value}, '0001-01-01T00:00:00Z')), ticks(coalesce({before_value}, '9999-12-31T23:59:59Z'))))")
    guard = "true" if not checks else checks[0] if len(checks) == 1 else "and(" + ", ".join(checks) + ")"
    ids = f"Validate_{operation}_controls"
    actions[ids] = parse(ids, "@" + guard, {"type": "boolean", "enum": [True]}, after(args_name))[1]
    failure(actions, ids, status=400, code="INVALID_ARGUMENTS", message="The operation arguments are invalid.")
    return args_name, ids, actions


def odata_quote(value: str) -> str:
    return f"replace({value}, '''', '''''')"


def metadata_schema() -> dict:
    return object_schema({"@odata.type": {"type": "string", "minLength": 1, "maxLength": 128}, "id": ID_SCHEMA, "name": {"type": "string", "maxLength": 512}, "contentType": {"type": "string", "maxLength": 256}, "size": {"type": "integer", "minimum": 0}, "isInline": {"type": "boolean"}})


def onedrive_metadata_schema() -> dict:
    props = {"Id": ID_SCHEMA, "Name": {"type": "string", "minLength": 1, "maxLength": 512}, "Size": {"type": "integer", "minimum": 0}, "MediaType": {"type": "string", "maxLength": 256}, "IsFolder": {"type": "boolean"}}
    for key in ("LastModified", "ETag", "Path", "NameNoExt", "DisplayName", "FileLocator"):
        props[key] = {"type": ["string", "null"], "maxLength": 2048}
    return object_schema(props, required=["Id", "Name", "Size", "MediaType", "IsFolder"])


def projection(action: str, fields: list[str]) -> dict:
    return {literal_key(k): field(action, k) for k in fields}


def select_projection(source: str, fields: list[str]) -> dict:
    return {literal_key(k): f"@item()?['{k}']" for k in fields}


def route_expr(validate: str) -> str:
    return (
        "@if(not(empty(body('" + validate + "')?['folderId'])), "
        "concat('mailFolders/', uriComponent(coalesce(body('" + validate + "')?['folderId'], 'unused')), '/messages'), "
        "if(equals(coalesce(body('" + validate + "')?['mailbox'], 'inbox'), 'sent'), 'mailFolders/sentitems/messages', "
        "if(equals(coalesce(body('" + validate + "')?['mailbox'], 'inbox'), 'all'), 'messages', 'mailFolders/inbox/messages')))"
    )


def filter_expr(validate: str) -> str:
    return (
        "@concat("
        "if(equals(body('" + validate + "')?['filters']?['isRead'], null), '', concat('isRead eq ', toLower(string(body('" + validate + "')?['filters']?['isRead'])))),"
        "if(equals(body('" + validate + "')?['filters']?['hasAttachments'], null), '', concat(if(equals(body('" + validate + "')?['filters']?['isRead'], null), '', ' and '), 'hasAttachments eq ', toLower(string(body('" + validate + "')?['filters']?['hasAttachments'])))),"
        "if(empty(body('" + validate + "')?['filters']?['receivedAfter']), '', concat(if(or(not(equals(body('" + validate + "')?['filters']?['isRead'], null)), not(equals(body('" + validate + "')?['filters']?['hasAttachments'], null))), ' and ', ''), 'receivedDateTime ge ', body('" + validate + "')?['filters']?['receivedAfter'])),"
        "if(empty(body('" + validate + "')?['filters']?['receivedBefore']), '', concat(if(or(not(equals(body('" + validate + "')?['filters']?['isRead'], null)), not(equals(body('" + validate + "')?['filters']?['hasAttachments'], null)), not(empty(body('" + validate + "')?['filters']?['receivedAfter']))), ' and ', ''), 'receivedDateTime lt ', body('" + validate + "')?['filters']?['receivedBefore']))"
        ")"
    )


def outlook_list_case() -> dict:
    validate, controls, actions = request_args("list_messages")
    actions["Compose_list_messages_route"] = secure({"type": "Compose", "runAfter": after(controls), "inputs": route_expr(validate)})
    actions["Compose_list_messages_filter"] = secure({"type": "Compose", "runAfter": after("Compose_list_messages_route"), "inputs": filter_expr(validate)})
    uri = "@concat('" + GRAPH_ROOT + "', outputs('Compose_list_messages_route'), '?$top=', string(body('" + validate + "')?['top']), '&$skip=', string(coalesce(body('" + validate + "')?['skip'], 0)), '&$select=" + SUMMARY_SELECT + "', if(empty(outputs('Compose_list_messages_filter')), '&$orderby=receivedDateTime%20desc', concat('&$filter=', uriComponent(outputs('Compose_list_messages_filter')))))"
    actions["Graph_list_messages"] = graph("Graph_list_messages", uri, after("Compose_list_messages_filter"))[1]
    failure(actions, "Graph_list_messages")
    actions["Respond_list_messages"] = success("@body('Graph_list_messages')", after("Graph_list_messages"))
    return {"case": "list_messages", "actions": actions}


def outlook_search_case() -> dict:
    validate, controls, actions = request_args("search_messages")
    actions["Compose_search_messages_route"] = secure({"type": "Compose", "runAfter": after(controls), "inputs": route_expr(validate)})
    uri = "@concat('" + GRAPH_ROOT + "', outputs('Compose_search_messages_route'), '?$search=', uriComponent(concat('\"', trim(body('" + validate + "')?['query']), '\"')), '&$top=', string(body('" + validate + "')?['top']), '&$select=" + SUMMARY_SELECT + "')"
    actions["Graph_search_messages"] = graph("Graph_search_messages", uri, after("Compose_search_messages_route"))[1]
    failure(actions, "Graph_search_messages")
    actions["Respond_search_messages"] = success("@body('Graph_search_messages')", after("Graph_search_messages"))
    return {"case": "search_messages", "actions": actions}


def outlook_get_case() -> dict:
    validate, controls, actions = request_args("get_message")
    uri = f"@concat('{GRAPH_ROOT}messages/', uriComponent(body('{validate}')?['messageId']), '?$select={DETAIL_SELECT}')"
    actions["Graph_get_message"] = graph("Graph_get_message", uri, after(controls), body_text=True)[1]
    failure(actions, "Graph_get_message")
    actions["Respond_get_message"] = success("@body('Graph_get_message')", after("Graph_get_message"))
    return {"case": "get_message", "actions": actions}


def folders_case() -> dict:
    validate, controls, actions = request_args("list_mail_folders")
    uri = f"@concat('{GRAPH_ROOT}mailFolders?$top=', string(body('{validate}')?['top']), '&$select={FOLDER_SELECT}')"
    actions["Graph_list_mail_folders"] = graph("Graph_list_mail_folders", uri, after(controls))[1]
    failure(actions, "Graph_list_mail_folders")
    actions["Respond_list_mail_folders"] = success("@body('Graph_list_mail_folders')", after("Graph_list_mail_folders"))
    return {"case": "list_mail_folders", "actions": actions}


def conversation_case() -> dict:
    validate, controls, actions = request_args("get_conversation")
    # Do not include $orderby with conversationId filter; Graph can reject this
    # as InefficientFilter. The Worker may sort each returned page after receipt.
    uri = "@concat('" + GRAPH_ROOT + "messages?$top=', string(body('" + validate + "')?['top']), '&$skip=', string(coalesce(body('" + validate + "')?['skip'], 0)), '&$select=" + DETAIL_SELECT + "&$filter=', uriComponent(concat('conversationId eq ''', replace(body('" + validate + "')?['conversationId'], '''', ''''''), '''')))"
    actions["Graph_get_conversation"] = graph("Graph_get_conversation", uri, after(controls), body_text=True)[1]
    failure(actions, "Graph_get_conversation")
    actions["Respond_get_conversation"] = success("@body('Graph_get_conversation')", after("Graph_get_conversation"))
    return {"case": "get_conversation", "actions": actions}


def onedrive_project_list(actions: dict, source: str, top: str, name: str, run_after: str, *, max_items: int = 100) -> None:
    # Each native action has a documented shape; do not guess array vs page
    # using empty(value), which misclassifies an empty folder page.
    actions[f"Select_{name}"] = secure({"type": "Select", "runAfter": after(run_after), "inputs": {"from": f"@take({source}, {top})", "select": select_projection(source, ONEDRIVE_RETURN_FIELDS)}})
    failure(actions, f"Select_{name}", code="INVALID_ONEDRIVE_METADATA", message="The upstream OneDrive metadata is invalid.")
    actions[f"Validate_{name}"] = parse(f"Validate_{name}", f"@body('Select_{name}')", {"type": "array", "maxItems": max_items, "items": onedrive_metadata_schema()}, after(f"Select_{name}"))[1]
    failure(actions, f"Validate_{name}", code="INVALID_ONEDRIVE_METADATA", message="The upstream OneDrive metadata is invalid.")


def onedrive_search_case() -> dict:
    validate, controls, actions = request_args("onedrive_search_files")
    native = {"OneDrive_find_files": onedrive("OneDrive_find_files", "FindFiles", {"query": f"@trim(body('{validate}')?['query'])", "id": ONEDRIVE_ROOT_BINDING_PLACEHOLDER, "findMode": ONEDRIVE_FIND_MODE_PLACEHOLDER, "maxFileCount": f"@body('{validate}')?['top']"})[1]}
    failure(native, "OneDrive_find_files")
    top = f"body('{validate}')?['top']"
    source = "body('OneDrive_find_files')"
    onedrive_project_list(native, source, top, "onedrive_search_files", "OneDrive_find_files")
    native["Respond_onedrive_search_files"] = success({"value": "@body('Validate_onedrive_search_files')", "truncated": f"@greaterOrEquals(length({source}), {top})"}, after("Validate_onedrive_search_files"))
    # Named Compose settings are editable in the non-solution cloud-flow designer.
    # Fixed empty defaults fail closed; never derive configuration from caller args.
    actions["OneDriveSearchRootId"] = secure({"type": "Compose", "inputs": "", "runAfter": after(controls)})
    actions["OneDriveSearchMode"] = secure({"type": "Compose", "inputs": "", "runAfter": after("OneDriveSearchRootId")})
    configured = "@and(not(empty(outputs('OneDriveSearchRootId'))), not(empty(outputs('OneDriveSearchMode'))))"
    actions["Check_onedrive_search_binding"] = condition({"equals": [configured, True]}, native, {"Reject_onedrive_search_binding": error(503, "ONEDRIVE_SEARCH_NOT_CONFIGURED", "The OneDrive search binding must be configured before use.")}, after("OneDriveSearchMode"))
    return {"case": "onedrive_search_files", "actions": actions}


def onedrive_list_folder_case() -> dict:
    validate, controls, actions = request_args("onedrive_list_folder")
    top = f"body('{validate}')?['top']"
    list_root = {"OneDrive_list_root": onedrive("OneDrive_list_root", "ListRootFolder", {})[1]}
    failure(list_root, "OneDrive_list_root")
    source_root = "body('OneDrive_list_root')"
    cap = ONEDRIVE_FOLDER_WINDOW_SIZE
    # ListRootFolder is an array operation, with no supported paginationPolicy.
    onedrive_project_list(list_root, f"take({source_root}, {cap})", top, "onedrive_root_folder", "OneDrive_list_root", max_items=cap)
    list_root["Respond_onedrive_list_root"] = success({"value": "@body('Validate_onedrive_root_folder')", "truncated": f"@or(greaterOrEquals(length({source_root}), {cap}), greater(length({source_root}), {top}))"}, after("Validate_onedrive_root_folder"))
    list_child = {"OneDrive_list_folder_v2": onedrive("OneDrive_list_folder_v2", "ListFolderV2", {"id": f"@body('{validate}')?['folderId']"})[1]}
    # Native ListFolderV2 auto-aggregates its actual continuation pages. This
    # threshold is a minimum, not a maximum: its final page may overshoot. Bound
    # the aggregated raw array before projection/transport and conservatively
    # report incompleteness at the boundary even if nextLink is absent. Never
    # synthesize skip tokens or follow native URLs with a different connector.
    # Merge rather than replace the secure input/output history configuration.
    list_child["OneDrive_list_folder_v2"]["runtimeConfiguration"]["paginationPolicy"] = {"minimumItemCount": cap}
    failure(list_child, "OneDrive_list_folder_v2")
    source_child = "body('OneDrive_list_folder_v2')?['value']"
    onedrive_project_list(list_child, f"take({source_child}, {cap})", top, "onedrive_list_folder", "OneDrive_list_folder_v2", max_items=cap)
    list_child["Respond_onedrive_list_folder"] = success({"value": "@body('Validate_onedrive_list_folder')", "truncated": f"@or(greaterOrEquals(length({source_child}), {cap}), greater(length({source_child}), {top}), not(empty(body('OneDrive_list_folder_v2')?['nextLink'])))"}, after("Validate_onedrive_list_folder"))
    actions["Route_onedrive_list_folder"] = condition({"equals": [f"@empty(body('{validate}')?['folderId'])", True]}, list_root, list_child, after(controls))
    return {"case": "onedrive_list_folder", "actions": actions}


def onedrive_get_metadata_case() -> dict:
    validate, controls, actions = request_args("onedrive_get_metadata")
    actions["OneDrive_get_metadata"] = onedrive("OneDrive_get_metadata", "GetFileMetadata", {"id": f"@body('{validate}')?['fileId']"}, after(controls))[1]
    failure(actions, "OneDrive_get_metadata")
    actions["Validate_onedrive_metadata"] = parse("Validate_onedrive_metadata", "@body('OneDrive_get_metadata')", onedrive_metadata_schema(), after("OneDrive_get_metadata"))[1]
    failure(actions, "Validate_onedrive_metadata", code="INVALID_ONEDRIVE_METADATA", message="The upstream OneDrive metadata is invalid.")
    actions["Check_onedrive_metadata_identity"] = condition({"equals": [field("Validate_onedrive_metadata", "Id"), field(validate, "fileId")]}, {"Respond_onedrive_metadata": success(projection("Validate_onedrive_metadata", ONEDRIVE_RETURN_FIELDS))}, {"Reject_onedrive_metadata": error(502, "INVALID_ONEDRIVE_METADATA", "The upstream OneDrive metadata is invalid.")}, after("Validate_onedrive_metadata"))
    return {"case": "onedrive_get_metadata", "actions": actions}


def content_type_gate_expr(media_type: str = "body('Validate_onedrive_content_metadata')?['MediaType']") -> dict:
    name = "toLower(body('Validate_onedrive_content_metadata')?['Name'])"
    media = f"toLower(trim(first(split({media_type}, ';'))))"
    generic = f"or(equals({media}, ''), equals({media}, 'application/octet-stream'))"
    pairs = (('pdf', 'application/pdf'),
             ('docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
             ('xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))
    clauses = [f"and(endsWith({name}, '.{extension}'), or({generic}, equals({media}, '{mime}')))" for extension, mime in pairs]
    return {"equals": ["@or(" + ", ".join(clauses) + ")", True]}


def onedrive_get_content_case() -> dict:
    validate, controls, actions = request_args("onedrive_get_content")
    actions["OneDrive_content_metadata"] = onedrive("OneDrive_content_metadata", "GetFileMetadata", {"id": f"@body('{validate}')?['fileId']"}, after(controls))[1]
    failure(actions, "OneDrive_content_metadata")
    actions["Validate_onedrive_content_metadata"] = parse("Validate_onedrive_content_metadata", "@body('OneDrive_content_metadata')", onedrive_metadata_schema(), after("OneDrive_content_metadata"))[1]
    failure(actions, "Validate_onedrive_content_metadata", code="INVALID_ONEDRIVE_METADATA", message="The upstream OneDrive metadata is invalid.")
    content_actions: dict = {}
    content_actions["OneDrive_get_content"] = onedrive("OneDrive_get_content", "GetFileContent", {"id": f"@body('{validate}')?['fileId']", "inferContentType": True})[1]
    failure(content_actions, "OneDrive_get_content")
    binary_schema = object_schema({"contentType": {"type": "string", "maxLength": 256}, "contentBytes": {"type": "string", "minLength": 4, "maxLength": MAX_BASE64}}, strict=True)
    # Use ordinary JSON keys: recreating the reserved $content-type/$content
    # envelope can preserve semantic binary content rather than a JSON object.
    # Copy the native base64 string directly; never encode or decode it here.
    binary_fields = {
        "contentType": "@body('OneDrive_get_content')?['$content-type']",
        "contentBytes": "@body('OneDrive_get_content')?['$content']",
    }
    content_actions["Validate_onedrive_binary"] = parse("Validate_onedrive_binary", binary_fields, binary_schema, after("OneDrive_get_content"))[1]
    failure(content_actions, "Validate_onedrive_binary", code="INVALID_ONEDRIVE_CONTENT", message="The upstream OneDrive content is invalid or exceeds the limit.")
    content_actions["Select_onedrive_content_alphabet"] = secure({"type": "Select", "runAfter": after("Validate_onedrive_binary"), "inputs": {"from": f"@chunk(body('Validate_onedrive_binary')?['contentBytes'], {BASE64_CHUNK})", "select": "@" + base64_alphabet_guard("item()")}})
    failure(content_actions, "Select_onedrive_content_alphabet", code="INVALID_ONEDRIVE_CONTENT", message="The upstream OneDrive content is invalid or exceeds the limit.")
    content_actions["Validate_onedrive_content_alphabet"] = parse("Validate_onedrive_content_alphabet", "@body('Select_onedrive_content_alphabet')", {"type": "array", "maxItems": (MAX_BASE64 + BASE64_CHUNK - 1) // BASE64_CHUNK, "items": {"type": "boolean", "enum": [True]}}, after("Select_onedrive_content_alphabet"))[1]
    failure(content_actions, "Validate_onedrive_content_alphabet", code="INVALID_ONEDRIVE_CONTENT", message="The upstream OneDrive content is invalid or exceeds the limit.")
    b64 = "body('Validate_onedrive_binary')?['contentBytes']"
    decoded = f"sub(mul(div(length({b64}), 4), 3), if(endsWith({b64}, '=='), 2, if(endsWith({b64}, '='), 1, 0)))"
    content_actions["Check_onedrive_content_size"] = condition({"and": [{"lessOrEquals": ["@" + decoded, MAX_ATTACHMENT_BYTES]}, {"equals": ["@" + decoded, field("Validate_onedrive_content_metadata", "Size")]}, content_type_gate_expr("body('Validate_onedrive_binary')?['contentType']"), {"equals": [f"@mod(length({b64}), 4)", 0]}, {"equals": [f"@or(equals(indexOf({b64}, '='), -1), equals(indexOf({b64}, '='), sub(length({b64}), if(endsWith({b64}, '=='), 2, 1))))", True]}]}, {"Respond_onedrive_content": success({"metadata": projection("Validate_onedrive_content_metadata", ONEDRIVE_RETURN_FIELDS), "contentBytes": "@body('Validate_onedrive_binary')?['contentBytes']"})}, {"Reject_onedrive_content": error(502, "INVALID_ONEDRIVE_CONTENT", "The upstream OneDrive content is invalid or exceeds the limit.")}, after("Validate_onedrive_content_alphabet"))
    type_gate = condition(content_type_gate_expr(), content_actions, {"Reject_onedrive_content_type": error(415, "UNSUPPORTED_ONEDRIVE_FILE_TYPE", "Only PDF, DOCX, and XLSX files can be read.")})
    size_gate = condition({"and": [{"equals": [field("Validate_onedrive_content_metadata", "IsFolder"), False]}, {"lessOrEquals": [field("Validate_onedrive_content_metadata", "Size"), MAX_ATTACHMENT_BYTES]}, {"equals": ["@greater(body('Validate_onedrive_content_metadata')?['Size'], 0)", True]}]}, {"Check_onedrive_content_type": type_gate}, {"Reject_onedrive_content_size": error(413, "ONEDRIVE_FILE_TOO_LARGE", "The OneDrive file is empty, a folder, or exceeds the 4 MiB limit.")})
    actions["Check_onedrive_content_identity"] = condition({"equals": [field("Validate_onedrive_content_metadata", "Id"), field(validate, "fileId")]}, {"Check_onedrive_metadata_size": size_gate}, {"Reject_onedrive_content_metadata": error(502, "INVALID_ONEDRIVE_METADATA", "The upstream OneDrive metadata is invalid.")}, after("Validate_onedrive_content_metadata"))
    return {"case": "onedrive_get_content", "actions": actions}



def augment_read_tools(definition: dict) -> dict:
    """Apply only revised mail/new read cases; existing attachment cases stay intact."""
    # Request-trigger outputs contain caller args/headers. Hide them from run
    # history without changing trigger identity, authentication or key guard.
    definition["triggers"]["manual"]["runtimeConfiguration"] = {
        "secureData": {"properties": ["outputs"]}
    }
    schema = definition["triggers"]["manual"]["inputs"]["schema"]
    schema["properties"]["operation"]["enum"] = list(FLOW_OPERATIONS)
    # Per-operation validators provide tighter bounds, including OneDrive top100.
    schema["properties"]["args"] = {"type": "object"}
    cases = definition["actions"]["スイッチ"]["cases"]
    for build in (outlook_list_case, outlook_get_case, outlook_search_case,
                  folders_case, conversation_case, onedrive_search_case,
                  onedrive_list_folder_case, onedrive_get_metadata_case,
                  onedrive_get_content_case):
        case = build()
        cases[case["case"]] = case
    return definition
