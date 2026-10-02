"""Draft-only augmentation of the canonical flow; no network access or output file.

Parse JSON has no regex support here. Guards use documented WDL string/collection
functions, and base64 validation works on bounded chunks plus canonical padding
checks, never a decoded binary value. All writes are fixed Office365 requests.
"""

from build_read_tools_flow import (
    BASE64_CHUNK, GRAPH_ROOT, ID_SCHEMA, after, base64_alphabet_guard, condition,
    error, failure, field, graph, id_guard, object_schema, parse, secure, success,
    uuid_v4_guard,
)

DRAFT_OPERATIONS = ("create_draft", "create_reply_draft", "add_draft_attachment")
MAX_BYTES = 2 * 1024 * 1024
MAX_BASE64 = 4 * ((MAX_BYTES + 2) // 3)
RECIPIENT_SCHEMA = {"type": "array", "maxItems": 50, "items": {"type": "string", "maxLength": 254}}
DRAFT_ARGUMENTS = {
    "create_draft": {
        "to": RECIPIENT_SCHEMA, "cc": RECIPIENT_SCHEMA, "bcc": RECIPIENT_SCHEMA,
        "subject": {"type": "string", "maxLength": 512},
        "body": {"type": "string", "maxLength": 20000},
    },
    "create_reply_draft": {
        "messageId": ID_SCHEMA, "body": {"type": "string", "maxLength": 20000},
    },
    "add_draft_attachment": {
        "draftId": ID_SCHEMA,
        "name": {"type": "string", "minLength": 1, "maxLength": 255},
        "contentType": {"type": "string", "minLength": 1, "maxLength": 127},
        "contentBytes": {"type": "string", "minLength": 4, "maxLength": MAX_BASE64},
    },
}
DRAFT_REQUIRED_ARGS = {
    "create_draft": ["to", "subject", "body"],
    "create_reply_draft": ["messageId", "body"],
    "add_draft_attachment": ["draftId", "name", "contentType", "contentBytes"],
}
AMBIGUOUS_MESSAGE = "The draft write outcome may be ambiguous; inspect Drafts before retrying."


def invalid_input(actions, name):
    failure(actions, name, status=400, code="INVALID_ARGUMENTS", message="The draft operation arguments are invalid.")


def ambiguous(actions, name):
    failure(actions, name, code="DRAFT_WRITE_AMBIGUOUS", message=AMBIGUOUS_MESSAGE)


def add_parse(actions, name, content, schema, dependency, *, on_failure=invalid_input):
    actions[name] = parse(name, content, schema, after(dependency))[1]
    on_failure(actions, name)
    return name


def add_guard(actions, name, expression, dependency, *, on_failure=invalid_input):
    return add_parse(actions, name, "@" + expression, {"type": "boolean", "enum": [True]}, dependency, on_failure=on_failure)


def compose(actions, name, value, dependency, *, on_failure=invalid_input):
    actions[name] = secure({"type": "Compose", "inputs": value, "runAfter": after(dependency) if dependency else {}})
    on_failure(actions, name)
    return name


def request_args(operation):
    envelope = f"Validate_{operation}_request"
    actions = {envelope: parse(envelope, "@triggerBody()", object_schema({
        "operation": {"type": "string", "enum": [operation]},
        "requestId": {"type": "string", "minLength": 36, "maxLength": 36},
        "args": {"type": "object"},
    }, strict=True))[1]}
    failure(actions, envelope, status=400, code="INVALID_REQUEST", message="The request envelope is invalid.")
    request_format = f"Validate_{operation}_request_format"
    add_guard(actions, request_format, uuid_v4_guard(f"body('{envelope}')?['requestId']"), envelope,
              on_failure=lambda a, n: failure(a, n, status=400, code="INVALID_REQUEST", message="The request envelope is invalid."))
    args = f"Validate_{operation}_args"
    add_parse(actions, args, f"@body('{envelope}')?['args']",
              object_schema(DRAFT_ARGUMENTS[operation], required=DRAFT_REQUIRED_ARGS[operation], strict=True), request_format)
    return args, actions


def ascii_token_guard(value, punctuation):
    """Alphanumerics plus exactly punctuation, independent of URI punctuation escaping."""
    remainder = value
    for char in punctuation:
        literal = char.replace("'", "''")
        remainder = f"replace({remainder}, '{literal}', '')"
    # Case-sensitive replace removes allowed punctuation before encoding. Forbid
    # every residual URI-safe symbol, including those a stricter encoder escapes,
    # so only ASCII alphanumerics can remain unchanged by uriComponent.
    checks = [f"equals(uriComponent({remainder}), {remainder})"]
    checks += [f"not(contains({remainder}, '{char.replace(chr(39), chr(39) * 2)}'))" for char in "-_.!~*'()"]
    return "and(" + ", ".join(checks) + ")"


def email_guard(value):
    # Conservative ordinary ASCII subset of zod.email(): no quoted locals,
    # Unicode, consecutive/edge dots, or edge hyphens in domain labels. Trim
    # before all checks and projection, matching the Worker normalization.
    local = f"first(split({value}, '@'))"
    domain = f"last(split({value}, '@'))"
    tld = f"last(split({domain}, '.'))"
    letters = f"toLower({tld})"
    for char in "abcdefghijklmnopqrstuvwxyz":
        letters = f"replace({letters}, '{char}', '')"
    # Some trim implementations also remove NEL/ASCII record separators;
    # JavaScript String.trim() does not. Never normalize those invalid inputs
    # into accepted email addresses. Ordinary surrounding whitespace is fine.
    checks = [f"not(contains(item(), '{chr(cp)}'))" for cp in (*range(0x1C, 0x20), 0x85)]
    checks += [
        f"greater(length({value}), 0)", f"lessOrEquals(length({value}), 254)",
        f"equals(length(split({value}, '@')), 2)", f"greater(length({local}), 0)",
        ascii_token_guard(local, "_+-'."), f"not(startsWith({local}, '.'))",
        f"not(endsWith({local}, '.'))", f"not(endsWith({local}, ''''))", f"not(contains({local}, '..'))",
        ascii_token_guard(domain, "-."), f"contains({domain}, '.')",
        f"not(startsWith({domain}, '.'))", f"not(endsWith({domain}, '.'))",
        f"not(startsWith({domain}, '-'))", f"not(endsWith({domain}, '-'))",
        f"not(contains({domain}, '..'))", f"not(contains({domain}, '.-'))", f"not(contains({domain}, '-.'))",
        f"greaterOrEquals(length({tld}), 2)", f"equals({letters}, '')",
    ]
    return "and(" + ", ".join(checks) + ")"


def filename_guard(value):
    encoded = f"uriComponent({value})"
    cleaned = encoded
    # UTF-8 C1 controls are C2 80..9F; these prefixes cannot match any
    # non-control character. Literal caller percent signs encode to %25.
    for forbidden in ("%0", "%1", "%7F", "%C2%8", "%C2%9", "%2F", "%5C"):
        cleaned = f"replace({cleaned}, '{forbidden}', '')"
    return f"equals(length({encoded}), length({cleaned}))"


def mime_guard(value):
    return "and(" + ", ".join([
        ascii_token_guard(value, "!#$&^_.+-/"),
        f"equals(length(split({value}, '/')), 2)",
        f"not(empty(first(split({value}, '/'))))", f"not(empty(last(split({value}, '/'))))",
    ]) + ")"


def decoded_size(value):
    return f"sub(mul(div(length({value}), 4), 3), if(endsWith({value}, '=='), 2, if(endsWith({value}, '='), 1, 0)))"


def canonical_padding_guard(value):
    # == requires the final sextet's low four bits zero; = requires low two
    # bits zero. Native WDL indexOf is case-insensitive, so test the exact legal
    # sextets with case-sensitive contains instead. Schema minLength=4 makes
    # both substring calls eager-safe.
    penultimate = f"substring({value}, sub(length({value}), 3), 1)"
    final = f"substring({value}, sub(length({value}), 2), 1)"
    return "and(" + ", ".join([
        f"equals(mod(length({value}), 4), 0)",
        f"or(equals(indexOf({value}, '='), -1), equals(indexOf({value}, '='), sub(length({value}), if(endsWith({value}, '=='), 2, 1))))",
        f"or(not(endsWith({value}, '==')), contains('AQgw', {penultimate}))",
        f"or(not(endsWith({value}, '=')), endsWith({value}, '=='), contains('AEIMQUYcgkosw048', {final}))",
        f"greater({decoded_size(value)}, 0)", f"lessOrEquals({decoded_size(value)}, {MAX_BYTES})",
    ]) + ")"


def post(actions, operation, uri, body_value, dependency):
    body_name = compose(actions, f"Compose_{operation}_body", body_value, dependency)
    name = f"Graph_{operation}"
    action = graph(name, uri, after(body_name))[1]
    action["inputs"]["parameters"].update({"Method": "POST", "Body": f"@string(outputs('{body_name}'))"})
    actions[name] = action
    ambiguous(actions, name)
    return name


def summary_result(actions, operation, source):
    result = f"Validate_{operation}_result"
    add_parse(actions, result, "@body('" + source + "')",
              object_schema({"id": ID_SCHEMA, "isDraft": {"type": "boolean", "enum": [True]}}), source, on_failure=ambiguous)
    controls = add_guard(actions, f"Validate_{operation}_result_id", id_guard(f"body('{result}')?['id']"), result, on_failure=ambiguous)
    actions[f"Respond_{operation}"] = success({"id": field(result, "id"), "isDraft": True}, after(controls))


def create_case():
    operation = "create_draft"
    args, actions = request_args(operation)
    count = lambda key: f"length(coalesce(body('{args}')?['{key}'], json('[]')))"
    dependency = add_guard(actions, f"Validate_{operation}_controls",
                           f"and(greater({count('to')}, 0), lessOrEquals(add(add({count('to')}, {count('cc')}), {count('bcc')}), 50))", args)
    recipients = {}
    for key in ("to", "cc", "bcc"):
        name = f"Select_{operation}_{key}"
        actions[name] = secure({
            "type": "Select", "runAfter": after(dependency),
            "inputs": {
                "from": f"@coalesce(body('{args}')?['{key}'], json('[]'))",
                "select": {"emailAddress": {"address": "@if(" + email_guard("trim(item())") + ", trim(item()), null)"}},
            },
        })
        invalid_input(actions, name)
        dependency = add_parse(actions, f"Validate_{operation}_{key}", f"@body('{name}')", {
            "type": "array", "maxItems": 50,
            "items": object_schema({"emailAddress": object_schema({"address": {"type": "string", "minLength": 1, "maxLength": 254}}, strict=True)}, strict=True),
        }, name)
        recipients[key + "Recipients"] = f"@body('{name}')"
    # Optional recipients are represented as empty arrays when absent. They
    # remain fixed Graph fields, never caller-controlled JSON or expressions.
    source = post(actions, operation, GRAPH_ROOT + "messages", {
        "subject": field(args, "subject"), "body": {"contentType": "Text", "content": field(args, "body")}, **recipients,
    }, dependency)
    summary_result(actions, operation, source)
    return {"case": operation, "actions": actions}


def reply_case():
    operation = "create_reply_draft"
    args, actions = request_args(operation)
    dependency = add_guard(actions, f"Validate_{operation}_controls", id_guard(f"body('{args}')?['messageId']"), args)
    uri = f"@concat('{GRAPH_ROOT}messages/', uriComponent(body('{args}')?['messageId']), '/createReply')"
    source = post(actions, operation, uri, {"message": {"body": {"contentType": "Text", "content": field(args, "body")}}}, dependency)
    summary_result(actions, operation, source)
    return {"case": operation, "actions": actions}


def attachment_case():
    operation = "add_draft_attachment"
    args, actions = request_args(operation)
    dependency = add_guard(actions, f"Validate_{operation}_controls", "and(" + ", ".join([
        id_guard(f"body('{args}')?['draftId']"), filename_guard(f"body('{args}')?['name']"), mime_guard(f"body('{args}')?['contentType']"),
    ]) + ")", args)
    b64 = f"body('{args}')?['contentBytes']"
    alphabet = f"Select_{operation}_alphabet"
    actions[alphabet] = secure({"type": "Select", "runAfter": after(dependency), "inputs": {
        "from": f"@chunk({b64}, {BASE64_CHUNK})", "select": "@" + base64_alphabet_guard("item()"),
    }})
    invalid_input(actions, alphabet)
    dependency = add_parse(actions, f"Validate_{operation}_alphabet", f"@body('{alphabet}')", {
        "type": "array", "maxItems": (MAX_BASE64 + BASE64_CHUNK - 1) // BASE64_CHUNK,
        "items": {"type": "boolean", "enum": [True]},
    }, alphabet)
    dependency = add_guard(actions, f"Validate_{operation}_padding", canonical_padding_guard(b64), dependency)
    size = compose(actions, f"Compose_{operation}_size", "@" + decoded_size(b64), dependency)
    uri = f"@concat('{GRAPH_ROOT}messages/', uriComponent(body('{args}')?['draftId'])"
    preflight = "Graph_draft_preflight"
    actions[preflight] = graph(preflight, uri + ", '?$select=id,isDraft')", after(size))[1]
    failure(actions, preflight, code="DRAFT_PREFLIGHT_FAILED", message="The draft could not be verified; no attachment upload was attempted.")
    verified = "Validate_draft_preflight"
    preflight_failure = lambda a, n: failure(a, n, code="INVALID_DRAFT_PREFLIGHT", message="The draft could not be verified; no attachment upload was attempted.")
    add_parse(actions, verified, f"@body('{preflight}')", object_schema({"id": ID_SCHEMA, "isDraft": {"type": "boolean"}}), preflight, on_failure=preflight_failure)
    validated = add_guard(actions, "Validate_draft_preflight_id", id_guard(f"body('{verified}')?['id']"), verified, on_failure=preflight_failure)
    upload = {}
    source = post(upload, operation, uri + ", '/attachments')", {
        "@@odata.type": "#microsoft.graph.fileAttachment", "name": field(args, "name"),
        "contentType": field(args, "contentType"), "contentBytes": field(args, "contentBytes"),
    }, None)
    result = f"Validate_{operation}_result"
    add_parse(upload, result, f"@body('{source}')", object_schema({
        "id": ID_SCHEMA, "name": {"type": "string", "minLength": 1, "maxLength": 255},
        # Graph's attachment size is Int32 metadata, not the raw-file length.
        "size": {"type": "integer", "minimum": 0, "maximum": 2147483647},
        "contentBytes": {"type": "string", "minLength": 4, "maxLength": MAX_BASE64},
    }), source, on_failure=ambiguous)
    # Equal numeric lengths plus documented case-sensitive contains() prove full
    # payload identity with the canonical, nonempty request (at most 2 MiB raw).
    # Do not rely on equals() string case behavior or Graph metadata overhead.
    # https://learn.microsoft.com/en-us/azure/logic-apps/expression-functions-reference#contains
    controls = add_guard(upload, f"Validate_{operation}_result_identity", "and(" + ", ".join([
        id_guard(f"body('{result}')?['id']"),
        f"equals(body('{result}')?['name'], body('{args}')?['name'])",
        f"equals(length(body('{result}')?['contentBytes']), length({b64}))",
        f"contains(body('{result}')?['contentBytes'], {b64})",
    ]) + ")", result, on_failure=ambiguous)
    upload[f"Respond_{operation}"] = success({
        "draftId": field(args, "draftId"), "attachmentId": field(result, "id"),
        # Return the verified raw-file byte count, never Graph size or content.
        "name": field(result, "name"), "size": "@" + decoded_size(b64),
    }, after(controls))
    actions["Check_draft_preflight"] = condition({"and": [
        {"equals": [field(verified, "id"), field(args, "draftId")]},
        {"equals": [field(verified, "isDraft"), True]},
    ]}, upload, {"Reject_draft_preflight": error(409, "DRAFT_NOT_VERIFIED", "The target is not the requested draft; no attachment upload was attempted.")}, after(validated))
    return {"case": operation, "actions": actions}


def augment_draft_tools(definition):
    """Append only draft cases/operations after all read augmentations."""
    definition["triggers"]["manual"]["inputs"]["schema"]["properties"]["operation"]["enum"].extend(DRAFT_OPERATIONS)
    cases = definition["actions"]["スイッチ"]["cases"]
    for build in (create_case, reply_case, attachment_case):
        case = build()
        cases[case["case"]] = case
    return definition
