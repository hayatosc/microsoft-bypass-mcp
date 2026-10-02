"""Offline draft contracts against the single generated definition.json.

Uses the existing strict Flow/Expressions emulator with synthetic connector
results only. Exact request expectations below are independent of the generator;
this is not a live Power Automate/Office365 run or an importer verification.
"""

import base64
from contextlib import nullcontext
from copy import deepcopy
import itertools
import json
import re
import unittest
from unittest.mock import patch
from urllib.parse import quote

from build_draft_tools_flow import DRAFT_OPERATIONS, ascii_token_guard
from test_attachment_flow import (
    FORBIDDEN_ID_CODEPOINTS, REQUEST_ID, ROOT, SOURCE, Expressions, Flow,
    string_length, validate, walk,
)

OPERATIONS = ("create_draft", "create_reply_draft", "add_draft_attachment")
CASES = {name: SOURCE["actions"]["スイッチ"]["cases"][name] for name in OPERATIONS}
GRAPH_ROOT = "https://graph.microsoft.com/v1.0/me/messages"
MAX_BYTES = 2097152
MAX_ENCODED = 2796204
CANARY = "synthetic-private-draft-canary"
VALID_ARGS = {
    "create_draft": {"to": ["alice@example.com"], "subject": "subject", "body": "plain text"},
    "create_reply_draft": {"messageId": "original-id", "body": "plain text"},
    "add_draft_attachment": {"draftId": "draft-id", "name": "sample.pdf", "contentType": "application/pdf", "contentBytes": "YWJj"},
}
AMBIGUOUS_ERROR = {"ok": False, "error": {
    "code": "DRAFT_WRITE_AMBIGUOUS",
    "message": "The draft write outcome may be ambiguous; inspect Drafts before retrying.",
}}


def action_maps(actions):
    yield actions
    for action in actions.values():
        if "actions" in action:
            yield from action_maps(action["actions"])
        if "else" in action:
            yield from action_maps(action["else"]["actions"])


def named_actions(operation):
    return {name: action for group in action_maps(CASES[operation]["actions"]) for name, action in group.items()}


def upstream(operation, args):
    if operation == "add_draft_attachment":
        return [
            {"id": args["draftId"], "isDraft": True},
            {"id": "attachment-id", "name": args["name"], "size": len(base64.b64decode(args["contentBytes"]))},
        ]
    return [{"id": "draft-id", "isDraft": True}]


def draft_flow(operation, args=None, responses=None, **kwargs):
    args = deepcopy(VALID_ARGS[operation] if args is None else args)
    return Flow(operation, args, upstream(operation, args) if responses is None else responses, **kwargs)


def rfc3986_uri_encoder():
    """Scoped portability fixture; leave the default emulator semantics unchanged."""
    original_call = Expressions.call

    def call(expressions, name, args):
        if name == "uriComponent":
            return quote(*args, safe="~")
        return original_call(expressions, name, args)

    return patch.object(Expressions, "call", call)


class DraftStaticContracts(unittest.TestCase):
    def test_exact_draft_surface_appended_after_read_surface(self):
        self.assertEqual(DRAFT_OPERATIONS, OPERATIONS)
        from build_read_tools_flow import FLOW_OPERATIONS
        self.assertEqual(SOURCE["triggers"]["manual"]["inputs"]["schema"]["properties"]["operation"]["enum"], list(FLOW_OPERATIONS) + list(OPERATIONS))
        self.assertEqual(list(SOURCE["actions"]["スイッチ"]["cases"])[-3:], list(OPERATIONS))
        expected_args = {
            "create_draft": ({"to", "cc", "bcc", "subject", "body"}, {"to", "subject", "body"}),
            "create_reply_draft": ({"messageId", "body"}, {"messageId", "body"}),
            "add_draft_attachment": ({"draftId", "name", "contentType", "contentBytes"}, {"draftId", "name", "contentType", "contentBytes"}),
        }
        id_schema = {"type": "string", "minLength": 1, "maxLength": 2048}
        create = CASES["create_draft"]["actions"]["Validate_create_draft_args"]["inputs"]["schema"]["properties"]
        for key in ("to", "cc", "bcc"):
            self.assertEqual(create[key], {"type": "array", "maxItems": 50, "items": {"type": "string", "maxLength": 254}})
        self.assertEqual(create["subject"], {"type": "string", "maxLength": 512})
        self.assertEqual(create["body"], {"type": "string", "maxLength": 20000})
        reply = CASES["create_reply_draft"]["actions"]["Validate_create_reply_draft_args"]["inputs"]["schema"]["properties"]
        self.assertEqual(reply, {"messageId": id_schema, "body": {"type": "string", "maxLength": 20000}})
        attachment = CASES["add_draft_attachment"]["actions"]["Validate_add_draft_attachment_args"]["inputs"]["schema"]["properties"]
        self.assertEqual(attachment, {
            "draftId": id_schema, "name": {"type": "string", "minLength": 1, "maxLength": 255},
            "contentType": {"type": "string", "minLength": 1, "maxLength": 127},
            "contentBytes": {"type": "string", "minLength": 4, "maxLength": MAX_ENCODED},
        })
        for operation, (fields, required) in expected_args.items():
            schema = CASES[operation]["actions"][f"Validate_{operation}_args"]["inputs"]["schema"]
            self.assertEqual(set(schema["properties"]), fields)
            self.assertEqual(set(schema["required"]), required)
            self.assertIs(schema["additionalProperties"], False)
        compact = re.sub(r"\s+", "", (ROOT / "src/features/outlook/drafts/schema.ts").read_text())
        for literal in ("MAX_DRAFT_RECIPIENTS=50", "MAX_DRAFT_SUBJECT_CHARACTERS=512", "MAX_DRAFT_BODY_CHARACTERS=20000", "MAX_DRAFT_ATTACHMENT_BYTES=2*1024*1024", "MAX_DRAFT_ATTACHMENT_NAME_CHARACTERS=255", "MAX_DRAFT_ATTACHMENT_CONTENT_TYPE_CHARACTERS=127"):
            self.assertIn(literal, compact)
        self.assertIn("/^[A-Za-z0-9!#$&^_.+-]+\\/[A-Za-z0-9!#$&^_.+-]+$(?![\\s\\S])/", compact)

    def test_fixed_exact_routes_methods_parameters_and_serialized_body(self):
        expected = {
            "Graph_create_draft": ("POST", GRAPH_ROOT, "create_draft"),
            "Graph_create_reply_draft": ("POST", "@concat('https://graph.microsoft.com/v1.0/me/messages/', uriComponent(body('Validate_create_reply_draft_args')?['messageId']), '/createReply')", "create_reply_draft"),
            "Graph_draft_preflight": ("GET", "@concat('https://graph.microsoft.com/v1.0/me/messages/', uriComponent(body('Validate_add_draft_attachment_args')?['draftId']), '?$select=id,isDraft')", None),
            "Graph_add_draft_attachment": ("POST", "@concat('https://graph.microsoft.com/v1.0/me/messages/', uriComponent(body('Validate_add_draft_attachment_args')?['draftId']), '/attachments')", "add_draft_attachment"),
        }
        actual = {name: action for operation in OPERATIONS for name, action in named_actions(operation).items() if action["type"] == "OpenApiConnection"}
        self.assertEqual(set(actual), set(expected))
        for name, (method, uri, operation) in expected.items():
            action = actual[name]
            parameters = {"Uri": uri, "Method": method, "CustomHeader1": 'Prefer: IdType="ImmutableId"', "ContentType": "application/json"}
            if operation:
                parameters["Body"] = f"@string(outputs('Compose_{operation}_body'))"
            self.assertEqual(action["inputs"], {
                "parameters": parameters,
                "host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_office365", "connectionName": "shared_office365", "operationId": "HttpRequest"},
                "authentication": "@parameters('$authentication')", "retryPolicy": {"type": "none"},
            })
        for operation in OPERATIONS:
            body = named_actions(operation)[f"Compose_{operation}_body"]["inputs"]
            if operation == "create_draft":
                self.assertEqual(body, {
                    "subject": "@body('Validate_create_draft_args')?['subject']",
                    "body": {"contentType": "Text", "content": "@body('Validate_create_draft_args')?['body']"},
                    **{key + "Recipients": f"@body('Select_create_draft_{key}')" for key in ("to", "cc", "bcc")},
                })
            elif operation == "create_reply_draft":
                self.assertEqual(body, {"message": {"body": {"contentType": "Text", "content": "@body('Validate_create_reply_draft_args')?['body']"}}})
            else:
                self.assertEqual(body, {"@@odata.type": "#microsoft.graph.fileAttachment", **{key: f"@body('Validate_add_draft_attachment_args')?['{key}']" for key in ("name", "contentType", "contentBytes")}})
                self.assertEqual(Expressions(None).resolve({"@@odata.type": body["@@odata.type"]}), {"@odata.type": "#microsoft.graph.fileAttachment"})
        self.assertNotRegex(json.dumps(CASES), r"(?i)(?:/send|/delete|/replyall|createreplyall|\$expand|\$value|nextlink)")

    def test_schema_guard_and_side_effect_dependency_order(self):
        for operation in OPERATIONS:
            actions = named_actions(operation)
            for name, predecessor in (
                (f"Validate_{operation}_request_format", f"Validate_{operation}_request"),
                (f"Validate_{operation}_args", f"Validate_{operation}_request_format"),
                (f"Validate_{operation}_controls", f"Validate_{operation}_args"),
                (f"Graph_{operation}", f"Compose_{operation}_body"),
                (f"Validate_{operation}_result", f"Graph_{operation}"),
            ):
                self.assertEqual(actions[name]["runAfter"], {predecessor: ["Succeeded"]})
            envelope = actions[f"Validate_{operation}_request"]["inputs"]["schema"]
            self.assertEqual(envelope, {"type": "object", "properties": {
                "operation": {"type": "string", "enum": [operation]}, "requestId": {"type": "string", "minLength": 36, "maxLength": 36}, "args": {"type": "object"},
            }, "required": ["operation", "requestId", "args"], "additionalProperties": False})
        actions = named_actions("create_draft")
        predecessor = "Validate_create_draft_controls"
        for key in ("to", "cc", "bcc"):
            self.assertEqual(actions[f"Select_create_draft_{key}"]["runAfter"], {predecessor: ["Succeeded"]})
            predecessor = f"Validate_create_draft_{key}"
            self.assertEqual(actions[predecessor]["runAfter"], {f"Select_create_draft_{key}": ["Succeeded"]})
        self.assertEqual(actions["Compose_create_draft_body"]["runAfter"], {predecessor: ["Succeeded"]})
        self.assertEqual(named_actions("create_reply_draft")["Compose_create_reply_draft_body"]["runAfter"], {"Validate_create_reply_draft_controls": ["Succeeded"]})
        actions = named_actions("add_draft_attachment")
        chain = ["Validate_add_draft_attachment_controls", "Select_add_draft_attachment_alphabet", "Validate_add_draft_attachment_alphabet", "Validate_add_draft_attachment_padding", "Compose_add_draft_attachment_size", "Graph_draft_preflight", "Validate_draft_preflight", "Validate_draft_preflight_id", "Check_draft_preflight"]
        for predecessor, name in zip(chain, chain[1:]):
            self.assertEqual(actions[name]["runAfter"], {predecessor: ["Succeeded"]})
        gate = actions["Check_draft_preflight"]
        self.assertEqual(gate["expression"], {"and": [
            {"equals": ["@body('Validate_draft_preflight')?['id']", "@body('Validate_add_draft_attachment_args')?['draftId']"]},
            {"equals": ["@body('Validate_draft_preflight')?['isDraft']", True]},
        ]})
        self.assertIn("Graph_add_draft_attachment", gate["actions"])
        self.assertNotIn("Graph_add_draft_attachment", gate["else"]["actions"])
        self.assertEqual(actions["Compose_add_draft_attachment_body"]["runAfter"], {})
        self.assertEqual(actions["Select_add_draft_attachment_alphabet"]["inputs"]["from"], "@chunk(body('Validate_add_draft_attachment_args')?['contentBytes'], 8192)")
        self.assertNotIn("body(", actions["Select_add_draft_attachment_alphabet"]["inputs"]["select"])
        for operation in OPERATIONS:
            actions = named_actions(operation)
            controls = "identity" if operation == "add_draft_attachment" else "id"
            self.assertEqual(actions[f"Respond_{operation}"]["runAfter"], {f"Validate_{operation}_result_{controls}": ["Succeeded"]})

    def test_canonical_padding_uses_case_sensitive_legal_sextets(self):
        actions = named_actions("add_draft_attachment")
        b64 = "body('Validate_add_draft_attachment_args')?['contentBytes']"
        padding = actions["Validate_add_draft_attachment_padding"]["inputs"]["content"]
        self.assertIn(f"contains('AQgw', substring({b64}, sub(length({b64}), 3), 1))", padding)
        self.assertIn(f"contains('AEIMQUYcgkosw048', substring({b64}, sub(length({b64}), 2), 1))", padding)
        self.assertNotIn("indexOf('", padding)
        # Both substring offsets must be safe even with eager WDL evaluation.
        self.assertEqual(actions["Validate_add_draft_attachment_args"]["inputs"]["schema"]["properties"]["contentBytes"]["minLength"], 4)

    def test_all_data_actions_secure_and_have_static_failure_handlers(self):
        for operation in OPERATIONS:
            for group in action_maps(CASES[operation]["actions"]):
                for name, action in group.items():
                    kind = action["type"]
                    if kind == "If":
                        self.assertNotIn("runtimeConfiguration", action)
                        continue
                    self.assertIn(kind, {"ParseJson", "Select", "Compose", "Response", "OpenApiConnection"})
                    properties = ["inputs"] if kind in {"ParseJson", "Compose", "Response"} else ["inputs", "outputs"]
                    self.assertEqual(action["runtimeConfiguration"], {"secureData": {"properties": properties}})
                    if kind != "Response":
                        handlers = [value for value in group.values() if value.get("runAfter", {}).get(name) == ["Failed", "TimedOut"]]
                        self.assertEqual(len(handlers), 1, name)
                        self.assertEqual(handlers[0]["type"], "Response")
                    else:
                        self.assertEqual(action["inputs"]["headers"], {"Content-Type": "application/json", "Cache-Control": "no-store"})
                        if action["inputs"]["statusCode"] >= 400:
                            body = action["inputs"]["body"]
                            self.assertEqual(set(body), {"ok", "error"})
                            self.assertNotIn("@", json.dumps(body))
                            self.assertRegex(body["error"]["code"], r"^[A-Z][A-Z_]+$")
        for node in walk(CASES):
            if isinstance(node, str) and node.startswith("@"):
                self.assertLessEqual(string_length(node), 8192)
            if isinstance(node, dict):
                self.assertTrue(set(node).isdisjoint({"pattern", "patternProperties", "trackedProperties"}))
        self.assertNotRegex(json.dumps(CASES), r"base64ToBinary|base64ToString|InitializeVariable|SetVariable|AppendToArrayVariable")

    def test_alternate_uri_encoder_is_scoped_and_delegates_other_functions(self):
        expressions = Expressions(None)
        original_call = Expressions.call
        self.assertEqual(expressions.call("uriComponent", ["!*'()~"]), "!*'()~")
        with rfc3986_uri_encoder():
            self.assertEqual(expressions.call("uriComponent", ["!*'()~"]), "%21%2A%27%28%29~")
            self.assertEqual(expressions.call("replace", ["AbAa", "a", "x"]), "AbAx")
            self.assertFalse(expressions.call("contains", ["AQgw", "G"]))
        self.assertIs(Expressions.call, original_call)
        self.assertEqual(expressions.call("uriComponent", ["!*'()~"]), "!*'()~")

    def test_ascii_token_helper_exact_alphabets_under_both_uri_encoders(self):
        alphanumeric = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
        for encoder, context in (("encodeURIComponent", nullcontext), ("RFC3986", rfc3986_uri_encoder)):
            with self.subTest(encoder=encoder), context():
                for punctuation in ("_+-'.", "-.", "!#$&^_.+-/"):
                    expression = "@" + ascii_token_guard("item()", punctuation)
                    values = ["", alphanumeric + punctuation] + [chr(cp) for cp in range(128)]
                    values += ["é", "☃", "😀", "\u0085", "\u00a0", "\u200b", "\u2028", "\ufeff", "%21", "%27", "%00"]
                    for value in values:
                        with self.subTest(punctuation=punctuation, value=value):
                            expected = all(char in alphanumeric + punctuation for char in value)
                            self.assertEqual(Expressions(None, item=value).resolve(expression), expected)

    def test_only_documented_extra_wdl_functions_and_safe_json_serialization(self):
        expressions = Expressions(None)
        self.assertEqual(expressions.resolve("@add(19, 31)"), 50)
        self.assertEqual(expressions.call("add", [1.5, 2]), 3.5)
        for function, args in (("add", [True, 1]), ("add", ["1", 2]), ("startsWith", [None, "a"]), ("startsWith", ["a", []]), ("last", [{}])):
            with self.assertRaises(ValueError):
                expressions.call(function, args)
        self.assertEqual(expressions.resolve("@startsWith('Example', 'ex')"), True)
        self.assertEqual(expressions.resolve("@startsWith('', '-')"), False)
        self.assertEqual(expressions.resolve("@last(split('a.b.c', '.'))"), "c")
        self.assertEqual(expressions.resolve("@last('abc')"), "c")
        value = {"body": 'quotes"\\\n@{expression}', "recipient": "o'connor@example.com"}
        self.assertEqual(json.loads(expressions.call("string", [value])), value)
        for schema in ({"type": "string", "format": "email"}, {"type": "string", "unknownKeyword": 1}):
            with self.assertRaises(AssertionError):
                validate("value", schema)
        with self.assertRaises(AssertionError):
            expressions.resolve("@matches('x', 'regex')")

    def test_native_wdl_string_case_semantics_and_operand_types(self):
        expressions = Expressions(None)
        for source, expected in (
            ("@startsWith('Example', 'EX')", True),
            ("@startsWith('Example', 'am')", False),
            ("@endsWith('Example', 'PLE')", True),
            ("@endsWith('Example', 'EX')", False),
            ("@endsWith('', '-')", False),
            ("@endsWith('Example', '')", True),
            ("@contains('AQgw', 'g')", True),
            ("@contains('AQgw', 'G')", False),
            ("@replace('AbAa', 'a', 'x')", "AbAx"),
            ("@replace('AbAa', 'A', 'x')", "xbxa"),
            ("@equals('Draft-ID', 'draft-id')", False),
            ("@equals('Draft-ID', 'Draft-ID')", True),
        ):
            with self.subTest(source=source):
                self.assertEqual(expressions.resolve(source), expected)
        alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
        for value, part, expected in (
            ("hello WORLD", "world", 6), ("HeLLo world", "LLo", 2),
            (alphabet, "g", 6), (alphabet, "G", 6), (alphabet, "w", 22),
            ("Yg==", "=", 2), ("YWg=", "=", 3), ("YWJj", "=", -1),
            ("", "", 0), ("", "a", -1), ("abc", "", 0),
        ):
            with self.subTest(value=value, part=part):
                offset = expressions.call("indexOf", [value, part])
                self.assertIs(type(offset), int)
                self.assertEqual(offset, expected)
        for function in ("startsWith", "endsWith", "indexOf"):
            for invalid in (None, True, 123, 1.5, [], {}):
                for args in ([invalid, "a"], ["a", invalid]):
                    with self.subTest(function=function, args=args), self.assertRaises(ValueError):
                        expressions.call(function, args)
        self.assertFalse(expressions.call("contains", [["Draft-ID"], "draft-id"]))
        self.assertFalse(expressions.call("contains", [{"Draft-ID": True}, "draft-id"]))


class DraftExecutionContracts(unittest.TestCase):
    def assert_error(self, flow, *, status=400, calls=0, ambiguous=False, code=None):
        result = flow.run()
        self.assertEqual(result["statusCode"], status)
        self.assertEqual(len(flow.calls), calls)
        self.assertEqual(set(result["body"]), {"ok", "error"})
        self.assertIs(result["body"]["ok"], False)
        self.assertNotIn(CANARY, json.dumps(result))
        if ambiguous:
            self.assertEqual(result["body"], AMBIGUOUS_ERROR)
        if code:
            self.assertEqual(result["body"]["error"]["code"], code)
        return result

    def test_create_request_json_and_projection_redact_all_other_data(self):
        text = CANARY + '"\\\r\n日本語 @{triggerBody()} <b>not html</b>'
        args = {"to": [" alice@example.com ", "o'connor+test@example.co.uk"], "cc": [" CC@example.org "], "bcc": ["bcc@example.net"], "subject": text, "body": text}
        flow = draft_flow("create_draft", args, [{"id": "new-draft", "isDraft": True, "body": text, "toRecipients": CANARY, "webLink": "https://outside.invalid/" + CANARY, "contentBytes": CANARY}])
        self.assertEqual(flow.run()["body"], {"ok": True, "requestId": REQUEST_ID, "operation": "create_draft", "data": {"id": "new-draft", "isDraft": True}})
        self.assertEqual(len(flow.calls), 1)
        self.assertEqual(flow.calls[0]["Uri"], GRAPH_ROOT)
        self.assertEqual(flow.calls[0]["Method"], "POST")
        self.assertIsInstance(flow.calls[0]["Body"], str)
        self.assertEqual(json.loads(flow.calls[0]["Body"]), {
            "subject": text, "body": {"contentType": "Text", "content": text},
            "toRecipients": [{"emailAddress": {"address": email}} for email in ("alice@example.com", "o'connor+test@example.co.uk")],
            "ccRecipients": [{"emailAddress": {"address": "CC@example.org"}}],
            "bccRecipients": [{"emailAddress": {"address": "bcc@example.net"}}],
        })
        flow = draft_flow("create_draft")
        self.assertEqual(flow.run()["statusCode"], 200)
        self.assertEqual(json.loads(flow.calls[0]["Body"])["ccRecipients"], [])
        self.assertEqual(json.loads(flow.calls[0]["Body"])["bccRecipients"], [])

    def test_sender_only_reply_json_and_encoded_route(self):
        args = {"messageId": "original/a?b#c%20", "body": CANARY + '"\\\n@{x}'}
        flow = draft_flow("create_reply_draft", args, [{"id": "reply-id", "isDraft": True, "body": CANARY, "recipients": CANARY}])
        self.assertEqual(flow.run()["body"]["data"], {"id": "reply-id", "isDraft": True})
        self.assertEqual(flow.calls[0]["Uri"], GRAPH_ROOT + "/original%2Fa%3Fb%23c%2520/createReply")
        self.assertEqual(flow.calls[0]["Method"], "POST")
        self.assertEqual(json.loads(flow.calls[0]["Body"]), {"message": {"body": {"contentType": "Text", "content": args["body"]}}})

    def test_attachment_preflight_upload_json_and_safe_result(self):
        args = dict(VALID_ARGS["add_draft_attachment"], draftId="draft/a?b#c%20", name='日本語 "@file.pdf', contentBytes=base64.b64encode(CANARY.encode()).decode())
        raw = upstream("add_draft_attachment", args)
        for item in raw:
            item.update(body=CANARY, contentBytes=args["contentBytes"], webLink="https://outside.invalid/" + CANARY)
        flow = draft_flow("add_draft_attachment", args, raw)
        self.assertEqual(flow.run()["body"]["data"], {"draftId": args["draftId"], "attachmentId": "attachment-id", "name": args["name"], "size": len(CANARY)})
        self.assertEqual([call["Method"] for call in flow.calls], ["GET", "POST"])
        path = GRAPH_ROOT + "/draft%2Fa%3Fb%23c%2520"
        self.assertEqual([call["Uri"] for call in flow.calls], [path + "?$select=id,isDraft", path + "/attachments"])
        self.assertNotIn("Body", flow.calls[0])
        self.assertEqual(json.loads(flow.calls[1]["Body"]), {"@odata.type": "#microsoft.graph.fileAttachment", **{key: args[key] for key in ("name", "contentType", "contentBytes")}})

    def test_strict_envelopes_uuid_args_and_no_route_overrides(self):
        for operation in OPERATIONS:
            for changes in ({"extra": CANARY}, {"args": None}, {"args": []}, {"requestId": None}, {"requestId": CANARY}):
                flow = draft_flow(operation, responses=[])
                flow.request.update(changes)
                self.assert_error(flow, code="INVALID_REQUEST")
            for key in ("url", "uri", "method", "nextLink", "operationId", "send", "replyAll", "from", "headers", "bodyContentType"):
                flow = draft_flow(operation, dict(VALID_ARGS[operation], **{key: CANARY}), [])
                self.assert_error(flow, code="INVALID_ARGUMENTS")
            required = set(VALID_ARGS[operation])
            for key in required:
                args = dict(VALID_ARGS[operation])
                del args[key]
                self.assert_error(draft_flow(operation, args, []))
            for args in (None, [], "string", True):
                flow = draft_flow(operation, responses=[])
                flow.request["args"] = args
                self.assert_error(flow, code="INVALID_REQUEST")
            for request_id in (True, 123, [], {}, "", REQUEST_ID[:-1], REQUEST_ID + "0"):
                flow = draft_flow(operation, responses=[])
                flow.request["requestId"] = request_id
                self.assert_error(flow, code="INVALID_REQUEST")
            for index, replacement in ((0, "g"), (0, "\u0661"), (8, "0"), (13, "0"), (18, "0"), (23, "0"), (14, "5"), (19, "7"), (35, "\n")):
                flow = draft_flow(operation, responses=[])
                flow.request["requestId"] = REQUEST_ID[:index] + replacement + REQUEST_ID[index + 1:]
                self.assert_error(flow, code="INVALID_REQUEST")
            flow = draft_flow(operation)
            flow.request["requestId"] = REQUEST_ID.upper()
            self.assertEqual(flow.run()["statusCode"], 200)

    def test_subject_body_utf16_bounds_and_types_before_any_connector(self):
        for operation in ("create_draft", "create_reply_draft"):
            for key, maximum in (("body", 20000), ("subject", 512)):
                if key not in VALID_ARGS[operation]:
                    continue
                for value in ("x" * (maximum + 1), "😀" * (maximum // 2 + 1), None, True, 123, [], {}):
                    self.assert_error(draft_flow(operation, dict(VALID_ARGS[operation], **{key: value}), []))
                for value in ("", "x" * maximum, "😀" * (maximum // 2)):
                    flow = draft_flow(operation, dict(VALID_ARGS[operation], **{key: value}))
                    self.assertEqual(flow.run()["statusCode"], 200)
                    request = json.loads(flow.calls[0]["Body"])
                    actual = request[key] if key == "subject" else (request["message"]["body"]["content"] if operation == "create_reply_draft" else request["body"]["content"])
                    self.assertEqual(actual, value)

    def test_recipient_counts_required_to_and_optional_list_types(self):
        for key in ("to", "cc", "bcc"):
            for value in (None, True, 123, "alice@example.com", {}, [None], [True], [123], [{}], ["a@example.com"] * 51):
                self.assert_error(draft_flow("create_draft", dict(VALID_ARGS["create_draft"], **{key: value}), []))
        self.assert_error(draft_flow("create_draft", dict(VALID_ARGS["create_draft"], to=[]), []))
        for count in (1, 50):
            flow = draft_flow("create_draft", dict(VALID_ARGS["create_draft"], to=["a@example.com"] * count, cc=[], bcc=[]))
            self.assertEqual(flow.run()["statusCode"], 200)
        args = dict(VALID_ARGS["create_draft"], to=["a@example.com"] * 20, cc=["b@example.com"] * 20, bcc=["c@example.com"] * 10)
        self.assertEqual(draft_flow("create_draft", args).run()["statusCode"], 200)
        args["bcc"].append("c@example.com")
        self.assert_error(draft_flow("create_draft", args, []))

    def test_oversized_raw_recipients_fail_parse_json_before_guards_or_connectors(self):
        maximum_address = "a" * 242 + "@example.com"
        self.assertEqual(len(maximum_address), 254)
        oversized = (
            "a" * 243 + "@example.com",
            "a" * 200000 + "@example.com",
            " " + maximum_address,
            maximum_address + " ",
            " " * 122 + "a@example.com" + " " * 122,
            " " * 200000 + "alice@example.com" + " " * 200000,
        )
        for key in ("to", "cc", "bcc"):
            for address in oversized:
                with self.subTest(key=key, raw_length=len(address)):
                    self.assertGreater(len(address), 254)
                    flow = draft_flow("create_draft", dict(VALID_ARGS["create_draft"], **{key: [address]}), [])
                    self.assert_error(flow, code="INVALID_ARGUMENTS")
                    self.assertEqual(flow.status["Validate_create_draft_args"], "Failed")
                    self.assertEqual(flow.status["Validate_create_draft_controls"], "Skipped")
                    for recipient_key in ("to", "cc", "bcc"):
                        self.assertEqual(flow.status[f"Select_create_draft_{recipient_key}"], "Skipped")
                    self.assertEqual(flow.status["Graph_create_draft"], "Skipped")

    def test_email_syntax_ascii_bounds_and_trim(self):
        valid = ["a@example.com", "o'connor@example.com", "A.Z_+-'b@sub.example.co.uk", "x@a-b.example", " a@example.com \r\n", "a" * 242 + "@example.com"]
        invalid = ["", " \n\t ", "a" * 243 + "@example.com", "a", "@example.com", "a@", "a@@example.com", "a@localhost", "a@domain.c", "a@domain.12", "a@domain.c0m", "a@.example.com", "a@exam..ple.com", "a@example.com.", "a@-example.com", "a@example-.com", "a@sub.-example.com", ".a@example.com", "a.@example.com", "a..b@example.com", "a'@example.com", '"a"@example.com', "a b@example.com", "ü@example.com", "a@exämple.com", "a@☃.com", "a@example.com\n" + CANARY, "a@exam/ple.com", "a@exam_ple.com"]
        invalid += [chr(cp) + "a@example.com" + chr(cp) for cp in (*range(0x1C, 0x20), 0x85)]
        for encoder, context in (("encodeURIComponent", nullcontext), ("RFC3986", rfc3986_uri_encoder)):
            with self.subTest(encoder=encoder), context():
                for address in valid:
                    flow = draft_flow("create_draft", dict(VALID_ARGS["create_draft"], to=[address]))
                    self.assertEqual(flow.run()["statusCode"], 200, address)
                    self.assertEqual(json.loads(flow.calls[0]["Body"])["toRecipients"][0]["emailAddress"]["address"], address.strip())
                for key in ("to", "cc", "bcc"):
                    for address in invalid:
                        self.assert_error(draft_flow("create_draft", dict(VALID_ARGS["create_draft"], **{key: [address]}), []))
                for cp in range(128):
                    char = chr(cp)
                    args = dict(VALID_ARGS["create_draft"], to=["x" + char + "y@example.com"])
                    allowed = char in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_+-'."
                    flow = draft_flow("create_draft", args, [{"id": "draft-id", "isDraft": True}] if allowed else [])
                    self.assertEqual(flow.run()["statusCode"], 200 if allowed else 400, repr(char))

    def test_ids_bounds_controls_encoding_and_unicode(self):
        invalid = ["", ".", "..", "x" * 2049, "😀" * 1025, None, True, 123, [], {}, "x\ud800y", "x\udfffy"]
        invalid += ["x" + chr(cp) + "y" for cp in FORBIDDEN_ID_CODEPOINTS]
        for operation, key in (("create_reply_draft", "messageId"), ("add_draft_attachment", "draftId")):
            for value in invalid:
                self.assert_error(draft_flow(operation, dict(VALID_ARGS[operation], **{key: value}), []))
            for value in ("x" * 2048, "😀" * 1024, "日本語-é-\u200b", "%00%20%C2%A0", "a/b?x=1#y"):
                args = dict(VALID_ARGS[operation], **{key: value})
                flow = draft_flow(operation, args)
                self.assertEqual(flow.run()["statusCode"], 200)
                self.assertTrue(all(call["Uri"].startswith(GRAPH_ROOT + "/" + quote(value, safe="~()*!.'-_")) for call in flow.calls))

    def test_filename_bounds_exact_forbidden_controls_and_percent_literals(self):
        invalid = ["", "x" * 256, "😀" * 128, "a/b", "a\\b", None, True, 123, [], {}, "a\ud800b"]
        invalid += ["x" + chr(cp) + "y" for cp in (*range(32), *range(127, 160))]
        for name in invalid:
            self.assert_error(draft_flow("add_draft_attachment", dict(VALID_ARGS["add_draft_attachment"], name=name), []))
        for name in ("x" * 255, "😀" * 127 + "x", "日本語-é \u200b.pdf", "%00%2F%5C%C2%80", "..", 'quote"@name.pdf'):
            flow = draft_flow("add_draft_attachment", dict(VALID_ARGS["add_draft_attachment"], name=name))
            self.assertEqual(flow.run()["statusCode"], 200)
            self.assertEqual(json.loads(flow.calls[1]["Body"])["name"], name)

    def test_mime_exact_worker_token_alphabet_and_length(self):
        for encoder, context in (("encodeURIComponent", nullcontext), ("RFC3986", rfc3986_uri_encoder)):
            with self.subTest(encoder=encoder), context():
                for cp in range(128):
                    char = chr(cp)
                    allowed = char in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!#$&^_.+-"
                    args = dict(VALID_ARGS["add_draft_attachment"], contentType="x" + char + "y/a")
                    flow = draft_flow("add_draft_attachment", args, upstream("add_draft_attachment", args) if allowed else [])
                    self.assertEqual(flow.run()["statusCode"], 200 if allowed else 400, repr(char))
                for content_type in ("", "a/", "/a", "a", "a/b/c", "text/plain; charset=utf-8", " text/plain", "text/plain\n", "a/é", "a/☃", "a/~", "a/*", "a/'", "a/" + "b" * 126, None, True, 123, [], {}):
                    self.assert_error(draft_flow("add_draft_attachment", dict(VALID_ARGS["add_draft_attachment"], contentType=content_type), []))
                for content_type in ("application/x!custom", "a/" + "b" * 125):
                    flow = draft_flow("add_draft_attachment", dict(VALID_ARGS["add_draft_attachment"], contentType=content_type))
                    self.assertEqual(flow.run()["statusCode"], 200)
                    self.assertEqual(json.loads(flow.calls[1]["Body"])["contentType"], content_type)

    def test_native_wdl_canonical_base64_padding_regressions(self):
        for text, valid in (("Yg==", True), ("Yw==", True), ("YWg=", True), ("Ya==", False), ("Yq==", False), ("YWa=", False)):
            with self.subTest(contentBytes=text):
                args = dict(VALID_ARGS["add_draft_attachment"], contentBytes=text)
                flow = draft_flow("add_draft_attachment", args, upstream("add_draft_attachment", args) if valid else [])
                if valid:
                    self.assertEqual(flow.run()["body"]["data"]["size"], len(base64.b64decode(text)))
                    self.assertEqual([call["Method"] for call in flow.calls], ["GET", "POST"])
                    self.assertEqual(json.loads(flow.calls[1]["Body"])["contentBytes"], text)
                else:
                    self.assert_error(flow, code="INVALID_ARGUMENTS")
                    self.assertEqual(flow.status["Validate_add_draft_attachment_padding"], "Failed")
                    self.assertEqual(flow.status["Graph_draft_preflight"], "Skipped")

    def test_base64_padding_and_unused_bits_match_canonical_reference(self):
        candidates = ["", "A", "A=", "AA=", "AAAA=", "====", "=YWJ", "Y=Jj", "AA=A", "AAA=AAAA", "YR==", "YWJ=", "YQ==", "YWI=", "YWJj"]
        candidates += [char + "A==" for char in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"]
        candidates += ["Y" + char + "==" for char in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"]
        candidates += ["YW" + char + "=" for char in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"]
        candidates += ["".join(chars) for chars in itertools.product("A+/=_", repeat=4)]
        for text in candidates:
            try:
                decoded = base64.b64decode(text, validate=True)
                valid = bool(decoded) and base64.b64encode(decoded).decode() == text
            except ValueError:
                valid = False
            args = dict(VALID_ARGS["add_draft_attachment"], contentBytes=text)
            flow = draft_flow("add_draft_attachment", args, upstream("add_draft_attachment", args) if valid else [])
            self.assertEqual(flow.run()["statusCode"], 200 if valid else 400, text)
            self.assertEqual(len(flow.calls), 2 if valid else 0, text)

    def test_base64_alphabet_whitespace_unicode_urlsafe_and_chunk_edges(self):
        for char in [chr(cp) for cp in range(128)] + ["é", "\u200b", "\u2028", "\ufeff", "☃", "😀", "\ud800"]:
            allowed = char in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
            args = dict(VALID_ARGS["add_draft_attachment"], contentBytes="AA" + char + "A")
            flow = draft_flow("add_draft_attachment", args, upstream("add_draft_attachment", args) if allowed else [])
            self.assertEqual(flow.run()["statusCode"], 200 if allowed else 400, repr(char))
        for text in (None, True, 123, [], {}, "YWJj\n", "YWJj\r\n", "YQ== ", "YQ==\u2028", "YW_j", "YW-j", "A" * (MAX_ENCODED + 1)):
            flow = draft_flow("add_draft_attachment", dict(VALID_ARGS["add_draft_attachment"], contentBytes=text), [])
            self.assert_error(flow)
            if not isinstance(text, str) or len(text) > MAX_ENCODED:
                self.assertEqual(flow.status["Select_add_draft_attachment_alphabet"], "Skipped")
        for offset in (8191, 8192, 8193, 16383):
            text = "A" * offset + "_" + "A" * (16384 - offset - 1)
            self.assert_error(draft_flow("add_draft_attachment", dict(VALID_ARGS["add_draft_attachment"], contentBytes=text), []))

    def test_exact_two_mib_raw_and_encoded_overages(self):
        for size in (1, 2, 3, 2097150, 2097151, MAX_BYTES, MAX_BYTES + 1, MAX_BYTES + 2, MAX_BYTES + 3):
            text = base64.b64encode(b"\xff" * size).decode()
            args = dict(VALID_ARGS["add_draft_attachment"], contentBytes=text)
            if size <= MAX_BYTES:
                flow = draft_flow("add_draft_attachment", args)
                self.assertEqual(flow.run()["body"]["data"]["size"], size)
                self.assertEqual(flow.bodies["Compose_add_draft_attachment_size"], size)
                self.assertEqual(json.loads(flow.calls[1]["Body"])["contentBytes"], text)
                self.assertTrue(all(flow.bodies["Select_add_draft_attachment_alphabet"]))
            else:
                self.assert_error(draft_flow("add_draft_attachment", args, []))
                if size <= MAX_BYTES + 1:
                    self.assertLessEqual(len(text), MAX_ENCODED)
        # Two extra raw bytes are possible at an encoded ceiling only when the
        # byte cap mod 3 == 1; here the cap mod 3 == 2 permits one extra byte.
        self.assertEqual(MAX_BYTES % 3, 2)

    def test_preflight_wrong_id_non_draft_malformed_or_failure_prevents_post(self):
        for raw in ({"id": "wrong-id", "isDraft": True}, {"id": "draft-id", "isDraft": False}):
            flow = draft_flow("add_draft_attachment", responses=[raw])
            self.assert_error(flow, status=409, calls=1, code="DRAFT_NOT_VERIFIED")
            self.assertEqual([call["Method"] for call in flow.calls], ["GET"])
        for raw in (None, [], {}, "wrong", {"id": "draft-id"}, {"id": "draft-id", "isDraft": 1}, {"id": "draft-id", "isDraft": "true"}, {"id": "bad id", "isDraft": True}, {"id": ".", "isDraft": True}, {"id": "x" * 2049, "isDraft": True}, ValueError(CANARY), TimeoutError(CANARY)):
            flow = draft_flow("add_draft_attachment", responses=[raw])
            self.assert_error(flow, status=502, calls=1)
            self.assertEqual([call["Method"] for call in flow.calls], ["GET"])

    def test_preflight_id_case_only_mismatch_prevents_post(self):
        requested = VALID_ARGS["add_draft_attachment"]["draftId"]
        returned = requested.upper()
        self.assertNotEqual(requested, returned)
        self.assertEqual(requested.lower(), returned.lower())
        flow = draft_flow("add_draft_attachment", responses=[{"id": returned, "isDraft": True}])
        self.assert_error(flow, status=409, calls=1, code="DRAFT_NOT_VERIFIED")
        self.assertEqual([call["Method"] for call in flow.calls], ["GET"])
        self.assertNotIn("Graph_add_draft_attachment", flow.status)

    def test_create_reply_malformed_results_and_post_failures_are_ambiguous(self):
        malformed = [None, [], {}, "wrong", {"id": "draft-id"}, {"isDraft": True}, {"id": "draft-id", "isDraft": False}, {"id": "draft-id", "isDraft": 1}, {"id": "draft-id", "isDraft": "true"}]
        malformed += [{"id": value, "isDraft": True, "body": CANARY} for value in ("", ".", "..", "bad id", "x\u0085y", "x\ud800y", "x" * 2049, None, True, 123, [], {})]
        for operation in ("create_draft", "create_reply_draft"):
            for raw in malformed + [ValueError(CANARY), TimeoutError(CANARY)]:
                self.assert_error(draft_flow(operation, responses=[raw]), status=502, calls=1, ambiguous=True)

    def test_attachment_result_id_name_size_and_post_failures_are_ambiguous(self):
        malformed = [None, [], {}, "wrong", {"id": "attachment-id", "name": "sample.pdf"}]
        malformed += [{"id": value, "name": "sample.pdf", "size": 3} for value in ("", ".", "..", "bad id", "x\u0085y", "x\ud800y", "x" * 2049, None, True, 123, [], {})]
        malformed += [{"id": "attachment-id", "name": name, "size": size} for name, size in (("wrong.pdf", 3), ("sample.pdf", 2), ("sample.pdf", 4), ("sample.pdf", 0), ("sample.pdf", MAX_BYTES + 1), ("sample.pdf", True), ("sample.pdf", "3"), ("sample.pdf", 3.5), (None, 3))]
        for raw in malformed + [ValueError(CANARY), TimeoutError(CANARY)]:
            self.assert_error(draft_flow("add_draft_attachment", responses=[{"id": "draft-id", "isDraft": True}, raw]), status=502, calls=2, ambiguous=True)

    def test_attachment_result_name_case_only_mismatch_is_ambiguous(self):
        requested = VALID_ARGS["add_draft_attachment"]["name"]
        returned = requested.upper()
        self.assertNotEqual(requested, returned)
        self.assertEqual(requested.lower(), returned.lower())
        responses = [{"id": "draft-id", "isDraft": True}, {"id": "attachment-id", "name": returned, "size": 3}]
        flow = draft_flow("add_draft_attachment", responses=responses)
        self.assert_error(flow, status=502, calls=2, ambiguous=True)
        self.assertEqual([call["Method"] for call in flow.calls], ["GET", "POST"])
        self.assertEqual(flow.status["Validate_add_draft_attachment_result_identity"], "Failed")
        self.assertEqual(flow.status["Respond_add_draft_attachment"], "Skipped")

    def test_all_fallible_data_actions_have_executable_sanitized_failure_paths(self):
        class FailingFlow(Flow):
            def execute(self, name, action):
                if name == self.fail_name:
                    raise self.fail_exception
                return super().execute(name, action)
        for operation in OPERATIONS:
            actions = named_actions(operation)
            for name, action in actions.items():
                if action["type"] not in {"ParseJson", "Select", "Compose", "OpenApiConnection"}:
                    continue
                for exception in (ValueError(CANARY), TimeoutError(CANARY)):
                    flow = draft_flow(operation)
                    flow.__class__ = FailingFlow
                    flow.fail_name, flow.fail_exception = name, exception
                    result = flow.run()
                    self.assertGreaterEqual(result["statusCode"], 400, name)
                    self.assertNotIn(CANARY, json.dumps(result))
                    if name == f"Graph_{operation}" or "_result" in name:
                        self.assertEqual(result["body"], AMBIGUOUS_ERROR, name)
                    else:
                        self.assertFalse(any(call["Method"] == "POST" for call in flow.calls), name)
                    self.assertEqual(set(result["body"]), {"ok", "error"})

    def test_gateway_guard_stays_active_for_every_write(self):
        for operation in OPERATIONS:
            for configured, supplied in (("", ""), ("synthetic-test-only-key", None), ("synthetic-test-only-key", "wrong")):
                flow = draft_flow(operation, responses=[])
                flow.parameters["McpGatewayKey"] = configured
                flow.headers = {} if supplied is None else {"X-MCP-Gateway-Key": supplied}
                self.assertIsNone(flow.run())
                self.assertEqual(flow.calls, [])


if __name__ == "__main__":
    unittest.main()
