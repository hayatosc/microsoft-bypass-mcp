"""Independent offline contracts for the expanded canonical read-only flow.

These tests execute synthetic requests against definition.json, not a parallel
flow or mocked generator. They do not establish connector import/runtime behavior.
"""

import base64
from copy import deepcopy
import json
import re
import unittest
from urllib.parse import parse_qs, quote, urlsplit

from test_attachment_flow import (
    BASELINE_SOURCE, DESTINATION, FORBIDDEN_ID_CODEPOINTS, GRAPH_ROOT,
    MAX_BASE64, MAX_BYTES, REQUEST_ID, ROOT, SOURCE, Expressions, Flow,
    actions_in, validate, walk,
)

OPERATIONS = {
    "list_messages", "search_messages", "get_message", "list_mail_folders",
    "get_conversation", "list_attachments", "get_attachment",
    "onedrive_search_files", "onedrive_list_folder", "onedrive_get_metadata",
    "onedrive_get_content",
}
TOOLS = {
    "outlook_list_messages", "outlook_search_messages", "outlook_get_message",
    "outlook_list_mail_folders", "outlook_get_conversation",
    "outlook_list_attachments", "outlook_inspect_attachment", "outlook_read_attachment",
    "onedrive_search_files", "onedrive_list_folder", "onedrive_get_metadata",
    "onedrive_inspect_file", "onedrive_read_file",
}
SUMMARY = "id,subject,from,receivedDateTime,sentDateTime,parentFolderId,conversationId,hasAttachments,importance,isRead,bodyPreview"
DETAIL = SUMMARY + ",toRecipients,ccRecipients,body"
FOLDERS = "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount"
NATIVE_FIELDS = {"Id", "Name", "Size", "MediaType", "IsFolder", "LastModified", "ETag"}
CASES = SOURCE["actions"]["スイッチ"]["cases"]
CANARY = "synthetic-private-upstream-canary"
VALID_ARGS = {
    "list_messages": {"top": 2},
    "search_messages": {"query": "synthetic", "top": 2},
    "get_message": {"messageId": "message-id"},
    "list_mail_folders": {"top": 2},
    "get_conversation": {"conversationId": "conversation-id", "top": 2},
    "list_attachments": {"messageId": "message-id", "top": 2, "skip": 0},
    "get_attachment": {"messageId": "message-id", "attachmentId": "attachment-id"},
    "onedrive_search_files": {"query": "synthetic", "top": 2},
    "onedrive_list_folder": {"top": 2},
    "onedrive_get_metadata": {"fileId": "file-id"},
    "onedrive_get_content": {"fileId": "file-id"},
}


def native_metadata(**changes):
    item = {"Id": "file-id", "Name": "sample.pdf", "Size": 3,
            "MediaType": "application/pdf", "IsFolder": False,
            "LastModified": None, "ETag": None}
    item.update(changes)
    return item


def binary(content="YWJj", content_type="application/pdf"):
    return {"$content-type": content_type, "$content": content}


def configured_flow(operation, args=None, responses=()):
    flow = Flow(operation, deepcopy(VALID_ARGS[operation] if args is None else args), responses)
    # Explicitly synthetic local bindings; no guessed live findMode or root ID.
    flow.parameters.update(OneDriveSearchRootId="synthetic-root", OneDriveSearchMode="synthetic-mode")
    return flow


def action_maps(actions):
    yield actions
    for action in actions.values():
        if "actions" in action:
            yield from action_maps(action["actions"])
        if "else" in action:
            yield from action_maps(action["else"]["actions"])
        if "cases" in action:
            for case in action["cases"].values():
                yield from action_maps(case["actions"])
            yield from action_maps(action["default"]["actions"])


def all_named_actions():
    return {name: action for group in action_maps(SOURCE["actions"]) for name, action in group.items()}


class ExpandedStaticContracts(unittest.TestCase):
    def test_exact_eleven_operations_and_thirteen_fixed_tools(self):
        self.assertEqual(len(OPERATIONS), 11)
        self.assertEqual(set(CASES), OPERATIONS)
        self.assertEqual(set(SOURCE["triggers"]["manual"]["inputs"]["schema"]["properties"]["operation"]["enum"]), OPERATIONS)
        source = (ROOT / "src/features/outlook/server.ts").read_text()
        surface = source.split("export const TOOL_NAMES = [", 1)[1].split("] as const", 1)[0]
        self.assertEqual(set(re.findall(r"'([^']+)'", surface)), TOOLS)
        self.assertEqual(len(TOOLS), 13)
        # No generic URL, method, query-option or connector controls in any case.
        for case in CASES.values():
            operation = case["case"]
            schema = case["actions"][f"Validate_{operation}_args"]["inputs"]["schema"]
            self.assertFalse(schema["additionalProperties"])
            self.assertTrue(set(schema["properties"]).isdisjoint({"url", "uri", "method", "body", "nextLink", "operationId", "findMode", "rootId"}))

    def test_json_has_no_duplicate_keys_and_all_dependencies_are_siblings(self):
        def unique_pairs(pairs):
            result = {}
            for name, value in pairs:
                self.assertNotIn(name, result)
                result[name] = value
            return result
        json.loads(DESTINATION.read_text(), object_pairs_hook=unique_pairs)
        names = []
        for group in action_maps(SOURCE["actions"]):
            for name, action in group.items():
                names.append(name)
                dependencies = action.get("runAfter", {})
                self.assertTrue(set(dependencies) <= set(group), (name, dependencies, set(group)))
                self.assertNotIn(name, dependencies)
                for statuses in dependencies.values():
                    self.assertTrue(set(statuses) <= {"Succeeded", "Failed", "TimedOut", "Skipped"})
        self.assertEqual(len(names), len(set(names)))

    def test_graph_and_native_connector_exact_read_allowlists(self):
        native = {
            "OneDrive_find_files": ("FindFiles", {"query", "id", "findMode", "maxFileCount"}),
            "OneDrive_list_root": ("ListRootFolder", set()),
            "OneDrive_list_folder_v2": ("ListFolderV2", {"id"}),
            "OneDrive_get_metadata": ("GetFileMetadata", {"id"}),
            "OneDrive_content_metadata": ("GetFileMetadata", {"id"}),
            "OneDrive_get_content": ("GetFileContent", {"id", "inferContentType"}),
        }
        graph_names = {"Graph_list_messages", "Graph_search_messages", "Graph_get_message", "Graph_list_mail_folders", "Graph_get_conversation", "Graph_list_attachments", "Graph_attachment_metadata", "Graph_attachment_content"}
        actual = {name: action for name, action in all_named_actions().items() if action["type"] == "OpenApiConnection"}
        self.assertEqual(set(actual), set(native) | graph_names)
        for name, action in actual.items():
            inputs = action["inputs"]
            self.assertEqual(inputs["authentication"], "@parameters('$authentication')")
            self.assertEqual(inputs["retryPolicy"], {"type": "none"})
            if name in native:
                operation, fields = native[name]
                self.assertEqual(inputs["host"], {"apiId": "/providers/Microsoft.PowerApps/apis/shared_onedriveforbusiness", "connectionName": "shared_onedriveforbusiness", "operationId": operation})
                self.assertEqual(set(inputs["parameters"]), fields)
                self.assertNotIn("Uri", inputs["parameters"])
            else:
                self.assertEqual(inputs["host"], {"apiId": "/providers/Microsoft.PowerApps/apis/shared_office365", "connectionName": "shared_office365", "operationId": "HttpRequest"})
                self.assertEqual(inputs["parameters"]["Method"], "GET")
                self.assertTrue(inputs["parameters"]["Uri"].startswith("@concat('" + GRAPH_ROOT))
                self.assertNotRegex(inputs["parameters"]["Uri"], r"(?i)(?:nextlink|/drive|/sites|/users|/groups|\$expand|\$value)")
        self.assertIs(actual["OneDrive_get_content"]["inputs"]["parameters"]["inferContentType"], True)

    def test_expanded_mail_uri_expressions_are_exactly_pinned(self):
        # Independent expectations prevent regeneration from silently adding
        # a generic route, arbitrary query option or nextLink fetch.
        expected = {
            "Graph_list_messages": "@concat('https://graph.microsoft.com/v1.0/me/', outputs('Compose_list_messages_route'), '?$top=', string(body('Validate_list_messages_args')?['top']), '&$skip=', string(coalesce(body('Validate_list_messages_args')?['skip'], 0)), '&$select=" + SUMMARY + "', if(empty(outputs('Compose_list_messages_filter')), '&$orderby=receivedDateTime%20desc', concat('&$filter=', uriComponent(outputs('Compose_list_messages_filter')))))",
            "Graph_search_messages": "@concat('https://graph.microsoft.com/v1.0/me/', outputs('Compose_search_messages_route'), '?$search=', uriComponent(concat('\"', trim(body('Validate_search_messages_args')?['query']), '\"')), '&$top=', string(body('Validate_search_messages_args')?['top']), '&$select=" + SUMMARY + "')",
            "Graph_get_message": "@concat('https://graph.microsoft.com/v1.0/me/messages/', uriComponent(body('Validate_get_message_args')?['messageId']), '?$select=" + DETAIL + "')",
            "Graph_list_mail_folders": "@concat('https://graph.microsoft.com/v1.0/me/mailFolders?$top=', string(body('Validate_list_mail_folders_args')?['top']), '&$select=" + FOLDERS + "')",
            "Graph_get_conversation": "@concat('https://graph.microsoft.com/v1.0/me/messages?$top=', string(body('Validate_get_conversation_args')?['top']), '&$skip=', string(coalesce(body('Validate_get_conversation_args')?['skip'], 0)), '&$select=" + DETAIL + "&$filter=', uriComponent(concat('conversationId eq ''', replace(body('Validate_get_conversation_args')?['conversationId'], '''', ''''''), '''')))",
        }
        actions = all_named_actions()
        for name, expression in expected.items():
            self.assertEqual(actions[name]["inputs"]["parameters"]["Uri"], expression)
        for operation in ("list_messages", "search_messages"):
            validate_name = "Validate_" + operation + "_args"
            expected_route = (
                "@if(not(empty(body('" + validate_name + "')?['folderId'])), "
                "concat('mailFolders/', uriComponent(coalesce(body('" + validate_name + "')?['folderId'], 'unused')), '/messages'), "
                "if(equals(coalesce(body('" + validate_name + "')?['mailbox'], 'inbox'), 'sent'), 'mailFolders/sentitems/messages', "
                "if(equals(coalesce(body('" + validate_name + "')?['mailbox'], 'inbox'), 'all'), 'messages', 'mailFolders/inbox/messages')))"
            )
            self.assertEqual(actions["Compose_" + operation + "_route"]["inputs"], expected_route)

    def test_security_settings_cover_all_supported_data_actions(self):
        allowed = {"ParseJson", "Response", "Select", "Compose", "OpenApiConnection", "If"}
        for action in actions_in(CASES):
            kind = action["type"]
            self.assertIn(kind, allowed)
            if kind == "If":
                self.assertNotIn("runtimeConfiguration", action)
                continue
            # Official Logic Apps secure-data contract: Compose, Parse JSON,
            # and Response use secure inputs, which also hide their outputs.
            properties = ["inputs"] if kind in {"Compose", "ParseJson", "Response"} else ["inputs", "outputs"]
            self.assertEqual(action["runtimeConfiguration"]["secureData"]["properties"], properties)
            if kind == "Response":
                self.assertEqual(action["inputs"]["headers"]["Cache-Control"], "no-store")
                if action["inputs"]["statusCode"] >= 400:
                    self.assertEqual(set(action["inputs"]["body"]), {"ok", "error"})
                    self.assertFalse(action["inputs"]["body"]["ok"])
                    self.assertNotIn("@", json.dumps(action["inputs"]["body"]))
        for group in action_maps(SOURCE["actions"]):
            for name, action in group.items():
                if action["type"] in {"OpenApiConnection", "ParseJson", "Select"}:
                    handlers = [candidate for candidate in group.values() if candidate.get("runAfter", {}).get(name) == ["Failed", "TimedOut"]]
                    self.assertEqual(len(handlers), 1, name)
                    self.assertEqual(handlers[0]["type"], "Response")

    def test_no_persistence_logging_or_private_resources(self):
        for node in walk(SOURCE):
            if isinstance(node, dict):
                self.assertTrue(set(node).isdisjoint({"trackedProperties", "connectionReferences", "connectionId", "callbackUrl", "tenantId", "subscriptionId"}))
                if "metadata" in node:
                    self.assertEqual(set(node), {"metadata", "contentBytes"})
                    self.assertEqual(set(node["metadata"]), NATIVE_FIELDS)
        for action in actions_in(CASES):
            self.assertNotIn(action["type"], {"Http", "ApiConnectionWebhook", "AppendToArrayVariable", "SetVariable", "InitializeVariable"})
        for key in ("$authentication", "$connections", "McpGatewayKey"):
            self.assertEqual(SOURCE["parameters"][key], BASELINE_SOURCE["parameters"][key])
        self.assertEqual(set(SOURCE["parameters"]) - set(BASELINE_SOURCE["parameters"]), {"OneDriveSearchRootId", "OneDriveSearchMode"})
        for key in ("OneDriveSearchRootId", "OneDriveSearchMode"):
            self.assertEqual(SOURCE["parameters"][key], {"type": "String", "defaultValue": ""})

    def test_bounded_native_content_and_separate_preflight_gates(self):
        actions = all_named_actions()
        self.assertEqual(actions["Validate_onedrive_binary"]["inputs"]["content"], {
            "$content-type": "@body('OneDrive_get_content')?['$content-type']",
            "$content": "@body('OneDrive_get_content')?['$content']",
        })
        self.assertEqual(actions["Validate_onedrive_binary"]["inputs"]["schema"]["properties"]["$content"]["maxLength"], MAX_BASE64)
        self.assertEqual(actions["Select_onedrive_content_alphabet"]["inputs"]["from"], "@chunk(body('Validate_onedrive_binary')?['$content'], 8192)")
        self.assertNotIn("body(", actions["Select_onedrive_content_alphabet"]["inputs"]["select"])
        self.assertEqual(actions["OneDrive_get_content"]["runAfter"], {})
        identity = actions["Check_onedrive_content_identity"]
        size = identity["actions"]["Check_onedrive_metadata_size"]
        type_gate = size["actions"]["Check_onedrive_content_type"]
        self.assertIn("OneDrive_get_content", type_gate["actions"])
        for gate in (identity, size, type_gate):
            self.assertEqual(gate["type"], "If")
            self.assertNotIn("OneDrive_get_content", gate["else"]["actions"])


class InterpreterContracts(unittest.TestCase):
    def test_only_known_types_functions_and_eager_safe_if_are_supported(self):
        self.assertEqual(Expressions(None).resolve("@if(false, uriComponent(coalesce(null, 'unused')), 'safe')"), "safe")
        self.assertEqual(Expressions(None).resolve("@if(true, 'safe', uriComponent(coalesce(null, 'unused')))"), "safe")
        self.assertEqual(Expressions(None).resolve("@string(false)"), "false")
        for value in (None, "value"):
            validate(value, {"type": ["string", "null"], "maxLength": 5})
        for value in (True, 123, {}, []):
            with self.assertRaises(ValueError):
                validate(value, {"type": ["string", "null"]})
        with self.assertRaises(AssertionError):
            Expressions(None).resolve("@notAnImplementedFunction('x')")
        with self.assertRaises(AssertionError):
            validate("x", {"type": "string", "unknownKeyword": True})

    def test_ticks_validate_calendar_values_and_preserve_fractional_ticks(self):
        one = Expressions(None).resolve("@ticks('2026-01-01T00:00:00.0000001Z')")
        two = Expressions(None).resolve("@ticks('2026-01-01T00:00:00.0000002Z')")
        self.assertEqual(two - one, 1)
        with self.assertRaises(ValueError):
            Expressions(None).resolve("@ticks('2026-02-29T00:00:00Z')")


class ExpandedExecutionContracts(unittest.TestCase):
    def assert_error(self, flow, status, *, calls=0, code=None):
        response = flow.run()
        self.assertEqual(response["statusCode"], status)
        self.assertEqual(len(flow.calls), calls)
        self.assertEqual(set(response["body"]), {"ok", "error"})
        self.assertFalse(response["body"]["ok"])
        if code:
            self.assertEqual(response["body"]["error"]["code"], code)
        self.assertNotIn(CANARY, json.dumps(response))
        return response

    def test_every_operation_rejects_extra_controls_invalid_envelopes_and_ids(self):
        for operation, args in VALID_ARGS.items():
            for controls in ({"url": "https://outside.invalid/"}, {"method": "POST"}, {"nextLink": "https://graph.microsoft.com/v1.0/me/messages"}, {"operationId": "DeleteFile"}):
                with self.subTest(operation=operation, controls=controls):
                    self.assert_error(configured_flow(operation, dict(args, **controls)), 400)
            for request_id in ("legacy-id", None, True, REQUEST_ID[:14] + "5" + REQUEST_ID[15:]):
                flow = configured_flow(operation)
                flow.request["requestId"] = request_id
                self.assert_error(flow, 400)
            for key in args:
                if key.endswith("Id"):
                    for bad in (".", "..", "bad id", "bad\x00id", "bad\u0085id", "x" * 2049, None):
                        self.assert_error(configured_flow(operation, dict(args, **{key: bad})), 400)

    def test_new_required_and_optional_ids_reject_all_forbidden_codepoints(self):
        targets = [("list_messages", "folderId"), ("search_messages", "folderId"), ("get_message", "messageId"), ("get_conversation", "conversationId"), ("onedrive_list_folder", "folderId"), ("onedrive_get_metadata", "fileId"), ("onedrive_get_content", "fileId")]
        invalid = ["x" + chr(cp) + "y" for cp in FORBIDDEN_ID_CODEPOINTS]
        invalid += [".", "..", "x\ud800y", "x\udfffy", "x" * 2049, "\U0001f600" * 1025]
        for operation, key in targets:
            for value in invalid:
                with self.subTest(operation=operation, key=key, value=repr(value)):
                    self.assert_error(configured_flow(operation, dict(VALID_ARGS[operation], **{key: value})), 400)
        for value in ("日本語-é-\u200b", "%00%1F%20%C2%A0%EF%BB%BF", "x" * 2048, "\U0001f600" * 1024):
            flow = configured_flow("get_message", {"messageId": value}, [{"id": value}])
            self.assertEqual(flow.run()["statusCode"], 200)
            self.assertIn("messages/" + quote(value, safe="~()*!.'-_") + "?$select=", flow.calls[0]["Uri"])

    def test_mailbox_scopes_and_folder_override_use_only_fixed_routes(self):
        scopes = [({}, "mailFolders/inbox/messages"), ({"mailbox": "inbox"}, "mailFolders/inbox/messages"), ({"mailbox": "sent"}, "mailFolders/sentitems/messages"), ({"mailbox": "all"}, "messages"), ({"mailbox": "all", "folderId": "folder/a?b#c%20"}, "mailFolders/folder%2Fa%3Fb%23c%2520/messages")]
        for operation in ("list_messages", "search_messages"):
            for extra, path in scopes:
                flow = configured_flow(operation, dict(VALID_ARGS[operation], **extra), [{"value": []}])
                self.assertEqual(flow.run()["statusCode"], 200)
                url = urlsplit(flow.calls[0]["Uri"])
                self.assertEqual(url.scheme + "://" + url.netloc + url.path, GRAPH_ROOT + path)
                params = parse_qs(url.query)
                self.assertEqual(params["$top"], ["2"])
                self.assertEqual(params["$select"], [SUMMARY])
                if operation == "list_messages":
                    self.assertEqual(params, {"$top": ["2"], "$skip": ["0"], "$select": [SUMMARY], "$orderby": ["receivedDateTime desc"]})
                else:
                    self.assertEqual(set(params), {"$top", "$select", "$search"})
        for mailbox in ("sentitems", "users/other", None, 3):
            self.assert_error(configured_flow("list_messages", {"top": 1, "mailbox": mailbox}), 400)

    def test_false_boolean_filters_and_all_combinations_are_preserved(self):
        cases = [({}, None), ({"isRead": False}, "isRead eq false"), ({"isRead": True}, "isRead eq true"), ({"hasAttachments": False}, "hasAttachments eq false"), ({"isRead": False, "hasAttachments": False}, "isRead eq false and hasAttachments eq false"), ({"isRead": True, "hasAttachments": False, "receivedAfter": "2026-01-01T00:00:00Z", "receivedBefore": "2026-02-01T00:00:00Z"}, "isRead eq true and hasAttachments eq false and receivedDateTime ge 2026-01-01T00:00:00Z and receivedDateTime lt 2026-02-01T00:00:00Z")]
        for filters, expected in cases:
            flow = configured_flow("list_messages", {"top": 2, "filters": filters}, [{"value": []}])
            self.assertEqual(flow.run()["statusCode"], 200)
            params = parse_qs(urlsplit(flow.calls[0]["Uri"]).query)
            self.assertEqual(params.get("$filter"), [expected] if expected is not None else None)
            self.assertEqual(params.get("$orderby"), None if expected else ["receivedDateTime desc"])
        for filters in ({"isRead": "false"}, {"isRead": 0}, {"hasAttachments": None}, {"arbitrary": "x"}, []):
            self.assert_error(configured_flow("list_messages", {"top": 2, "filters": filters}), 400)

    def test_dates_are_calendar_parsed_and_rejected_before_graph(self):
        for date in ("2026-02-29T00:00:00Z", "2026-13-01T00:00:00Z", "2026-01-01T24:00:00Z", "2026-01-01T00:00:00Z or true", "2026-01-01T00:00:00+01:00", "2026-01-01T00:00:00'Z", "2026-01-01T00:00:00Z\n", "x" * 20, None):
            for field in ("receivedAfter", "receivedBefore"):
                self.assert_error(configured_flow("list_messages", {"top": 2, "filters": {field: date}}), 400)
        for after, before in (("2026-02-01T00:00:00Z", "2026-01-01T00:00:00Z"), ("2026-01-01T00:00:00Z", "2026-01-01T00:00:00Z")):
            self.assert_error(configured_flow("list_messages", {"top": 2, "filters": {"receivedAfter": after, "receivedBefore": before}}), 400)
        for filters in ({}, {"receivedAfter": "2024-02-29T00:00:00Z"}, {"receivedBefore": "2026-12-31T23:59:59.1234567Z"}):
            flow = configured_flow("list_messages", {"top": 2, "filters": filters}, [{"value": []}])
            self.assertEqual(flow.run()["statusCode"], 200)

    def test_date_range_comparison_preserves_graph_hundred_nanosecond_precision(self):
        start, end = "2026-01-01T00:00:00.1000001Z", "2026-01-01T00:00:00.1000002Z"
        flow = configured_flow("list_messages", {"top": 2, "filters": {"receivedAfter": start, "receivedBefore": end}}, [{"value": []}])
        self.assertEqual(flow.run()["statusCode"], 200)
        self.assert_error(configured_flow("list_messages", {"top": 2, "filters": {"receivedAfter": end, "receivedBefore": start}}), 400)
        self.assert_error(configured_flow("list_messages", {"top": 2, "filters": {"receivedAfter": "2026-01-01T00:00:00.10000001Z"}}), 400)

    def test_search_encodes_query_without_adding_query_parameters(self):
        query = '  quoted"&$expand=attachments/../?#  '
        flow = configured_flow("search_messages", {"query": query, "top": 3, "mailbox": "all"}, [{"value": []}])
        self.assertEqual(flow.run()["statusCode"], 200)
        params = parse_qs(urlsplit(flow.calls[0]["Uri"]).query)
        self.assertEqual(params, {"$search": ['"' + query.strip() + '"'], "$top": ["3"], "$select": [SUMMARY]})
        for query in ("", " \n\t ", "x" * 513, None, [], True):
            for operation in ("search_messages", "onedrive_search_files"):
                self.assert_error(configured_flow(operation, {"query": query, "top": 2}), 400)

    def test_mail_and_native_counts_and_offsets_are_typed_and_bounded(self):
        for operation, args in VALID_ARGS.items():
            if "top" in args:
                maximum = 100 if operation.startswith("onedrive_") else 50
                for top in (0, maximum + 1, True, "1", 1.5):
                    self.assert_error(configured_flow(operation, dict(args, top=top)), 400)
        for operation in ("list_messages", "get_conversation", "list_attachments"):
            for skip in (-1, 10001, True, "1", 0.5):
                self.assert_error(configured_flow(operation, dict(VALID_ARGS[operation], skip=skip)), 400)
        for operation in ("list_messages", "get_conversation"):
            flow = configured_flow(operation, dict(VALID_ARGS[operation], skip=10000), [{"value": []}])
            self.assertEqual(flow.run()["statusCode"], 200)
            self.assertEqual(parse_qs(urlsplit(flow.calls[0]["Uri"]).query)["$skip"], ["10000"])

    def test_conversation_quotes_are_escaped_and_sent_items_remain_in_scope(self):
        conversation = "same'conversation/&$filter=unrelated"
        messages = [{"id": "inbox-id", "parentFolderId": "inbox", "conversationId": conversation}, {"id": "sent-id", "parentFolderId": "sentitems", "conversationId": conversation}]
        flow = configured_flow("get_conversation", {"conversationId": conversation, "top": 50}, [{"value": messages}])
        response = flow.run()
        self.assertEqual(response["body"]["data"]["value"], messages)
        url = urlsplit(flow.calls[0]["Uri"])
        self.assertEqual(url.path, "/v1.0/me/messages")
        self.assertEqual(parse_qs(url.query), {"$top": ["50"], "$skip": ["0"], "$select": [DETAIL], "$filter": ["conversationId eq 'same''conversation/&$filter=unrelated'"]})
        self.assertEqual(flow.calls[0]["CustomHeader2"], 'Prefer: outlook.body-content-type="text"')

    def test_folder_discovery_and_message_detail_fixed_routes(self):
        flow = configured_flow("list_mail_folders", responses=[{"value": []}])
        self.assertEqual(flow.run()["statusCode"], 200)
        self.assertEqual(flow.calls[0]["Uri"], GRAPH_ROOT + "mailFolders?$top=2&$select=" + FOLDERS)
        flow = configured_flow("get_message", {"messageId": "a/b?c%"}, [{"id": "a/b?c%"}])
        self.assertEqual(flow.run()["statusCode"], 200)
        self.assertEqual(flow.calls[0]["Uri"], GRAPH_ROOT + "messages/a%2Fb%3Fc%25?$select=" + DETAIL)

    def test_nextlinks_remain_data_and_never_trigger_another_network_request(self):
        for operation in ("list_messages", "search_messages", "get_conversation", "list_mail_folders"):
            flow = configured_flow(operation, responses=[{"value": [], "@odata.nextLink": "https://outside.invalid/" + CANARY}])
            self.assertEqual(flow.run()["statusCode"], 200)
            self.assertEqual(len(flow.calls), 1)
            self.assertNotIn(CANARY, flow.calls[0]["Uri"])
        flow = configured_flow("onedrive_list_folder", {"top": 2, "folderId": "folder-id"}, [{"value": [], "nextLink": "https://outside.invalid/" + CANARY}])
        data = flow.run()["body"]["data"]
        self.assertEqual(data, {"value": [], "truncated": True})
        self.assertEqual(len(flow.calls), 1)
        self.assertNotIn(CANARY, json.dumps(data))

    def test_search_binding_is_explicit_and_unset_values_fail_closed(self):
        for root, mode in (("", ""), ("synthetic-root", ""), ("", "synthetic-mode")):
            flow = Flow("onedrive_search_files", {"query": "synthetic", "top": 1})
            flow.parameters.update(OneDriveSearchRootId=root, OneDriveSearchMode=mode)
            self.assert_error(flow, 503, code="ONEDRIVE_SEARCH_NOT_CONFIGURED")
        flow = configured_flow("onedrive_search_files", {"query": "  file phrase  ", "top": 7}, [[]])
        self.assertEqual(flow.run()["statusCode"], 200)
        self.assertEqual(flow.call_operations, ["FindFiles"])
        self.assertEqual(flow.calls, [{"query": "file phrase", "id": "synthetic-root", "findMode": "synthetic-mode", "maxFileCount": 7}])

    def test_empty_native_arrays_and_folder_pages_are_valid(self):
        for operation, args, raw, native_operation in (
            ("onedrive_search_files", {"query": "file", "top": 2}, [], "FindFiles"),
            ("onedrive_list_folder", {"top": 2}, [], "ListRootFolder"),
            ("onedrive_list_folder", {"folderId": "folder-id", "top": 2}, {"value": []}, "ListFolderV2"),
        ):
            flow = configured_flow(operation, args, [raw])
            self.assertEqual(flow.run()["body"]["data"], {"value": [], "truncated": False})
            self.assertEqual(flow.call_operations, [native_operation])

    def test_native_pages_bound_results_and_preserve_truncation(self):
        values = [native_metadata(Id="file-" + str(index), Path=CANARY, FileLocator=CANARY) for index in range(5)]
        for operation, args, raw in (
            ("onedrive_search_files", {"query": "file", "top": 2}, values),
            ("onedrive_list_folder", {"top": 2}, values),
            ("onedrive_list_folder", {"folderId": "folder-id", "top": 2}, {"value": values}),
            ("onedrive_list_folder", {"folderId": "folder-id", "top": 2}, {"value": values[:2], "nextLink": "https://outside.invalid/"}),
        ):
            flow = configured_flow(operation, args, [raw])
            data = flow.run()["body"]["data"]
            self.assertEqual(len(data["value"]), 2)
            self.assertIs(data["truncated"], True)
            self.assertTrue(all(set(item) == NATIVE_FIELDS for item in data["value"]))
            self.assertNotIn(CANARY, json.dumps(data))
            self.assertEqual(len(flow.calls), 1)
        # Hitting search's configured count is conservatively incomplete, even
        # when the connector returns no explicit continuation indicator.
        flow = configured_flow("onedrive_search_files", {"query": "file", "top": 2}, [values[:2]])
        self.assertIs(flow.run()["body"]["data"]["truncated"], True)

    def test_native_metadata_projection_nullable_fields_and_identity(self):
        flow = configured_flow("onedrive_get_metadata", responses=[native_metadata(Path=CANARY, FileLocator=CANARY)])
        data = flow.run()["body"]["data"]
        self.assertEqual(set(data), NATIVE_FIELDS)
        self.assertIsNone(data["LastModified"])
        self.assertNotIn(CANARY, json.dumps(data))
        for changes in ({"Id": "wrong"}, {"Size": -1}, {"Size": True}, {"Size": "3"}, {"LastModified": {}}, {"ETag": 123}):
            self.assert_error(configured_flow("onedrive_get_metadata", responses=[native_metadata(**changes)]), 502, calls=1)

    def test_native_malformed_page_shapes_and_items_fail_safely(self):
        for raw in (None, {}, {"value": []}, "wrong", [native_metadata(Size="3")], [native_metadata(Id="")]):
            self.assert_error(configured_flow("onedrive_search_files", responses=[raw]), 502, calls=1)
        for raw in (None, {}, [], {"value": None}, {"value": [native_metadata(IsFolder="false")]}):
            self.assert_error(configured_flow("onedrive_list_folder", {"folderId": "folder-id", "top": 2}, [raw]), 502, calls=1)

    def test_content_identity_size_folder_and_type_are_checked_before_download(self):
        cases = [({"Id": "wrong"}, 502), ({"Size": MAX_BYTES + 1}, 413), ({"Size": 0}, 413), ({"IsFolder": True}, 413), ({"Name": "file.txt", "MediaType": "text/plain"}, 415), ({"Name": "file.pdf", "MediaType": "text/plain"}, 415), ({"Name": "file.txt", "MediaType": "application/not-a-pdf"}, 415), ({"Name": "file.pdf.exe", "MediaType": "application/pdf"}, 415)]
        for changes, status in cases:
            flow = configured_flow("onedrive_get_content", responses=[native_metadata(**changes)])
            self.assert_error(flow, status, calls=1)
            self.assertEqual(flow.call_operations, ["GetFileMetadata"])

    def test_content_success_uses_native_id_and_sanitized_projection(self):
        flow = configured_flow("onedrive_get_content", {"fileId": "file/a?b#c"}, [native_metadata(Id="file/a?b#c", Path=CANARY), binary()])
        response = flow.run()
        self.assertEqual(response["statusCode"], 200)
        self.assertEqual(flow.call_operations, ["GetFileMetadata", "GetFileContent"])
        self.assertEqual(flow.calls, [{"id": "file/a?b#c"}, {"id": "file/a?b#c", "inferContentType": True}])
        self.assertEqual(response["body"]["data"], {"metadata": native_metadata(Id="file/a?b#c"), "contentBytes": "YWJj"})
        self.assertNotIn(CANARY, json.dumps(response))

    def test_binary_envelope_type_and_exact_native_size_are_rechecked(self):
        invalid = [None, {}, {"$content": "YWJj"}, binary(content_type="text/plain"), binary("YQ=="), binary("YWJjZA=="), {"$content-type": "application/pdf", "$content": []}]
        for raw in invalid:
            self.assert_error(configured_flow("onedrive_get_content", responses=[native_metadata(), raw]), 502, calls=2)
        for text in ("", "A", "A=", "AA=", "AAAA=", "====", "YWJj\n", "YWJj\r\n", "YWJj\u2028", "=YWJ", "Y=Jj", "YW_j", "YW-j", "YW.j", "Y☃Jj"):
            self.assert_error(configured_flow("onedrive_get_content", responses=[native_metadata(), binary(text)]), 502, calls=2)
        for data in (b"x", b"xy", b"xyz"):
            text = base64.b64encode(data).decode()
            flow = configured_flow("onedrive_get_content", responses=[native_metadata(Size=len(data)), binary(text)])
            self.assertEqual(flow.run()["statusCode"], 200)

    def test_binary_projection_discards_extra_native_fields_without_fetching_them(self):
        # The binary connector result is projected into a JSON object before
        # Parse JSON, avoiding binary media-type coercion. Extra native fields
        # must be discarded rather than copied into the Worker transport.
        flow = configured_flow("onedrive_get_content", responses=[native_metadata(), dict(binary(), downloadUrl="https://outside.invalid/" + CANARY, rawError=CANARY)])
        response = flow.run()
        self.assertEqual(response["statusCode"], 200)
        self.assertEqual(response["body"]["data"], {"metadata": native_metadata(), "contentBytes": "YWJj"})
        self.assertEqual(flow.call_operations, ["GetFileMetadata", "GetFileContent"])
        self.assertNotIn(CANARY, json.dumps(response))

    def test_content_limit_and_post_metadata_growth_cannot_bypass_four_mib(self):
        for size, expected in ((MAX_BYTES, 200), (MAX_BYTES + 1, 502), (MAX_BYTES + 2, 502), (MAX_BYTES + 3, 502)):
            content = base64.b64encode(b"\xff" * size).decode()
            flow = configured_flow("onedrive_get_content", responses=[native_metadata(Size=MAX_BYTES), binary(content)])
            self.assertEqual(flow.run()["statusCode"], expected)
            self.assertEqual(len(flow.calls), 2)

    def test_all_upstream_failures_and_timeouts_are_sanitized(self):
        for operation in OPERATIONS - {"list_attachments", "get_attachment"}:
            for error in (ValueError(CANARY), TimeoutError(CANARY)):
                self.assert_error(configured_flow(operation, responses=[error]), 502, calls=1, code="UPSTREAM_ERROR")
        for error in (ValueError(CANARY), TimeoutError(CANARY)):
            self.assert_error(configured_flow("onedrive_get_content", responses=[native_metadata(), error]), 502, calls=2, code="UPSTREAM_ERROR")


if __name__ == "__main__":
    unittest.main()
