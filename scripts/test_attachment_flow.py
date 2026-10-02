"""Offline existing-flow extension contract tests. No mailbox/connector/network access.

The tiny interpreter executes only the checked-in WDL subset against synthetic
Graph responses. It tests dependency paths and expressions, NOT the Microsoft
runtime or package importer. Static checks independently pin every network path.
"""

import base64
from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import itertools
import json
from pathlib import Path
import re
import unittest
from urllib.parse import quote

from build_attachment_flow import ATTACHMENT_OPERATIONS, BASELINE, BASE64_CHUNK, DESTINATION, FORBIDDEN_ID_CODEPOINTS, GRAPH_ROOT, ID_SCHEMA, MAIL_OPERATIONS, MAX_BASE64, MAX_BYTES, build_definition

from build_read_tools_flow import ARGUMENTS, REQUIRED_ARGS
from build_draft_tools_flow import DRAFT_ARGUMENTS, DRAFT_OPERATIONS, DRAFT_REQUIRED_ARGS

ROOT = Path(__file__).resolve().parents[1]
SOURCE = json.loads(DESTINATION.read_text(encoding="utf-8"))
BASELINE_SOURCE = json.loads(BASELINE.read_text(encoding="utf-8"))
ATTACHMENT_CASES = {name: SOURCE["actions"]["スイッチ"]["cases"][name] for name in ATTACHMENT_OPERATIONS}
BASELINE_SHA256 = "688fd5a8a1e83288710e965f9597c33b99a7ff702c89c18ca571e67eb3f075d8"
REQUEST_ID = "12345678-1234-4234-8234-123456789abc"
WORKER_ID_PATTERN = r"^(?!\.{1,2}$)[^\s\u0000-\u001f\u007f]+$(?![\s\S])"


def string_length(value):
    return len(value.encode("utf-16-le", errors="surrogatepass")) // 2


def walk(value):
    yield value
    if isinstance(value, dict):
        for child in value.values():
            yield from walk(child)
    elif isinstance(value, list):
        for child in value:
            yield from walk(child)


def actions_in(value):
    return [node for node in walk(value) if isinstance(node, dict) and "type" in node and ("inputs" in node or "actions" in node or "cases" in node)]


def validate(value, schema):
    """Fail-closed local validator for precisely the JSON Schema subset we use."""
    allowed = {"type", "properties", "required", "additionalProperties", "minimum", "maximum", "minLength", "maxLength", "pattern", "enum", "items", "maxItems"}
    if not set(schema) <= allowed:
        raise AssertionError("The test interpreter does not understand this schema keyword")
    kind = schema.get("type")
    matches = {
        "object": isinstance(value, dict), "array": isinstance(value, list),
        "string": isinstance(value, str), "integer": type(value) is int,
        "boolean": type(value) is bool,
    }
    matches["null"] = value is None
    kinds = kind if isinstance(kind, list) else [kind] if kind else []
    if kinds and not any(matches[candidate] for candidate in kinds):
        raise ValueError("invalid type")
    if isinstance(kind, list):
        kind = next(candidate for candidate in kinds if matches[candidate])
    if "enum" in schema and value not in schema["enum"]:
        raise ValueError("invalid enum")
    if kind == "object":
        props = schema.get("properties", {})
        if not set(schema.get("required", [])) <= set(value):
            raise ValueError("missing required fields")
        if schema.get("additionalProperties") is False and not set(value) <= set(props):
            raise ValueError("additional fields")
        for key in value.keys() & props.keys():
            validate(value[key], props[key])
    if kind == "array":
        if len(value) > schema.get("maxItems", len(value)):
            raise ValueError("too many items")
        for item in value:
            validate(item, schema["items"])
    if kind == "string":
        if not schema.get("minLength", 0) <= string_length(value) <= schema.get("maxLength", string_length(value)):
            raise ValueError("invalid string length")
        if "pattern" in schema and re.search(schema["pattern"], value) is None:
            raise ValueError("invalid string pattern")
    if kind == "integer" and not schema.get("minimum", value) <= value <= schema.get("maximum", value):
        raise ValueError("invalid numeric range")


class Expressions:
    TOKEN = re.compile(r"\s*(?:('(?:''|[^'])*')|([A-Za-z_][A-Za-z0-9_]*)|(-?\d+)|([(),?\[\]]))")

    def __init__(self, engine, item=None):
        self.engine = engine
        self.item = item

    def resolve(self, value):
        if isinstance(value, dict):
            return {self.resolve(key): self.resolve(child) for key, child in value.items()}
        if isinstance(value, list):
            return [self.resolve(child) for child in value]
        if not isinstance(value, str) or not value.startswith("@"):
            return value
        if value.startswith("@@"):
            return value[1:]
        source = value[2:-1] if value.startswith("@{") and value.endswith("}") else value[1:]
        self.tokens = []
        while source.strip():
            match = self.TOKEN.match(source)
            if not match:
                raise AssertionError("Unknown WDL expression in offline interpreter")
            self.tokens.append(next(group for group in match.groups() if group is not None))
            source = source[match.end():]
        self.position = 0
        result = self.parse()
        if self.position != len(self.tokens):
            raise AssertionError("Unconsumed WDL tokens")
        return result

    def pop(self, expected=None):
        value = self.tokens[self.position]
        self.position += 1
        if expected is not None and value != expected:
            raise AssertionError("Unexpected WDL token")
        return value

    def parse(self):
        token = self.pop()
        if token.startswith("'"):
            result = token[1:-1].replace("''", "'")
        elif re.fullmatch(r"-?\d+", token):
            result = int(token)
        elif token in {"true", "false", "null"}:
            result = {"true": True, "false": False, "null": None}[token]
        else:
            self.pop("(")
            args = []
            if self.tokens[self.position] != ")":
                while True:
                    # Conservatively evaluate all arguments, so validation must
                    # not depend on a null-sensitive branch being short-circuited.
                    args.append(self.parse())
                    if self.tokens[self.position] != ",":
                        break
                    self.pop(",")
            self.pop(")")
            result = self.call(token, args)
        while self.position < len(self.tokens) and self.tokens[self.position] in {"?", "["}:
            if self.tokens[self.position] == "?":
                self.pop("?")
            self.pop("[")
            key = self.pop()[1:-1]
            self.pop("]")
            result = result.get(key) if isinstance(result, dict) else None
        return result

    def call(self, name, args):
        def wdl_string(value):
            if value is None:
                return ""
            if isinstance(value, (dict, list, bool)):
                return json.dumps(value, separators=(",", ":"))
            return str(value)

        def take(value, count):
            if not isinstance(value, (str, list)) or type(count) is not int:
                raise ValueError("invalid take operands")
            return value[:max(count, 0)]

        def add(a, b):
            if type(a) not in (int, float) or type(b) not in (int, float):
                raise ValueError("invalid add operands")
            return a + b

        def starts_with(value, prefix):
            if not isinstance(value, str) or not isinstance(prefix, str):
                raise ValueError("invalid startsWith operands")
            return value.lower().startswith(prefix.lower())

        def ends_with(value, suffix):
            if not isinstance(value, str) or not isinstance(suffix, str):
                raise ValueError("invalid endsWith operands")
            return value.lower().endswith(suffix.lower())

        def index_of(value, part):
            if not isinstance(value, str) or not isinstance(part, str):
                raise ValueError("invalid indexOf operands")
            # Native WDL is case-insensitive. Lowercasing preserves the integer
            # offsets for the ASCII strings used by our base64 guards.
            return value.lower().find(part.lower())

        def last(value):
            if not isinstance(value, (str, list)) or not value:
                raise ValueError("invalid last collection")
            return value[-1]

        def ticks(value):
            # This interpreter supports only the UTC ISO values admitted by
            # these cases, retaining the seventh fractional tick digit.
            match = re.fullmatch(r"(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})(?:\.(\d{1,7}))?Z", value)
            if not match:
                raise ValueError("invalid timestamp")
            parsed = datetime.fromisoformat(match[1]).replace(tzinfo=timezone.utc)
            span = parsed - datetime(1, 1, 1, tzinfo=timezone.utc)
            return (span.days * 86400 + span.seconds) * 10_000_000 + int((match[2] or "").ljust(7, "0"))
        fns = {
            "body": lambda name: self.engine.bodies[name],
            "outputs": lambda name: self.engine.bodies[name],
            "triggerBody": lambda: self.engine.request,
            "triggerOutputs": lambda: {"headers": self.engine.headers},
            "parameters": lambda name: self.engine.parameters[name],
            "item": lambda: self.item,
            "concat": lambda *values: "".join(map(wdl_string, values)),
            "json": json.loads, "addProperty": lambda value, key, child: dict(value, **{key: child}),
            "string": wdl_string, "uriComponent": lambda value: quote(value, safe="~()*!.'-_"),
            "coalesce": lambda *values: next((value for value in values if value is not None), None),
            "equals": lambda a, b: a == b, "not": lambda value: not value,
            "and": lambda *values: all(values), "empty": lambda value: value is None or value == "" or value == [] or value == {},
            "or": lambda *values: any(values), "contains": lambda value, part: part in value,
            "replace": lambda value, old, new: value.replace(old, new), "toLower": lambda value: value.lower(),
            "substring": lambda value, start, length: value[start:start + length],
            "indexOf": index_of,
            "chunk": lambda value, size: [value[index:index + size] for index in range(0, len(value), size)],
            "length": lambda value: string_length(value) if isinstance(value, str) else len(value), "div": lambda a, b: a // b, "mul": lambda a, b: a * b,
            "sub": lambda a, b: a - b, "mod": lambda a, b: a % b, "if": lambda cond, yes, no: yes if cond else no,
            "endsWith": ends_with,
            # Documented WDL functions used by draft guards:
            # https://learn.microsoft.com/azure/logic-apps/workflow-definition-language-functions-reference
            "add": add, "startsWith": starts_with, "last": last,
            "trim": lambda value: value.strip(), "take": take,
            "greater": lambda a, b: a > b, "greaterOrEquals": lambda a, b: a >= b,
            "less": lambda a, b: a < b, "lessOrEquals": lambda a, b: a <= b,
            "ticks": ticks, "first": lambda value: value[0],
            "split": lambda value, separator: value.split(separator),
        }
        if name not in fns:
            raise AssertionError("Unknown WDL function")
        return fns[name](*args)


class Flow:
    def __init__(self, operation, args, graph_responses=(), *, definition=None):
        self.definition = deepcopy(definition or SOURCE)
        self.request = {"operation": operation, "requestId": REQUEST_ID, "args": args}
        self.headers = {"X-MCP-Gateway-Key": "synthetic-test-only-key"}
        self.parameters = {name: deepcopy(value.get("defaultValue")) for name, value in self.definition["parameters"].items()}
        self.parameters["McpGatewayKey"] = "synthetic-test-only-key"
        self.queue = list(graph_responses)
        self.bodies = {}
        self.status = {}
        self.calls = []
        self.call_operations = []
        self.responses = []

    def condition(self, expression):
        name, args = next(iter(expression.items()))
        if name == "and":
            return all(self.condition(child) for child in args)
        values = Expressions(self).resolve(args)
        if name == "equals":
            return values[0] == values[1]
        if name == "lessOrEquals":
            return values[0] <= values[1]
        raise AssertionError("Unknown WDL condition")

    def run_actions(self, actions):
        pending = dict(actions)
        while pending:
            progress = False
            for name, action in list(pending.items()):
                deps = action.get("runAfter", {})
                if not set(deps) <= self.status.keys():
                    continue
                progress = True
                del pending[name]
                if any(self.status[key] not in statuses for key, statuses in deps.items()):
                    self.status[name] = "Skipped"
                    continue
                try:
                    self.execute(name, action)
                    self.status[name] = "Succeeded"
                except ValueError:
                    self.status[name] = "Failed"
                except TimeoutError:
                    self.status[name] = "TimedOut"
            if not progress:
                raise AssertionError("Unresolvable action dependency")

    def execute(self, name, action):
        resolve = Expressions(self).resolve
        kind = action["type"]
        if kind == "ParseJson":
            content = resolve(action["inputs"]["content"])
            validate(content, resolve(action["inputs"]["schema"]))
            self.bodies[name] = content
        elif kind == "OpenApiConnection":
            inputs = resolve(action["inputs"]["parameters"])
            self.calls.append(inputs)
            self.call_operations.append(action["inputs"]["host"]["operationId"])
            if not self.queue:
                raise AssertionError("Unexpected Graph fetch")
            result = self.queue.pop(0)
            if isinstance(result, Exception):
                raise result
            self.bodies[name] = result
        elif kind == "Select":
            self.bodies[name] = [Expressions(self, item).resolve(action["inputs"]["select"]) for item in resolve(action["inputs"]["from"])]
        elif kind == "Compose":
            self.bodies[name] = resolve(action["inputs"])
        elif kind == "Response":
            self.responses.append(resolve(action["inputs"]))
        elif kind == "If":
            self.run_actions(action["actions"] if self.condition(action["expression"]) else action["else"]["actions"])
        elif kind == "Switch":
            key = resolve(action["expression"])
            case = next((case for case in action["cases"].values() if case["case"] == key), action["default"])
            self.run_actions(case["actions"])
        else:
            raise AssertionError("Unsupported action in offline interpreter")

    def run(self, *, allow_no_response=False):
        # Execute actions after HTTP trigger admission. Platform rejection bodies
        # are not reproduced by this interpreter; trigger schema is checked below.
        for guard in self.definition["triggers"]["manual"]["conditions"]:
            if not Expressions(self).resolve(guard["expression"]):
                return None
        self.run_actions(self.definition["actions"])
        if allow_no_response and not self.responses:
            return None
        if len(self.responses) != 1:
            raise AssertionError("Expected exactly one HTTP response")
        return self.responses[0]


def metadata(**updates):
    value = {"@odata.type": "#microsoft.graph.fileAttachment", "id": "attachment-id", "name": "sample.pdf", "contentType": "application/pdf", "size": 3, "isInline": False}
    value.update(updates)
    return value


def get_flow(*responses, **kwargs):
    return Flow("get_attachment", {"messageId": "message-id", "attachmentId": "attachment-id"}, responses, **kwargs)


class StaticContracts(unittest.TestCase):
    def test_power_automate_schema_and_expression_constraints(self):
        for node in walk(SOURCE):
            if isinstance(node, dict):
                self.assertTrue(set(node).isdisjoint({"pattern", "patternProperties"}))
            if isinstance(node, str) and node.startswith("@"):
                self.assertLessEqual(string_length(node), 8192)
        # Cloud-flow limits are 500 actions and 8 nesting levels, not 250.
        # https://learn.microsoft.com/power-automate/limits-and-config#flow-definition-limits
        self.assertLessEqual(len(actions_in(SOURCE)), 500)
        def depth(actions, level=1):
            maximum = level
            for action in actions.values():
                groups = [action["actions"]] if "actions" in action else []
                if "else" in action:
                    groups.append(action["else"]["actions"])
                if "cases" in action:
                    groups.extend(case["actions"] for case in action["cases"].values())
                    groups.append(action["default"]["actions"])
                maximum = max(maximum, *(depth(group, level + 1) for group in groups)) if groups else maximum
            return maximum
        self.assertLessEqual(depth(SOURCE["actions"]), 8)

    def test_format_guards_run_after_typed_bounded_schemas(self):
        for operation in ATTACHMENT_OPERATIONS:
            actions = ATTACHMENT_CASES[operation]["actions"]
            envelope, request_format, args, ids = (f"Validate_{operation}_{suffix}" for suffix in ("request", "request_format", "args", "ids"))
            self.assertEqual(actions[request_format]["runAfter"], {envelope: ["Succeeded"]})
            self.assertEqual(actions[args]["runAfter"], {request_format: ["Succeeded"]})
            self.assertEqual(actions[ids]["runAfter"], {args: ["Succeeded"]})
            graph = "Graph_list_attachments" if operation == "list_attachments" else "Graph_attachment_metadata"
            self.assertEqual(actions[graph]["runAfter"], {ids: ["Succeeded"]})
            for name in (request_format, ids):
                self.assertEqual(actions[name]["inputs"]["schema"], {"type": "boolean", "enum": [True]})
        actions = ATTACHMENT_CASES["list_attachments"]["actions"]
        self.assertEqual(actions["Validate_projected_attachment_list"]["runAfter"], {"Select_attachment_metadata": ["Succeeded"]})
        self.assertEqual(actions["Respond_list_attachments"]["runAfter"], {"Validate_projected_attachment_list": ["Succeeded"]})

    def test_content_alphabet_transforms_only_bounded_chunks(self):
        actions = ATTACHMENT_CASES["get_attachment"]["actions"]["Check_attachment_identity"]["actions"]["Check_attachment_type"]["actions"]["Check_attachment_size"]["actions"]
        self.assertEqual(actions["Select_attachment_content_alphabet"]["runAfter"], {"Validate_attachment_content": ["Succeeded"]})
        self.assertEqual(actions["Select_attachment_content_alphabet"]["inputs"]["from"], f"@chunk(body('Validate_attachment_content')?['contentBytes'], {BASE64_CHUNK})")
        self.assertNotIn("body(", actions["Select_attachment_content_alphabet"]["inputs"]["select"])
        self.assertEqual(actions["Check_content_size_and_identity"]["runAfter"], {"Validate_attachment_content_alphabet": ["Succeeded"]})
        self.assertEqual(actions["Validate_attachment_content"]["inputs"]["schema"]["properties"]["contentBytes"]["maxLength"], MAX_BASE64)
        self.assertLess(BASE64_CHUNK * 12, 131072)

    def test_literal_at_schema_values_and_input_keys_are_escaped(self):
        schemas = [action["inputs"]["schema"] for action in actions_in(ATTACHMENT_CASES) if action["type"] == "ParseJson"]
        serialized_literals = [value for schema in schemas for value in walk(schema) if isinstance(value, str) and value.startswith("@")]
        self.assertEqual(serialized_literals, ["@@odata.type"] * 4)
        for schema in schemas:
            resolved = Expressions(None).resolve(schema)
            for node in walk(resolved):
                if isinstance(node, dict) and "required" in node:
                    self.assertTrue(set(node["required"]) <= node["properties"].keys())
        serialized_keys = [key for node in walk(ATTACHMENT_CASES) if isinstance(node, dict) for key in node if key.startswith("@")]
        self.assertCountEqual(serialized_keys, ["@@odata.type"] * 6 + ["@@odata.nextLink"] * 2)
        value = {"@@odata.type": "@@odata.type", "nested": ["@@", "@@@value"]}
        self.assertEqual(Expressions(None).resolve(value), {"@odata.type": "@odata.type", "nested": ["@", "@@value"]})

    def test_generated_source_is_current_and_baseline_fixture_unchanged(self):
        self.assertEqual(SOURCE, build_definition())
        self.assertEqual(hashlib.sha256(BASELINE.read_bytes()).hexdigest(), BASELINE_SHA256)

    def test_operation_and_argument_schema_parity_with_worker_types(self):
        ts = (ROOT / "src/lib/power-automate.ts").read_text(encoding="utf-8")
        start = ts.index("export interface PowerAutomateOperations")
        end = ts.index("\n}", start)
        entries = re.findall(r"\b(\w+):\s*\{\s*args:\s*\{([^}]+)\}", ts[start:end])
        actual, required = {}, {}
        for name, fields in entries:
            parsed = re.findall(r"(\w+)(\?)?:\s*([^;\n]+)", fields)
            actual[name] = {key: ("string" if value.strip().startswith("'") else "object" if value.strip() == "MailListFilters" else value.strip().split(" | ")[0]) for key, optional, value in parsed}
            required[name] = {key for key, optional, value in parsed if not optional}
        all_arguments = {**ARGUMENTS, **DRAFT_ARGUMENTS}
        all_required = {**REQUIRED_ARGS, **DRAFT_REQUIRED_ARGS}
        expected = {name: {key: {"integer": "number", "array": "string[]"}.get(schema["type"], schema["type"]) for key, schema in fields.items()} for name, fields in all_arguments.items()}
        self.assertEqual(actual, expected)
        self.assertEqual(required, {name: set(keys) for name, keys in all_required.items()})
        trigger = SOURCE["triggers"]["manual"]["inputs"]["schema"]
        self.assertEqual(set(trigger["properties"]["operation"]["enum"]), set(expected))
        cases = SOURCE["actions"]["スイッチ"]["cases"]
        self.assertEqual(set(cases), set(expected))
        for name in expected:
            schema = cases[name]["actions"][f"Validate_{name}_args"]["inputs"]["schema"]
            self.assertEqual(schema["properties"], all_arguments[name])
            self.assertEqual(set(schema["required"]), required[name])
            self.assertFalse(schema["additionalProperties"])
            envelope = cases[name]["actions"][f"Validate_{name}_request"]["inputs"]["schema"]
            self.assertEqual(envelope["properties"]["operation"]["enum"], [name])
            self.assertFalse(envelope["additionalProperties"])

    def test_expansion_preserves_original_attachments_and_all_authentication(self):
        # Mail branches are intentionally v2. Attachment cases stay byte-for-
        # byte equivalent to their original canonical, sorted JSON contracts.
        hashes = {
            "list_attachments": "b47f9ee3526ff045892ede06bd3bc7941a24a1d0d621421b6c46f88156f2f6bc",
            "get_attachment": "ae381560356c8335ebf89337619edd15953020108d157473f2cbdf319a279894",
        }
        for name, expected in hashes.items():
            canonical = json.dumps(ATTACHMENT_CASES[name], sort_keys=True, separators=(",", ":")).encode()
            self.assertEqual(hashlib.sha256(canonical).hexdigest(), expected)
        original_trigger, expanded_trigger = deepcopy(BASELINE_SOURCE["triggers"]), deepcopy(SOURCE["triggers"])
        # Explicit v2 privacy addition. Every other trigger setting, including
        # the gateway-key condition and authentication, remains identical.
        self.assertEqual(expanded_trigger["manual"].pop("runtimeConfiguration"), {"secureData": {"properties": ["outputs"]}})
        del original_trigger["manual"]["inputs"]["schema"]
        del expanded_trigger["manual"]["inputs"]["schema"]
        self.assertEqual(expanded_trigger, original_trigger)
        for name, original in BASELINE_SOURCE["parameters"].items():
            self.assertEqual(SOURCE["parameters"][name], original)
        self.assertEqual(set(SOURCE["actions"]), {"スイッチ"})
        self.assertEqual(SOURCE["actions"]["スイッチ"]["default"], BASELINE_SOURCE["actions"]["スイッチ"]["default"])

    def test_trigger_schema_defers_v2_operation_bounds_to_case_validators(self):
        schema = SOURCE["triggers"]["manual"]["inputs"]["schema"]
        for op, args in [("list_messages", {}), ("search_messages", {"query": "synthetic"}), ("get_message", {}), ("list_attachments", {"messageId": "id", "top": 1, "skip": 0}), ("get_attachment", {"messageId": "id", "attachmentId": "id"})]:
            validate({"operation": op, "requestId": "legacy-non-uuid", "args": args}, schema)
        for request in [None, [], {}, {"operation": "write_message", "requestId": REQUEST_ID, "args": {}}]:
            with self.assertRaises(ValueError):
                validate(request, schema)

    def test_action_names_are_unique_and_flow_source_is_not_duplicated(self):
        names = [name for node in walk(SOURCE["actions"]) if isinstance(node, dict) for name, value in node.items() if isinstance(value, dict) and isinstance(value.get("type"), str) and value["type"] in {"ParseJson", "Response", "OpenApiConnection", "Select", "Compose", "If", "Switch"}]
        self.assertEqual(len(names), len(set(names)))
        self.assertEqual(list((ROOT / "power-automate").glob("*/definition.json")), [DESTINATION])

    def test_attachment_limits_and_id_pattern_match_worker_schema(self):
        ts = (ROOT / "src/features/outlook/attachments/schema.ts").read_text(encoding="utf-8")
        compact = re.sub(r"\s+", "", re.sub(r"(?m)^\s*//.*$", "", ts))
        self.assertIn("MAX_ATTACHMENT_BYTES=4*1024*1024", compact)
        self.assertIn("z.string().min(1).max(2048).regex(/" + WORKER_ID_PATTERN + "/)", compact)
        self.assertIn("limit:z.number().int().min(1).max(50)", compact)
        self.assertIn("offset:z.number().int().min(0).max(10000)", compact)

    def test_exact_uri_expressions_are_pinned(self):
        expected = {
            "Graph_list_attachments": "@concat('https://graph.microsoft.com/v1.0/me/messages/',uriComponent(body('Validate_list_attachments_args')?['messageId']),'/attachments?$select=id,name,contentType,size,isInline&$top=',string(body('Validate_list_attachments_args')?['top']),'&$skip=',string(body('Validate_list_attachments_args')?['skip']))",
            "Graph_attachment_metadata": "@concat('https://graph.microsoft.com/v1.0/me/messages/',uriComponent(body('Validate_get_attachment_args')?['messageId']),'/attachments/',uriComponent(body('Validate_get_attachment_args')?['attachmentId']),'?$select=id,name,contentType,size,isInline')",
            "Graph_attachment_content": "@concat('https://graph.microsoft.com/v1.0/me/messages/',uriComponent(body('Validate_get_attachment_args')?['messageId']),'/attachments/',uriComponent(body('Validate_get_attachment_args')?['attachmentId']))",
        }
        actual = {}
        for node in walk(ATTACHMENT_CASES):
            if isinstance(node, dict):
                for name, value in node.items():
                    if isinstance(value, dict) and value.get("type") == "OpenApiConnection":
                        actual[name] = re.sub(r"\s+", "", value["inputs"]["parameters"]["Uri"])
        self.assertEqual(actual, expected)

    def test_exact_fixed_get_network_allowlist(self):
        read_cases = {name: case for name, case in SOURCE["actions"]["スイッチ"]["cases"].items() if name not in DRAFT_OPERATIONS}
        network = [node for node in actions_in(read_cases) if node["type"] == "OpenApiConnection" and node["inputs"]["host"]["operationId"] == "HttpRequest"]
        self.assertEqual(len(network), 8)
        expected = [
            "", "", "messages/", "messages/", "messages/", "messages/",
            "mailFolders?$top=", "messages?$top=",
        ]
        prefixes = []
        for action in network:
            params = action["inputs"]["parameters"]
            self.assertEqual(params["Method"], "GET")
            self.assertEqual(action["inputs"]["host"], {"apiId": "/providers/Microsoft.PowerApps/apis/shared_office365", "connectionName": "shared_office365", "operationId": "HttpRequest"})
            if "retryPolicy" in action["inputs"]:
                self.assertEqual(action["inputs"]["retryPolicy"], {"type": "none"})
            uri = params["Uri"]
            match = re.search(r"@concat\(\s*'([^']+)'", uri)
            self.assertTrue(match)
            self.assertTrue(match[1].startswith(GRAPH_ROOT))
            prefixes.append(match[1][len(GRAPH_ROOT):])
            self.assertNotIn("nextLink", uri)
            self.assertNotIn("$expand", uri)
            self.assertNotIn("$value", uri)
            for field in ("messageId", "attachmentId"):
                if field in uri:
                    self.assertRegex(uri, rf"uriComponent\((?:body\('[^']+'\)|triggerBody\(\)\?\['args'\])\?\['{field}'\]\)")
        self.assertCountEqual(prefixes, expected)
        allowed_types = {"Request", "ParseJson", "Response", "OpenApiConnection", "Select", "Compose", "If", "Switch"}
        self.assertTrue(all(node["type"] in allowed_types for node in actions_in(SOURCE)))

    def test_no_private_resource_metadata_or_urls(self):
        allowed_urls = {SOURCE["$schema"], GRAPH_ROOT}
        for node in walk(SOURCE):
            if isinstance(node, dict):
                self.assertTrue(set(node).isdisjoint({"connectionReferences", "trackedProperties", "tenantId", "subscriptionId", "callbackUrl", "connectionId"}))
                if "metadata" in node:
                    # This is the new document payload, never flow resource metadata.
                    self.assertEqual(set(node), {"metadata", "contentBytes"})
                    self.assertEqual(set(node["metadata"]), {"Id", "Name", "Size", "MediaType", "IsFolder", "LastModified", "ETag"})
            if isinstance(node, str):
                self.assertNotRegex(node, r"(?i)(?:[?&]sig=|Bearer |SharedAccessSignature|logic\.azure\.com|environment\.api\.powerplatform\.com)")
                for url in re.findall(r"https?://[^\s'\"]+", node):
                    self.assertTrue(any(url.startswith(allowed) for allowed in allowed_urls))
        self.assertEqual(SOURCE["parameters"]["McpGatewayKey"], {"type": "SecureString", "defaultValue": ""})
        for name in ("$authentication", "$connections"):
            self.assertEqual(SOURCE["parameters"][name]["defaultValue"], {})

    def test_attachment_secure_history_settings_and_sanitized_error_envelopes(self):
        for action in actions_in(ATTACHMENT_CASES):
            kind = action["type"]
            if kind in {"If", "Switch"}:
                self.assertNotIn("runtimeConfiguration", action)
                continue
            expected = ["inputs"] if kind in {"ParseJson", "Response"} else ["inputs", "outputs"]
            self.assertEqual(action["runtimeConfiguration"]["secureData"]["properties"], expected)
            if kind == "OpenApiConnection":
                self.assertEqual(action["inputs"]["retryPolicy"], {"type": "none"})
            if kind == "Response":
                inputs = action["inputs"]
                self.assertEqual(inputs["headers"]["Cache-Control"], "no-store")
                if inputs["statusCode"] >= 400:
                    self.assertEqual(set(inputs["body"]), {"ok", "error"})
                    self.assertFalse(inputs["body"]["ok"])
                    self.assertNotIn("@", json.dumps(inputs["body"]))
            if kind in {"OpenApiConnection", "ParseJson", "Select"}:
                # Every fallible data action has a sibling static failure path.
                dependency = next(name for node in walk(SOURCE) if isinstance(node, dict) for name, value in node.items() if value is action)
                handlers = [node for node in actions_in(SOURCE) if node.get("runAfter", {}).get(dependency) == ["Failed", "TimedOut"]]
                self.assertEqual(len(handlers), 1)
                self.assertEqual(handlers[0]["type"], "Response")

    def test_endpoint_outputs_cannot_echo_connector_or_gateway_authentication(self):
        for action in actions_in(SOURCE):
            if action["type"] == "Response":
                body = json.dumps(action["inputs"]["body"])
                for marker in ("parameters(", "triggerOutputs(", "outputs(", "authentication", "McpGatewayKey", "X-MCP-Gateway-Key"):
                    self.assertNotIn(marker, body)
        self.assertNotIn("McpGatewayKey", json.dumps(SOURCE["actions"]))
        self.assertNotIn("X-MCP-Gateway-Key", json.dumps(SOURCE["actions"]))


class ExecutionContracts(unittest.TestCase):
    def test_uuid_v4_format_and_type_are_rejected_before_arguments_or_graph(self):
        invalid = [None, True, 123, [], {}, "", REQUEST_ID[:-1], REQUEST_ID + "0"]
        invalid += [REQUEST_ID[:index] + replacement + REQUEST_ID[index + 1:] for index, replacement in (
            (0, "-"), (0, "g"), (0, "\u0661"), (8, "0"), (13, "0"), (18, "0"), (23, "0"), (14, "5"), (19, "7"), (19, "c"), (35, "\n"),
        )]
        for operation in ATTACHMENT_OPERATIONS:
            for request_id in invalid:
                flow = Flow(operation, {})
                flow.request["requestId"] = request_id
                response = flow.run()
                self.assertEqual(response["statusCode"], 400)
                self.assertEqual(response["body"]["error"]["code"], "INVALID_REQUEST")
                self.assertEqual(flow.calls, [])
        for request_id in (REQUEST_ID, REQUEST_ID.upper(), "ABCDEF01-2345-4678-BABC-DEF012345678"):
            flow = Flow("list_attachments", {"messageId": "id", "top": 1, "skip": 0}, [{"value": []}])
            flow.request["requestId"] = request_id
            self.assertEqual(flow.run()["statusCode"], 200)

    def test_all_forbidden_id_codepoints_and_surrogates_fail_closed(self):
        bad_ids = ["x" + chr(codepoint) + "y" for codepoint in FORBIDDEN_ID_CODEPOINTS]
        bad_ids += [".", "..", "x\ud800y", "x\udfffy"]
        for value in bad_ids:
            targets = [
                Flow("list_attachments", {"messageId": value, "top": 1, "skip": 0}),
                Flow("get_attachment", {"messageId": value, "attachmentId": "id"}),
                Flow("get_attachment", {"messageId": "id", "attachmentId": value}),
            ]
            for flow in targets:
                response = flow.run()
                self.assertEqual(response["statusCode"], 400)
                self.assertEqual(response["body"]["error"]["code"], "INVALID_ARGUMENTS")
                self.assertEqual(flow.calls, [])
            flow = Flow("list_attachments", {"messageId": "id", "top": 1, "skip": 0}, [{"value": [metadata(id=value)]}])
            response = flow.run()
            self.assertEqual(response["statusCode"], 502)
            self.assertEqual(response["body"]["error"]["code"], "INVALID_ATTACHMENT_METADATA")
            self.assertEqual(len(flow.calls), 1)

    def test_id_bounds_percent_literals_and_non_whitespace_unicode(self):
        valid = ["x" * 2048, "\U0001f600" * 1024, "日本語-é-\u200b", "%00%1F%20%C2%A0%EF%BB%BF", "a/b?x=1#y", "...", "%0%1"]
        for value in valid:
            flow = Flow("get_attachment", {"messageId": value, "attachmentId": value}, [metadata(id=value), metadata(id=value, contentBytes="YWJj")])
            self.assertEqual(flow.run()["statusCode"], 200)
            encoded = quote(value, safe="~()*!.'-_")
            self.assertTrue(all("messages/" + encoded + "/attachments/" + encoded in call["Uri"] for call in flow.calls))
        for value in ("x" * 2049, "\U0001f600" * 1025, None, True, 123, [], {}):
            flow = Flow("get_attachment", {"messageId": "id", "attachmentId": value})
            self.assertEqual(flow.run()["statusCode"], 400)
            self.assertEqual(flow.calls, [])
            self.assertEqual(flow.status[f"Validate_get_attachment_ids"], "Skipped")

    def test_base64_short_alphabet_and_padding_combinations(self):
        for length in range(5):
            for chars in itertools.product("A+/=_", repeat=length):
                content = "".join(chars)
                valid = re.fullmatch(r"[A-Za-z0-9+/]*={0,2}", content) is not None and len(content) % 4 == 0
                response = get_flow(metadata(), metadata(contentBytes=content)).run()
                self.assertEqual(response["statusCode"], 200 if valid else 502, repr(content))

    def test_base64_ascii_unicode_types_and_chunk_boundaries(self):
        for character in [chr(value) for value in range(128)] + ["é", "\u200b", "\u2028", "\ufeff", "\U0001f600", "\ud800"]:
            content = "AA" + character + "A"
            valid = character in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
            self.assertEqual(get_flow(metadata(), metadata(contentBytes=content)).run()["statusCode"], 200 if valid else 502, repr(character))
        for content in (None, True, 123, [], {}, "A" * (MAX_BASE64 + 1)):
            flow = get_flow(metadata(), metadata(contentBytes=content))
            self.assertEqual(flow.run()["statusCode"], 502)
            self.assertEqual(flow.status["Select_attachment_content_alphabet"], "Skipped")
        for offset in (BASE64_CHUNK - 1, BASE64_CHUNK, BASE64_CHUNK + 1, BASE64_CHUNK * 2 - 1):
            content = "A" * offset + "_" + "A" * (BASE64_CHUNK * 2 - offset - 1)
            self.assertEqual(get_flow(metadata(), metadata(contentBytes=content)).run()["statusCode"], 502)

    def test_exact_maximum_content_with_heavy_uri_escaping(self):
        for byte in (0xFF, 0xFB):
            encoded = base64.b64encode(bytes([byte]) * MAX_BYTES).decode("ascii")
            flow = get_flow(metadata(size=MAX_BYTES), metadata(size=MAX_BYTES, contentBytes=encoded))
            self.assertEqual(flow.run()["statusCode"], 200)
            chunks = flow.bodies["Select_attachment_content_alphabet"]
            self.assertTrue(all(value is True for value in chunks))
            self.assertLessEqual(len(chunks), (MAX_BASE64 + BASE64_CHUNK - 1) // BASE64_CHUNK)

    def test_escaped_metadata_type_remains_required_at_every_parse_gate(self):
        missing_type = metadata()
        del missing_type["@odata.type"]
        flows = [
            (Flow("list_attachments", {"messageId": "message-id", "top": 1, "skip": 0}, [{"value": [missing_type]}]), "INVALID_ATTACHMENT_METADATA", 1),
            (get_flow(missing_type), "INVALID_ATTACHMENT_METADATA", 1),
            (get_flow(metadata(), dict(missing_type, contentBytes="YWJj")), "INVALID_ATTACHMENT_CONTENT", 2),
        ]
        for flow, code, calls in flows:
            response = flow.run()
            self.assertEqual(response["statusCode"], 502)
            self.assertEqual(response["body"]["error"]["code"], code)
            self.assertEqual(len(flow.calls), calls)

    def test_gateway_fails_closed(self):
        for configured, supplied in [("", ""), ("", None), ("synthetic-test-only-key", None), ("synthetic-test-only-key", "wrong")]:
            flow = Flow("list_messages", {"top": 1})
            flow.parameters["McpGatewayKey"] = configured
            flow.headers = {} if supplied is None else {"X-MCP-Gateway-Key": supplied}
            self.assertIsNone(flow.run())
            self.assertEqual(flow.calls, [])

    def test_metadata_only_listing_and_nextlink_is_opaque(self):
        item = metadata(contentBytes="must-not-escape", sourceUrl="https://outside.invalid/file")
        flow = Flow("list_attachments", {"messageId": "message/id?&%=", "top": 2, "skip": 4}, [{"value": [item], "@odata.nextLink": "https://outside.invalid/untrusted-next-page", "extra": "must-not-escape"}])
        response = flow.run()
        self.assertEqual(response["statusCode"], 200)
        self.assertEqual(response["body"]["data"]["value"], [metadata()])
        self.assertEqual(response["body"]["data"]["@odata.nextLink"], "https://outside.invalid/untrusted-next-page")
        self.assertEqual(len(flow.calls), 1)
        self.assertEqual(flow.calls[0]["Uri"], GRAPH_ROOT + "messages/message%2Fid%3F%26%25%3D/attachments?$select=id,name,contentType,size,isInline&$top=2&$skip=4")
        self.assertNotIn("contentBytes", flow.calls[0]["Uri"])

    def test_empty_list_and_no_nextlink(self):
        flow = Flow("list_attachments", {"messageId": "message-id", "top": 50, "skip": 10000}, [{"value": []}])
        self.assertEqual(flow.run()["body"]["data"], {"value": [], "@odata.nextLink": ""})

    def test_file_preflight_then_content_and_projection(self):
        flow = get_flow(metadata(), metadata(contentBytes="YWJj", sourceUrl="must-not-escape", item={"body": "must-not-escape"}))
        response = flow.run()
        self.assertEqual(response["statusCode"], 200)
        self.assertEqual(response["body"], {"ok": True, "requestId": REQUEST_ID, "operation": "get_attachment", "data": metadata(contentBytes="YWJj")})
        self.assertEqual([call["Uri"] for call in flow.calls], [GRAPH_ROOT + "messages/message-id/attachments/attachment-id?$select=id,name,contentType,size,isInline", GRAPH_ROOT + "messages/message-id/attachments/attachment-id"])

    def test_each_id_is_a_separate_encoded_path_segment(self):
        flow = Flow("get_attachment", {"messageId": "a/b?x=1", "attachmentId": "%2e%2e/#"}, [metadata(id="%2e%2e/#"), metadata(id="%2e%2e/#", contentBytes="YWJj")])
        self.assertEqual(flow.run()["statusCode"], 200)
        for call in flow.calls:
            self.assertIn("messages/a%2Fb%3Fx%3D1/attachments/%252e%252e%2F%23", call["Uri"])

    def test_non_file_types_never_fetch_content_or_external_urls(self):
        for kind in ["#microsoft.graph.itemAttachment", "#microsoft.graph.referenceAttachment", "#microsoft.graph.unknownAttachment"]:
            flow = get_flow(metadata(**{"@odata.type": kind, "sourceUrl": "https://outside.invalid"}))
            self.assertEqual(flow.run()["statusCode"], 415)
            self.assertEqual(len(flow.calls), 1)

    def test_oversized_file_never_fetches_content(self):
        flow = get_flow(metadata(size=MAX_BYTES + 1))
        response = flow.run()
        self.assertEqual(response["statusCode"], 413)
        self.assertEqual(response["body"]["error"]["code"], "ATTACHMENT_TOO_LARGE")
        self.assertEqual(len(flow.calls), 1)

    def test_bad_metadata_never_fetches_content(self):
        values = [metadata(size=-1), metadata(size="3"), metadata(size=None), metadata(id="different"), {"@odata.type": "#microsoft.graph.fileAttachment"}]
        for value in values:
            flow = get_flow(value)
            self.assertEqual(flow.run()["statusCode"], 502)
            self.assertEqual(len(flow.calls), 1)

    def test_full_fetch_type_identity_and_size_are_rechecked(self):
        for content in [metadata(**{"@odata.type": "#microsoft.graph.referenceAttachment", "contentBytes": "YWJj"}), metadata(id="different", contentBytes="YWJj"), metadata(size=MAX_BYTES + 1, contentBytes="YWJj"), metadata(contentBytes="invalid base64!")]:
            flow = get_flow(metadata(), content)
            self.assertEqual(flow.run()["statusCode"], 502)
            self.assertEqual(len(flow.calls), 2)

    def test_exact_decoded_size_boundary_even_when_metadata_lies(self):
        for size, expected in [(MAX_BYTES, 200), (MAX_BYTES + 1, 502), (MAX_BYTES + 2, 502), (MAX_BYTES + 3, 502)]:
            encoded = base64.b64encode(b"x" * size).decode("ascii")
            flow = get_flow(metadata(size=MAX_BYTES), metadata(size=MAX_BYTES, contentBytes=encoded))
            self.assertEqual(flow.run()["statusCode"], expected)

    def test_base64_padding_length_and_true_end_of_input(self):
        for content in ("", "Zg==", "Zm8=", "Zm9v"):
            self.assertEqual(get_flow(metadata(), metadata(contentBytes=content)).run()["statusCode"], 200)
        for content in ("A", "A=", "AA=", "AAAA=", "===", "YWJj\n", "YWJj\r\n", "YWJj\u2028", "=YWJ", "Y=Jj"):
            self.assertEqual(get_flow(metadata(), metadata(contentBytes=content)).run()["statusCode"], 502)

    def test_invalid_arguments_fail_before_graph(self):
        base = {"messageId": "message-id", "top": 1, "skip": 0}
        bad = [dict(base, top=n) for n in (0, 51, 1.5, "1", True)] + [dict(base, skip=n) for n in (-1, 10001, 1.5, "1", True)]
        bad += [dict(base, messageId=value) for value in ("", ".", "..", "white space", "id\x00", "id\n", "id\r\n", "id\u2028", "x" * 2049, None)]
        bad += [dict(base, url="https://outside.invalid"), {"messageId": "message-id", "top": 1}]
        for args in bad:
            flow = Flow("list_attachments", args)
            self.assertEqual(flow.run()["statusCode"], 400)
            self.assertEqual(flow.calls, [])
        for args in ({"messageId": "message-id"}, {"messageId": "message-id", "attachmentId": ".."}, {"messageId": "message-id", "attachmentId": "id", "method": "POST"}):
            flow = Flow("get_attachment", args)
            self.assertEqual(flow.run()["statusCode"], 400)
            self.assertEqual(flow.calls, [])

    def test_invalid_attachment_envelopes_are_sanitized_inside_their_cases(self):
        for operation, args in [("list_attachments", {"messageId": "id", "top": 1, "skip": 0}), ("get_attachment", {"messageId": "id", "attachmentId": "id"})]:
            for change in [{"requestId": "private-invalid-id"}, {"extra": "private-input"}, {"args": None}]:
                flow = Flow(operation, args)
                flow.request.update(change)
                response = flow.run()
                self.assertEqual(response["statusCode"], 400)
                self.assertEqual(response["body"], {"ok": False, "error": {"code": "INVALID_REQUEST", "message": "The request envelope is invalid."}})
                self.assertEqual(flow.calls, [])

    def test_graph_failures_and_timeouts_have_sanitized_502(self):
        for error in (ValueError("private Graph error detail"), TimeoutError("private URL")):
            for flow in [get_flow(error), get_flow(metadata(), error), Flow("list_attachments", {"messageId": "id", "top": 1, "skip": 0}, [error])]:
                response = flow.run()
                self.assertEqual(response["statusCode"], 502)
                self.assertEqual(response["body"], {"ok": False, "error": {"code": "UPSTREAM_ERROR", "message": "The upstream read could not be completed."}})

    def test_malformed_list_cannot_escape(self):
        for body in [{}, {"value": None}, {"value": [metadata(size="3")]}, {"value": [metadata()] * 51}]:
            flow = Flow("list_attachments", {"messageId": "id", "top": 50, "skip": 0}, [body])
            self.assertEqual(flow.run()["statusCode"], 502)

    def test_legacy_message_operations_keep_success_envelopes_and_queries(self):
        for operation, args, data in [("list_messages", {"top": 2}, {"value": []}), ("search_messages", {"query": 'two words"&$expand=attachments', "top": 3}, {"value": []}), ("get_message", {"messageId": "message/id"}, {"id": "message/id", "body": {"contentType": "text", "content": "synthetic"}})]:
            flow = Flow(operation, args, [data])
            response = flow.run()
            self.assertEqual(response["body"], {"ok": True, "requestId": REQUEST_ID, "operation": operation, "data": data})
            self.assertEqual(flow.calls[0]["Method"], "GET")
            if operation == "search_messages":
                self.assertIn("%22two%20words%22%26%24expand%3Dattachments%22", flow.calls[0]["Uri"])
                self.assertNotIn("&$expand=", flow.calls[0]["Uri"])
            if operation == "get_message":
                self.assertIn("messages/message%2Fid?$select=", flow.calls[0]["Uri"])
                self.assertEqual(flow.calls[0]["CustomHeader2"], 'Prefer: outlook.body-content-type="text"')

    def test_v2_mail_requires_typed_top_uuid_and_message_id_but_keeps_switch_fallback(self):
        # The Worker already supplies bounded top and UUIDs. The expanded flow
        # deliberately replaces loose direct-trigger legacy behavior.
        for operation, args in [("list_messages", {}), ("search_messages", {"query": "synthetic"})]:
            flow = Flow(operation, args)
            self.assertEqual(flow.run()["body"]["error"]["code"], "INVALID_ARGUMENTS")
            self.assertEqual(flow.calls, [])
            flow = Flow(operation, dict(args, top=20))
            flow.request["requestId"] = "legacy-non-uuid"
            self.assertEqual(flow.run()["body"]["error"]["code"], "INVALID_REQUEST")
            self.assertEqual(flow.calls, [])
        for args in ({}, {"messageId": ""}):
            flow = Flow("get_message", args)
            self.assertEqual(flow.run()["statusCode"], 400)
            self.assertEqual(flow.calls, [])
        flow = Flow("unsupported", {})
        self.assertEqual(flow.run()["body"]["error"]["code"], "INVALID_OPERATION")

    def test_v2_mail_upstream_failures_have_explicit_sanitized_responses(self):
        for operation, args in [("list_messages", {"top": 20}), ("search_messages", {"query": "synthetic", "top": 20}), ("get_message", {"messageId": "id"})]:
            for error in (ValueError("private synthetic failure"), TimeoutError("private synthetic timeout")):
                flow = Flow(operation, args, [error])
                response = flow.run()
                self.assertEqual(response["statusCode"], 502)
                self.assertEqual(response["body"], {"ok": False, "error": {"code": "UPSTREAM_ERROR", "message": "The upstream read could not be completed."}})
                self.assertEqual(len(flow.calls), 1)


if __name__ == "__main__":
    unittest.main()
