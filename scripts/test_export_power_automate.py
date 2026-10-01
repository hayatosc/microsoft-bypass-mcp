"""Offline contract and redaction checks; never call the live university flow."""

import copy
import json
from pathlib import Path
import tempfile
import unittest
import zipfile

from export_power_automate import (
    FLOW_NAME, OPERATIONS, PARAMETER, SAFE_CONDITION,
    read_export, sanitize, validate_definition, walk,
)

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "scripts/fixtures/microsoft-bypass-flow.pre-attachments.json"
FAKE_KEY = "synthetic-test-only-gateway-key"


def fixture_resource():
    definition = json.loads(SOURCE.read_text(encoding="utf-8"))
    del definition["parameters"][PARAMETER]
    definition["triggers"]["manual"]["conditions"][0]["expression"] = (
        "@equals(triggerOutputs()?['headers']?['X-MCP-Gateway-Key'], '" + FAKE_KEY + "')"
    )
    definition["metadata"] = {"creator": {"id": "private-owner", "tenantId": "private-tenant"}}
    definition["triggers"]["manual"]["metadata"] = {"operationMetadataId": "private-operation"}
    return {
        "id": "private-flow-id", "name": "private-flow-name", "type": "Microsoft.Flow/flows",
        "properties": {
            "displayName": FLOW_NAME,
            "definition": definition,
            "connectionReferences": {"shared_office365": {"connectionName": "private-connection"}},
        },
    }


class ExportTests(unittest.TestCase):
    def test_sanitization_preserves_every_non_configuration_field(self):
        actual = sanitize(fixture_resource())
        self.assertEqual(actual, json.loads(SOURCE.read_text(encoding="utf-8")))
        serialized = json.dumps(actual)
        for value in [FAKE_KEY, "private-owner", "private-tenant", "private-operation", "private-flow-id", "private-connection"]:
            self.assertNotIn(value, serialized)
        self.assertFalse(any(isinstance(value, dict) and "metadata" in value for value in walk(actual)))

    def test_escaped_quote_in_key_is_removed(self):
        resource = fixture_resource()
        resource["properties"]["definition"]["triggers"]["manual"]["conditions"][0]["expression"] = (
            "@equals(triggerOutputs()?['headers']?['X-MCP-Gateway-Key'], 'synthetic''quoted-key')"
        )
        self.assertEqual(sanitize(resource)["triggers"]["manual"]["conditions"][0]["expression"], SAFE_CONDITION)

    def test_reused_gateway_secret_is_rejected(self):
        resource = fixture_resource()
        resource["properties"]["definition"]["description"] = FAKE_KEY
        with self.assertRaises(ValueError):
            sanitize(resource)

    def test_changed_gateway_guard_requires_private_review(self):
        resource = fixture_resource()
        resource["properties"]["definition"]["triggers"]["manual"]["conditions"] = []
        with self.assertRaises(ValueError):
            sanitize(resource)

    def test_unexpected_flow_name_is_rejected(self):
        resource = fixture_resource()
        resource["properties"]["displayName"] = "another flow"
        with self.assertRaises(ValueError):
            sanitize(resource)

    def test_retained_private_identifiers_and_urls_are_rejected(self):
        for value in [
            "owner@example.test", "11111111-2222-4333-8444-555555555555",
            "https://example.test/callback?sig=synthetic", "Bearer synthetic-token",
        ]:
            with self.subTest(value=value):
                resource = fixture_resource()
                resource["properties"]["definition"]["description"] = value
                with self.assertRaises(ValueError):
                    sanitize(resource)

    def test_zip_is_read_without_extracting_paths(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "flow.zip"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("Microsoft.Flow/flows/test/definition.json", json.dumps(fixture_resource()))
                archive.writestr("../not-extracted.txt", "unrelated archive entry")
            self.assertEqual(sanitize(read_export(source)), json.loads(SOURCE.read_text(encoding="utf-8")))
            self.assertEqual(list(Path(directory).iterdir()), [source])

    def test_multiple_flows_are_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "flow.zip"
            with zipfile.ZipFile(source, "w") as archive:
                for name in ["one", "two"]:
                    archive.writestr(f"Microsoft.Flow/flows/{name}/definition.json", "{}")
            with self.assertRaises(ValueError):
                read_export(source)


class FlowContractTests(unittest.TestCase):
    def setUp(self):
        self.definition = json.loads(SOURCE.read_text(encoding="utf-8"))

    def test_actual_read_only_operation_contract(self):
        validate_definition(self.definition)
        schema = self.definition["triggers"]["manual"]["inputs"]["schema"]
        self.assertEqual(set(schema["properties"]["operation"]["enum"]), OPERATIONS)
        self.assertEqual(set(schema["required"]), {"operation", "requestId", "args"})
        top = schema["properties"]["args"]["properties"]["top"]
        self.assertEqual((top["minimum"], top["maximum"]), (1, 50))

    def test_redacted_configuration_is_fail_closed(self):
        self.assertEqual(self.definition["parameters"][PARAMETER], {"type": "SecureString", "defaultValue": ""})
        self.assertEqual(self.definition["triggers"]["manual"]["conditions"], [{"expression": SAFE_CONDITION}])
        self.assertEqual(self.definition["parameters"]["$authentication"]["defaultValue"], {})
        self.assertEqual(self.definition["parameters"]["$connections"]["defaultValue"], {})

    def test_write_operations_and_external_endpoints_are_rejected(self):
        for change in [{"Method": "POST"}, {"Uri": "https://example.test/other"}]:
            modified = copy.deepcopy(self.definition)
            request = next(value for value in walk(modified["actions"]) if isinstance(value, dict) and value.get("type") == "OpenApiConnection")
            request["inputs"]["parameters"].update(change)
            with self.assertRaises(ValueError):
                validate_definition(modified)

    def test_additional_action_types_are_rejected(self):
        self.definition["actions"]["unexpected"] = {"type": "Http", "inputs": {"method": "POST"}}
        with self.assertRaises(ValueError):
            validate_definition(self.definition)

    def test_known_live_quirks_are_preserved(self):
        self.assertEqual(self.definition["contentVersion"], "undefined")
        responses = [value for value in walk(self.definition["actions"]) if isinstance(value, dict) and value.get("type") == "Response"]
        missing = next(response for response in responses if isinstance(response["inputs"]["body"], dict) and response["inputs"]["body"].get("error", {}).get("code") == "MISSING_MESSAGE_ID")
        self.assertEqual(missing["inputs"]["statusCode"], 200)
        self.assertFalse(missing["inputs"]["body"]["ok"])
        self.assertFalse(any(response["inputs"]["statusCode"] == 502 for response in responses))


if __name__ == "__main__":
    unittest.main()
