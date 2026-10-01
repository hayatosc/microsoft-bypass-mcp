#!/usr/bin/env python3
"""Extract a reviewable, fail-closed flow definition from a private legacy export.

Never prints source values. Keep the input ZIP outside the repository. This is
intentionally specific to the three-operation flow, not a general secret scanner.
"""

import argparse
import copy
import json
from pathlib import Path
import re
import zipfile

FLOW_NAME = "microsoft bypass flow"
OPERATIONS = {"list_messages", "search_messages", "get_message"}
PARAMETER = "McpGatewayKey"
SAFE_CONDITION = (
    "@and(not(empty(parameters('McpGatewayKey'))),"
    "equals(triggerOutputs()?['headers']?['X-MCP-Gateway-Key'],"
    "parameters('McpGatewayKey')))"
)
LITERAL_CONDITION = re.compile(
    r"^@equals\(triggerOutputs\(\)\?\['headers'\]\?\['X-MCP-Gateway-Key'\],\s*"
    r"'((?:''|[^'])*)'\)$"
)
MAX_BYTES = 10 * 1024 * 1024


def read_export(source):
    """Read exactly one workflow resource; do not extract any archive paths."""
    source = Path(source)
    if source.suffix.lower() == ".zip":
        with zipfile.ZipFile(source) as archive:
            entries = [
                item for item in archive.infolist()
                if re.fullmatch(r"Microsoft\.Flow/flows/[^/]+/definition\.json", item.filename)
            ]
            if len(entries) != 1 or entries[0].file_size > MAX_BYTES:
                raise ValueError("Expected one reasonably sized flow definition in the ZIP")
            return json.loads(archive.read(entries[0]))
    if source.stat().st_size > MAX_BYTES:
        raise ValueError("Flow definition is too large")
    return json.loads(source.read_text(encoding="utf-8"))


def remove_metadata(value):
    if isinstance(value, dict):
        return {key: remove_metadata(child) for key, child in value.items() if key != "metadata"}
    if isinstance(value, list):
        return [remove_metadata(child) for child in value]
    return value


def walk(value):
    yield value
    if isinstance(value, dict):
        for child in value.values():
            yield from walk(child)
    elif isinstance(value, list):
        for child in value:
            yield from walk(child)


def validate_definition(definition):
    """Reject unexpected operations, connector writes, identities, and URLs."""
    trigger = definition["triggers"]["manual"]
    schema = trigger["inputs"]["schema"]
    if set(definition["triggers"]) != {"manual"} or (trigger["type"], trigger["kind"]) != ("Request", "Http"):
        raise ValueError("Unexpected trigger structure; review the export privately")
    if set(schema["properties"]["operation"]["enum"]) != OPERATIONS:
        raise ValueError("Unexpected operation allowlist")
    switches = [value for value in definition["actions"].values() if value.get("type") == "Switch"]
    if len(switches) != 1 or set(switches[0]["cases"]) != OPERATIONS:
        raise ValueError("Unexpected operation switch")
    def actions(action_map):
        for action in action_map.values():
            yield action
            yield from actions(action.get("actions", {}))
            for branch in [action.get("else", {}), action.get("default", {})]:
                yield from actions(branch.get("actions", {}))
            for branch in action.get("cases", {}).values():
                yield from actions(branch.get("actions", {}))

    all_actions = list(actions(definition["actions"]))
    if any(action.get("type") not in {"Switch", "If", "Response", "OpenApiConnection"} for action in all_actions):
        raise ValueError("Unexpected action type; review the export privately")
    requests = [action for action in all_actions if action["type"] == "OpenApiConnection"]
    if len(requests) != 3:
        raise ValueError("Expected three connector read actions")
    for request in requests:
        inputs = request["inputs"]
        host = inputs["host"]
        if inputs["parameters"]["Method"] != "GET" or host != {
            "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
            "connectionName": "shared_office365",
            "operationId": "HttpRequest",
        }:
            raise ValueError("Unexpected connector action; review the export privately")
        if not inputs["parameters"]["Uri"].startswith("@concat(\r\n  'https://graph.microsoft.com/v1.0/me/"):
            raise ValueError("Unexpected Graph endpoint expression")
    for value in walk(definition):
        if not isinstance(value, str):
            continue
        if re.search(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}", value, re.I):
            raise ValueError("Retained identifier requires private review")
        if re.search(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}", value):
            raise ValueError("Retained email address requires private review")
        if re.search(r"(?:[?&]sig=|Bearer\s+|[?&]token=|-----BEGIN .*PRIVATE KEY-----)", value, re.I):
            raise ValueError("Potential credential requires private review")
        for host in re.findall(r"https?://([^/'\s?]+)", value):
            if host not in {"graph.microsoft.com", "schema.management.azure.com"}:
                raise ValueError("Unexpected external URL requires private review")


def sanitize(resource):
    properties = resource["properties"]
    if properties["displayName"] != FLOW_NAME:
        raise ValueError("Unexpected flow name")
    definition = remove_metadata(copy.deepcopy(properties["definition"]))
    trigger = definition["triggers"]["manual"]
    conditions = trigger["conditions"]
    if len(conditions) != 1 or set(conditions[0]) != {"expression"}:
        raise ValueError("Unexpected gateway guard; review the export privately")
    match = LITERAL_CONDITION.fullmatch(conditions[0]["expression"])
    if not match or not match.group(1):
        raise ValueError("Expected the existing literal gateway-key guard")
    original_key = match.group(1).replace("''", "'")
    if PARAMETER in definition["parameters"]:
        raise ValueError("Gateway parameter already exists; review before replacing it")
    definition["parameters"][PARAMETER] = {"defaultValue": "", "type": "SecureString"}
    conditions[0]["expression"] = SAFE_CONDITION
    if any(original_key in value or match.group(1) in value for value in walk(definition) if isinstance(value, str)):
        raise ValueError("Gateway key appears outside the expected guard; refusing output")
    validate_definition(definition)
    return definition


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", help="Private Power Automate package ZIP or exported definition.json")
    parser.add_argument("destination", type=Path, help="Sanitized definition.json to write")
    args = parser.parse_args()
    try:
        definition = sanitize(read_export(args.source))
    except (KeyError, TypeError, ValueError, OSError, zipfile.BadZipFile) as error:
        # Do not include exception text: JSON decoder errors can contain input data.
        parser.exit(1, f"Export validation failed ({type(error).__name__}); inspect the source privately.\n")
    args.destination.parent.mkdir(parents=True, exist_ok=True)
    args.destination.write_text(json.dumps(definition, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("Wrote sanitized workflow definition; review the diff before committing.")


if __name__ == "__main__":
    main()
