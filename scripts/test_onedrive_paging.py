"""Focused offline OneDrive pagination/projection contracts using synthetic data.

The existing WDL emulator does not implement Microsoft's connector runtime.
NativePagingFlow models the already-verified native auto-aggregation contract to
exercise our post-aggregation guards, including a final-page threshold overshoot.
This is not tenant/import/live validation and does not guess or fetch skip tokens.
"""

from copy import deepcopy
import json
import unittest

from build_attachment_flow import build_definition
from test_attachment_flow import Expressions, Flow, actions_in, validate


CAP = 1000
CANARY = "synthetic-private-onedrive-paging-canary"
URL_CANARY = "https://outside.invalid/" + CANARY
FIELDS = {"Id", "Name", "Size", "MediaType", "IsFolder", "LastModified", "ETag"}


def native_items(count, start=0):
    return [{"Id": "owned-file-" + str(index), "Name": "report.pdf", "Size": 3,
             "MediaType": "application/pdf", "IsFolder": False,
             "LastModified": None, "ETag": None, "Path": URL_CANARY,
             "FileLocator": URL_CANARY, "contentBytes": CANARY}
            for index in range(start, start + count)]


def definition():
    # Same canonical generator, never a separate flow source or private export.
    return build_definition()


def folder_actions(source=None):
    source = source or definition()
    return source["actions"]["スイッチ"]["cases"]["onedrive_list_folder"]["actions"]


def flow(raw, *, folder_id=None, top=CAP):
    args = {"top": top}
    if folder_id is not None:
        args["folderId"] = folder_id
    return Flow("onedrive_list_folder", args, [raw], definition=definition())


class NativePagingFlow(Flow):
    """Synthetic native aggregation; never interprets nextLink as a network URL.

    Only ListFolderV2 has the actual connector pagination policy. We provide its
    aggregated action output to the existing emulator, which then executes the
    unchanged WDL projection, validation and response/security paths.
    """

    def __init__(self, pages):
        super().__init__("onedrive_list_folder", {"folderId": "owned-folder", "top": CAP},
                         definition=definition())
        self.native_pages = deepcopy(pages)
        self.native_page_reads = 0
        self.native_aggregate_count = 0

    def execute(self, name, action):
        if name == "OneDrive_list_folder_v2":
            threshold = action["runtimeConfiguration"]["paginationPolicy"]["minimumItemCount"]
            values = []
            next_link = None
            for page in self.native_pages:
                values.extend(page["value"])
                next_link = page.get("nextLink")
                self.native_page_reads += 1
                if len(values) >= threshold or not next_link:
                    break
            self.native_aggregate_count = len(values)
            self.queue.append({"value": values, "nextLink": next_link})
        super().execute(name, action)


class OneDrivePagingStaticContracts(unittest.TestCase):
    def test_static_native_policy_only_on_list_folder_v2_preserves_secure_history(self):
        policies = []
        for action in actions_in(definition()):
            runtime = action.get("runtimeConfiguration", {})
            if "paginationPolicy" in runtime:
                policies.append(action)
                self.assertEqual(action["type"], "OpenApiConnection")
                self.assertEqual(action["inputs"]["host"], {
                    "apiId": "/providers/Microsoft.PowerApps/apis/shared_onedriveforbusiness",
                    "connectionName": "shared_onedriveforbusiness", "operationId": "ListFolderV2",
                })
                self.assertEqual(runtime, {
                    "secureData": {"properties": ["inputs", "outputs"]},
                    "paginationPolicy": {"minimumItemCount": CAP},
                })
                self.assertEqual(set(action["inputs"]["parameters"]), {"id"})
        self.assertEqual(len(policies), 1)
        route = folder_actions()["Route_onedrive_list_folder"]
        root = route["actions"]["OneDrive_list_root"]
        self.assertEqual(root["inputs"]["host"]["operationId"], "ListRootFolder")
        self.assertEqual(root["inputs"]["parameters"], {})
        self.assertEqual(root["runtimeConfiguration"], {"secureData": {"properties": ["inputs", "outputs"]}})

    def test_folder_top_and_projection_are_bounded_to_1000_only(self):
        cases = definition()["actions"]["スイッチ"]["cases"]
        schema = cases["onedrive_list_folder"]["actions"]["Validate_onedrive_list_folder_args"]["inputs"]["schema"]
        self.assertEqual(schema["properties"]["top"], {"type": "integer", "minimum": 1, "maximum": CAP})
        self.assertEqual(set(schema["properties"]), {"folderId", "top"})
        self.assertFalse(schema["additionalProperties"])
        for top in (1, 100, 101, CAP):
            validate({"top": top}, schema)
        for top in (0, CAP + 1, "1000", True):
            with self.assertRaises(ValueError):
                validate({"top": top}, schema)
        route = folder_actions()["Route_onedrive_list_folder"]
        for branch, prefix, source in (
            (route["actions"], "onedrive_root_folder", "body('OneDrive_list_root')"),
            (route["else"]["actions"], "onedrive_list_folder", "body('OneDrive_list_folder_v2')?['value']"),
        ):
            self.assertEqual(branch["Select_" + prefix]["inputs"]["from"],
                             "@take(take(" + source + ", 1000), body('Validate_onedrive_list_folder_args')?['top'])")
            self.assertEqual(branch["Validate_" + prefix]["inputs"]["schema"]["maxItems"], CAP)
            self.assertEqual(set(branch["Select_" + prefix]["inputs"]["select"]), FIELDS)
        # No widened search/Outlook list validators or binary operations.
        maxima = [action["inputs"]["schema"]["maxItems"]
                  for action in actions_in(cases)
                  if action["type"] == "ParseJson" and action["inputs"]["schema"].get("type") == "array"]
        self.assertEqual(maxima.count(CAP), 2)

    def test_find_files_search_settings_projection_and_no_continuation_are_unchanged(self):
        actions = definition()["actions"]["スイッチ"]["cases"]["onedrive_search_files"]["actions"]
        schema = actions["Validate_onedrive_search_files_args"]["inputs"]["schema"]
        self.assertEqual(schema["properties"]["top"], {"type": "integer", "minimum": 1, "maximum": 100})
        for key in ("OneDriveSearchRootId", "OneDriveSearchMode"):
            self.assertEqual(actions[key]["type"], "Compose")
            self.assertEqual(actions[key]["inputs"], "")
            self.assertEqual(actions[key]["runtimeConfiguration"], {"secureData": {"properties": ["inputs"]}})
        native = actions["Check_onedrive_search_binding"]["actions"]
        find = native["OneDrive_find_files"]
        self.assertEqual(find["inputs"]["host"]["operationId"], "FindFiles")
        self.assertEqual(find["inputs"]["parameters"], {
            "query": "@trim(body('Validate_onedrive_search_files_args')?['query'])",
            "id": "@outputs('OneDriveSearchRootId')", "findMode": "@outputs('OneDriveSearchMode')",
            "maxFileCount": "@body('Validate_onedrive_search_files_args')?['top']",
        })
        self.assertNotIn("paginationPolicy", find["runtimeConfiguration"])
        self.assertEqual(native["Validate_onedrive_search_files"]["inputs"]["schema"]["maxItems"], 100)
        self.assertEqual(native["Select_onedrive_search_files"]["inputs"]["from"],
                         "@take(body('OneDrive_find_files'), body('Validate_onedrive_search_files_args')?['top'])")
        self.assertEqual(native["Respond_onedrive_search_files"]["inputs"]["body"]["data"], {
            "value": "@body('Validate_onedrive_search_files')",
            "truncated": "@greaterOrEquals(length(body('OneDrive_find_files')), body('Validate_onedrive_search_files_args')?['top'])",
        })

    def test_folder_action_has_no_guessed_skiptoken_or_arbitrary_url_following(self):
        actions = folder_actions()
        serialized = json.dumps(actions)
        for marker in ("skiptoken", "$skip", "HttpRequest", "Until", "nextCursor", "InitializeVariable"):
            self.assertNotIn(marker, serialized)
        network = [action for action in actions_in(actions) if action["type"] == "OpenApiConnection"]
        self.assertEqual({action["inputs"]["host"]["operationId"] for action in network}, {"ListRootFolder", "ListFolderV2"})
        self.assertTrue(all("Uri" not in action["inputs"]["parameters"] for action in network))
        for action in actions_in(actions):
            if action["type"] == "Response" and action["inputs"]["statusCode"] == 200:
                self.assertEqual(set(action["inputs"]["body"]["data"]), {"value", "truncated"})


class OneDrivePagingExecutionContracts(unittest.TestCase):
    def assert_safe_data(self, engine, count, truncated):
        response = engine.run()
        self.assertEqual(response["statusCode"], 200)
        data = response["body"]["data"]
        self.assertEqual(len(data["value"]), count)
        self.assertIs(data["truncated"], truncated)
        self.assertTrue(all(set(item) == FIELDS for item in data["value"]))
        self.assertNotIn(CANARY, json.dumps(data))
        self.assertEqual(len(engine.calls), 1)
        return data

    def test_native_multiple_pages_autoaggregate_beyond_the_first_page(self):
        pages = [
            {"value": native_items(120), "nextLink": URL_CANARY},
            {"value": native_items(130, 120)},
        ]
        engine = NativePagingFlow(pages)
        data = self.assert_safe_data(engine, 250, False)
        self.assertEqual(engine.native_page_reads, 2)
        self.assertEqual(engine.call_operations, ["ListFolderV2"])
        self.assertEqual(engine.calls, [{"id": "owned-folder"}])
        self.assertEqual([item["Id"] for item in data["value"]], ["owned-file-" + str(index) for index in range(250)])

    def test_native_threshold_overshoot_is_bounded_before_projection(self):
        engine = NativePagingFlow([
            {"value": native_items(600), "nextLink": URL_CANARY},
            {"value": native_items(600, 600)},
        ])
        data = self.assert_safe_data(engine, CAP, True)
        self.assertEqual(engine.native_aggregate_count, 1200)
        self.assertEqual(engine.native_page_reads, 2)
        self.assertEqual(data["value"][-1]["Id"], "owned-file-999")
        self.assertEqual(len(engine.bodies["Select_onedrive_list_folder"]), CAP)

    def test_exact_threshold_stops_with_conservative_incompleteness_and_no_exposed_token(self):
        engine = NativePagingFlow([
            {"value": native_items(500), "nextLink": URL_CANARY},
            {"value": native_items(500, 500), "nextLink": URL_CANARY},
            {"value": native_items(50, 1000)},
        ])
        self.assert_safe_data(engine, CAP, True)
        self.assertEqual(engine.native_page_reads, 2)
        self.assertEqual(engine.native_aggregate_count, CAP)

    def test_root_and_child_raw_count_boundary_is_conservative_even_without_nextlink(self):
        for folder_id in (None, "owned-folder"):
            for count in (0, 250, 999, 1000, 1001, 1300):
                with self.subTest(folder_id=folder_id, count=count):
                    values = native_items(count)
                    engine = flow(values if folder_id is None else {"value": values}, folder_id=folder_id)
                    self.assert_safe_data(engine, min(count, CAP), count >= CAP)
                    self.assertEqual(engine.call_operations, ["ListRootFolder" if folder_id is None else "ListFolderV2"])

    def test_lower_top_and_native_continuation_are_bounded_without_fetching_urls(self):
        for raw, folder_id, top, count, truncated in (
            (native_items(250), None, 100, 100, True),
            ({"value": native_items(250)}, "owned-folder", 100, 100, True),
            ({"value": native_items(2), "nextLink": URL_CANARY}, "owned-folder", CAP, 2, True),
            ({"value": [], "nextLink": URL_CANARY}, "owned-folder", CAP, 0, True),
            ({"value": [], "nextLink": ""}, "owned-folder", CAP, 0, False),
            ({"value": [], "nextLink": None}, "owned-folder", CAP, 0, False),
        ):
            engine = flow(raw, folder_id=folder_id, top=top)
            self.assert_safe_data(engine, count, truncated)
            self.assertNotIn(CANARY, json.dumps(engine.calls))

    def test_invalid_args_and_malformed_aggregates_keep_existing_static_error_paths(self):
        for args in ({"top": CAP + 1}, {"top": 0}, {"top": CAP, "nextLink": URL_CANARY},
                     {"top": CAP, "cursor": CANARY}):
            engine = Flow("onedrive_list_folder", args, definition=definition())
            response = engine.run()
            self.assertEqual(response["statusCode"], 400)
            self.assertEqual(engine.calls, [])
            self.assertNotIn(CANARY, json.dumps(response))
        for raw in (None, {}, {"value": None}, {"value": [{"Id": CANARY}]},
                    {"value": [dict(native_items(1)[0], Size="3")]}):
            engine = flow(raw, folder_id="owned-folder")
            response = engine.run()
            self.assertEqual(response["statusCode"], 502)
            self.assertNotIn(CANARY, json.dumps(response))
            self.assertEqual(len(engine.calls), 1)

    def test_search_count_100_stays_incomplete_and_has_no_paging(self):
        source = definition()
        actions = source["actions"]["スイッチ"]["cases"]["onedrive_search_files"]["actions"]
        actions["OneDriveSearchRootId"]["inputs"] = "synthetic-root"
        actions["OneDriveSearchMode"]["inputs"] = "synthetic-mode"
        engine = Flow("onedrive_search_files", {"query": "synthetic", "top": 100},
                      [native_items(120)], definition=source)
        self.assert_safe_data(engine, 100, True)
        self.assertEqual(engine.call_operations, ["FindFiles"])
        self.assertEqual(engine.calls[0]["maxFileCount"], 100)
        with self.assertRaises(ValueError):
            validate({"query": "synthetic", "top": 101},
                     Expressions(None).resolve(actions["Validate_onedrive_search_files_args"]["inputs"]["schema"]))


if __name__ == "__main__":
    unittest.main()
