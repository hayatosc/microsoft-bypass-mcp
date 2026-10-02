"""Synthetic lexical/semantic and canonical-generation compatibility contracts.

No network, native WDL runtime or designer access. Interpolation tests use an
explicit test-only bridge to the unchanged existing public WDL subset emulator.
Canonical hashes pin the original public source and validated string(null)
candidate without depending on staging paths or creating another flow source.
"""

from copy import deepcopy
import hashlib
import itertools
import json
import re
import unittest
from unittest.mock import patch

import build_attachment_flow as builder
from clipboard_compat import REPLACEMENT, TransformError, transform_json, transform_string
from test_attachment_flow import Expressions, Flow, SOURCE, string_length, walk

ORIGINAL_SHA256 = "2835abf5329db9c485919958a7496854a3f1e602c704427020bf9edf711ebd3d"
CANDIDATE_SHA256 = "0d35f92d41144c5dc6daa1c670ee131e783245e549bf05878985d64d02a68a7a"


def resolve_with_interpolation(value):
    """Independent quote-aware boundary scanner; not a native WDL oracle.

    Expressions lacks embedded interpolation and whole @{...} string coercion.
    Use its unchanged evaluator per expression, never re-scan resolved results.
    """
    if isinstance(value, dict):
        return {resolve_with_interpolation(k): resolve_with_interpolation(v)
                for k, v in value.items()}
    if isinstance(value, list):
        return [resolve_with_interpolation(v) for v in value]
    if not isinstance(value, str) or value.startswith("@@"):
        return Expressions(None).resolve(value)
    if value.startswith("@") and not value.startswith("@{"):
        return Expressions(None).resolve(value)
    result, position = [], 0
    while position < len(value):
        if value.startswith("@@", position):
            result.append("@")
            position += 2
        elif value.startswith("@{", position):
            start = position + 2
            end, quoted = start, False
            while end < len(value):
                char = value[end]
                if char == "'":
                    if quoted and value.startswith("''", end):
                        end += 2
                        continue
                    quoted = not quoted
                if char == "}" and not quoted:
                    break
                end += 1
            if end == len(value):
                raise AssertionError("bad interpolation fixture")
            resolved = Expressions(None).resolve("@" + value[start:end])
            result.append(Expressions(None).call("string", [resolved]))
            position = end + 1
        else:
            result.append(value[position])
            position += 1
    return "".join(result)


def assert_typed_equal(test, before, after):
    test.assertIs(type(before), type(after))
    if isinstance(before, dict):
        test.assertEqual(list(before), list(after))
        for key in before:
            assert_typed_equal(test, before[key], after[key])
    elif isinstance(before, list):
        test.assertEqual(len(before), len(after))
        for a, b in zip(before, after):
            assert_typed_equal(test, a, b)
    else:
        test.assertEqual(before, after)


def render(definition):
    return (json.dumps(definition, ensure_ascii=False, indent=2) + "\n").encode("utf-8")


def pre_compat_definition():
    # Scoped bypass of ONLY the final lexical pass, not any generator logic.
    with patch.object(builder, "transform_json", side_effect=lambda value: value):
        return builder.build_definition()


class ClipboardTransformTests(unittest.TestCase):
    def test_empty_string_type_regression_uses_existing_emulator(self):
        expressions = Expressions(None)
        result = expressions.call("string", [None])
        self.assertIs(type(result), str)
        self.assertEqual(result, "")
        before = expressions.resolve("@''")
        after = expressions.resolve(transform_json("@''"))
        assert_typed_equal(self, before, after)
        self.assertIsNot(after, None)
        self.assertIs(expressions.resolve("@null"), None)

    def test_only_complete_empty_literal_tokens(self):
        source = "@concat('', '''', 'don''t', 'x''''y', 'literal string(null)', '')"
        expected = "@concat(string(null), '''', 'don''t', 'x''''y', 'literal string(null)', string(null))"
        self.assertEqual(transform_json(source), expected)
        self.assertEqual(Expressions(None).resolve(source), Expressions(None).resolve(expected))

    def test_quote_runs_are_not_multiple_empty_tokens(self):
        for count in range(1, 16):
            source = "@" + "'" * (count * 2)
            expected = "@string(null)" if count == 1 else source
            self.assertEqual(transform_json(source), expected)
            self.assertEqual(Expressions(None).resolve(expected), "'" * (count - 1))

    def test_multiple_empty_args_nested_functions_and_whitespace(self):
        source = "@concat( '' , replace(concat('a', ''), 'a', ''), if(true, '', coalesce(null, '')) ) \n"
        expected = "@concat( string(null) , replace(concat('a', string(null)), 'a', string(null)), if(true, string(null), coalesce(null, string(null))) ) \n"
        transformed, offsets = transform_string(source)
        self.assertEqual(transformed, expected)
        self.assertEqual(len(offsets), 5)
        assert_typed_equal(self, Expressions(None).resolve(source), Expressions(None).resolve(transformed))

    def test_quoted_at_braces_commas_parentheses_remain_quoted(self):
        source = "@concat('literal @{concat('''', '''')} } ), [] ?', '')"
        expected = source[:-3] + "string(null))"
        self.assertEqual(transform_json(source), expected)
        self.assertEqual(Expressions(None).resolve(source), Expressions(None).resolve(expected))

    def test_plain_values_types_schema_literals_and_names(self):
        value = {
            "''": "ordinary '' and alice@example.com", "don't": "''",
            "@odata.type": "schema property", "@@odata.nextLink": "@@odata.type",
            "null": None, "bool": False, "int": 3, "float": 1.5,
            "list": [None, True, 0, [], {}, "", "email@example.com ''"],
            "schema": {"type": "object", "properties": {"''": {"type": "string"}},
                       "required": ["''"], "enum": ["''", "don''t"]},
        }
        changes = []
        copied = transform_json(value, changes=changes)
        assert_typed_equal(self, value, copied)
        self.assertEqual(changes, [])
        self.assertIsNot(copied, value)
        copied["list"].append("not in input")
        self.assertEqual(len(value["list"]), 7)

    def test_escaped_leading_at_is_opaque(self):
        for value in ("@@concat('', '')", "@@{''}", "@@@{''}", "@@", "@@replace('',)",
                      "@@literal @{''}"):
            self.assertEqual(transform_json(value), value)
            self.assertEqual(Expressions(None).resolve(value), value[1:])

    def test_interpolation_whole_multiple_and_surrounding_literal_text(self):
        cases = {
            "@{''}": "@{string(null)}",
            "@{ concat('', '''') }": "@{ concat(string(null), '''') }",
            "before '' alice@example.com @{replace('a-b', '-', '')} / @{concat('', 'don''t')} after ''":
                "before '' alice@example.com @{replace('a-b', '-', string(null))} / @{concat(string(null), 'don''t')} after ''",
            "quoted '' @{concat('}', '')} braces { stay }":
                "quoted '' @{concat('}', string(null))} braces { stay }",
            "literal @@{''} and @{''}": "literal @@{''} and @{string(null)}",
            "literal @@@{''}": "literal @@@{string(null)}",
        }
        for source, expected in cases.items():
            with self.subTest(source=source):
                self.assertEqual(transform_json(source), expected)
                assert_typed_equal(self, resolve_with_interpolation(source), resolve_with_interpolation(expected))

    def test_interpolation_type_coercion_is_distinct_from_whole_expressions(self):
        for expression in ("null", "false", "17", "''", "coalesce(null, '')"):
            whole = "@" + expression
            interpolation = "@{" + expression + "}"
            assert_typed_equal(self, Expressions(None).resolve(whole),
                               Expressions(None).resolve(transform_json(whole)))
            before = resolve_with_interpolation(interpolation)
            after = resolve_with_interpolation(transform_json(interpolation))
            assert_typed_equal(self, before, after)
            self.assertIs(type(after), str)
        self.assertIs(type(Expressions(None).resolve("@17")), int)
        self.assertEqual(resolve_with_interpolation("@{17}"), "17")

    def test_expression_keys_and_empty_property_selectors(self):
        source = {"@concat('key', '')": {"@{concat('v', '')}": "@''"}, "literal''": 1}
        expected = {"@concat('key', string(null))": {"@{concat('v', string(null))}": "@string(null)"}, "literal''": 1}
        changes = []
        self.assertEqual(transform_json(source, changes=changes), expected)
        self.assertEqual(sum(c.location == "key" for c in changes), 2)
        assert_typed_equal(self, resolve_with_interpolation(source), resolve_with_interpolation(expected))
        self.assertEqual(transform_json("@triggerBody()?['']"), "@triggerBody()?[string(null)]")
        self.assertEqual(transform_json("@triggerBody()['don''t']"), "@triggerBody()['don''t']")

    def test_collision_with_transformed_or_unchanged_key_in_both_orders(self):
        for first, second in (("@''", "@string(null)"), ("@{''}", "@{string(null)}")):
            for keys in ((first, second), (second, first)):
                source = {"nested": dict(zip(keys, [1, 2]))}
                before = deepcopy(source)
                changes = []
                with self.assertRaisesRegex(TransformError, "collision"):
                    transform_json(source, changes=changes)
                self.assertEqual(source, before)
                self.assertEqual(changes, [])

    def test_property_access_and_number_syntax_without_refactoring(self):
        cases = {
            "@body('a').nested?['']": "@body('a').nested?[string(null)]",
            "@body('a')?.nested[0]": "@body('a')?.nested[0]",
            "@fn(-1, 1.25, 1e-3, true, false, null, '')":
                "@fn(-1, 1.25, 1e-3, true, false, null, string(null))",
        }
        for source, expected in cases.items():
            self.assertEqual(transform_json(source), expected)

    def test_malformed_recognized_expressions_fail_closed(self):
        invalid = (
            "@replace('a', '-', )", "@concat('',)", "@concat(, '')", "@concat('', 'x'",
            "@concat('', 'unterminated)", "@concat('', 'x') trailing", "@concat('' '')",
            "@concat(''", "@body('a')?", "@body('a')?[]", "@body('a')['']junk",
            "@fn('', +1)", "@fn('', 01)", "@fn('', 1e)", "@fn('', {})",
            "@fn('', \"double quoted\")", "@{''", "prefix @{concat('',)} suffix",
            "@{concat('', @{''})}", "prefix @{''} then @{broken(}", "@{}", "@{ }",
        )
        for source in invalid:
            with self.subTest(source=source):
                definition = {"first": "@''", "nested": [source]}
                before = deepcopy(definition)
                changes = []
                with self.assertRaises(TransformError):
                    transform_json(definition, changes=changes)
                self.assertEqual(definition, before)
                self.assertEqual(changes, [])

    def test_malformed_expression_keys_fail_closed(self):
        for key in ("@concat('',)", "prefix @{''", "@body('a')?[]"):
            with self.assertRaises(TransformError):
                transform_json({key: "ordinary"})

    def test_excessive_nesting_is_an_explicit_refusal(self):
        expression = "@" + "concat(" * 1500 + "''" + ")" * 1500
        with self.assertRaises(TransformError):
            transform_json(expression)
        nested = None
        for _ in range(1500):
            nested = [nested]
        with self.assertRaisesRegex(TransformError, "JSON nesting"):
            transform_json(nested)

    def test_non_json_values_fail_without_partial_audit(self):
        for value in (float("nan"), float("inf"), float("-inf"), (1, 2), {1: "@''"}, object()):
            changes = ["existing audit sentinel"]
            with self.assertRaises(TransformError):
                transform_json(["@''", value], changes=changes)
            self.assertEqual(changes, ["existing audit sentinel"])

    def test_json_pointer_escaping_and_audit_offsets(self):
        changes = []
        source = {"a/b~c": ["@concat('', '', '''')", {"@''": None}]}
        transform_json(source, changes=changes)
        self.assertEqual([c.path for c in changes], ["/a~1b~0c/0", "/a~1b~0c/1/@''"])
        self.assertEqual(changes[0].token_offsets, (8, 12))
        self.assertEqual([c.location for c in changes], ["value", "key"])

    def test_idempotence_including_interpolation_and_keys(self):
        source = {"@''": ["@{''}", "before @{concat('', '')} after ''", "@@{''}", None]}
        once = transform_json(source)
        changes = []
        twice = transform_json(once, changes=changes)
        assert_typed_equal(self, once, twice)
        self.assertEqual(changes, [])

    def test_representative_adversarial_expression_semantics(self):
        expressions = (
            "@replace('a--b', '-', '')", "@replace('don''t', '''', '')",
            "@coalesce(null, '')", "@if(false, '', 'value')", "@equals('', string(null))",
            "@concat('', '日本語', '', '''', '', 'don''t')", "@string('')",
            "@length('')", "@empty('')", "@endsWith('abc', '')", "@contains('abc', '')",
            "@startsWith('', '')", "@indexOf('', '')", "@json(concat('[', '', ']'))",
            "@json(concat('{', '', '}'))", "@if(true, null, '')", "@if(true, 3, '')",
            "@if(true, false, '')", "@addProperty(json('{}'), 'empty', '')",
        )
        for source in expressions:
            with self.subTest(source=source):
                assert_typed_equal(self, Expressions(None).resolve(source),
                                   Expressions(None).resolve(transform_json(source)))

    def test_deterministic_quote_and_nested_expression_matrix(self):
        literals = ("", "'", "''", "don't", "@{''}", "}),(''", "日本語", "string(null)")
        for a, b in itertools.product(literals, repeat=2):
            quoted_a = "'" + a.replace("'", "''") + "'"
            quoted_b = "'" + b.replace("'", "''") + "'"
            source = "@concat(" + quoted_a + ", '', " + quoted_b + ", coalesce(null, ''))"
            transformed = transform_json(source)
            assert_typed_equal(self, Expressions(None).resolve(source), Expressions(None).resolve(transformed))
            self.assertEqual(transform_json(transformed), transformed)

    def test_empty_property_selector_semantics_with_existing_call_bridge(self):
        # The existing emulator only parses literal bracket keys. Resolve the
        # new selector with its unchanged string function, then use that key.
        engine = Flow("synthetic", {})
        engine.request[""] = [None, False, 3, ""]
        original = Expressions(engine).resolve("@triggerBody()?['']")
        key = Expressions(engine).resolve("@string(null)")
        assert_typed_equal(self, original, engine.request[key])


class CanonicalGenerationTests(unittest.TestCase):
    def test_full_action_isolation_against_independent_quote_tokenizer(self):
        original = pre_compat_definition()
        before = deepcopy(original)
        changes = []
        transformed = transform_json(original["actions"], changes=changes)
        assert_typed_equal(self, original, before)
        assert_typed_equal(self, transformed, SOURCE["actions"])
        expected_paths = {}

        def compare(a, b, path=""):
            self.assertIs(type(a), type(b), path)
            if isinstance(a, dict):
                self.assertEqual(list(a), list(b), path)  # No canonical keys change.
                for key in a:
                    compare(a[key], b[key], path + "/" + key.replace("~", "~0").replace("/", "~1"))
            elif isinstance(a, list):
                self.assertEqual(len(a), len(b), path)
                for index, (x, y) in enumerate(zip(a, b)):
                    compare(x, y, path + "/" + str(index))
            elif isinstance(a, str) and a.startswith("@") and not a.startswith("@@"):
                # Canonical input has only whole expressions, including @{...}.
                # Independently tokenize quoted literals, NOT production spans
                # and NOT global apostrophe-pair replacement. This deliberately
                # does not claim to parse arbitrary interpolated ordinary text.
                if a.startswith("@{"):
                    self.assertTrue(a.endswith("}"))
                tokens = list(re.finditer(r"'(?:''|[^'])*'|[^']+", a))
                self.assertEqual("".join(t[0] for t in tokens), a)
                offsets = tuple(t.start() for t in tokens if t[0] == "''")
                expected = "".join(REPLACEMENT if t[0] == "''" else t[0] for t in tokens)
                self.assertEqual(b, expected, path)
                if offsets:
                    expected_paths[path] = offsets
            else:
                self.assertEqual(a, b, path)

        compare(original["actions"], transformed)
        self.assertEqual({c.path: c.token_offsets for c in changes}, expected_paths)
        self.assertEqual((len(changes), sum(c.replacements for c in changes)), (37, 1149))
        self.assertTrue(all(c.location == "value" for c in changes))

    def test_generation_is_deterministic_and_matches_validated_candidate_bytes(self):
        first = render(builder.build_definition())
        self.assertEqual(first, render(builder.build_definition()))
        self.assertEqual(first, builder.DESTINATION.read_bytes())
        self.assertEqual(hashlib.sha256(first).hexdigest(), CANDIDATE_SHA256)
        self.assertEqual(hashlib.sha256(render(pre_compat_definition())).hexdigest(), ORIGINAL_SHA256)

    def test_all_non_actions_remain_exactly_pre_transform_output(self):
        before = pre_compat_definition()
        after = builder.build_definition()
        assert_typed_equal(self, {k: v for k, v in before.items() if k != "actions"},
                           {k: v for k, v in after.items() if k != "actions"})

    def test_pass_runs_last_and_does_not_visit_trigger_parameters_or_other_top_level_values(self):
        from build_draft_tools_flow import augment_draft_tools

        def augment(definition):
            result = augment_draft_tools(definition)
            # Synthetic malformed expressions must remain opaque outside actions.
            result["triggers"]["synthetic"] = {"expression": "@concat('',)"}
            result["parameters"]["synthetic"] = {"defaultValue": "@''"}
            result["synthetic"] = {"@''": "prefix @{broken(}"}
            result["actions"]["synthetic"] = {"type": "Compose", "inputs": "@''"}
            return result

        with patch("build_draft_tools_flow.augment_draft_tools", side_effect=augment), \
                patch.object(builder, "transform_json", wraps=transform_json) as final_pass:
            result = builder.build_definition()
        final_pass.assert_called_once()
        self.assertEqual(final_pass.call_args.args[0]["synthetic"]["inputs"], "@''")
        self.assertEqual(result["actions"]["synthetic"]["inputs"], "@string(null)")
        self.assertEqual(result["triggers"]["synthetic"], {"expression": "@concat('',)"})
        self.assertEqual(result["parameters"]["synthetic"], {"defaultValue": "@''"})
        self.assertEqual(result["synthetic"], {"@''": "prefix @{broken(}"})

    def test_canonical_actions_are_idempotent_and_have_no_remaining_empty_literal_tokens(self):
        changes = []
        assert_typed_equal(self, SOURCE["actions"], transform_json(SOURCE["actions"], changes=changes))
        self.assertEqual(changes, [])
        # Literal strings and escaped apostrophes still contain apostrophe pairs;
        # a substring search would falsely reject correct expression content.
        self.assertIn("''", json.dumps(SOURCE["actions"]))

    def test_transformed_canonical_keeps_public_expression_length_bound(self):
        lengths = [string_length(value) for value in walk(SOURCE)
                   if isinstance(value, str) and value.startswith("@")]
        self.assertLessEqual(max(lengths), 8192)


if __name__ == "__main__":
    unittest.main()
