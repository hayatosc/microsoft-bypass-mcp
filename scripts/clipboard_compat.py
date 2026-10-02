"""Lexical WDL empty-string compatibility for classic-designer clipboard paste.

Only complete '' tokens in recognized expressions become string(null). Source
slices preserve every other character; no expressions are evaluated. The
canonical generator applies transform_json only to its completed actions tree.
This conservative syntax parser is not a WDL runtime/signature validator:
unsupported recognized syntax fails closed, without mutating input or audit.

Supports calls, doubled-apostrophe literals, null/booleans/numbers, [] / ?[] and
. / ?. access, whole @expressions and @{expression} interpolation in text.
Bare @names (e.g. @odata.type) are literals. Leading @@ strings are opaque;
embedded @@ pairs escape an @, matching the existing public emulator contract.
"""

from dataclasses import dataclass
import math
import re

REPLACEMENT = "string(null)"
IDENTIFIER = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")
NUMBER = re.compile(r"-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?")
WHOLE_START = re.compile(
    r"@\s*(?:[A-Za-z_][A-Za-z0-9_]*\s*\(|'|-?[0-9]|(?:true|false|null)\b)"
)


class TransformError(ValueError):
    """Refuse ambiguous/invalid input rather than partially rewriting it."""


@dataclass(frozen=True)
class Change:
    path: str  # RFC 6901 pointer relative to the supplied subtree, original key
    location: str  # "value" or "key"
    token_offsets: tuple[int, ...]  # offsets in the original decoded JSON string

    @property
    def replacements(self):
        return len(self.token_offsets)


class _Expression:
    def __init__(self, text, position):
        self.text = text
        self.position = position
        self.empty_offsets = []

    def error(self, reason):
        # Do not echo expression contents in errors.
        raise TransformError(f"{reason} at string offset {self.position}")

    def space(self):
        while self.position < len(self.text) and self.text[self.position].isspace():
            self.position += 1

    def peek(self):
        self.space()
        return self.text[self.position:self.position + 1]

    def expect(self, token):
        if self.peek() != token:
            self.error(f"expected {token!r}")
        self.position += 1

    def quoted(self):
        start = self.position
        self.position += 1
        while self.position < len(self.text):
            if self.text[self.position] != "'":
                self.position += 1
            elif self.text.startswith("''", self.position):
                self.position += 2  # One escaped apostrophe, NOT an empty token.
            else:
                self.position += 1
                if self.position == start + 2:
                    self.empty_offsets.append(start)
                return
        self.error("unterminated quoted literal")

    def expression(self):
        char = self.peek()
        if char == "'":
            self.quoted()
        elif match := NUMBER.match(self.text, self.position):
            self.position = match.end()
        elif match := IDENTIFIER.match(self.text, self.position):
            self.position = match.end()
            if match[0] not in {"true", "false", "null"}:
                self.expect("(")
                if self.peek() != ")":
                    self.expression()
                    while self.peek() == ",":
                        self.position += 1
                        self.expression()
                self.expect(")")
        else:
            self.error("expected expression operand")
        while self.peek() in {"?", "[", "."}:
            if self.peek() == "?":
                self.position += 1
                if self.peek() not in {"[", "."}:
                    self.error("expected property access after ?")
            if self.peek() == "[":
                self.position += 1
                self.expression()
                self.expect("]")
            else:
                self.expect(".")
                match = IDENTIFIER.match(self.text, self.position)
                if not match:
                    self.error("expected property name")
                self.position = match.end()


def transform_string(text):
    """Return (rewritten text, original empty-token offsets); never evaluate WDL."""
    offsets = []
    try:
        if text.startswith("@@"):
            return text, ()
        if WHOLE_START.match(text):
            parser = _Expression(text, 1)
            parser.expression()
            if parser.peek():
                parser.error("unexpected trailing expression text")
            offsets.extend(parser.empty_offsets)
        else:
            position = 0
            while position < len(text):
                if text.startswith("@@", position):
                    position += 2
                elif text.startswith("@{", position):
                    parser = _Expression(text, position + 2)
                    parser.expression()
                    parser.expect("}")
                    offsets.extend(parser.empty_offsets)
                    position = parser.position
                else:
                    position += 1
    except RecursionError:
        raise TransformError("expression nesting exceeds local parser limit") from None
    pieces, previous = [], 0
    for offset in offsets:
        pieces.extend((text[previous:offset], REPLACEMENT))
        previous = offset + 2
    pieces.append(text[previous:])
    return "".join(pieces), tuple(offsets)


def _pointer(parent, token):
    return parent + "/" + str(token).replace("~", "~0").replace("/", "~1")


def transform_json(value, *, changes=None):
    """Copy a JSON subtree, replacing expression empty tokens in values and keys.

    Optionally append Change records only on full success. Preserve object order;
    refuse key collisions and non-JSON values. The caller owns subtree selection.
    """
    pending_changes = []

    def string(value, path, location):
        try:
            result, offsets = transform_string(value)
        except TransformError as error:
            raise TransformError(f"{location} at JSON pointer {path!r}: {error}") from None
        if offsets:
            pending_changes.append(Change(path, location, offsets))
        return result

    def visit(value, path):
        if isinstance(value, str):
            return string(value, path, "value")
        if isinstance(value, list):
            return [visit(child, _pointer(path, index)) for index, child in enumerate(value)]
        if isinstance(value, dict):
            result = {}
            for key, child in value.items():
                if not isinstance(key, str):
                    raise TransformError("JSON object keys must be strings")
                child_path = _pointer(path, key)
                transformed_key = string(key, child_path, "key")
                if transformed_key in result:
                    raise TransformError(f"transformed key collision at JSON pointer {path!r}")
                result[transformed_key] = visit(child, child_path)
            return result
        if value is None or type(value) in (bool, int):
            return value
        if type(value) is float and math.isfinite(value):
            return value
        raise TransformError(f"non-JSON value at JSON pointer {path!r}")

    try:
        result = visit(value, "")
    except RecursionError:
        raise TransformError("JSON nesting exceeds local traversal limit") from None
    if changes is not None:
        changes.extend(pending_changes)
    return result
