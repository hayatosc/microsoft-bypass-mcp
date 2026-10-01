# Power Automate compatibility fixture

`microsoft-bypass-flow.pre-attachments.json` is the unchanged, sanitized
three-operation export originally committed in `0da7009`. It is historical test
input and generator baseline, not a second flow or a deployment artifact.

The canonical definition is
[`power-automate/microsoft-bypass-flow/definition.json`](../../power-automate/microsoft-bypass-flow/definition.json).
Run `python3 scripts/build_attachment_flow.py` to extend the baseline into that
file. Tests pin this fixture's checksum and require the generated source minus
its two attachment cases and enum entries to equal it exactly. Do not refresh
this fixture from the generated definition or commit raw exports here.
