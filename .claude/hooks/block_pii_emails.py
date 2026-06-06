#!/usr/bin/env python3
"""PreToolUse guard: block edits that would write a real email address into
mailcompiler source/doc files.

mailcompiler processes databases of real people. Embedding a real address into
committed code/comments/tests leaks PII permanently. This hook lets the harness
mechanically reject such an edit, regardless of model behavior.

Reads the PreToolUse JSON payload on stdin. Exit 0 = allow; exit 2 = block
(stderr is shown to Claude as the reason). Only enforced for text/source files;
reserved placeholder domains (example.com, .invalid, .test, ...) are allowed.
"""
import json
import re
import sys

# Files we guard. Other paths (binary, data fixtures) are left alone.
GUARDED_SUFFIXES = (".py", ".pyi", ".md", ".rst", ".sh", ".txt", ".cfg",
                    ".toml", ".ini", ".yaml", ".yml")

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")

# RFC 2606 / 6761 reserved, non-routable, and explicit example domains.
PLACEHOLDER_DOMAINS = {
    "example.com", "example.org", "example.net", "example.edu", "example.mil",
    "localhost",
}
PLACEHOLDER_TLDS = (".example", ".invalid", ".test", ".localhost")


def is_placeholder(addr):
    domain = addr.split("@", 1)[1].lower()
    return domain in PLACEHOLDER_DOMAINS or domain.endswith(PLACEHOLDER_TLDS)


def new_text(tool_name, tool_input):
    """The text this tool would introduce."""
    if tool_name == "Write":
        return tool_input.get("content", "") or ""
    if tool_name == "Edit":
        return tool_input.get("new_string", "") or ""
    if tool_name == "MultiEdit":
        return "\n".join(e.get("new_string", "") or ""
                         for e in tool_input.get("edits", []))
    return ""


def main():
    try:
        payload = json.load(sys.stdin)
    except Exception:
        sys.exit(0)                       # fail open: never break the toolchain

    tool_input = payload.get("tool_input", {}) or {}
    path = (tool_input.get("file_path") or "").lower()
    if path and not path.endswith(GUARDED_SUFFIXES):
        sys.exit(0)

    text = new_text(payload.get("tool_name", ""), tool_input)
    offenders = sorted({m.group(0) for m in EMAIL_RE.finditer(text)
                        if not is_placeholder(m.group(0))})
    if offenders:
        sys.stderr.write(
            "BLOCKED: this edit puts what looks like a real email address into "
            "a source file:\n  " + "\n  ".join(offenders) + "\n"
            "Source code must never contain real contact data (PII). Use a "
            "synthetic placeholder instead (example.com / .invalid, or a "
            "fictional name). Program output may contain real data; source code "
            "may not.\n")
        sys.exit(2)
    sys.exit(0)


if __name__ == "__main__":
    main()
