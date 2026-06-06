# CLAUDE.md - mailcompiler

## CRITICAL: never put real data in code

IMPORTANT: mailcompiler processes databases of real people (names, emails,
phones). NEVER copy any of that data into source code, comments, docstrings,
test fixtures, commit messages, or PR text. This is a hard line, not a style
preference - code is committed and shared, so embedded PII leaks permanently.

- In any example (comment, docstring, test, commit), use ONLY synthetic
  placeholders: reserved domains (`example.com`, `example.mil`, `.invalid`) and
  clearly-fictional names (Wile E. Coyote, Jane Q. Public, Foo Bar, McFly).
- Never lift a value you saw while inspecting a dataset into an artifact. When
  documenting a parser/transform, invent the input - do not paste a real one.
- Program OUTPUT (built DBs, exports, `-v` discard logs) legitimately contains
  real data - that is fine. The rule is about source/committed files.

## Overview

mailcompiler is a single-module Python CLI (`mc`) that aggregates contacts from
many sources (mbox/PST mailboxes, CSV, Outlook CSV/XLSX, LinkedIn, vCard),
de-duplicates and cleans them, and exports (CSV/XLSX/Outlook/vCard). The whole
implementation is `mailcompiler/mailcompiler.py`.

## Conventions

- Run `flake8` on changed files before declaring done; CI fails on any warning.
  Config: `.flake8` (max-line-length 120).
- **Unicode is intentional and required** - contact names contain accented and
  non-Latin characters (e.g. diacritics, CJK), and the name-parsing code is
  deliberately Unicode-aware. Do NOT impose an ASCII-only rule or strip
  non-ASCII letters from name handling.
- Edit surgically; match the surrounding style.

## Key behaviors (so changes stay consistent)

- **Merge, never wipe:** importing folds into the existing `-o` DB by default;
  `--force` overwrites overlapping fields. A `json -> same json` import
  normalizes in place.
- **`--reconcile`** cleans + merges a json DB: drops junk/role/bot addresses,
  scrubs control characters, merges duplicates by shared email and by name,
  recovers/normalizes names, picks the best primary email, and requires both a
  first AND last name (records missing either are dropped).
- **Name parsing** (`split_name` and helpers) applies semantic rules: honorific
  prefix/suffix stripping, nobiliary-particle surnames (`van der ...`),
  bracket/paren tag removal, Unicode-aware casing (hyphen/apostrophe/Mc), and
  surname recovery + cross-validation from the email local-part for weak names.
- **Mailbox ingest** skips Spam- and Trash-labeled messages (not real
  correspondence).
- **`-v`/`--verbose`** prints every discard/action: on import, each skipped
  email and why; with `--reconcile`, every drop/merge/field change.
