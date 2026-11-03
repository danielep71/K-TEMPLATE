#!/usr/bin/env python3
"""Reject tracked binary files outside the approved paths and the allowlist.

A file is binary when .gitattributes unsets its text attribute (the
office-doc, opaque and vba-form-bin rules, the byte-preserved regulatory
sources and the variant files of file-based test cases), or when a file Git
would treat as text contains a NUL byte anywhere, also past the first 8000
bytes, which are all that Git inspects. A binary file passes only if it is:

- under docs/methodology/sources/, the byte-preserved regulatory sources;
- in a variant directory of a file-based test case,
  tests/fixtures/<case-id>/<variant>/, where tools/check_test_cases.py admits
  only the files the data contract names;
- src/workbook/Template.xlsx, the macro-free workbook template;
- a UserForm .frx with its .frm beside it; or
- listed in .github/binary-allowlist.json with its current SHA-256.

.gitignore keeps such files out by default; this check catches one added on
purpose, for example with git add -f.

An Office Open XML package that passes (a .xlsx, .xltx, .docx or similar
file, as a ZIP with [Content_Types].xml) also fails when it carries macros, a
vbaProject.bin part or a reference to one, or personal document metadata: an
author or last editor in docProps/core.xml, or a company or manager in
docProps/app.xml. Each part is parsed as XML in its declared encoding, so a
UTF-16 part or a character reference hides nothing, and a part that is not
XML fails too. The byte-preserved regulatory sources under
docs/methodology/sources/ are exempt: they stay as published. It never
changes a file.
"""

from __future__ import annotations

import argparse
import fnmatch
import hashlib
import json
import re
import subprocess
import sys
import zipfile
from collections.abc import Sequence
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

ROOT = Path(__file__).resolve().parents[1]
ALLOWLIST = Path(".github/binary-allowlist.json")
# fnmatch's * also matches "/": the second pattern needs a variant directory.
APPROVED_PATHS = ("docs/methodology/sources/*", "tests/fixtures/*/*/*",
                  "src/workbook/Template.xlsx")
# Read in chunks: Git looks at the first 8000 bytes only, so a NUL byte later in
# the file would otherwise stay unseen.
CHUNK_BYTES = 1 << 16
# Office packages are inspected for macros and personal metadata, except the
# byte-preserved regulatory sources.
SOURCE_COPIES = "docs/methodology/sources/"
# A VBA part, its content type or relationship, or a macro-enabled main part.
MACRO_PART = re.compile(r"vbaProject|macroEnabled", re.I)
# docProps fields that name a person or an organization, by part.
PERSONAL_FIELDS = {"docProps/core.xml": ("creator", "lastModifiedBy"),
                   "docProps/app.xml": ("Company", "Manager")}


def git(root: Path, *args: str, stdin: str | None = None) -> str:
    return subprocess.run(["git", *args], cwd=root, input=stdin, text=True,
                          capture_output=True, check=True).stdout


def tracked_files(root: Path) -> list[str]:
    return [name for name in git(root, "ls-files", "-z").split("\0") if name]


def text_attributes(root: Path, names: Sequence[str]) -> dict[str, str]:
    """The value of the text attribute for each path: set, unset, auto or unspecified."""
    fields = git(root, "check-attr", "-z", "--stdin", "text",
                 stdin="\0".join(names) + "\0").split("\0")
    return {fields[index]: fields[index + 2] for index in range(0, len(fields) - 2, 3)}


def has_nul(path: Path) -> bool:
    """Whether the file holds a NUL byte anywhere."""
    with path.open("rb") as handle:
        while chunk := handle.read(CHUNK_BYTES):
            if b"\0" in chunk:
                return True
    return False


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def load_allowlist(root: Path, findings: list[str]) -> dict[str, str]:
    """Map each allowlisted path to its SHA-256; report a malformed file."""
    path = root / ALLOWLIST
    if not path.is_file():
        return {}
    try:
        data: Any = json.loads(path.read_text(encoding="utf-8"))
    except (UnicodeDecodeError, ValueError) as error:
        findings.append(f"{ALLOWLIST.as_posix()}: not valid UTF-8 JSON ({error})")
        return {}
    files = data.get("files") if isinstance(data, dict) else None
    if not isinstance(data, dict) or data.get("schema_version") != 1 or not isinstance(files, dict):
        findings.append(f"{ALLOWLIST.as_posix()}: needs schema_version 1 and a files object")
        return {}
    allowed: dict[str, str] = {}
    for name, entry in files.items():
        digest = entry.get("sha256") if isinstance(entry, dict) else None
        reason = entry.get("reason") if isinstance(entry, dict) else None
        if not isinstance(digest, str) or len(digest) != 64 or not isinstance(reason, str) \
                or not reason.strip():
            findings.append(f"{ALLOWLIST.as_posix()}: {name} needs a 64-character sha256 "
                            "and a reason")
            continue
        allowed[name] = digest.lower()
    return allowed


def local_name(tag: str) -> str:
    """An element or attribute name without its namespace."""
    return tag.rpartition("}")[2]


def read_part(data: bytes) -> ElementTree.Element | None:
    """The parsed part, read in its declared encoding, or None when it is not XML.

    The parser honours a byte-order mark and the XML declaration, so a UTF-16
    part is read like a UTF-8 one, and it resolves character references.
    """
    try:
        return ElementTree.fromstring(data)
    except ElementTree.ParseError:
        return None


def macro_reference(part: ElementTree.Element) -> bool:
    """Whether a name, attribute value or text in the part refers to macros."""
    return any(MACRO_PART.search(text)
               for element in part.iter()
               for text in (element.tag, element.text or "", *element.attrib.values()))


def personal_fields(part: ElementTree.Element, fields: Sequence[str]) -> list[str]:
    """The fields that have a non-blank value in the part, whatever their namespace."""
    return [field for field in fields
            if any(local_name(element.tag) == field and (element.text or "").strip()
                   for element in part.iter())]


def package_problems(path: Path) -> list[str]:
    """Macros or personal metadata in an Office Open XML package; none for other files."""
    try:
        with zipfile.ZipFile(path) as package:
            names = package.namelist()
            if "[Content_Types].xml" not in names:
                return []
            data = {name: package.read(name) for name in names
                    if name == "[Content_Types].xml" or name.endswith(".rels")
                    or name in PERSONAL_FIELDS}
    except (zipfile.BadZipFile, OSError):
        return []
    problems = []
    parts = {name: read_part(content) for name, content in data.items()}
    unreadable = sorted(name for name, part in parts.items() if part is None)
    if unreadable:
        problems.append(f"{', '.join(unreadable)} cannot be read as XML, so it cannot be "
                        "checked for macros or personal metadata")
    if any(MACRO_PART.search(name) for name in names) or \
            any(part is not None and macro_reference(part)
                for name, part in parts.items() if name not in PERSONAL_FIELDS):
        problems.append("carries macros (a vbaProject part, a reference to one or a "
                        "macro-enabled content type); commit "
                        "the macro-free package and keep the code as exported source")
    for name, fields in PERSONAL_FIELDS.items():
        part = parts.get(name)
        found = personal_fields(part, fields) if part is not None else []
        if found:
            problems.append(f"{name} names a person or organization ({', '.join(found)}); "
                            "clear it, as DISTRIBUTION.md describes")
    return problems


def approved(name: str, tracked: set[str]) -> bool:
    if any(fnmatch.fnmatchcase(name, pattern) for pattern in APPROVED_PATHS):
        return True
    return name.endswith(".frx") and f"{name[:-4]}.frm" in tracked


def check(root: Path) -> tuple[list[str], int]:
    """Return the findings and the number of binary files found."""
    findings: list[str] = []
    names = tracked_files(root)
    tracked = set(names)
    allowed = load_allowlist(root, findings)
    attributes = text_attributes(root, names) if names else {}
    binaries = 0
    for name in names:
        path = root / name
        if not path.is_file():
            continue
        if not name.startswith(SOURCE_COPIES):
            findings.extend(f"{name}: {problem}" for problem in package_problems(path))
        # An allowlisted file is checked against its recorded hash whatever its
        # content is now, so a replacement never passes by no longer looking binary.
        if name in allowed:
            binaries += 1
            if sha256(path) != allowed[name]:
                findings.append(f"{name}: content changed since it was allowlisted; review it "
                                f"and update its sha256 in {ALLOWLIST.as_posix()}")
            continue
        declared = attributes.get(name) == "unset"
        if not declared and not has_nul(path):
            continue
        binaries += 1
        if approved(name, tracked):
            continue
        kind = "binary file" if declared else "binary content in a file treated as text"
        findings.append(f"{name}: {kind} outside the approved paths; remove it, or review it "
                        f"and add it to {ALLOWLIST.as_posix()}")
    for name in sorted(set(allowed) - tracked):
        findings.append(f"{ALLOWLIST.as_posix()}: {name} is not a tracked file")
    return findings, binaries


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, binaries = check(args.root.resolve())
    for finding in findings:
        print(f"FAIL: {finding}")
    status = "FAIL" if findings else "PASS"
    print(f"binaries: {status} ({binaries} binary file(s) tracked)")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
