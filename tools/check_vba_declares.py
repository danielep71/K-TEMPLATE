#!/usr/bin/env python3
"""Check native Declare statements against their register, and how their errors are captured.

docs/native-declarations.json is the register of the native exports the
project declares. ``Err.LastDllError`` holds the error of the most recent
``Declare`` call; ``Err.Clear``, any ``On Error`` statement and the next
native call replace it, so it is trustworthy only when it is copied the
moment the call returns. The rules, applied in each compilation environment
of check_vba_conditionals.py, so that one branch of an ``#If VBA7`` split
cannot drift while another stays correct:

- every ``Declare`` names an export of the register: its ``Alias`` if it has
  one, else its own name, compared ignoring case, from the library the
  register gives for it;
- ``GetLastError`` is never declared: it does not see the error that VBA has
  already copied into ``Err.LastDllError``;
- ``Err.LastDllError`` is read only as ``<local> = Err.LastDllError``, into a
  ``Dim`` or ``Static`` variable of the procedure, in the statement directly
  after a call to a ``Declare``; a ``.LastDllError`` read inside a
  ``With Err`` block, or through another qualifier, is rejected, since the
  ``With`` statement already stands between the call and the read;
- a call to an export the register marks ``zero_is_ambiguous``, whose zero
  result may mean success or failure, assigns its result, and nothing else,
  to a local variable in the call statement, is directly preceded by ``SetLastError 0`` (or
  ``Call SetLastError(0)``) through a ``Declare`` of the ``SetLastError``
  export, and is directly followed by that capture.

A ``Declare`` is private to its module (check_vba_public_api.py rejects a
public one in a standard module), so a call resolves to a ``Declare`` of the
calling module. The check reads statements in order, not control flow: a jump
into the middle of a sequence, and an error replaced inside the call itself,
are not seen. It does not compare a ``Declare``'s parameters or return type
with the export's signature. docs/VBA_HOUSE_STYLE.md, Static checks, lists
its limits.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections.abc import Iterator, Sequence
from dataclasses import dataclass, field
from pathlib import Path

from check_vba_conditionals import reachable_sources
from check_vba_identifiers import (
    END_PROC,
    NAME,
    PROC,
    STATEMENT_WORDS,
    SUFFIX,
    closing_parenthesis,
    declared_names,
)
from vba_lexer import (
    Finding,
    logical_units,
    read_component,
    report,
    split_statements,
    tracked_components,
)

ROOT = Path(__file__).resolve().parents[1]
REGISTER = "docs/native-declarations.json"
REGISTER_KEYS = frozenset({"schema_version", "declarations"})
ENTRY_KEYS = frozenset({"export", "library", "zero_is_ambiguous"})
# Kept for the signature of each export, which the check does not read yet:
# an entry that holds one fails, so that no signature is recorded unchecked.
RESERVED_KEYS = frozenset({"arguments", "returns"})
# An exported name, or an ordinal such as #12.
EXPORT = re.compile(r"[A-Za-z_][A-Za-z0-9_@$?]*|#[0-9]+")
DECLARE = re.compile(rf"^\s*(?:(?:Public|Private|Global)\s+)?Declare\s+(?:PtrSafe\s+)?"
                     rf"(?:Sub|Function)\s+(?P<name>{NAME}){SUFFIX}", re.I)
LIBRARY = re.compile(r'\s+Lib\s+"(?P<library>[^"]*)"(?:\s+Alias\s+"(?P<alias>[^"]*)")?', re.I)
LOCAL = re.compile(r"^\s*(?:Dim|Static)\s+(?P<list>.*)$", re.I)
CONST = re.compile(r"^\s*Const\s", re.I)
LABEL = re.compile(rf"^\s*(?:(?P<name>{NAME})|\d+):(?!=)")
LINE_NUMBER = re.compile(r"^\s*\d+(?=\s)")
# Any read of a LastDllError member: Err.LastDllError, .LastDllError in a
# With Err block, or one through another qualifier such as VBA.Err.
LAST_DLL_ERROR = re.compile(r"\.\s*LastDllError\b", re.I)
PLAIN_LAST_DLL_ERROR = re.compile(r"(?<![\w.!])Err\s*\.\s*LastDllError\b", re.I)
CAPTURE = re.compile(rf"^\s*(?:Let\s+)?(?P<local>{NAME})\s*=\s*Err\s*\.\s*LastDllError\s*$",
                     re.I)
CLEAR = re.compile(rf"^\s*(?:Call\s+)?(?P<name>{NAME})\s*(?:\(\s*0&?\s*\)|\s0&?)\s*$", re.I)
ON_ERROR = re.compile(r"^\s*On\s+Error\b", re.I)
# A name that is not the member of an object or of a With block.
WORD = re.compile(rf"(?<![\w.!])(?P<name>{NAME})")


@dataclass(frozen=True)
class Entry:
    """One export of the register."""

    export: str
    library: str
    zero_is_ambiguous: bool


def export_key(export: str, library: str) -> tuple[str, str]:
    """How an export is matched: ignoring case, and a library's .dll suffix."""
    library = library.strip().casefold()
    return export.casefold(), library.removesuffix(".dll")


def read_register(root: Path) -> tuple[dict[tuple[str, str], Entry], list[Finding]]:
    """The register's entries by export and library, and what is wrong with it."""
    path = root / REGISTER
    if not path.is_file():
        return {}, [Finding(REGISTER, None, "is missing; restore it with an empty "
                                            "declarations list")]
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (UnicodeDecodeError, ValueError) as error:
        return {}, [Finding(REGISTER, None, f"is not valid JSON: {error}")]
    if not isinstance(data, dict) or data.keys() != REGISTER_KEYS or \
            data["schema_version"] != 1 or not isinstance(data["declarations"], list):
        return {}, [Finding(REGISTER, None, "needs exactly schema_version 1 and a "
                                            "declarations list")]
    entries: dict[tuple[str, str], Entry] = {}
    findings: list[Finding] = []
    for index, item in enumerate(data["declarations"]):
        where = f"{REGISTER}: declarations[{index}]"
        problems = entry_problems(item)
        if not problems:
            entry = Entry(item["export"], item["library"], item["zero_is_ambiguous"])
            key = export_key(entry.export, entry.library)
            if key in entries:
                problems.append(f"{entry.export} from {entry.library} is listed twice")
            elif key[0] == "getlasterror":
                problems.append("GetLastError is never declared, so it has no entry")
            else:
                entries[key] = entry
        findings.extend(Finding(where, None, problem) for problem in problems)
    return entries, findings


def entry_problems(item: object) -> list[str]:
    """What is wrong with one entry of the declarations list."""
    if not isinstance(item, dict):
        return ["is not an object"]
    problems = [f"{key} is reserved for the export's signature, which the check does not "
                "read yet; remove it" for key in sorted(item.keys() & RESERVED_KEYS)]
    problems += [f"unknown key {key!r}" for key in sorted(item.keys() - ENTRY_KEYS -
                                                           RESERVED_KEYS)]
    problems += [f"{key} is missing" for key in sorted(ENTRY_KEYS - item.keys())]
    if problems:
        return problems
    if not isinstance(item["export"], str) or not EXPORT.fullmatch(item["export"]):
        problems.append("export is not an export name or an ordinal such as #12")
    library = item["library"]
    if not isinstance(library, str) or not library.strip() or library != library.strip() \
            or '"' in library:
        problems.append("library is not a library name, such as kernel32")
    if not isinstance(item["zero_is_ambiguous"], bool):
        problems.append("zero_is_ambiguous is not true or false")
    return problems


@dataclass(frozen=True)
class Declared:
    """One Declare statement: its VBA name and the export it imports."""

    line: int
    name: str
    export: str
    library: str


@dataclass
class Procedure:
    name: str
    locals: set[str] = field(default_factory=set)
    # Executable statements, and labels as "<label>:", in source order.
    body: list[tuple[int, str]] = field(default_factory=list)


@dataclass
class Module:
    declared: dict[str, Declared] = field(default_factory=dict)
    procedures: list[Procedure] = field(default_factory=list)
    findings: list[Finding] = field(default_factory=list)


def declaration(path: str, lines: list[str], start: int, end: int,
                module: Module) -> None:
    """Record the Declare statement of the logical line ``start``..``end``.

    The statements were read with string literals blanked; the library and
    the alias are read again from the source, where they are kept."""
    for _, _, kind, code in logical_units(lines[start - 1:end]):
        for statement in split_statements(code) if kind == "code" else []:
            head = DECLARE.match(statement)
            if head is None:
                continue
            names = LIBRARY.match(statement, head.end())
            if names is None:
                module.findings.append(Finding(path, start, "cannot read the Lib clause of "
                                                            f"Declare {head.group('name')}"))
                continue
            module.declared[head.group("name").casefold()] = Declared(
                start, head.group("name"), names.group("alias") or head.group("name"),
                names.group("library"))


def scan(path: str, source: str) -> Module:
    """The Declare statements and the procedure bodies of one environment's source."""
    lines = source.split("\n")
    module = Module()
    current: Procedure | None = None
    for start, end, kind, code in logical_units(lines, blank_strings=True):
        if kind == "directive":
            continue
        label = LABEL.match(code) or LINE_NUMBER.match(code)
        if label and (label.groupdict().get("name") or "").casefold() not in STATEMENT_WORDS:
            if current is not None:
                current.body.append((start, label.group().strip().rstrip(":") + ":"))
            code = code[label.end():]
        for statement in split_statements(code):
            if not statement:
                continue
            opener = PROC.match(statement)
            if DECLARE.match(statement):
                declaration(path, lines, start, end, module)
            elif opener:
                current = Procedure(opener.group("name"))
                module.procedures.append(current)
            elif END_PROC.match(statement):
                current = None
            elif current is not None:
                local_or_statement(current, start, statement)
    return module


def local_or_statement(procedure: Procedure, line: int, statement: str) -> None:
    """Record a Dim or Static as locals; skip a Const; keep any other statement."""
    local = LOCAL.match(statement)
    if local:
        procedure.locals |= {name.casefold() for name in declared_names(local.group("list"))
                             or []}
    elif not CONST.match(statement):
        procedure.body.append((line, statement))


def called(statement: str, declared: dict[str, Declared]) -> list[Declared]:
    """The Declares a statement names, each once, in order."""
    found: dict[str, Declared] = {}
    for match in WORD.finditer(statement):
        item = declared.get(match.group("name").casefold())
        if item is not None:
            found.setdefault(item.name.casefold(), item)
    return list(found.values())


def is_clear(statement: str | None, declared: dict[str, Declared]) -> bool:
    """Whether a statement is SetLastError 0, through any Declare of that export."""
    match = CLEAR.match(statement or "")
    item = declared.get(match.group("name").casefold()) if match else None
    return item is not None and item.export.casefold() == "setlasterror"


def native_call(statement: str | None, declared: dict[str, Declared]) -> bool:
    """Whether a statement calls a Declare other than SetLastError."""
    return statement is not None and any(item.export.casefold() != "setlasterror"
                                         for item in called(statement, declared))


def capture_problem(procedure: Procedure, index: int, declared: dict[str, Declared]
                    ) -> str | None:
    """What is wrong with the Err.LastDllError read of ``procedure.body[index]``."""
    statement = procedure.body[index][1]
    if len(LAST_DLL_ERROR.findall(statement)) != len(PLAIN_LAST_DLL_ERROR.findall(statement)):
        return (f"{procedure.name} reads LastDllError through a With block or a qualifier "
                "other than Err; write it as <local> = Err.LastDllError")
    capture = CAPTURE.match(statement)
    if capture is None:
        return (f"{procedure.name} reads Err.LastDllError inside a larger statement; copy "
                "it by itself to a local variable (<local> = Err.LastDllError)")
    if capture.group("local").casefold() not in procedure.locals:
        return (f"{procedure.name} copies Err.LastDllError to {capture.group('local')}, "
                "which is not a Dim or Static variable of the procedure")
    if index == 0:
        return f"{procedure.name} reads Err.LastDllError before any native call"
    line, before = procedure.body[index - 1]
    if before.endswith(":"):
        return (f"{procedure.name} reads Err.LastDllError after the label of line {line}, "
                "which a jump can reach without the native call")
    if ON_ERROR.match(before):
        return (f"{procedure.name} reads Err.LastDllError after the On Error statement of "
                f"line {line}, which clears it; read it directly after the native call")
    if not native_call(before, declared):
        return (f"{procedure.name} reads Err.LastDllError after the statement of line "
                f"{line}, not directly after a native call")
    return None


def assigned_call(statement: str, name: str) -> str | None:
    """The variable of ``<variable> = Name`` or ``<variable> = Name(<arguments>)``,
    with balanced parentheses and nothing after them; None for any other statement."""
    head = re.match(rf"^\s*(?:Let\s+)?(?P<local>{NAME})\s*=\s*{re.escape(name)}{SUFFIX}(?!\w)\s*",
                    statement, re.I)
    if head is None:
        return None
    rest = statement[head.end():]
    if rest.startswith("("):
        end = closing_parenthesis(rest, 0)
        if end is None:
            return None
        rest = rest[end + 1:]
    return head.group("local") if not rest.strip() else None


def ambiguous_problems(procedure: Procedure, index: int, item: Declared,
                       declared: dict[str, Declared]) -> list[str]:
    """What is wrong around a call to an export whose zero result is ambiguous."""
    body = procedure.body
    statement = body[index][1]
    before = body[index - 1][1] if index else None
    after = body[index + 1][1] if index + 1 < len(body) else None
    call = f"{procedure.name} calls {item.name} ({item.export}), whose zero result is ambiguous,"
    result = assigned_call(statement, item.name)
    problems = []
    if result is None or result.casefold() not in procedure.locals:
        problems.append(f"{call} without assigning its result to a local variable in the "
                        "call statement")
    if not is_clear(before, declared):
        problems.append(f"{call} without SetLastError 0 directly before it")
    capture = CAPTURE.match(after or "")
    if capture is None or capture.group("local").casefold() not in procedure.locals:
        problems.append(f"{call} without <local> = Err.LastDllError directly after it")
    return problems


def procedure_findings(path: str, procedure: Procedure, declared: dict[str, Declared],
                       register: dict[tuple[str, str], Entry]) -> Iterator[Finding]:
    """The capture findings of one procedure in one environment."""
    for index, (line, statement) in enumerate(procedure.body):
        if LAST_DLL_ERROR.search(statement):
            problem = capture_problem(procedure, index, declared)
            if problem:
                yield Finding(path, line, problem)
        for item in called(statement, declared):
            entry = register.get(export_key(item.export, item.library))
            if entry is not None and entry.zero_is_ambiguous:
                yield from (Finding(path, line, problem)
                            for problem in ambiguous_problems(procedure, index, item, declared))


def analyze_source(path: str, source: str, register: dict[tuple[str, str], Entry]
                   ) -> list[Finding]:
    """The findings of one environment's source."""
    module = scan(path, source)
    findings = list(module.findings)
    for item in module.declared.values():
        if item.export.casefold() == "getlasterror":
            findings.append(Finding(path, item.line, f"Declare {item.name} imports "
                                                     "GetLastError; copy Err.LastDllError "
                                                     "directly after the native call instead"))
        if export_key(item.export, item.library) not in register:
            findings.append(Finding(path, item.line, f"Declare {item.name} imports "
                                                     f"{item.export} from {item.library}, "
                                                     f"which {REGISTER} does not list"))
    for procedure in module.procedures:
        findings.extend(procedure_findings(path, procedure, module.declared, register))
    return findings


def analyze(path: str, text: str, register: dict[tuple[str, str], Entry]
            ) -> tuple[list[Finding], set[int]]:
    """Findings of every environment, each listed once with the environments it
    occurs in; and the lines of the Declare statements found."""
    # Unreadable directives are vba-conditionals findings, not repeated here.
    sources, _ = reachable_sources(path, text)
    environments: dict[Finding, list[str]] = {}
    declares: set[int] = set()
    for environment, source in sources.items():
        declares |= {item.line for item in scan(path, source).declared.values()}
        for finding in analyze_source(path, source, register):
            names = environments.setdefault(finding, [])
            if environment not in names:
                names.append(environment)
    return [Finding(item.path, item.line, f"{item.message} [{', '.join(names)}]")
            for item, names in environments.items()], declares


def check(root: Path) -> tuple[list[Finding], int, int, int]:
    register, findings = read_register(root)
    components = tracked_components(root)
    declares = 0
    for path in components:
        found, lines = analyze(path, read_component(root, path), register)
        findings.extend(found)
        declares += len(lines)
    return findings, len(components), declares, len(register)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components, declares, registered = check(args.root.resolve())
    return report("vba-declares", findings,
                  f"{components} VBA component(s), {declares} Declare statement(s), "
                  f"{registered} registered export(s)")


if __name__ == "__main__":
    sys.exit(main())
