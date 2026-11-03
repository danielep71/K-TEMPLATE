#!/usr/bin/env python3
"""Check the evidence of an Excel run of the VBA harness against one commit.

The harness runs only in Excel, by hand or through tools/devbuild/excel_run.py
(INSTALLATION.md, Recording an Excel run). This reads what that run left in an
evidence folder and checks it against the candidate commit, without Excel.
The evidence layout, the suite, the modes and the supported host are
configured once, in .github/excel-evidence-policy.json as the candidate commit
holds it, and docs/methodology/TEST_HARNESS.md describes them:

- source identity: the modules exported back from the workbook equal, line
  for line, the commit's VBA components, each form's .frx equals the
  commit's byte for byte, and none is missing or extra. Code outside string
  literals and comments is compared without letter case, because the VBE
  re-cases every use of a name (HB-01): a re-cased identifier is a note, not a
  finding. In a document module, the attributes Excel writes itself are not
  compared (DOCUMENT_ATTRIBUTES);
- the three logs saved with TEST_Harness.SaveLog: each holds, record for
  record, what the commit's suite writes in its mode (no record missing,
  extra or out of order, no ERROR), with cleanup PASS and no recovery: the
  normal run passes with every case and assertion; the injected-failure run
  fails only its designated assertion, with the expected value the suite
  injects, an actual value that differs and the same values in its ASSERT,
  FAILURE and PRIMARY_FAILURE records; the skipped-case run lacks only the
  last case and fails only harness.completeness; all three come from one
  host of the policy's bitness; a normal run that is INCOMPLETE is rejected as such;
- certification.log, written by TEST_Harness.EndCertification: each mode
  ended as it must, the two failure modes reaching their caller as
  ERR_TEST_FAILED, and CERTIFICATION=PASS with COMPLETE;
- each project suite the policy lists under "suites": its normal-run log
  holds what the suite writes, passes and is COMPLETE, from the same host,
  and its own certification log records that run and CERTIFICATION=PASS.
  Every suite is read as the policy's Office bitness compiles it, so a case
  or assertion inside #If Win64 counts only on the host it runs on;
- host.txt: the Windows build, the references (exactly those of the policy)
  and that Debug > Compile VBAProject completed;
- results.json, written by TEST_Harness.SaveResults, when the commit holds
  accuracy contracts (tests/accuracy/contracts.json): tools/check_accuracy.py
  measures them on it, against the commit's own test cases, so the results
  must be of the commit, cover every output of the contracts' cases, lie
  within each output's tolerance and within each contract's threshold;
- with --workbook, the binding of the export to the workbook it came from:
  export-manifest.json (tools/devbuild/export_manifest.py) records the
  SHA-256 of the workbook and of every exported file, and the workbook's
  build stamp names the commit. The binding is VERIFIED, FAIL, or UNVERIFIED
  when no workbook or manifest is given; --require-binding fails anything
  but VERIFIED.

Without an exported folder, --workbook proves source identity instead by
rebuild: the commit is checked out in a temporary worktree and built with
tools/devbuild/build_workbook.py, and the workbook Excel ran, closed without
saving, must hold the same parts as that build, in the same order and byte for
byte once decompressed (binding VERIFIED, or FAIL). The parts, not the
compressed package, are compared, because another platform's zlib may pack
the same parts into other bytes.
Excel compiles the project from source on open, so the code that ran is the
commit's. Such a run, as tools/devbuild/excel_run.py records it without Trust
access to the VBA project, may record compile=ON_DEMAND: only the modules the
run called were compiled.

It proves only that this evidence is consistent with the commit. It does not
run Excel, and a PASS validates a calculation only as far as the commit's
accuracy contracts and their reference values do.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import posixpath
import re
import subprocess
import sys
import tempfile
import zipfile
from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Any
from xml.etree import ElementTree

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))

import check_accuracy  # noqa: E402
import check_vba_conditionals  # noqa: E402
import vba_lexer  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent / "devbuild"))

import build_workbook  # noqa: E402

VBA_SUFFIXES = (".bas", ".cls", ".frm")
# A form's binary resources, exported beside its .frm and compared byte for byte.
RESOURCE_SUFFIXES = (".frx",)

# The module attributes of a document module (ThisWorkbook or a sheet, a .cls in
# build_workbook.DOCUMENT_HOME) that Excel writes itself when it exports one:
# they follow from the object the module belongs to, not from its source. The
# guided import pastes a document module's code without its Attribute lines,
# so whatever the commit's file holds, the export holds Excel's own lines,
# which it may add, drop or spell differently. VB_Name is not among them: it
# is the code name, and a different one is a different component.
DOCUMENT_ATTRIBUTES = ("VB_Base", "VB_GlobalNameSpace", "VB_Creatable", "VB_PredeclaredId",
                       "VB_Exposed", "VB_TemplateDerived", "VB_Customizable")
DOCUMENT_ATTRIBUTE = re.compile(rf"^Attribute ({'|'.join(DOCUMENT_ATTRIBUTES)})\s*=",
                                re.IGNORECASE)
IDENTIFIER = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")

# The evidence policy: the suite, its modes, the evidence layout and the
# supported host. It is read from the commit being checked.
POLICY = ".github/excel-evidence-policy.json"
# The modes the harness and the example suite implement, in run order; only
# NORMAL may pass, the others are deliberate failures.
MODE_KINDS = ("NORMAL", "INJECTED_FAILURE", "SKIPPED_CASE")
OFFICE_BITNESS = ("64-bit", "32-bit")
# The compilation environment of check_vba_conditionals.py for each bitness.
BITNESS_ENVIRONMENT = {"64-bit": "vba7-win64", "32-bit": "vba7-win32"}
MANIFEST_FORMAT = "vba-export-1"
BUILD_CONFIG = "tools/devbuild/workbook.json"
ENVIRONMENT_FIELDS = ("host", "version", "build", "os", "office", "runtime")
HOST_FIELDS = ("windows", "references", "compile")
# compile= of a run whose whole project could not be compiled: VBA compiled on
# demand the modules the run called. Accepted only with identity by rebuild.
COMPILE_ON_DEMAND = "ON_DEMAND"
REBUILT_CONFIGURATION = "development"
# What tools/check_accuracy.py reads of a commit to measure its contracts.
CASE_PATHS = ("docs/methodology", "tests/fixtures", "tests/expected",
              check_accuracy.REGISTRY.parent.as_posix())

VB_OBJECT_ERROR = -2147221504
# ERR_TEST_FAILED as vbObjectError + n, or as the literal &H8004xxxx equal to it.
TEST_FAILED = re.compile(r"^\s*Public Const ERR_TEST_FAILED\s+As Long = "
                         r"(?:vbObjectError \+ (\d+)|&H([0-9A-F]{8}))\b", re.M | re.I)
RECORD = re.compile(r"^([A-Z_]+)=(.*)$")
CONSTANT = re.compile(r"^\s*Private Const (\w+)\s+As Long = (\d+)\b", re.M)
PROCEDURE = re.compile(r"^(?:Public|Private) Sub (\w+)\(.*?^End Sub$", re.M | re.S)
BEGIN_CASE = re.compile(r'TEST_Harness\.BeginCase "([^"]+)"')
ASSERTION = re.compile(r'TEST_Harness\.Assert\w+ "([^"]+)"')
CALL = re.compile(r"^\s+(Case\w+)$", re.M)
INJECTED_VALUE = re.compile(r'If mMode = "INJECTED_FAILURE" Then\s+\w+ = ([-+.\dE]+)\s')


class PolicyError(Exception):
    """The evidence policy is missing or malformed."""


@dataclass(frozen=True)
class ProjectSuite:
    """A project suite: run once in NORMAL mode by its own certification entry point,
    which saves its log and its certification log in the evidence folder."""

    name: str
    source: str
    entry_point: str
    expected_cases: int
    expected_assertions: int
    log: str
    certification: str


@dataclass(frozen=True)
class Policy:
    """What an Excel run must leave as evidence (.github/excel-evidence-policy.json)."""

    harness: str
    suite: str
    suite_source: str
    entry_point: str
    expected_cases: int
    expected_assertions: int
    logs: dict[str, str]
    must_pass: frozenset[str]
    certification: str
    exported: str
    host: str
    manifest: str
    results: str
    references: tuple[str, ...]
    office: str
    project_suites: tuple[ProjectSuite, ...] = ()


def parse_policy(text: str) -> Policy:
    """The policy in ``text``; PolicyError names the first problem."""
    try:
        data = json.loads(text)
        suite, evidence, host = data["suite"], data["evidence"], data["host"]
        modes = [(str(item["mode"]), str(item["log"]), item["must_pass"])
                 for item in data["modes"]]
        policy = Policy(
            str(data["harness"]), str(suite["name"]), str(suite["source"]),
            str(suite["certification_entry_point"]), suite["expected_cases"],
            suite["expected_assertions"], {mode: log for mode, log, _ in modes},
            frozenset(mode for mode, _, must in modes if must is True),
            str(evidence["certification_log"]), str(evidence["exported"]),
            str(evidence["host_record"]), str(evidence["export_manifest"]),
            str(evidence["results"]),
            tuple(str(item) for item in host["references"]), str(host["office"]),
            tuple(ProjectSuite(str(item["name"]), str(item["source"]),
                               str(item["certification_entry_point"]), item["expected_cases"],
                               item["expected_assertions"], str(item["log"]),
                               str(item["certification_log"]))
                  for item in data.get("suites", [])))
    except (ValueError, KeyError, TypeError) as error:
        raise PolicyError(f"{POLICY}: {error!r}") from error
    problems = policy_problems(data, modes, policy)
    if problems:
        raise PolicyError(f"{POLICY}: {problems[0]}")
    return policy


def policy_problems(data: dict[str, Any], modes: list[tuple[str, str, Any]],
                    policy: Policy) -> list[str]:
    """What is wrong with a policy that parsed, in the order it is reported."""
    problems = []
    # Version 1 has no project suites; version 2 adds the optional "suites" list.
    if data.get("schema_version") not in (1, 2) or \
            (data.get("schema_version") == 1 and "suites" in data):
        problems.append("schema_version must be 2, or 1 without suites")
    if [mode for mode, _, _ in modes] != list(MODE_KINDS):
        problems.append(f"modes must be {', '.join(MODE_KINDS)}, in that order")
    if any(not isinstance(must, bool) for _, _, must in modes) or \
            policy.must_pass != frozenset({"NORMAL"}):
        problems.append("only NORMAL must pass; must_pass is true or false")
    files = [*policy.logs.values(), policy.certification, policy.host, policy.manifest,
             policy.results, policy.exported,
             *(name for item in policy.project_suites for name in (item.log, item.certification))]
    if len(set(files)) != len(files) or not all(files):
        problems.append("the logs and evidence files must be distinct and named")
    if any("/" in name or "\\" in name or name in (".", "..") for name in files):
        problems.append("the logs and evidence files must be plain names in the evidence "
                        "folder")
    sources = [policy.harness, policy.suite_source,
               *(item.source for item in policy.project_suites)]
    names = [policy.suite, *(item.name for item in policy.project_suites)]
    if len(set(sources)) != len(sources) or len(set(names)) != len(names) or not all(names):
        problems.append("every suite needs its own name and source, apart from the harness")
    for source in sources:
        path = PurePosixPath(source)
        if path.is_absolute() or ".." in path.parts or path.suffix != ".bas":
            problems.append(f"{source}: harness and suite sources must be .bas files "
                            "inside the repository")
    for suite in (policy, *policy.project_suites):
        for name in ("expected_cases", "expected_assertions"):
            value = getattr(suite, name)
            if not isinstance(value, int) or isinstance(value, bool) or value < 1:
                problems.append(f"suite {name} must be a positive integer")
    if not policy.references or len(set(policy.references)) != len(policy.references):
        problems.append("host references must be listed once each")
    if policy.office not in OFFICE_BITNESS:
        problems.append(f"host office must be one of {', '.join(OFFICE_BITNESS)}")
    return problems


def working_policy(root: Path) -> Policy:
    """The policy of the working tree."""
    path = root / POLICY
    if not path.is_file():
        raise PolicyError(f"{POLICY}: missing")
    return parse_policy(path.read_text(encoding="utf-8"))


def policy_at(root: Path, commit: str) -> Policy:
    """The policy of a commit, which judges the evidence that names it."""
    try:
        text = git(root, "show", f"{commit}:{POLICY}").decode("utf-8")
    except RuntimeError as error:
        raise PolicyError(f"{POLICY}: missing at {commit[:12]}") from error
    return parse_policy(text)


@dataclass
class Record:
    """One log line: its key, its first value and its named fields."""

    key: str
    value: str
    fields: dict[str, str]
    line: int


@dataclass
class Suite:
    """What the commit's example suite runs: case and assertion identifiers in order,
    the identifiers of the last case (skipped in SKIPPED_CASE), the assertions of
    the case whose expected value INJECTED_FAILURE changes and the value it sets."""

    cases: list[str]
    assertions: list[str]
    last_case: list[str]
    injected: list[str]
    expected_cases: int
    expected_assertions: int
    injected_expected: str = ""


@dataclass
class Evidence:
    findings: list[str] = field(default_factory=list)
    hashes: dict[str, str] = field(default_factory=dict)
    environment: dict[str, str] = field(default_factory=dict)
    host: dict[str, str] = field(default_factory=dict)
    binding: str = "UNVERIFIED"
    binding_notes: list[str] = field(default_factory=list)
    accuracy: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


def git(root: Path, *args: str) -> bytes:
    completed = subprocess.run(["git", "-C", str(root), *args], capture_output=True, check=False)
    if completed.returncode:
        raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())
    return completed.stdout


def components(root: Path, commit: str) -> dict[str, bytes]:
    """The commit's VBA components and form resources, by file name, as Git stores
    them: text components with LF line ends, resources byte for byte."""
    listed = git(root, "ls-tree", "-r", "-z", "--name-only", commit).split(b"\0")
    paths = [item.decode("utf-8") for item in listed
             if item.decode("utf-8").lower().endswith(VBA_SUFFIXES + RESOURCE_SUFFIXES)]
    return {Path(path).name: git(root, "show", f"{commit}:{path}") for path in paths}


def source_at(root: Path, commit: str, path: str, evidence: Evidence) -> bytes:
    """The file at ``path`` in the commit, exactly as the policy names it; a
    finding, and no bytes, when the commit has no such file."""
    try:
        return git(root, "show", f"{commit}:{path}")
    except RuntimeError:
        evidence.findings.append(f"{path}: not in commit {commit[:12]}, but {POLICY} names it")
        return b""


def parse_record(text: str, line: int) -> Record | None:
    match = RECORD.match(text)
    if not match:
        return None
    first, *rest = match.group(2).split("; ")
    fields = {}
    for part in rest:
        name, _, value = part.partition("=")
        fields[name] = value
    if "=" in first:
        name, _, value = first.partition("=")
        fields[name] = value
        first = ""
    return Record(match.group(1), first, fields, line)


def read_log(path: Path, problems: list[str]) -> list[Record]:
    if not path.is_file():
        problems.append(f"{path.name}: missing")
        return []
    records = []
    for number, text in enumerate(path.read_text(encoding="utf-8", errors="replace")
                                  .splitlines(), 1):
        record = parse_record(text, number)
        if record is None:
            problems.append(f"{path.name}:{number}: not a KEY=value record")
        else:
            records.append(record)
    if not records:
        problems.append(f"{path.name}: incomplete log: no record")
    return records


def read_suite(source: bytes, office: str = "64-bit") -> Suite:
    """The suite as a host of ``office`` bitness compiles it: lines inside an #If
    branch that is inactive there are blank."""
    text = source.decode("ascii", errors="replace").replace("\r\n", "\n")
    environment = BITNESS_ENVIRONMENT.get(office, BITNESS_ENVIRONMENT["64-bit"])
    text = check_vba_conditionals.reachable_sources("suite.bas", text)[0][environment]
    bodies = {match.group(1): match.group(0) for match in PROCEDURE.finditer(text)}
    constants = {name: int(value) for name, value in CONSTANT.findall(text)}
    cases: list[str] = []
    assertions: list[str] = []
    last_case: list[str] = []
    injected: list[str] = []
    injected_expected = ""
    for name in CALL.findall(bodies.get("RunSuite", "")):
        body = bodies.get(name, "")
        case = BEGIN_CASE.findall(body)
        names = [f"{case[0] if case else name}/{item}" for item in ASSERTION.findall(body)]
        cases += case
        assertions += names
        last_case = case + names
        if 'mMode = "INJECTED_FAILURE"' in body:
            injected += names
            value = INJECTED_VALUE.search(body)
            injected_expected = value.group(1) if value else ""
    return Suite(cases, assertions, last_case, injected, constants.get("EXPECTED_CASES", -1),
                 constants.get("EXPECTED_ASSERTIONS", -1), injected_expected)


def document_modules(root: Path, commit: str) -> set[str]:
    """The file names of the commit's document modules: the .cls files of the
    build's document home."""
    listed = git(root, "ls-tree", "-r", "-z", "--name-only", commit, "--",
                 build_workbook.DOCUMENT_HOME).split(b"\0")
    return {PurePosixPath(item.decode("utf-8")).name for item in listed
            if item.decode("utf-8").lower().endswith(".cls")}


@dataclass
class SourceLine:
    """One line of a module: its number, its text, and its text with the letters
    of code outside strings and comments in lower case."""

    number: int
    text: str
    key: str


def source_lines(data: bytes, document: bool) -> list[SourceLine]:
    """The lines of a module (LF line ends) as they are compared, without the
    attributes Excel writes for a document module. A string literal, a [...]
    expression and a comment, continued or not, keep their letter case."""
    lines = []
    continued = False
    # Latin-1 maps each byte to one character, so no two bytes compare equal.
    for number, text in enumerate(data.decode("latin-1").split("\n"), 1):
        if document and DOCUMENT_ATTRIBUTE.match(text):
            continue
        if continued:
            continued = bool(vba_lexer.CONTINUATION.search(text))
            lines.append(SourceLine(number, text, text))
            continue
        # strip_code keeps the columns of the code before a comment and blanks
        # the content of strings and brackets: only code keeps its characters.
        code = vba_lexer.strip_code(text, blank_strings=True)
        continued = len(vba_lexer.strip_code(text)) < len(text.rstrip()) and \
            bool(vba_lexer.CONTINUATION.search(text))
        key = "".join(char.lower() if "A" <= char <= "Z" and index < len(code)
                      and code[index] == char else char for index, char in enumerate(text))
        lines.append(SourceLine(number, text, key))
    return lines


def source_difference(actual: bytes, expected: bytes, document: bool
                      ) -> tuple[int, list[str], list[str]]:
    """How an exported module differs from the commit's: the line of the export
    where it first differs other than in the letter case of code (0 when it does
    not), the identifiers re-cased, as "commit spelling -> exported spelling",
    and the document-module attributes that differ."""
    left, right = source_lines(actual, document), source_lines(expected, document)
    recased: dict[str, None] = {}
    for line, other in zip(left, right, strict=False):
        if line.key != other.key:
            return line.number, [], []
        if line.text != other.text:
            for match in IDENTIFIER.finditer(line.text):
                spelling = other.text[match.start():match.end()]
                if spelling != match.group():
                    recased[f"{spelling} -> {match.group()}"] = None
    if len(left) > len(right):
        return left[len(right)].number, [], []
    if len(left) < len(right):
        return (left[-1].number + 1 if left else 1), [], []
    changed = set(actual.decode("latin-1").split("\n")) ^ \
        set(expected.decode("latin-1").split("\n"))
    attributes = sorted({found.group(1) for found in map(DOCUMENT_ATTRIBUTE.match, changed)
                         if found}) if document else []
    return 0, list(recased), attributes


def check_sources(root: Path, commit: str, folder: Path, evidence: Evidence,
                  policy: Policy) -> dict[str, bytes]:
    expected = components(root, commit)
    documents = document_modules(root, commit)
    exported = folder / policy.exported
    present = {path.name: path for path in exported.iterdir()} if exported.is_dir() else {}
    if not expected:
        evidence.findings.append(f"commit {commit} holds no VBA component")
    for name, data in sorted(expected.items()):
        evidence.hashes[name] = git_blob(data)
        if name not in present:
            evidence.findings.append(f"{policy.exported}/{name}: not exported from the workbook")
            continue
        actual = present[name].read_bytes()
        if name.lower().endswith(RESOURCE_SUFFIXES):
            if actual != data:
                evidence.findings.append(f"{policy.exported}/{name}: differs from {name} at "
                                         f"{commit[:12]} (binary resource, compared byte for "
                                         "byte)")
            continue
        line, recased, attributes = source_difference(actual.replace(b"\r\n", b"\n"), data,
                                                      name in documents)
        if line:
            evidence.findings.append(f"{policy.exported}/{name}: differs from {name} at "
                                     f"{commit[:12]} (first difference at line {line})")
        if recased:
            evidence.notes.append(f"{policy.exported}/{name}: {len(recased)} identifier(s) "
                                  "re-cased outside strings and comments, as the VBE re-cases "
                                  f"a name (HB-01): {', '.join(recased)}")
        if attributes:
            evidence.notes.append(f"{policy.exported}/{name}: document-module attributes "
                                  f"Excel writes, not compared: {', '.join(attributes)}")
    for name in sorted(set(present) - set(expected)):
        evidence.findings.append(f"{policy.exported}/{name}: not a VBA component of the commit")
    return expected


def git_blob(data: bytes) -> str:
    """The Git blob id of ``data``, as ``git hash-object`` gives it."""
    # Git names objects by SHA-1; this is an identifier, not a security control.
    return hashlib.sha1(b"blob %d\0" % len(data) + data, usedforsecurity=False).hexdigest()


@dataclass
class Plan:
    """What one mode's run must log: its cases and assertions in order, the
    failures it must report (assertions, or harness identifiers) and its RESULT."""

    mode: str
    cases: list[str]
    assertions: list[str]
    failing: list[str]
    result: dict[str, str]


def plans(suite: Suite, name: str, modes: Sequence[str] = MODE_KINDS) -> list[Plan]:
    """The plan of each of ``modes`` for the suite called ``name``."""
    skipped = set(suite.last_case)
    cases = [case for case in suite.cases if case not in skipped]
    assertions = [item for item in suite.assertions if item not in skipped]

    def result(verdict: str, operation: str, complete: bool, kept: tuple[list[str], list[str]],
               failures: int) -> dict[str, str]:
        return {"RESULT": verdict, "suite": name, "operation": operation,
                "completeness": "COMPLETE" if complete else "INCOMPLETE",
                "cases": str(len(kept[0])), "assertions": str(len(kept[1])),
                "failures": str(failures), "cleanup": "PASS", "recovery": "NONE"}

    every = (suite.cases, suite.assertions)
    every_plan = [Plan("NORMAL", *every, [], result("PASS", "PASS", True, every, 0)),
                  Plan("INJECTED_FAILURE", *every, suite.injected,
                       result("FAIL", "FAIL", True, every, len(suite.injected))),
                  Plan("SKIPPED_CASE", cases, assertions, ["harness.completeness"],
                       result("FAIL", "PASS", False, (cases, assertions), 1))]
    return [plan for plan in every_plan if plan.mode in modes]


def sequence(plan: Plan, name: str) -> list[tuple[str, str]]:
    """Every record the run of the suite called ``name`` writes, as (key, first
    value), in order."""
    records = [("RUN", name), ("ENVIRONMENT", "")]
    for case in plan.cases:
        records.append(("CASE", case))
        for item in (name for name in plan.assertions if name.startswith(f"{case}/")):
            records.append(("ASSERT", item))
            if item in plan.failing:
                records.append(("FAILURE", item))
    records += [("FAILURE", item) for item in plan.failing if "/" not in item]
    records += [("CASES", str(len(plan.cases))), ("ASSERTIONS", str(len(plan.assertions))),
                ("FAILURES", str(len(plan.failing)))]
    records += [("PRIMARY_FAILURE", plan.failing[0])] if plan.failing else []
    return [*records, ("CLEANUP", "PASS"), ("RECOVERY", "NONE"),
            ("RESULT", plan.result["RESULT"])]


def check_sequence(name: str, records: list[Record], expected: list[tuple[str, str]],
                   problems: list[str]) -> bool:
    """The log holds exactly the expected records, in order: none missing, extra,
    repeated or contradicting the run's outcome."""
    for record, (key, value) in zip(records, expected, strict=False):
        if (record.key, record.value) != (key, value):
            problems.append(f"{name}:{record.line}: {record.key}={record.value}, "
                            f"expected {key}={value}")
            return False
    if len(records) < len(expected):
        key, value = expected[len(records)]
        problems.append(f"{name}: incomplete log: it ends before {key}={value}")
        return False
    if len(records) > len(expected):
        extra = records[len(expected)]
        problems.append(f"{name}:{extra.line}: {extra.key}={extra.value} after RESULT "
                        "(more than one run?)")
        return False
    return True


def expect(name: str, record: Record, wanted: dict[str, str], problems: list[str]) -> None:
    for key, value in wanted.items():
        actual = record.value if key == record.key else record.fields.get(key)
        if actual != value:
            problems.append(f"{name}:{record.line}: {record.key} {key}={actual}, "
                            f"expected {value}")


def check_fields(name: str, records: list[Record], plan: Plan, suite: Suite,
                 problems: list[str]) -> None:
    expected = {"CASES": str(suite.expected_cases), "ASSERTIONS": str(suite.expected_assertions)}
    for record in records:
        if record.key == "RUN":
            expect(name, record, {"mode": plan.mode, "expected_cases": expected["CASES"],
                                  "expected_assertions": expected["ASSERTIONS"]}, problems)
        elif record.key == "ASSERT":
            status = "FAIL" if record.value in plan.failing else "PASS"
            expect(name, record, {"status": status}, problems)
        elif record.key in expected:
            expect(name, record, {"expected": expected[record.key]}, problems)
        elif record.key == "FAILURES":
            expect(name, record, {"errors": "0"}, problems)
        elif record.key == "RESULT" and {"RESULT": record.value, **record.fields} != plan.result:
            problems.append(f"{name}:{record.line}: RESULT={record.value} with "
                            f"{record.fields}, expected {plan.result}")


def check_injected(name: str, records: list[Record], plan: Plan, suite: Suite,
                   problems: list[str]) -> None:
    """The designated failure's ASSERT, FAILURE and PRIMARY_FAILURE records carry
    numeric expected and actual values that differ, the expected one being the value
    the suite injects, and all three carry the same values."""
    if plan.mode != "INJECTED_FAILURE" or len(suite.injected) != 1:
        return
    found = [record for record in records
             if record.key in ("ASSERT", "FAILURE", "PRIMARY_FAILURE")
             and record.value == suite.injected[0]]
    for record in found:
        values = (record.fields.get("expected", ""), record.fields.get("actual", ""))
        try:
            expected, actual = float(values[0]), float(values[1])
        except ValueError:
            problems.append(f"{name}:{record.line}: {record.key} lacks numeric expected= "
                            "and actual= values")
            continue
        if suite.injected_expected and expected != float(suite.injected_expected):
            problems.append(f"{name}:{record.line}: {record.key} expected={values[0]}, but the "
                            f"suite injects {suite.injected_expected}")
        if expected == actual:
            problems.append(f"{name}:{record.line}: {record.key} expected equals actual, so "
                            "the assertion cannot have failed")
    if len({(record.fields.get("expected"), record.fields.get("actual"))
            for record in found}) > 1:
        problems.append(f"{name}: the injected failure's values differ between ASSERT, "
                        "FAILURE and PRIMARY_FAILURE")


def suite_logs(folder: Path, suite: Suite, evidence: Evidence, name: str,
               logs: dict[str, str]) -> list[dict[str, str]]:
    """Check the log of each mode in ``logs`` for the suite called ``name``; return
    the ENVIRONMENT records the logs hold."""
    environments = []
    for plan in plans(suite, name, tuple(logs)):
        filename = logs[plan.mode]
        records = read_log(folder / filename, evidence.findings)
        if not records:
            continue
        result = records[-1]
        if plan.mode == "NORMAL" and result.key == "RESULT" and \
                result.fields.get("completeness") != "COMPLETE":
            evidence.findings.append(f"{filename}:{result.line}: the run is "
                                     f"{result.fields.get('completeness') or 'not COMPLETE'}: "
                                     "a mandatory case or assertion did not run")
        if check_sequence(filename, records, sequence(plan, name), evidence.findings):
            check_fields(filename, records, plan, suite, evidence.findings)
            check_injected(filename, records, plan, suite, evidence.findings)
        if len(records) > 1 and records[1].key == "ENVIRONMENT":
            environments.append(records[1].fields)
    return environments


def check_logs(folder: Path, suites: Sequence[tuple[Suite, str, dict[str, str]]],
               evidence: Evidence) -> None:
    """Each suite's logs, as (suite, name, log by mode), all from one host."""
    environments = [environment for suite, name, logs in suites
                    for environment in suite_logs(folder, suite, evidence, name, logs)]
    for environment in environments:
        missing = [name for name in ENVIRONMENT_FIELDS if not environment.get(name)]
        if missing:
            evidence.findings.append(f"ENVIRONMENT lacks {', '.join(missing)}")
    if len({tuple(sorted(item.items())) for item in environments}) > 1:
        evidence.findings.append("the logs come from different hosts (ENVIRONMENT differs)")
    if environments:
        evidence.environment = environments[0]


def check_host(folder: Path, evidence: Evidence, policy: Policy,
               rebuilt: bool = False) -> None:
    path = folder / policy.host
    if not path.is_file():
        evidence.findings.append(f"{policy.host}: missing")
        return
    host = {}
    for text in path.read_text(encoding="utf-8").splitlines():
        key, separator, value = text.partition("=")
        if separator and not text.startswith("#"):
            host[key.strip()] = value.strip()
    evidence.host = host
    for key in HOST_FIELDS:
        if not host.get(key):
            evidence.findings.append(f"{policy.host}: no {key}=")
    references = tuple(item.strip() for item in host.get("references", "").split(";"))
    if host.get("references") and sorted(references) != sorted(policy.references):
        evidence.findings.append(f"{policy.host}: references must be exactly "
                                 f"{'; '.join(policy.references)}")
    allowed = ("PASS", COMPILE_ON_DEMAND) if rebuilt else ("PASS",)
    if host.get("compile") and host["compile"] not in allowed:
        evidence.findings.append(f"{policy.host}: compile={host['compile']}, expected "
                                 f"{' or '.join(allowed)}"
                                 + ("" if rebuilt else f" ({COMPILE_ON_DEMAND} only with "
                                    "identity by rebuild)"))
    if evidence.environment and evidence.environment.get("office") != policy.office:
        evidence.findings.append(f"not the supported target: Office is not {policy.office}")


def test_failed_number(source: bytes) -> int | None:
    """ERR_TEST_FAILED of the commit's harness, as the number VBA raises."""
    match = TEST_FAILED.search(source.decode("ascii", errors="replace"))
    if not match:
        return None
    if match.group(1):
        return VB_OBJECT_ERROR + int(match.group(1))
    value = int(match.group(2), 16)
    return value - (1 << 32) if value >= 1 << 31 else value


def certification_plan(failed: int | None, name: str,
                       modes: Sequence[str]) -> list[tuple[str, dict[str, str]]]:
    """Every record of the certification log of the suite called ``name``, after
    CERTIFY, as (key, fields). Only NORMAL must pass."""
    records = []
    for mode in modes:
        must_pass = mode == "NORMAL"
        complete = "INCOMPLETE" if mode == "SKIPPED_CASE" else "COMPLETE"
        records.append(("MODE", {"MODE": mode, "expected": "PASS" if must_pass else "FAIL",
                                 "error": "0" if must_pass else str(failed),
                                 "result": "PASS" if must_pass else "FAIL",
                                 "completeness": complete, "status": "PASS"}))
    records.append(("CERTIFICATION", {"CERTIFICATION": "PASS", "suite": name,
                                      "completeness": "COMPLETE",
                                      "modes": str(len(modes)),
                                      "failures": "0"}))
    return records


def check_certification(folder: Path, failed: int | None, evidence: Evidence,
                        name: str, modes: Sequence[str], filename: str) -> None:
    """The certification log ``filename`` of the suite called ``name`` holds CERTIFY,
    one MODE record per mode in order, each ending as it must, and
    CERTIFICATION=PASS: written before anything was raised."""
    records = read_log(folder / filename, evidence.findings)
    if not records:
        return
    if (records[0].key, records[0].value) != ("CERTIFY", name):
        evidence.findings.append(f"{filename}:{records[0].line}: {records[0].key}="
                                 f"{records[0].value}, expected CERTIFY={name}")
        return
    expected = certification_plan(failed, name, modes)
    for record, (key, fields) in zip(records[1:], expected, strict=False):
        actual = {record.key: record.value, **record.fields}
        if record.key != key or actual != fields:
            evidence.findings.append(f"{filename}:{record.line}: {record.key}="
                                     f"{record.value} with {record.fields}, expected {fields}")
    if len(records) - 1 < len(expected):
        key, fields = expected[len(records) - 1]
        evidence.findings.append(f"{filename}: incomplete log: it ends before "
                                 f"{key}={fields[key]}")
    elif len(records) - 1 > len(expected):
        extra = records[len(expected) + 1]
        evidence.findings.append(f"{filename}:{extra.line}: {extra.key}={extra.value} "
                                 "after CERTIFICATION (more than one certification?)")


SPREADSHEET = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RELATIONSHIP = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
PACKAGE = "{http://schemas.openxmlformats.org/package/2006/relationships}"
REFERENCE = re.compile(r"^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Z]{1,3})\$?([0-9]+)$")


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def cell_text(package: zipfile.ZipFile, part: str, cell: str) -> str | None:
    """The text of one cell of a worksheet part: inline, shared or formula string."""
    root = ElementTree.fromstring(package.read(part))
    for element in root.iter(f"{SPREADSHEET}c"):
        if element.get("r") != cell:
            continue
        kind = element.get("t")
        if kind == "inlineStr":
            return "".join(node.text or "" for node in element.iter(f"{SPREADSHEET}t"))
        value = element.find(f"{SPREADSHEET}v")
        if value is None or value.text is None:
            return None
        if kind == "s":
            shared = ElementTree.fromstring(package.read("xl/sharedStrings.xml"))
            item = list(shared.iter(f"{SPREADSHEET}si"))[int(value.text)]
            return "".join(node.text or "" for node in item.iter(f"{SPREADSHEET}t"))
        return value.text
    return None


def read_stamp(workbook: Path, name: str) -> str | None:
    """The value of the defined name ``name`` of a workbook, read from its parts."""
    with zipfile.ZipFile(workbook) as package:
        book = ElementTree.fromstring(package.read("xl/workbook.xml"))
        target = next((item.text or "" for item in book.iter(f"{SPREADSHEET}definedName")
                       if item.get("name", "").casefold() == name.casefold()
                       and item.get("localSheetId") is None), None)
        match = REFERENCE.match(target or "")
        if not match:
            return None
        tab = (match.group(1) or "").replace("''", "'") or match.group(2)
        sheet = next((item for item in book.iter(f"{SPREADSHEET}sheet")
                      if item.get("name") == tab), None)
        if sheet is None:
            return None
        relations = ElementTree.fromstring(package.read("xl/_rels/workbook.xml.rels"))
        relation = next((item for item in relations.iter(f"{PACKAGE}Relationship")
                         if item.get("Id") == sheet.get(RELATIONSHIP)), None)
        if relation is None:
            return None
        location = relation.get("Target", "")
        part = location.lstrip("/") if location.startswith("/") else \
            posixpath.normpath(str(PurePosixPath("xl") / location))
        return cell_text(package, part, f"{match.group(3)}{match.group(4)}")


def load_manifest(path: Path) -> tuple[str, dict[str, str]]:
    """The workbook digest and the exported files' digests of export-manifest.json."""
    data = json.loads(path.read_text(encoding="utf-8"))
    if data.get("format") != MANIFEST_FORMAT:
        raise ValueError(f"format is {data.get('format')!r}, expected {MANIFEST_FORMAT}")
    files = {str(item["file"]): str(item["sha256"]).lower() for item in data["files"]}
    return str(data["workbook"]["sha256"]).lower(), files


def check_rebuild(root: Path, commit: str, workbook: Path, evidence: Evidence) -> None:
    """Source identity without an export: the workbook equals a fresh build of the commit."""
    for name, data in sorted(components(root, commit).items()):
        evidence.hashes[name] = git_blob(data)
    evidence.binding = "FAIL"
    if not workbook.is_file():
        evidence.findings.append(f"workbook {workbook}: missing")
        return
    with tempfile.TemporaryDirectory(prefix="excel-run-rebuild-") as temporary:
        tree = Path(temporary) / "tree"
        target = Path(temporary) / "rebuilt.xlsm"
        try:
            git(root, "worktree", "add", "--detach", "--quiet", str(tree), commit)
        except RuntimeError as error:
            evidence.findings.append(f"cannot check out {commit[:12]} to rebuild it: {error}")
            return
        try:
            build_workbook.build(tree, REBUILT_CONFIGURATION, target)
            rebuilt = target.read_bytes()
        except (build_workbook.BuildError, ValueError, OSError, RuntimeError) as error:
            evidence.findings.append(f"cannot rebuild {commit[:12]}: {error}")
            return
        finally:
            try:
                git(root, "worktree", "remove", "--force", str(tree))
            except RuntimeError:
                pass
    try:
        actual = package_parts(workbook.read_bytes())
    except (zipfile.BadZipFile, OSError, ValueError) as error:
        evidence.findings.append(f"workbook {workbook.name}: not a readable package: {error}")
        return
    difference = parts_difference(actual, package_parts(rebuilt))
    if difference:
        evidence.findings.append(
            f"workbook {workbook.name}: {difference} the {REBUILT_CONFIGURATION} build of "
            f"{commit[:12]}: it was saved after the build, or not built from that commit")
        return
    evidence.binding = "VERIFIED"
    evidence.binding_notes.append(f"workbook parts byte-identical to the "
                                  f"{REBUILT_CONFIGURATION} build of {commit[:12]}; no export "
                                  "needed")


def package_parts(data: bytes) -> list[tuple[str, bytes]]:
    """The files of a ZIP package, in order, decompressed."""
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        return [(info.filename, package.read(info)) for info in package.infolist()]


def parts_difference(actual: list[tuple[str, bytes]], expected: list[tuple[str, bytes]]) -> str:
    """How the parts differ, as the start of a sentence that ends by naming the build;
    empty when they are the same."""
    names = [name for name, _ in actual]
    if names != [name for name, _ in expected]:
        return f"its parts ({', '.join(names)}) differ from those of"
    for (name, data), (_, other) in zip(actual, expected):
        if data != other:
            return f"{name} (SHA-256 {sha256(data)[:16]}...) differs from"
    return ""


def check_binding(root: Path, commit: str, folder: Path, workbook: Path | None,
                  evidence: Evidence, policy: Policy) -> None:
    """VERIFIED when the manifest matches the export and the workbook, and the
    workbook's stamp names the commit; FAIL on any mismatch; UNVERIFIED without
    a manifest or a workbook. A manifest that contradicts the export fails even
    without --workbook."""
    findings: list[str] = []
    path = folder / policy.manifest
    if not path.is_file():
        evidence.binding = "UNVERIFIED"
        evidence.binding_notes.append(f"no {policy.manifest}: the export is not bound to a "
                                      "workbook")
        return
    try:
        workbook_digest, files = load_manifest(path)
    except (OSError, ValueError, KeyError, TypeError) as error:
        evidence.binding = "FAIL"
        evidence.findings.append(f"{policy.manifest}: not a valid export manifest: {error!r}")
        return
    exported = folder / policy.exported
    present = {item.name: item for item in exported.iterdir() if item.is_file()} \
        if exported.is_dir() else {}
    for name in sorted(set(files) | set(present)):
        if name not in present:
            findings.append(f"{policy.manifest}: {name} is recorded but not in {policy.exported}/")
        elif name not in files:
            findings.append(f"{policy.manifest}: {policy.exported}/{name} is not recorded")
        elif sha256(present[name].read_bytes()) != files[name]:
            findings.append(f"{policy.manifest}: {policy.exported}/{name} changed after the export")
    if workbook is not None:
        findings += workbook_findings(root, commit, workbook, workbook_digest, policy.manifest)
    else:
        evidence.binding_notes.append("no --workbook: the manifest is not checked against "
                                      "a workbook")
    evidence.findings += findings
    evidence.binding = "FAIL" if findings else ("VERIFIED" if workbook else "UNVERIFIED")


def workbook_findings(root: Path, commit: str, workbook: Path, digest: str,
                      manifest: str) -> list[str]:
    if not workbook.is_file():
        return [f"workbook {workbook}: not a file"]
    findings = []
    actual = sha256(workbook.read_bytes())
    if actual != digest:
        findings.append(f"workbook {workbook.name}: sha256 {actual}, but {manifest} records "
                        f"{digest}")
    try:
        layout = json.loads(git(root, "show", f"{commit}:{BUILD_CONFIG}").decode("utf-8"))
        name = str(layout["stamp"]["commit"])
    except (RuntimeError, ValueError, KeyError, TypeError):
        return [*findings, f"{BUILD_CONFIG} at {commit[:12]} names no commit stamp"]
    try:
        stamp = read_stamp(workbook, name)
    except (OSError, KeyError, ValueError, IndexError, zipfile.BadZipFile,
            ElementTree.ParseError) as error:
        return [*findings, f"workbook {workbook.name}: build stamp not readable: {error!r}"]
    if stamp != commit:
        findings.append(f"workbook {workbook.name}: build stamp {name} is {stamp!r}, expected "
                        f"{commit} from a clean build")
    return findings


def copy_cases(root: Path, commit: str, target: Path) -> None:
    """The commit's files under CASE_PATHS, written below ``target``."""
    listed = git(root, "ls-tree", "-r", "-z", "--name-only", commit, "--", *CASE_PATHS)
    for name in filter(None, listed.decode("utf-8").split("\0")):
        path = PurePosixPath(name)
        if path.is_absolute() or ".." in path.parts:
            continue
        destination = target.joinpath(*path.parts)
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(git(root, "show", f"{commit}:{name}"))


def check_results(root: Path, commit: str, folder: Path, evidence: Evidence,
                  policy: Policy) -> None:
    """Measure the commit's accuracy contracts on the run's results file."""
    registry = check_accuracy.REGISTRY.as_posix()
    try:
        git(root, "cat-file", "-e", f"{commit}:{registry}")
    except RuntimeError:
        evidence.accuracy.append(f"not measured: commit {commit[:12]} has no {registry}")
        return
    path = folder / policy.results
    if not path.is_file():
        evidence.findings.append(f"{policy.results}: missing, but {registry} at {commit[:12]} "
                                 "holds the accuracy contracts it must measure")
        return
    with tempfile.TemporaryDirectory(prefix="excel-run-cases-") as temporary:
        cases = Path(temporary)
        copy_cases(root, commit, cases)
        found, _, lines = check_accuracy.check(cases, path, commit, widening=False)
    evidence.accuracy += lines
    evidence.findings += [f"{policy.results}: {item}" for item in found]


def suite_findings(policy: Policy | ProjectSuite, suite: Suite, source: bytes,
                   label: str) -> list[str]:
    """Where a suite's source disagrees with the policy, or cannot be checked. The
    example suite, the policy itself, must also inject exactly one failure."""
    findings = []
    path = policy.suite_source if isinstance(policy, Policy) else policy.source
    where = f"{path} at {label}"
    if not suite.cases:
        return [f"no suite in {where}"]
    if isinstance(policy, Policy) and len(suite.injected) != 1:
        findings.append(f"{where}: INJECTED_FAILURE must change exactly one assertion, "
                        f"found {suite.injected}")
    elif isinstance(policy, Policy) and not suite.injected_expected:
        findings.append(f"{where}: INJECTED_FAILURE sets no numeric expected value")
    for name, stated in (("cases", suite.expected_cases),
                         ("assertions", suite.expected_assertions)):
        expected = getattr(policy, f"expected_{name}")
        if stated != expected:
            findings.append(f"{where}: EXPECTED_{name.upper()} is {stated}, but {POLICY} "
                            f"expects {expected}")
    text = source.decode("ascii", errors="replace")
    if not re.search(rf"^Public Sub {re.escape(policy.entry_point)}\(", text, re.M):
        findings.append(f"{where}: no Public Sub {policy.entry_point}, the certification "
                        f"entry point of {POLICY}")
    return findings


def check(root: Path, commit: str, folder: Path, workbook: Path | None = None,
          require_binding: bool = False) -> Evidence:
    evidence = Evidence()
    try:
        policy = policy_at(root, commit)
    except PolicyError as error:
        evidence.findings.append(str(error))
        return evidence
    # Without an export, a given workbook proves identity by rebuild instead.
    rebuilt = workbook is not None and not (folder / policy.exported).exists()
    if rebuilt:
        assert workbook is not None
        check_rebuild(root, commit, workbook, evidence)
    else:
        check_sources(root, commit, folder, evidence, policy)
    # The harness and the suite are read at the paths the policy names, never
    # by file name, so a policy cannot point at a module the commit lacks.
    harness = source_at(root, commit, policy.harness, evidence)
    failed = test_failed_number(harness) if harness else None
    if harness and failed is None:
        evidence.findings.append(f"{policy.harness}: no ERR_TEST_FAILED constant")
    runs = []
    entries: list[tuple[Policy | ProjectSuite, str, str, dict[str, str], str]] = [
        (policy, policy.suite, policy.suite_source, policy.logs, policy.certification),
        *((item, item.name, item.source, {"NORMAL": item.log}, item.certification)
          for item in policy.project_suites)]
    for item, name, path, logs, certification in entries:
        source = source_at(root, commit, path, evidence)
        suite = read_suite(source, policy.office)
        if source:
            evidence.findings += suite_findings(item, suite, source, commit[:12])
        runs.append((suite, name, logs))
        if harness and failed is not None:
            check_certification(folder, failed, evidence, name, tuple(logs), certification)
    check_logs(folder, runs, evidence)
    check_host(folder, evidence, policy, rebuilt)
    check_results(root, commit, folder, evidence, policy)
    if not rebuilt:
        check_binding(root, commit, folder, workbook, evidence, policy)
    if require_binding and evidence.binding == "UNVERIFIED":
        evidence.findings.append("binding UNVERIFIED, but --require-binding needs VERIFIED: "
                                 + "; ".join(evidence.binding_notes))
    return evidence


def report(commit: str, evidence: Evidence) -> str:
    environment = "; ".join(f"{key}={evidence.environment.get(key, '')}"
                            for key in ENVIRONMENT_FIELDS)
    lines = [f"Candidate commit: {commit}",
             f"Environment: {environment}",
             f"Windows: {evidence.host.get('windows', '')}",
             f"References: {evidence.host.get('references', '')}",
             f"Compile: {evidence.host.get('compile', '')}",
             "Imported components (Git blob ids at the commit):"]
    lines += [f"  {name} {blob}" for name, blob in sorted(evidence.hashes.items())]
    lines += [f"Note: {note}" for note in evidence.notes]
    lines += [f"Accuracy: {line}" for line in evidence.accuracy]
    notes = f" ({'; '.join(evidence.binding_notes)})" if evidence.binding_notes else ""
    lines.append(f"Binding: {evidence.binding}{notes}")
    lines += [f"FAIL: {finding}" for finding in evidence.findings]
    status = "FAIL" if evidence.findings else "PASS"
    lines.append(f"excel-run: {status} (evidence checked against the commit; Excel is not run "
                 "here, and a calculation is validated only as far as its accuracy contracts "
                 "are)")
    return "\n".join(lines)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("evidence", type=Path, help="folder holding the run's evidence")
    parser.add_argument("--commit", required=True, help="the commit the workbook was built from")
    parser.add_argument("--workbook", type=Path,
                        help="the workbook the export came from, checked against "
                             f"the export manifest named in {POLICY} and its build stamp; "
                             "without an exported folder, compared part by part with a "
                             "rebuild of the commit")
    parser.add_argument("--require-binding", action="store_true",
                        help="fail unless the binding is VERIFIED (release evidence)")
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    root = args.root.resolve()
    try:
        commit = git(root, "rev-parse", "--verify", f"{args.commit}^{{commit}}").decode().strip()
    except RuntimeError as error:
        print(f"FAIL: commit {args.commit} not found: {error}")
        return 1
    evidence = check(root, commit, args.evidence, args.workbook, args.require_binding)
    print(report(commit, evidence))
    return 1 if evidence.findings else 0


if __name__ == "__main__":
    sys.exit(main())
