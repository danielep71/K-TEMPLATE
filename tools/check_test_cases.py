#!/usr/bin/env python3
"""Check the numerical test cases against docs/methodology/TEST_CASES.md.

Static: it checks the format of the fixtures under tests/fixtures/ and of
tests/expected/*.json, the references between them, and their links to the
methodology guides: the calculation contracts (CALC- IDs), the data contract
(DS- IDs, file names and encodings) and the source register (SRC- and IMPL-
IDs). It never computes or compares a value.

A fixture is a JSON file, <case-id>.json, or a file-based case, a directory
<case-id>/ with a case.json manifest and one subdirectory per variant holding
input files named as in the data contract. In an invalid case only the harness
metadata is checked: its calculation inputs, or the content of its files, may
break any input rule, because that is what the case tests.
"""

from __future__ import annotations

import argparse
import codecs
import datetime
import json
import math
import re
import sys
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
METHODOLOGY = Path("docs/methodology")
FIXTURES = Path("tests/fixtures")
EXPECTED = Path("tests/expected")
DERIVATIONS = "tests/derivations/"
MANIFEST = "case.json"
DEFAULT_ENCODING = "UTF-8"

CASE_ID = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
SNAKE = re.compile(r"^[a-z][a-z0-9]*(?:_[a-z0-9]+)*$")
CURRENCY = re.compile(r"^[A-Z]{3}$")
DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
CALCULATION_ID = re.compile(r"^CALC-\d+$")
DATASET_ID = re.compile(r"^DS-\d+$")
# A rule is a step of the case's calculation contract: CALC-01.2 is step 2.
RULE = re.compile(r"^(CALC-\d+)\.\d+$")
# A declaration is a table row, or for contracts a heading, whose first cell
# is the ID in a code span; "*Illustrative:*" before it marks an example entry.
ROW = re.compile(r"^\|\s*(\*Illustrative:\*\s*)?`([A-Z]+-\d+)`\s*\|")
HEADING = re.compile(r"^###\s+(\*Illustrative:\*\s*)?`(CALC-\d+)`")
CODE_SPAN = re.compile(r"`([^`]+)`")
CATEGORIES = frozenset({"nominal", "boundary", "invalid"})
CLASSES = frozenset({"published", "independent", "illustrative"})
UNITS = frozenset({"1", "years"})
SUFFIXES = ("_amount", "_weight", "_factor", "_years", "_count", "_rate", "_date", "_flag",
            "_code", "_id")

# The harness metadata, valid in every category; the other keys are the
# calculation inputs, which an invalid case may break.
HARNESS_KEYS = frozenset({"schema_version", "kind", "id", "synthetic", "description",
                          "category", "reporting_date"})
FIXTURE_KEYS = HARNESS_KEYS | {"calculation", "reporting_currency_code", "datasets"}
RECORD_KEYS = frozenset({"id", "inputs"})
EXPECTED_KEYS = frozenset({"schema_version", "kind", "fixture", "outputs"})
EXPECTED_FILE_KEYS = frozenset({"schema_version", "kind", "fixture", "variants"})
OUTPUT_KEYS = frozenset({"quantity", "scope", "rule", "reference"})
NUMBER_KEYS = OUTPUT_KEYS | {"value", "unit", "tolerance"}
TEXT_KEYS = OUTPUT_KEYS | {"text"}
SCOPE_KEYS = frozenset({"dataset", "record"})
REFERENCE_KEYS = frozenset({"class", "source", "locator", "derivation", "script",
                            "derived_by", "reviewed_by", "date"})

# Returned by load_json when a file cannot be read; distinct from a JSON null.
LOAD_FAILED = object()


@dataclass
class Methodology:
    """IDs declared by the methodology guides; True marks an illustrative entry."""

    calculations: dict[str, bool]
    datasets: dict[str, bool]
    sources: dict[str, bool]
    # The files each dataset is read from, with their encodings (Datasets table).
    files: dict[str, dict[str, str]] = field(default_factory=dict)


@dataclass
class Case:
    """What an expected file needs to know about its fixture."""

    category: str
    calculation: str | None
    records: dict[str, set[str]]
    # A file-based case: its datasets and variants; its records are not listed.
    variants: set[str] | None = None
    classes: set[str] = field(default_factory=set)
    expected_files: int = 0


@dataclass
class Report:
    findings: list[str] = field(default_factory=list)

    def add(self, where: str, message: str) -> None:
        self.findings.append(f"{where}: {message}")


def is_number(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def is_text(value: object) -> bool:
    return isinstance(value, str) and value.strip() != ""


def is_date(value: object) -> bool:
    if not isinstance(value, str) or not DATE.match(value):
        return False
    try:
        datetime.date.fromisoformat(value)
    except ValueError:
        return False
    return True


def keys_problem(data: Mapping[str, Any], allowed: frozenset[str]) -> str | None:
    missing = sorted(allowed - data.keys())
    unknown = sorted(data.keys() - allowed)
    parts = ([f"missing {', '.join(missing)}"] if missing else []) + (
        [f"unknown {', '.join(unknown)}"] if unknown else [])
    return "; ".join(parts) or None


def rel(path: Path, root: Path) -> str:
    return path.relative_to(root).as_posix()


def load_json(path: Path, root: Path, report: Report) -> Any:
    """Parse strict JSON: duplicate keys, NaN, Infinity and overflowing numbers fail."""

    def no_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f"duplicate key {key!r}")
            result[key] = value
        return result

    def no_constants(name: str) -> None:
        raise ValueError(f"non-standard number {name}")

    # A literal such as 1e400 parses to infinity without reaching parse_constant.
    def finite_float(text: str) -> float:
        value = float(text)
        if not math.isfinite(value):
            raise ValueError(f"number {text} is not finite")
        return value

    def finite_int(text: str) -> int:
        value = int(text)
        try:
            float(value)
        except OverflowError:
            raise ValueError(f"number {text[:20]}... is not finite") from None
        return value

    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=no_duplicates,
                          parse_constant=no_constants, parse_float=finite_float,
                          parse_int=finite_int)
    except (OSError, UnicodeDecodeError, ValueError) as error:
        report.add(rel(path, root), f"not valid UTF-8 JSON ({error})")
        return LOAD_FAILED


def declared(root: Path, name: str, section: str | None, pattern: re.Pattern[str],
             report: Report) -> dict[str, bool]:
    """The IDs a methodology guide declares, in one ## section or in the whole file."""
    path = root / METHODOLOGY / name
    if not path.is_file():
        report.add((METHODOLOGY / name).as_posix(), "is missing")
        return {}
    ids: dict[str, bool] = {}
    inside = section is None
    for line in path.read_text(encoding="utf-8").splitlines():
        if section is not None and line.startswith("## "):
            inside = line[3:].strip() == section
        elif inside and (match := pattern.match(line)):
            ids[match.group(2)] = match.group(1) is not None
    return ids


def cells(line: str) -> list[str]:
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def dataset_files(root: Path, report: Report) -> dict[str, dict[str, str]]:
    """Map each dataset of the data contract to its file names and their encoding.

    The Datasets table gives the names in code spans in its "File or sheet"
    column; a name without an extension is a sheet. An "Encoding" column, when
    present, gives a Python codec name; an empty cell, or no column, is UTF-8.
    """
    path = root / METHODOLOGY / "DATA_CONTRACT.md"
    if not path.is_file():
        return {}
    files: dict[str, dict[str, str]] = {}
    inside = False
    header: list[str] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.startswith("## "):
            inside = line[3:].strip() == "Datasets"
            header = []
        elif inside and line.startswith("|") and not header:
            header = cells(line)
        elif inside and (match := ROW.match(line)) and "File or sheet" in header:
            row = cells(line) + [""] * len(header)
            encoding = (row[header.index("Encoding")].strip("`") if "Encoding" in header
                        else "") or DEFAULT_ENCODING
            try:
                codecs.lookup(encoding)
            except LookupError:
                report.add((METHODOLOGY / "DATA_CONTRACT.md").as_posix(),
                           f"{match.group(2)} has an unknown encoding {encoding!r}")
                continue
            names = CODE_SPAN.findall(row[header.index("File or sheet")])
            files[match.group(2)] = {name: encoding for name in names if "." in name}
    return files


def read_methodology(root: Path, report: Report) -> Methodology:
    return Methodology(
        calculations=declared(root, "CALCULATION_CONTRACTS.md", None, HEADING, report),
        datasets=declared(root, "DATA_CONTRACT.md", "Datasets", ROW, report),
        sources=declared(root, "SOURCES.md", "Register", ROW, report),
        files=dataset_files(root, report))


def input_value_problem(key: str, value: object) -> str | None:
    """Check a key's unit suffix and its value's domain."""
    suffix = next((s for s in SUFFIXES if key.endswith(s)), None)
    if suffix is None:
        return f"input {key!r} has no unit suffix"
    if value is None:
        return None
    if suffix in ("_amount", "_weight", "_factor", "_years", "_rate"):
        valid = is_number(value)
    elif suffix == "_count":
        valid = isinstance(value, int) and not isinstance(value, bool) and value >= 0
    elif suffix == "_date":
        valid = is_date(value)
    elif suffix == "_flag":
        valid = isinstance(value, bool)
    elif key.endswith("currency_code"):
        valid = isinstance(value, str) and bool(CURRENCY.match(value))
    else:
        valid = is_text(value)
    return None if valid else f"input {key!r} has an invalid value for its suffix {suffix}"


def check_inputs(inputs: object, where: str, report: Report) -> None:
    if not isinstance(inputs, dict):
        report.add(where, "inputs is not an object")
        return
    for key, value in inputs.items():
        if not SNAKE.match(key):
            report.add(where, f"input {key!r} is not lower snake case")
        elif problem := input_value_problem(key, value):
            report.add(where, problem)
    if any(key.endswith("_amount") for key in inputs) and inputs.get("currency_code") is None:
        report.add(where, "inputs hold an amount but no currency_code")


def check_records(records: object, where: str, seen: set[str], report: Report) -> set[str]:
    """Check one dataset's records; return the IDs of the valid ones."""
    if not isinstance(records, list) or not records:
        report.add(where, "is not a non-empty list")
        return set()
    valid: set[str] = set()
    for index, record in enumerate(records):
        at = f"{where}[{index}]"
        if not isinstance(record, dict):
            report.add(at, "is not an object")
        elif problem := keys_problem(record, RECORD_KEYS):
            report.add(at, problem)
        elif not is_text(record["id"]):
            report.add(at, "id is not a text")
        elif record["id"] in seen:
            report.add(at, f"id {record['id']!r} is not unique in the fixture")
        else:
            seen.add(record["id"])
            check_inputs(record["inputs"], at, report)
            valid.add(record["id"])
    return valid


def harness_problems(data: dict[str, Any], stem: str, kind: str = "test-fixture") -> list[str]:
    """The metadata every case keeps valid, whatever its category."""
    problems = []
    if data["schema_version"] != 1:
        problems.append("schema_version is not 1")
    if data["kind"] != kind:
        problems.append(f"kind is not {kind!r}")
    if data["id"] != stem or not CASE_ID.match(stem):
        problems.append("id is not the kebab-case file name" if kind == "test-fixture"
                        else "id is not the kebab-case directory name")
    if data["synthetic"] is not True:
        problems.append("synthetic is not true")
    if not is_text(data["description"]):
        problems.append("description is not a text")
    if data["category"] not in CATEGORIES:
        problems.append(f"unknown category {data['category']!r}")
    if not is_date(data["reporting_date"]):
        problems.append("reporting_date is not an ISO date")
    return problems


def declared_problem(kind: str, value: object, known: Mapping[str, bool],
                     pattern: re.Pattern[str], case_id: str, guide: str) -> str | None:
    """An ID that the methodology declares; an illustrative one only in an illustrative case."""
    if not isinstance(value, str) or not pattern.match(value):
        return f"{kind} {value!r} is not an ID of the form {pattern.pattern.strip('^$')}"
    if value not in known:
        return f"{kind} {value} is not declared in {guide}"
    if known[value] and not case_id.startswith("illustrative-"):
        return f"{kind} {value} is illustrative and may be used only by an illustrative- case"
    return None


def check_calculation(data: dict[str, Any], case_id: str, method: Methodology,
                      where: str, report: Report) -> None:
    """Check the calculation and the reporting currency."""
    if problem := declared_problem("calculation", data["calculation"], method.calculations,
                                   CALCULATION_ID, case_id, "CALCULATION_CONTRACTS.md"):
        report.add(where, problem)
    code = data["reporting_currency_code"]
    if not isinstance(code, str) or not CURRENCY.match(code):
        report.add(where, "reporting_currency_code is not a currency code")


def check_calculation_inputs(data: dict[str, Any], case_id: str, method: Methodology,
                             where: str, report: Report) -> Case:
    """Check the calculation, the currency and the datasets with their records."""
    check_calculation(data, case_id, method, where, report)
    records: dict[str, set[str]] = {}
    datasets = data["datasets"]
    if not isinstance(datasets, dict) or not datasets:
        report.add(where, "datasets is not a non-empty object")
        datasets = {}
    seen: set[str] = set()
    for dataset, items in datasets.items():
        if problem := declared_problem("dataset", dataset, method.datasets, DATASET_ID,
                                       case_id, "DATA_CONTRACT.md, Datasets"):
            report.add(where, problem)
        records[dataset] = check_records(items, f"{where}: datasets.{dataset}", seen, report)
    calculation = data["calculation"] if isinstance(data["calculation"], str) else None
    return Case(str(data["category"]), calculation, records)


def check_fixture(path: Path, root: Path, method: Methodology, report: Report) -> Case | None:
    where = rel(path, root)
    data = load_json(path, root, report)
    if data is LOAD_FAILED:
        return None
    if not isinstance(data, dict):
        report.add(where, "is not an object")
        return None
    invalid = data.get("category") == "invalid"
    shown = {key: value for key, value in data.items() if key in HARNESS_KEYS} if invalid else data
    if problem := keys_problem(shown, HARNESS_KEYS if invalid else FIXTURE_KEYS):
        report.add(where, problem)
        return None
    for problem in harness_problems(data, path.stem):
        report.add(where, problem)
    if invalid:
        # The calculation inputs are what an invalid case tests: they are not
        # checked, and its outputs are not cross-checked against them.
        return Case("invalid", None, {})
    return check_calculation_inputs(data, path.stem, method, where, report)


def check_variant(directory: Path, root: Path, allowed: Mapping[str, str], invalid: bool,
                  report: Report) -> None:
    """A variant holds only data-contract files, readable in their encoding."""
    where = rel(directory, root)
    entries = sorted(directory.iterdir())
    if not entries:
        report.add(where, "is an empty variant")
    for entry in entries:
        if not entry.is_file():
            report.add(rel(entry, root), "is not a file; a variant holds files only")
        elif entry.name not in allowed:
            report.add(rel(entry, root), "is not a file the data contract lists for the "
                                         "case's datasets")
        elif not invalid:
            # An invalid case may test a file that its encoding cannot decode.
            try:
                entry.read_bytes().decode(allowed[entry.name])
            except UnicodeDecodeError as error:
                report.add(rel(entry, root), f"is not {allowed[entry.name]} ({error.reason} "
                                             f"at byte {error.start})")


def check_file_fixture(directory: Path, root: Path, method: Methodology,
                       report: Report) -> Case | None:
    """Check a file-based case: its manifest, then each variant's files."""
    manifest = directory / MANIFEST
    where = rel(manifest, root)
    if not manifest.is_file():
        report.add(rel(directory, root), f"is a directory without {MANIFEST}")
        return None
    data = load_json(manifest, root, report)
    if data is LOAD_FAILED:
        return None
    if not isinstance(data, dict):
        report.add(where, "is not an object")
        return None
    if problem := keys_problem(data, FIXTURE_KEYS):
        report.add(where, problem)
        return None
    for problem in harness_problems(data, directory.name, "test-file-fixture"):
        report.add(where, problem)
    check_calculation(data, directory.name, method, where, report)
    datasets, allowed = manifest_files(data["datasets"], directory.name, method, where, report)
    variants: set[str] = set()
    for entry in sorted(directory.iterdir()):
        if entry.name == MANIFEST:
            continue
        if not entry.is_dir():
            report.add(rel(entry, root), f"is not {MANIFEST} or a variant directory")
        elif not CASE_ID.match(entry.name):
            report.add(rel(entry, root), "is not a kebab-case variant name")
        else:
            variants.add(entry.name)
            check_variant(entry, root, allowed, data["category"] == "invalid", report)
    if not variants:
        report.add(rel(directory, root), "has no variant directory")
    calculation = data["calculation"] if isinstance(data["calculation"], str) else None
    return Case(str(data["category"]), calculation, {name: set() for name in datasets},
                variants)


def manifest_files(datasets: object, case_id: str, method: Methodology, where: str,
                   report: Report) -> tuple[list[str], dict[str, str]]:
    """The manifest's datasets, and the files they allow with their encodings."""
    if not isinstance(datasets, list) or not datasets or \
            len({str(name) for name in datasets}) != len(datasets):
        report.add(where, "datasets is not a non-empty list of distinct DS- IDs")
        return [], {}
    allowed: dict[str, str] = {}
    for dataset in datasets:
        if problem := declared_problem("dataset", dataset, method.datasets, DATASET_ID,
                                       case_id, "DATA_CONTRACT.md, Datasets"):
            report.add(where, problem)
        elif not method.files.get(dataset):
            report.add(where, f"dataset {dataset} names no file in the Datasets table of "
                              "DATA_CONTRACT.md")
        else:
            allowed.update(method.files[dataset])
    return [str(name) for name in datasets], allowed


def tolerance_problem(tolerance: object) -> str | None:
    if not isinstance(tolerance, dict) or set(tolerance) != {"absolute", "relative"}:
        return "tolerance must hold exactly absolute and relative"
    if not all(is_number(bound) and bound >= 0 for bound in tolerance.values()):
        return "tolerance bounds must be finite numbers of at least 0"
    return None


def form_problems(output: dict[str, Any]) -> list[str]:
    if "text" in output:
        if problem := keys_problem(output, TEXT_KEYS):
            return [problem]
        return [] if isinstance(output["text"], str) else ["text is not a text"]
    if problem := keys_problem(output, NUMBER_KEYS):
        return [problem]
    problems = []
    if not is_number(output["value"]):
        problems.append("value is not a finite number")
    unit = output["unit"]
    if not isinstance(unit, str) or not (unit in UNITS or CURRENCY.match(unit)):
        problems.append(f"unit {unit!r} is not a currency code, 1 or years")
    if problem := tolerance_problem(output["tolerance"]):
        problems.append(problem)
    return problems


def source_problems(kind: str, source: object, reference: Mapping[str, Any],
                    sources: Mapping[str, bool]) -> list[str]:
    problems = []
    if source is not None and source not in sources:
        problems.append(f"source {source!r} is not in the register (SOURCES.md)")
    elif isinstance(source, str) and sources[source] and kind != "illustrative":
        problems.append(f"source {source} is an illustrative register entry")
    if kind == "published" and (not isinstance(source, str) or source.startswith("IMPL-")
                                or not is_text(reference["locator"])):
        problems.append("a published value needs a register source (not IMPL-) and a locator")
    if kind == "independent" and not (isinstance(source, str) and source.startswith("IMPL-")
                                      and all(is_text(reference[key]) for key in
                                              ("locator", "derivation", "derived_by"))):
        problems.append("an independent value needs an IMPL- source, locator, derivation "
                        "and derived_by")
    if kind == "illustrative" and not is_text(reference["derivation"]):
        problems.append("an illustrative value needs a derivation")
    return problems


def reference_problems(reference: object, sources: Mapping[str, bool], root: Path) -> list[str]:
    if not isinstance(reference, dict):
        return ["reference is not an object"]
    if problem := keys_problem(reference, REFERENCE_KEYS):
        return [f"reference: {problem}"]
    kind = reference["class"]
    if kind not in CLASSES:
        return [f"unknown reference class {kind!r}"]
    problems = source_problems(kind, reference["source"], reference, sources)
    problems.extend(script_problems(reference["script"], root))
    if not is_date(reference["date"]):
        problems.append("reference date is not an ISO date")
    for key in ("derived_by", "reviewed_by"):
        if reference[key] is not None and not is_text(reference[key]):
            problems.append(f"{key} is not a text or null")
    return problems


def script_problems(script: object, root: Path) -> list[str]:
    if script is None:
        return []
    if (not isinstance(script, str) or not script.startswith(DERIVATIONS) or ".." in script
            or not (root / script).is_file()):
        return [f"script {script!r} is not an existing file under {DERIVATIONS}"]
    return []


def scope_problems(scope: object, case: Case) -> list[str]:
    """A whole-calculation output has scope null; a record output names a fixture record."""
    if scope is None:
        return []
    if not isinstance(scope, dict) or set(scope) != SCOPE_KEYS:
        return ["scope is null or an object with exactly dataset and record"]
    if case.category == "invalid":
        return [] if all(is_text(value) for value in scope.values()) else [
            "scope dataset and record are texts"]
    if case.variants is not None:
        # The records of a file-based case are rows of its files, not listed.
        if scope["dataset"] not in case.records or not is_text(scope["record"]):
            return [f"scope dataset {scope['dataset']!r} is not a dataset of the case, or "
                    "its record is not a text"]
        return []
    if scope["record"] not in case.records.get(scope["dataset"], set()):
        return [f"scope {scope['dataset']}/{scope['record']} is not a record of the fixture"]
    return []


def rule_problem(rule: object, kind: object, case: Case) -> str | None:
    if rule is None:
        return None if kind == "illustrative" else "only an illustrative value may have no rule"
    match = RULE.match(rule) if isinstance(rule, str) else None
    if match is None:
        return f"rule {rule!r} is not a calculation step such as CALC-01.1"
    if case.calculation is not None and match.group(1) != case.calculation:
        return f"rule {rule} is not a step of the case's calculation {case.calculation}"
    return None


def check_output(output: object, case_id: str, case: Case, method: Methodology,
                 root: Path) -> list[str]:
    if not isinstance(output, dict):
        return ["output is not an object"]
    problems = form_problems(output)
    if problems:
        return problems
    quantity = output["quantity"]
    if not isinstance(quantity, str) or not SNAKE.match(quantity):
        problems.append(f"quantity {quantity!r} is not lower snake case")
    elif quantity.startswith("illustrative_") and not case_id.startswith("illustrative-"):
        problems.append("an illustrative_ quantity outside an illustrative- case")
    problems.extend(scope_problems(output["scope"], case))
    reference = output["reference"]
    kind = reference.get("class") if isinstance(reference, dict) else None
    if problem := rule_problem(output["rule"], kind, case):
        problems.append(problem)
    problems.extend(reference_problems(reference, method.sources, root))
    if isinstance(kind, str):
        case.classes.add(kind)
    return problems


def check_outputs(outputs: object, where: str, case_id: str, case: Case, method: Methodology,
                  root: Path, report: Report) -> None:
    """Check one list of outputs: a case's, or one variant's."""
    items = outputs if isinstance(outputs, list) else []
    if not items:
        report.add(where, "outputs is not a non-empty list")
    for index, output in enumerate(items):
        for problem in check_output(output, case_id, case, method, root):
            report.add(f"{where}: outputs[{index}]", problem)
    statuses = [output for output in items if isinstance(output, dict) and "text" in output
                and str(output.get("quantity", "")).endswith("_status")]
    if case.category == "invalid" and not statuses:
        report.add(where, "an invalid case needs a _status text output")


def check_variant_outputs(variants: object, where: str, case_id: str, case: Case,
                          method: Methodology, root: Path, report: Report) -> None:
    """Every variant of a file-based case has outputs, and every output a variant."""
    if not isinstance(variants, dict) or not variants:
        report.add(where, "variants is not a non-empty object")
        return
    for variant in sorted((case.variants or set()) - variants.keys()):
        report.add(where, f"variant {variant} has no expected outputs")
    for variant, outputs in variants.items():
        if variant not in (case.variants or set()):
            report.add(where, f"variant {variant!r} has no directory under "
                              f"{FIXTURES.as_posix()}/{case_id}/")
            continue
        check_outputs(outputs, f"{where}: variants.{variant}", case_id, case, method, root,
                      report)


def check_expected(path: Path, root: Path, cases: Mapping[str, Case], method: Methodology,
                   report: Report) -> None:
    where = rel(path, root)
    case_id = path.stem
    if case_id not in cases:
        report.add(where, "is not named <case-id>.json for an existing fixture")
        return
    data = load_json(path, root, report)
    if data is LOAD_FAILED:
        return
    if not isinstance(data, dict):
        report.add(where, "is not an object")
        return
    case = cases[case_id]
    file_based = case.variants is not None
    if problem := keys_problem(data, EXPECTED_FILE_KEYS if file_based else EXPECTED_KEYS):
        report.add(where, problem + (" (a file-based case lists outputs by variant)"
                                     if file_based else ""))
        return
    if data["schema_version"] != 1 or data["kind"] != "test-expected" or \
            data["fixture"] != case_id:
        report.add(where, "envelope does not match: schema_version 1, kind 'test-expected', "
                          "fixture as in the file name")
    case.expected_files += 1
    if file_based:
        check_variant_outputs(data["variants"], where, case_id, case, method, root, report)
    else:
        check_outputs(data["outputs"], where, case_id, case, method, root, report)


def check_cases(cases: Mapping[str, Case], report: Report) -> None:
    for case_id, case in sorted(cases.items()):
        where = (f"{FIXTURES.as_posix()}/{case_id}/{MANIFEST}" if case.variants is not None
                 else f"{FIXTURES.as_posix()}/{case_id}.json")
        if case.expected_files == 0:
            report.add(where, "has no expected file")
        elif (case.classes == {"illustrative"}) != case_id.startswith("illustrative-"):
            report.add(where, "the illustrative- prefix must be used exactly when every "
                              "expected value is illustrative")


def check(root: Path) -> tuple[list[str], int, int]:
    """Return the findings and the numbers of fixtures and expected files checked."""
    report = Report()
    directory = root / FIXTURES
    fixtures = sorted(directory.glob("*.json"))
    file_cases = sorted(path for path in directory.glob("*") if path.is_dir())
    expected = sorted((root / EXPECTED).glob("*.json"))
    if not fixtures and not file_cases and not expected:
        return [], 0, 0
    method = read_methodology(root, report)
    cases: dict[str, Case] = {}
    for path in fixtures:
        if (case := check_fixture(path, root, method, report)) is not None:
            cases[path.stem] = case
    for path in file_cases:
        if (path.parent / f"{path.name}.json").exists():
            report.add(rel(path, root), f"is a directory, and {path.name}.json is a second "
                                        "fixture of the same case")
        elif (case := check_file_fixture(path, root, method, report)) is not None:
            cases[path.name] = case
    for path in expected:
        check_expected(path, root, cases, method, report)
    check_cases(cases, report)
    return report.findings, len(fixtures) + len(file_cases), len(expected)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT,
                        help="repository root (default: this checkout)")
    args = parser.parse_args(argv)
    findings, fixtures, expected = check(args.root.resolve())
    for finding in findings:
        print(f"FAIL: {finding}")
    status = "FAIL" if findings else "PASS"
    found = (f"{fixtures} fixture(s), {expected} expected file(s)" if fixtures or expected
             else "no test cases found")
    print(f"test-cases: {status} ({found}; format only, no value computed)")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
