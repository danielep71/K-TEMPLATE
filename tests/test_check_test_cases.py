"""Valid and invalid test-case fixtures for tools/check_test_cases.py.

Each test writes a temporary repository holding minimal methodology guides,
one fixture and its expected file, then changes one thing. The data is
synthetic and has no regulatory meaning.
"""

from __future__ import annotations

import contextlib
import copy
import io
import json
import sys
import tempfile
import unittest
from collections.abc import Callable
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_test_cases  # noqa: E402

CONTRACTS = "# Contracts\n\n### `CALC-02` Real calculation\n\n### *Illustrative:* `CALC-01` Example\n"
DATA = ("# Data\n\n## Datasets\n\n| ID | Description |\n| --- | --- |\n"
        "| `DS-02` | Real positions |\n| *Illustrative:* `DS-01` | Example |\n\n"
        "## Fields\n\n| `DS-09` | Not in the Datasets section |\n")
SOURCES = ("# Sources\n\n## Register\n\n| ID | Title |\n| --- | --- |\n"
           "| `SRC-02` | A regulation |\n| `IMPL-01` | An external implementation |\n"
           "| *Illustrative:* `SRC-01` | Example |\n")
FIXTURE: dict[str, Any] = {
    "schema_version": 1, "kind": "test-fixture", "id": "case-one", "synthetic": True,
    "description": "A synthetic case", "category": "nominal", "calculation": "CALC-02",
    "reporting_date": "2026-01-02", "reporting_currency_code": "EUR",
    "datasets": {"DS-02": [
        {"id": "P1", "inputs": {"currency_code": "EUR", "notional_amount": 600000,
                                "maturity_years": 2.5, "trade_count": 3, "start_date": "2025-01-02",
                                "netted_flag": False, "position_id": "X1", "risk_weight": 0.2}},
        {"id": "P2", "inputs": {"currency_code": "EUR", "notional_amount": 400000}},
    ]},
}
OUTPUT: dict[str, Any] = {
    "quantity": "total_notional_amount", "scope": None, "rule": "CALC-02.1",
    "value": 1000000, "unit": "EUR", "tolerance": {"absolute": 0.000001, "relative": 0},
    "reference": {"class": "published", "source": "SRC-02", "locator": "Article 1(1)",
                  "derivation": None, "script": None, "derived_by": None,
                  "reviewed_by": None, "date": "2026-11-02"},
}
EXPECTED: dict[str, Any] = {"schema_version": 1, "kind": "test-expected", "fixture": "case-one",
                            "outputs": [OUTPUT]}

Edit = Callable[[dict[str, Any], dict[str, Any]], None]


def run(edit: Edit | None = None, name: str = "case-one",
        extra: dict[str, str] | None = None) -> list[str]:
    """Write the repository, apply edit(fixture, expected), and check it."""
    fixture, expected = copy.deepcopy(FIXTURE), copy.deepcopy(EXPECTED)
    if edit is not None:
        edit(fixture, expected)
    files = {"docs/methodology/CALCULATION_CONTRACTS.md": CONTRACTS,
             "docs/methodology/DATA_CONTRACT.md": DATA,
             "docs/methodology/SOURCES.md": SOURCES,
             f"tests/fixtures/{name}.json": json.dumps(fixture),
             f"tests/expected/{name}.json": json.dumps(expected)}
    files.update(extra or {})
    with tempfile.TemporaryDirectory(prefix="test-cases-") as directory:
        root = Path(directory)
        for path, text in files.items():
            (root / path).parent.mkdir(parents=True, exist_ok=True)
            (root / path).write_text(text, encoding="utf-8")
        return check_test_cases.check(root)[0]


def output(expected: dict[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = expected["outputs"][0]
    return result


class ValidCasesTest(unittest.TestCase):
    def test_valid_case_passes(self) -> None:
        self.assertEqual(run(), [])

    def test_repository_cases_pass(self) -> None:
        findings, fixtures, expected = check_test_cases.check(ROOT)
        self.assertEqual(findings, [])
        self.assertGreaterEqual(fixtures, 1)
        self.assertEqual(fixtures, expected)

    def test_no_cases_pass_without_reading_the_guides(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            self.assertEqual(check_test_cases.check(Path(directory)), ([], 0, 0))

    def test_record_scope_and_text_output(self) -> None:
        def edit(fixture: dict[str, Any], expected: dict[str, Any]) -> None:
            output(expected)["scope"] = {"dataset": "DS-02", "record": "P2"}
            status = {key: OUTPUT[key] for key in ("quantity", "scope", "rule", "reference")}
            expected["outputs"].append({**status, "quantity": "run_status", "text": "OK"})
        self.assertEqual(run(edit), [])

    def test_illustrative_case_may_use_illustrative_entries(self) -> None:
        def edit(fixture: dict[str, Any], expected: dict[str, Any]) -> None:
            fixture.update(id="illustrative-one", calculation="CALC-01")
            fixture["datasets"] = {"DS-01": fixture["datasets"].pop("DS-02")}
            expected["fixture"] = "illustrative-one"
            output(expected).update(quantity="illustrative_total_amount", rule=None)
            output(expected)["reference"].update({"class": "illustrative", "source": None,
                                                  "locator": None,
                                                  "derivation": "Sum of the two notionals"})
        self.assertEqual(run(edit, name="illustrative-one"), [])

    def test_independent_reference(self) -> None:
        def edit(fixture: dict[str, Any], expected: dict[str, Any]) -> None:
            output(expected)["reference"].update({
                "class": "independent", "source": "IMPL-01", "locator": "Run of 2026-11-02",
                "derivation": "Computed by the external tool", "derived_by": "Tool 1.0"})
        self.assertEqual(run(edit), [])


class FindingTest(unittest.TestCase):
    def assertFinding(self, findings: list[str], text: str) -> None:
        self.assertTrue(any(text in finding for finding in findings), findings)


class FileTest(FindingTest):
    def test_strict_json(self) -> None:
        for raw, text in (('{"a": 1, "a": 2}', "duplicate key"), ('{"a": NaN}', "non-standard"),
                          ('{"a": 1e400}', "not finite"), ("[", "not valid UTF-8 JSON")):
            with self.subTest(raw=raw):
                self.assertFinding(run(extra={"tests/fixtures/case-one.json": raw}), text)

    def test_fixture_must_be_an_object(self) -> None:
        self.assertFinding(run(extra={"tests/fixtures/case-one.json": "null"}), "is not an object")

    def test_expected_needs_a_fixture_and_the_reverse(self) -> None:
        self.assertFinding(run(extra={"tests/expected/other.json": "{}"}),
                           "for an existing fixture")
        orphan = json.dumps({**FIXTURE, "id": "orphan"})
        self.assertFinding(run(extra={"tests/fixtures/orphan.json": orphan}), "has no expected file")

    def test_missing_guide(self) -> None:
        self.assertFinding(run(extra={"docs/methodology/SOURCES.md": ""}),
                           "source 'SRC-02' is not in the register")
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for folder, data in (("tests/fixtures", FIXTURE), ("tests/expected", EXPECTED)):
                (root / folder).mkdir(parents=True)
                (root / folder / "case-one.json").write_text(json.dumps(data), encoding="utf-8")
            self.assertFinding(check_test_cases.check(root)[0],
                               "docs/methodology/DATA_CONTRACT.md: is missing")


class FixtureTest(FindingTest):
    def test_envelope(self) -> None:
        for key, value, text in (("schema_version", 2, "schema_version is not 1"),
                                 ("kind", "fixture", "kind is not 'test-fixture'"),
                                 ("id", "case-two", "id is not the kebab-case file name"),
                                 ("synthetic", False, "synthetic is not true"),
                                 ("description", " ", "description is not a text"),
                                 ("category", "smoke", "unknown category 'smoke'"),
                                 ("reporting_date", "2026-02-30", "reporting_date is not"),
                                 ("reporting_currency_code", "eur", "is not a currency code")):
            with self.subTest(key=key):
                self.assertFinding(run(lambda f, e, k=key, v=value: f.update({k: v})), text)

    def test_missing_and_unknown_keys(self) -> None:
        self.assertFinding(run(lambda f, e: f.pop("datasets")), "missing datasets")
        self.assertFinding(run(lambda f, e: f.update(extra=1)), "unknown extra")

    def test_calculation_and_datasets_are_declared(self) -> None:
        self.assertFinding(run(lambda f, e: f.update(calculation="CALC-09")),
                           "calculation CALC-09 is not declared")
        self.assertFinding(run(lambda f, e: f.update(calculation="calc-2")),
                           "is not an ID of the form")
        self.assertFinding(run(lambda f, e: f.update(datasets={"DS-09": FIXTURE["datasets"]
                                                               ["DS-02"]})),
                           "dataset DS-09 is not declared")
        self.assertFinding(run(lambda f, e: f.update(calculation="CALC-01")),
                           "CALC-01 is illustrative")
        self.assertFinding(run(lambda f, e: f.update(datasets={})), "not a non-empty object")

    def test_records(self) -> None:
        def records(value: Any) -> Edit:
            return lambda f, e: f["datasets"].update({"DS-02": value})
        self.assertFinding(run(records([])), "is not a non-empty list")
        self.assertFinding(run(records(["P1"])), "is not an object")
        self.assertFinding(run(records([{"id": "P1"}])), "missing inputs")
        self.assertFinding(run(records([{"id": " ", "inputs": {}}])), "id is not a text")
        self.assertFinding(run(records([{"id": "P1", "inputs": {}}, {"id": "P1", "inputs": {}}])),
                           "is not unique")
        self.assertFinding(run(records([{"id": "P1", "inputs": []}])), "inputs is not an object")

    def test_input_suffixes_and_values(self) -> None:
        for key, value, text in (("notional", 1, "has no unit suffix"),
                                 ("NotionalAmount", 1, "not lower snake case"),
                                 ("notional_amount", "1", "invalid value for its suffix _amount"),
                                 ("trade_count", -1, "suffix _count"),
                                 ("trade_count", True, "suffix _count"),
                                 ("start_date", "2026-13-01", "suffix _date"),
                                 ("netted_flag", 1, "suffix _flag"),
                                 ("currency_code", "EURO", "suffix _code"),
                                 ("position_id", "", "suffix _id")):
            with self.subTest(key=key, value=value):
                def edit(f: dict[str, Any], e: dict[str, Any], k: str = key,
                         v: object = value) -> None:
                    f["datasets"]["DS-02"][0]["inputs"][k] = v
                self.assertFinding(run(edit), text)
        self.assertFinding(run(lambda f, e: f["datasets"]["DS-02"][1]["inputs"]
                               .pop("currency_code")), "amount but no currency_code")

    def test_null_input_is_not_applicable(self) -> None:
        def edit(f: dict[str, Any], e: dict[str, Any]) -> None:
            f["datasets"]["DS-02"][0]["inputs"]["maturity_years"] = None
        self.assertEqual(run(edit), [])

    def test_invalid_case_checks_only_the_harness_metadata(self) -> None:
        def edit(f: dict[str, Any], e: dict[str, Any]) -> None:
            f.update(category="invalid", calculation="unknown", datasets={"bad": "x"})
            f.pop("reporting_currency_code")
            e["outputs"] = [{"quantity": "run_status", "text": "REJECTED_UNKNOWN_CALCULATION",
                             "scope": {"dataset": "bad", "record": "R9"}, "rule": "CALC-02.1",
                             "reference": OUTPUT["reference"]}]
        self.assertEqual(run(edit), [])

        def no_status(f: dict[str, Any], e: dict[str, Any]) -> None:
            f["category"] = "invalid"
        self.assertFinding(run(no_status), "needs a _status text output")
        self.assertFinding(run(lambda f, e: f.update(category="invalid", synthetic=False)),
                           "synthetic is not true")


class ExpectedTest(FindingTest):
    def test_envelope(self) -> None:
        self.assertFinding(run(lambda f, e: e.update(kind="expected")), "envelope does not match")
        self.assertFinding(run(lambda f, e: e.update(fixture="other")), "envelope does not match")
        self.assertFinding(run(lambda f, e: e.pop("outputs")), "missing outputs")
        self.assertFinding(run(lambda f, e: e.update(outputs=[])), "not a non-empty list")
        self.assertFinding(run(lambda f, e: e.update(outputs=["x"])), "output is not an object")

    def test_number_form(self) -> None:
        for key, value, text in (("value", "1", "value is not a finite number"),
                                 ("unit", "eur", "is not a currency code, 1 or years"),
                                 ("tolerance", {"absolute": 0}, "exactly absolute and relative"),
                                 ("tolerance", {"absolute": -1, "relative": 0}, "at least 0")):
            with self.subTest(key=key):
                self.assertFinding(run(lambda f, e, k=key, v=value: output(e).update({k: v})),
                                   text)
        self.assertFinding(run(lambda f, e: output(e).pop("tolerance")), "missing tolerance")
        self.assertFinding(run(lambda f, e: output(e).update(text=1)), "unknown")

    def test_quantity_scope_and_rule(self) -> None:
        cases: list[tuple[Edit, str]] = [
            (lambda f, e: output(e).update(quantity="Total"), "not lower snake case"),
            (lambda f, e: output(e).update(quantity="illustrative_total_amount"),
             "illustrative_ quantity outside"),
            (lambda f, e: output(e).update(scope={"dataset": "DS-02", "record": "P9"}),
             "DS-02/P9 is not a record"),
            (lambda f, e: output(e).update(scope={"dataset": "DS-02"}), "exactly dataset and"),
            (lambda f, e: output(e).update(rule=None), "only an illustrative value may have"),
            (lambda f, e: output(e).update(rule="CALC-02"), "is not a calculation step"),
            (lambda f, e: output(e).update(rule="CALC-01.1"), "not a step of the case's"),
        ]
        for edit, text in cases:
            with self.subTest(text=text):
                self.assertFinding(run(edit), text)

    def test_reference(self) -> None:
        def ref(**values: object) -> Edit:
            return lambda f, e: output(e)["reference"].update(values)
        cases: list[tuple[Edit, str]] = [
            (lambda f, e: output(e).update(reference=[]), "reference is not an object"),
            (lambda f, e: output(e)["reference"].pop("date"), "reference: missing date"),
            (ref(**{"class": "guess"}), "unknown reference class"),
            (ref(source="SRC-09"), "not in the register"),
            (ref(source="SRC-01"), "illustrative register entry"),
            (ref(locator=None), "published value needs"),
            (ref(source="IMPL-01"), "published value needs"),
            (ref(**{"class": "independent"}), "independent value needs"),
            (ref(**{"class": "illustrative"}), "illustrative value needs a derivation"),
            (ref(script="tools/x.py"), "is not an existing file under tests/derivations/"),
            (ref(script="tests/derivations/missing.py"), "is not an existing file"),
            (ref(date="02/11/2026"), "reference date is not"),
            (ref(derived_by=" "), "derived_by is not a text or null"),
        ]
        for edit, text in cases:
            with self.subTest(text=text):
                self.assertFinding(run(edit), text)

    def test_existing_derivation_script(self) -> None:
        self.assertEqual(run(lambda f, e: output(e)["reference"].update(
            script="tests/derivations/total.py"),
            extra={"tests/derivations/total.py": "print(1000000)\n"}), [])

    def test_illustrative_prefix_matches_the_classes(self) -> None:
        def all_illustrative(f: dict[str, Any], e: dict[str, Any]) -> None:
            output(e)["reference"].update({"class": "illustrative", "derivation": "Example"})
        self.assertFinding(run(all_illustrative), "illustrative- prefix must be used exactly")


class MainTest(unittest.TestCase):
    def test_verdicts(self) -> None:
        for directory, code, verdict in ((ROOT, 0, "test-cases: PASS (1 fixture(s)"),):
            with self.subTest(directory=directory), \
                    contextlib.redirect_stdout(io.StringIO()) as printed:
                self.assertEqual(check_test_cases.main(["--root", str(directory)]), code)
                self.assertIn(verdict, printed.getvalue())
        with tempfile.TemporaryDirectory() as empty, \
                contextlib.redirect_stdout(io.StringIO()) as printed:
            self.assertEqual(check_test_cases.main(["--root", empty]), 0)
            self.assertIn("test-cases: PASS (no test cases found", printed.getvalue())


if __name__ == "__main__":
    unittest.main()
