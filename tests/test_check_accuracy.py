"""Valid and invalid accuracy contracts and results for tools/check_accuracy.py.

Each test builds a throw-away repository from this one's methodology guides
and test cases, adds a non-illustrative calculation CALC-02 with synthetic
cases whose values are marked independent, and writes a registry. The values
are fixtures for the check only: no contract here guarantees anything.
"""

from __future__ import annotations

import contextlib
import copy
import io
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from collections.abc import Iterator
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_accuracy  # noqa: E402

QUANTITY = "total_exposure_amount"
CANDIDATE = "a" * 40
ILLUSTRATIVE = {
    "id": "ACC-01", "calculation": "CALC-01", "regime": "illustrative",
    "quantities": ["illustrative_total_notional_amount"], "metric": "absolute",
    "threshold": 0.000001, "provenance": "Format example", "status": "draft",
    "decision": None, "calibration_cases": ["illustrative-format-example"],
    "holdout_cases": ["illustrative-file-import-example"],
}
FROZEN = {
    "id": "ACC-02", "calculation": "CALC-02", "regime": "CRR3", "quantities": [QUANTITY],
    "metric": "relative", "threshold": 0.0001, "provenance": "Synthetic fixture",
    "status": "frozen", "decision": "#1", "calibration_cases": ["demo-calibration"],
    "holdout_cases": ["demo-holdout-1", "demo-holdout-2"],
}


def registry(*contracts: dict[str, Any], minimum: int = 2) -> dict[str, Any]:
    return {"schema_version": 1, "kind": "accuracy-contracts", "min_holdout_cases": minimum,
            "contracts": [copy.deepcopy(item) for item in contracts]}


def demo_case(root: Path, case: str, value: float, reference: str = "independent") -> None:
    fixture = {"schema_version": 1, "kind": "test-fixture", "id": case, "synthetic": True,
               "description": "Synthetic", "category": "nominal", "calculation": "CALC-02",
               "reporting_date": "2026-01-02", "reporting_currency_code": "EUR",
               "datasets": {"DS-01": [{"id": "P1", "inputs": {"currency_code": "EUR"}}]}}
    expected = {"schema_version": 1, "kind": "test-expected", "fixture": case, "outputs": [
        {"quantity": QUANTITY, "scope": None, "rule": "CALC-02.1", "value": value,
         "unit": "EUR", "tolerance": {"absolute": 0.01, "relative": 0},
         "reference": {"class": reference}}]}
    (root / "tests/fixtures" / f"{case}.json").write_text(json.dumps(fixture), "utf-8")
    (root / "tests/expected" / f"{case}.json").write_text(json.dumps(expected), "utf-8")


def git(root: Path, *args: str) -> None:
    subprocess.run(["git", "-C", str(root), "-c", "user.name=Fixture",
                    "-c", "user.email=fixture@example.invalid", "-c", "commit.gpgsign=false",
                    "-c", "tag.gpgsign=false", *args], capture_output=True, check=True)


@contextlib.contextmanager
def repository(content: dict[str, Any] | None) -> Iterator[Path]:
    """This repository's guides and cases, CALC-02 and its three cases, and ``content``."""
    with tempfile.TemporaryDirectory(prefix="accuracy-") as temporary:
        root = Path(temporary)
        for name in ("docs/methodology", "tests/fixtures", "tests/expected"):
            shutil.copytree(ROOT / name, root / name)
        contracts = root / "docs/methodology/CALCULATION_CONTRACTS.md"
        contracts.write_text(contracts.read_text("utf-8") + "\n### `CALC-02` Demo\n", "utf-8")
        demo_case(root, "demo-calibration", 100.0)
        demo_case(root, "demo-holdout-1", 200.0)
        demo_case(root, "demo-holdout-2", 0.0)
        if content is not None:
            (root / "tests/accuracy").mkdir(parents=True)
            (root / "tests/accuracy/contracts.json").write_text(json.dumps(content), "utf-8")
        yield root


def findings(content: dict[str, Any] | None, results: dict[str, Any] | None = None,
             commit: str | None = CANDIDATE) -> list[str]:
    with repository(content) as root:
        path = None
        if results is not None:
            path = root / "results.json"
            path.write_text(json.dumps(results), "utf-8")
        return check_accuracy.check(root, path, commit)[0]


def results(**values: float) -> dict[str, Any]:
    """A results file with ``values`` for the CALC-02 cases (underscores for hyphens)."""
    return {"schema_version": 1, "kind": "test-results", "commit": CANDIDATE, "results": [
        {"case": case.replace("_", "-"), "variant": None, "quantity": QUANTITY, "scope": None,
         "value": value} for case, value in values.items()]}


class RegistryTest(unittest.TestCase):
    def assertFinding(self, found: list[str], text: str) -> None:
        self.assertTrue(any(text in item for item in found), found)

    def test_the_repository_registry_passes(self) -> None:
        self.assertEqual(check_accuracy.check(ROOT)[:2], ([], 1))
        with contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(check_accuracy.main(["--root", str(ROOT)]), 0)
        self.assertIn("accuracy: PASS (1 contract(s); registry only", output.getvalue())

    def test_valid_draft_and_frozen_contracts_pass(self) -> None:
        self.assertEqual(findings(registry(ILLUSTRATIVE, FROZEN)), [])

    def test_no_registry_passes_unless_results_are_given(self) -> None:
        self.assertEqual(findings(None), [])
        self.assertFinding(findings(None, results(demo_calibration=100)),
                           "is missing, so there is nothing to measure")

    def test_malformed_contracts(self) -> None:
        cases: dict[str, dict[str, Any]] = {
            "id is not ACC-<number>": {"id": "X-1"},
            "regime is not a text": {"regime": " "},
            "quantities is not a non-empty list": {"quantities": []},
            "metric is not absolute or relative": {"metric": "squared"},
            "threshold is not a finite number of at least 0": {"threshold": -1},
            "status is not draft or frozen": {"status": "final"},
            "a frozen contract names the decision": {"decision": None},
            "holdout_cases is not a list of case IDs": {"holdout_cases": "demo-holdout-1"},
        }
        for text, change in cases.items():
            with self.subTest(text=text):
                self.assertFinding(findings(registry({**FROZEN, **change})), text)
        self.assertFinding(findings(registry({**ILLUSTRATIVE, "decision": "#2"})),
                           "a draft contract has no decision")
        self.assertFinding(findings(registry(FROZEN, FROZEN)), "id appears twice")
        # A list id cannot be compared for repeats: a finding, not a traceback.
        self.assertFinding(findings(registry({**FROZEN, "id": ["ACC-01"]}, FROZEN)),
                           "id is not ACC-<number>")
        self.assertFinding(findings(registry({**FROZEN, "extra": 1})), "keys are not")
        self.assertFinding(findings(registry({**FROZEN, "quantities": [QUANTITY, QUANTITY]})),
                           "quantities lists a name twice")
        self.assertFinding(findings(registry(FROZEN, minimum=0)),
                           "min_holdout_cases is not a whole number of at least 1")
        self.assertFinding(findings({"schema_version": 2}), "envelope is not")

    def test_cases_must_fit_the_contract(self) -> None:
        cases: dict[str, dict[str, Any]] = {
            "calculation CALC-09 is not declared": {"calculation": "CALC-09"},
            "case missing-case has no fixture": {"calibration_cases": ["missing-case"]},
            "case illustrative-format-example tests CALC-01, not CALC-02":
                {"calibration_cases": ["illustrative-format-example"]},
            "case demo-calibration has no expected number for other_amount":
                {"quantities": [QUANTITY, "other_amount"]},
            "case demo-holdout-1 is both calibration and holdout":
                {"calibration_cases": ["demo-holdout-1"]},
            "lists no case": {"calibration_cases": [], "holdout_cases": [], "status": "draft",
                              "decision": None},
        }
        for text, change in cases.items():
            with self.subTest(text=text):
                self.assertFinding(findings(registry({**FROZEN, **change})), text)

    def test_a_repeated_case_does_not_count_twice(self) -> None:
        repeated = {**FROZEN, "holdout_cases": ["demo-holdout-1"] * 3}
        found = findings(registry(repeated, minimum=3))
        self.assertTrue(any("holdout_cases lists a case twice" in item for item in found), found)
        self.assertTrue(any("calibration_cases lists a case twice" in item for item in findings(
            registry({**FROZEN, "calibration_cases": ["demo-calibration"] * 2}))))

    def test_a_frozen_contract_needs_independent_holdout_cases(self) -> None:
        self.assertFinding(findings(registry(FROZEN, minimum=3)),
                           "a frozen contract needs at least 3 holdout case(s), has 2")
        self.assertFinding(findings(registry({**ILLUSTRATIVE, "status": "frozen",
                                              "decision": "#1"}, minimum=1)),
                           "calculation CALC-01 is illustrative, so the contract stays draft")
        with repository(registry(FROZEN)) as root:
            demo_case(root, "demo-holdout-2", 0.0, reference="illustrative")
            self.assertTrue(any("holdout output demo-holdout-2 total_exposure_amount is not "
                                "published or independent" in item
                                for item in check_accuracy.check(root)[0]))
        self.assertEqual(findings(registry({**FROZEN, "status": "draft", "decision": None},
                                           minimum=5)), [])


class ResultsTest(unittest.TestCase):
    def test_a_malformed_results_file_outside_the_root_is_a_finding(self) -> None:
        # An Excel run's results file lives in its evidence folder, not in the checkout.
        with repository(registry(FROZEN)) as root, \
                tempfile.TemporaryDirectory() as outside:
            path = Path(outside) / "results.json"
            path.write_text("{broken", "utf-8")
            found, _, lines = check_accuracy.check(root, path, CANDIDATE)
        self.assertEqual(lines, [])
        self.assertEqual(len(found), 1, found)
        self.assertIn(f"{path}: not valid UTF-8 JSON", found[0])

    def test_results_within_every_bound_pass_and_are_reported(self) -> None:
        good = results(demo_calibration=100.005, demo_holdout_1=200.001, demo_holdout_2=0)
        with repository(registry(FROZEN)) as root:
            path = root / "results.json"
            path.write_text(json.dumps(good), "utf-8")
            found, contracts, lines = check_accuracy.check(root, path, CANDIDATE)
        self.assertEqual((found, contracts), ([], 1))
        self.assertEqual([line.split(" error ")[0] for line in lines],
                         ["ACC-02 calibration: worst relative", "ACC-02 holdout: worst relative"])
        worst = [float(line.split(" error ")[1].split(",")[0]) for line in lines]
        self.assertAlmostEqual(worst[0], 5e-05, places=12)
        self.assertAlmostEqual(worst[1], 5e-06, places=12)
        self.assertTrue(all(line.endswith("threshold 0.0001") for line in lines))

    def test_a_result_outside_its_own_tolerance_fails(self) -> None:
        found = findings(registry(FROZEN), results(demo_calibration=100.02, demo_holdout_1=200,
                                                   demo_holdout_2=0))
        self.assertTrue(any("demo-calibration total_exposure_amount is 100.02, expected 100.0 "
                            "within 0.01" in item for item in found), found)

    def test_the_worst_error_is_checked_per_split(self) -> None:
        # Within each case's own tolerance of 0.01, but above the contract's absolute
        # threshold of 0.001 on the holdout only.
        loose = {**FROZEN, "metric": "absolute", "threshold": 0.001}
        found = findings(registry(loose), results(demo_calibration=100.0005,
                                                  demo_holdout_1=200.005, demo_holdout_2=0))
        self.assertEqual(len(found), 1, found)
        self.assertIn("worst absolute error 0.00499", found[0])
        self.assertIn("over holdout cases exceeds the threshold 0.001", found[0])

    def test_a_relative_error_at_zero_and_a_missing_result(self) -> None:
        found = findings(registry({**FROZEN, "holdout_cases": ["demo-holdout-2"],
                                   "status": "draft", "decision": None}),
                         results(demo_calibration=100, demo_holdout_2=0.001))
        self.assertTrue(any("worst relative error inf over holdout cases" in item
                            for item in found), found)
        missing = findings(registry(FROZEN), results(demo_calibration=100))
        self.assertTrue(any("no result for demo-holdout-1 total_exposure_amount" in item
                            for item in missing), missing)

    def test_results_must_be_of_the_candidate(self) -> None:
        stale = results(demo_calibration=100, demo_holdout_1=200, demo_holdout_2=0)
        found = findings(registry(FROZEN), stale, commit="b" * 40)
        self.assertTrue(any("results are of commit aaaaaaaaaaaa, not of the candidate "
                            "bbbbbbbbbbbb" in item for item in found), found)
        for commit in (None, "HEAD", "A" * 40):
            with self.subTest(commit=commit):
                bad = {**stale, "commit": commit}
                self.assertTrue(any("commit is not a full lower-case commit SHA" in item
                                    for item in findings(registry(FROZEN), bad)))
        # Outside a Git checkout, with no --commit, the candidate is unknown.
        self.assertTrue(any("the candidate commit is unknown" in item
                            for item in findings(registry(FROZEN), stale, commit=None)))

    def test_a_malformed_results_file_fails(self) -> None:
        self.assertTrue(any("envelope is not schema_version 1, kind 'test-results'" in item
                            for item in findings(registry(FROZEN), {"kind": "other"})))
        bad = results(demo_calibration=100)
        bad["results"].append({"case": "demo-calibration"})
        self.assertTrue(any("results[1] is not case, variant, quantity, scope" in item
                            for item in findings(registry(FROZEN), bad)))
        for field, value in (("variant", ["lf"]), ("case", {"id": 1}), ("quantity", 7)):
            with self.subTest(field=field):
                bad = results(demo_calibration=100)
                bad["results"][0][field] = value
                self.assertTrue(any("results[0] case and quantity are not text, or its "
                                    "variant is not text or null" in item
                                    for item in findings(registry(FROZEN), bad)))

    def test_main_reports_the_measurement(self) -> None:
        with repository(registry(FROZEN)) as root, \
                contextlib.redirect_stdout(io.StringIO()) as output:
            path = root / "results.json"
            path.write_text(json.dumps(results(demo_calibration=1, demo_holdout_1=200,
                                               demo_holdout_2=0)), "utf-8")
            self.assertEqual(check_accuracy.main(["--root", str(root), "--results",
                                                  str(path), "--commit", CANDIDATE]), 1)
        self.assertIn("accuracy: FAIL (1 contract(s); measured against the results)",
                      output.getvalue())


@unittest.skipUnless(shutil.which("git"), "needs Git")
class WideningTest(unittest.TestCase):
    def tagged(self, before: dict[str, Any], now: dict[str, Any] | None,
               widening: bool = True) -> list[str]:
        """Commit and tag ``before`` as v1.0.0, then check ``now`` in the working tree."""
        with repository(registry(before)) as root:
            git(root, "init", "-q")
            git(root, "add", "-A")
            git(root, "commit", "-qm", "Release")
            git(root, "tag", "-a", "v1.0.0", "-m", "v1.0.0")
            contracts = [] if now is None else [now]
            (root / "tests/accuracy/contracts.json").write_text(
                json.dumps(registry(*contracts)), "utf-8")
            return check_accuracy.check(root, widening=widening)[0]

    def test_the_tag_comparison_can_be_skipped_for_a_copy_of_a_commit(self) -> None:
        # check_excel_run.py measures a copy of a commit's cases, with no history.
        widened = {**FROZEN, "threshold": 0.001}
        self.assertTrue(self.tagged(FROZEN, widened))
        self.assertEqual(self.tagged(FROZEN, widened, widening=False), [])
        self.assertEqual(self.tagged(FROZEN, {**FROZEN, "status": "draft", "decision": None},
                                     widening=False), [])

    def test_a_frozen_threshold_widens_only_with_a_new_decision(self) -> None:
        self.assertEqual(self.tagged(FROZEN, FROZEN), [])
        self.assertEqual(self.tagged(FROZEN, {**FROZEN, "threshold": 0.00001}), [])
        widened = self.tagged(FROZEN, {**FROZEN, "threshold": 0.001})
        self.assertTrue(any("ACC-02 threshold 0.001 is larger than 0.0001 at v1.0.0; a weaker "
                            "or different guarantee needs a new decision" in item
                            for item in widened), widened)
        self.assertEqual(self.tagged(FROZEN, {**FROZEN, "threshold": 0.001,
                                              "decision": "#2"}), [])

    def test_any_other_change_to_a_frozen_guarantee_needs_a_new_decision(self) -> None:
        changes: dict[str, dict[str, Any]] = {
            "metric is 'absolute', not 'relative'": {"metric": "absolute"},
            "regime is 'CRR2', not 'CRR3'": {"regime": "CRR2"},
            "drops holdout cases demo-holdout-2": {"holdout_cases": ["demo-holdout-1"]},
            "drops calibration cases demo-calibration": {"calibration_cases": []},
        }
        for text, change in changes.items():
            with self.subTest(text=text):
                found = self.tagged(FROZEN, {**FROZEN, **change})
                self.assertTrue(any(f"ACC-02 {text} at v1.0.0" in item for item in found),
                                found)
                self.assertFalse(any("ACC-02" in item and "v1.0.0" in item for item in
                                     self.tagged(FROZEN, {**FROZEN, **change,
                                                          "decision": "#2"})))
        # Moving a holdout case to calibration drops it from the holdout.
        moved = {**FROZEN, "calibration_cases": ["demo-calibration", "demo-holdout-2"],
                 "holdout_cases": ["demo-holdout-1"]}
        self.assertTrue(any("drops holdout cases demo-holdout-2" in item
                            for item in self.tagged(FROZEN, moved)))
        # Adding a case only strengthens the guarantee.
        self.assertEqual(self.tagged({**FROZEN, "holdout_cases": ["demo-holdout-1"]}, FROZEN),
                         [])

    def test_a_frozen_contract_is_neither_removed_nor_unfrozen(self) -> None:
        for now in (None, {**FROZEN, "status": "draft", "decision": None}):
            with self.subTest(now=now):
                found = self.tagged(FROZEN, now)
                self.assertTrue(any("ACC-02 was frozen at v1.0.0" in item for item in found),
                                found)
        draft = {**FROZEN, "status": "draft", "decision": None}
        self.assertEqual(self.tagged(draft, None), [])

    def test_deleting_the_registry_does_not_drop_a_released_frozen_contract(self) -> None:
        for before, dropped in ((FROZEN, True), ({**FROZEN, "status": "draft",
                                                  "decision": None}, False)):
            with self.subTest(status=before["status"]), \
                    repository(registry(before)) as root:
                git(root, "init", "-q")
                git(root, "add", "-A")
                git(root, "commit", "-qm", "Release")
                git(root, "tag", "-a", "v1.0.0", "-m", "v1.0.0")
                (root / "tests/accuracy/contracts.json").unlink()
                found = check_accuracy.check(root)[0]
                self.assertEqual(any("ACC-02 was frozen at v1.0.0" in item for item in found),
                                 dropped, found)


if __name__ == "__main__":
    unittest.main()
