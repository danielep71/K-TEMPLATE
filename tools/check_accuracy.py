#!/usr/bin/env python3
"""Check the accuracy contracts of the test cases, and measure them against an Excel run.

An accuracy contract (docs/methodology/TEST_CASES.md, Accuracy contracts) is the
worst error a calculation guarantees for one regime, over the outputs of its
quantities in two lists of cases: the calibration cases, which may be used to
set the threshold, and the holdout cases, which never are. The registry is
tests/accuracy/contracts.json.

Without a results file, the check reads the registry against the cases:
- every contract is well formed;
- its calculation is declared by the calculation contracts;
- each of its cases tests that calculation and has an expected number for each
  of its quantities, and no case is both calibration and holdout;
- a frozen contract names the decision that froze it, is not illustrative, and
  has at least min_holdout_cases holdout cases, every expected value of which
  is published or independent;
- a contract frozen at the newest vX.Y.Z tag that HEAD contains is still
  frozen, and its threshold is no larger unless it names another decision.

With --results, a results file of an Excel run, it also fails:
- results of another revision: their commit is the full SHA of the candidate,
  --commit when given, HEAD otherwise;
- an output with no result;
- a result outside its expected output's own tolerance;
- a contract whose worst error, among its calibration or among its holdout
  outputs, exceeds its threshold.

The error is |a - e| for an absolute contract and |a - e| / |e| for a relative
one; a relative error at e = 0 is 0 when a = 0 and infinite otherwise. A
contract proves only what its cases and their references prove: an
illustrative contract validates nothing.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import subprocess
import sys
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

from check_test_cases import (  # noqa: E402
    LOAD_FAILED,
    Report,
    load_json,
    read_methodology,
)

ROOT = Path(__file__).resolve().parents[1]
REGISTRY = Path("tests/accuracy/contracts.json")
FIXTURES = Path("tests/fixtures")
EXPECTED = Path("tests/expected")
MANIFEST = "case.json"
REGISTRY_KEYS = frozenset({"schema_version", "kind", "min_holdout_cases", "contracts"})
CONTRACT_KEYS = frozenset({"id", "calculation", "regime", "quantities", "metric", "threshold",
                           "provenance", "status", "decision", "calibration_cases",
                           "holdout_cases"})
RESULTS_KEYS = frozenset({"schema_version", "kind", "commit", "results"})
RESULT_KEYS = frozenset({"case", "variant", "quantity", "scope", "value"})
CONTRACT_ID = re.compile(r"^ACC-\d+$")
SNAKE = re.compile(r"^[a-z][a-z0-9]*(?:_[a-z0-9]+)*$")
INDEPENDENT = frozenset({"published", "independent"})
METRICS = frozenset({"absolute", "relative"})
STATUSES = frozenset({"draft", "frozen"})
TAG = re.compile(r"^v\d+\.\d+\.\d+$")
SHA = re.compile(r"^[0-9a-f]{40}$")

# A result's identity: case, variant (None for a JSON fixture), quantity and scope.
Key = tuple[str, str | None, str, str | None]


@dataclass(frozen=True)
class Output:
    """One expected number of a case."""

    key: Key
    value: float
    absolute: float
    relative: float
    reference: str


@dataclass
class Contract:
    id: str
    calculation: str
    quantities: list[str]
    metric: str
    threshold: float
    status: str
    decision: str | None
    calibration: list[str]
    holdout: list[str]


@dataclass
class Registry:
    contracts: list[Contract] = field(default_factory=list)
    min_holdout: int = 1


def scope_text(scope: object) -> str | None:
    if isinstance(scope, dict):
        return f"{scope.get('dataset')}/{scope.get('record')}"
    return None


def label(key: Key) -> str:
    case, variant, quantity, scope = key
    where = case if variant is None else f"{case}[{variant}]"
    return f"{where} {quantity}" + (f" ({scope})" if scope else "")


def is_number(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and \
        math.isfinite(value)


def case_outputs(root: Path, case: str, report: Report) -> tuple[str | None, list[Output]]:
    """A case's calculation and its expected numbers; (None, []) when it has no fixture."""
    fixture = root / FIXTURES / f"{case}.json"
    manifest = root / FIXTURES / case / MANIFEST
    path = fixture if fixture.is_file() else manifest
    data = load_json(path, root, report) if path.is_file() else LOAD_FAILED
    expected = load_json(root / EXPECTED / f"{case}.json", root, report) \
        if (root / EXPECTED / f"{case}.json").is_file() else LOAD_FAILED
    if not isinstance(data, dict) or not isinstance(expected, dict):
        return None, []
    calculation = data.get("calculation")
    groups: dict[str | None, Any] = {None: expected.get("outputs")} if "outputs" in expected \
        else dict(expected.get("variants") or {})
    outputs = []
    for variant, items in groups.items():
        for item in items if isinstance(items, list) else []:
            tolerance = item.get("tolerance") if isinstance(item, dict) else None
            if not isinstance(item, dict) or not is_number(item.get("value")) or \
                    not isinstance(tolerance, dict):
                continue
            key = (case, variant, str(item.get("quantity")), scope_text(item.get("scope")))
            outputs.append(Output(key, float(item["value"]), float(tolerance["absolute"]),
                                  float(tolerance["relative"]),
                                  str((item.get("reference") or {}).get("class"))))
    return calculation if isinstance(calculation, str) else None, outputs


def contract_problems(data: Mapping[str, Any]) -> list[str]:
    """What is wrong with one contract's own fields."""
    problems = []
    if not isinstance(data.get("id"), str) or not CONTRACT_ID.match(data["id"]):
        problems.append("id is not ACC-<number>")
    if not isinstance(data.get("regime"), str) or not data["regime"].strip():
        problems.append("regime is not a text")
    quantities = data.get("quantities")
    if not isinstance(quantities, list) or not quantities or \
            not all(isinstance(item, str) and SNAKE.match(item) for item in quantities):
        problems.append("quantities is not a non-empty list of lower snake case names")
    elif len(set(quantities)) != len(quantities):
        problems.append("quantities lists a name twice")
    if data.get("metric") not in METRICS:
        problems.append("metric is not absolute or relative")
    if not is_number(data.get("threshold")) or data["threshold"] < 0:
        problems.append("threshold is not a finite number of at least 0")
    if not isinstance(data.get("provenance"), str) or not data["provenance"].strip():
        problems.append("provenance is not a text")
    if data.get("status") not in STATUSES:
        problems.append("status is not draft or frozen")
    decision = data.get("decision")
    if data.get("status") == "frozen" and (not isinstance(decision, str) or not decision.strip()):
        problems.append("a frozen contract names the decision that froze it")
    if data.get("status") == "draft" and decision is not None:
        problems.append("a draft contract has no decision; decision is null")
    for name in ("calibration_cases", "holdout_cases"):
        cases = data.get(name)
        if not isinstance(cases, list) or not all(isinstance(item, str) for item in cases):
            problems.append(f"{name} is not a list of case IDs")
        elif len(set(cases)) != len(cases):
            # Each case counts once, so a repeated ID would inflate the holdout.
            problems.append(f"{name} lists a case twice")
    return problems


def read_registry(root: Path, text: str | None, report: Report, where: str) -> Registry | None:
    """The registry in ``text`` (the working tree's when None); None when absent or unusable."""
    if text is None:
        path = root / REGISTRY
        if not path.is_file():
            return None
        data = load_json(path, root, report)
    else:
        try:
            data = json.loads(text)
        except ValueError:
            return None
    if data is LOAD_FAILED:
        return None
    if not isinstance(data, dict) or data.keys() != REGISTRY_KEYS or \
            data["schema_version"] != 1 or data["kind"] != "accuracy-contracts":
        report.add(where, "envelope is not schema_version 1, kind 'accuracy-contracts', "
                          "min_holdout_cases and contracts")
        return None
    minimum = data["min_holdout_cases"]
    if not isinstance(minimum, int) or isinstance(minimum, bool) or minimum < 1:
        report.add(where, "min_holdout_cases is not a whole number of at least 1")
        return None
    if not isinstance(data["contracts"], list):
        report.add(where, "contracts is not a list")
        return None
    registry = Registry(min_holdout=minimum)
    seen: set[str] = set()
    for index, item in enumerate(data["contracts"]):
        name = item.get("id") if isinstance(item, dict) else None
        at = f"{where}: contracts[{index}]" + (f" {name}" if isinstance(name, str) else "")
        if not isinstance(item, dict) or item.keys() != CONTRACT_KEYS:
            report.add(at, f"keys are not {', '.join(sorted(CONTRACT_KEYS))}")
            continue
        problems = contract_problems(item)
        if item["id"] in seen:
            problems.append("id appears twice")
        seen.add(str(item["id"]))
        for problem in problems:
            report.add(at, problem)
        if problems:
            continue
        registry.contracts.append(Contract(
            item["id"], str(item["calculation"]), list(item["quantities"]), item["metric"],
            float(item["threshold"]), item["status"], item["decision"],
            list(item["calibration_cases"]), list(item["holdout_cases"])))
    return registry


def contract_outputs(root: Path, contract: Contract, calculations: Mapping[str, bool],
                     report: Report, where: str) -> dict[str, list[Output]]:
    """Each split's expected outputs of the contract, after checking its cases."""
    if contract.calculation not in calculations:
        report.add(where, f"calculation {contract.calculation} is not declared by the "
                          "calculation contracts")
    if calculations.get(contract.calculation) and contract.status == "frozen":
        report.add(where, f"calculation {contract.calculation} is illustrative, so the "
                          "contract stays draft")
    both = set(contract.calibration) & set(contract.holdout)
    for case in sorted(both):
        report.add(where, f"case {case} is both calibration and holdout")
    if not contract.calibration and not contract.holdout:
        report.add(where, "lists no case")
    splits: dict[str, list[Output]] = {}
    for split, cases in (("calibration", contract.calibration), ("holdout", contract.holdout)):
        splits[split] = []
        for case in cases:
            calculation, outputs = case_outputs(root, case, report)
            if calculation is None:
                report.add(where, f"case {case} has no fixture and expected file")
                continue
            if calculation != contract.calculation:
                report.add(where, f"case {case} tests {calculation}, not "
                                  f"{contract.calculation}")
            mine = [output for output in outputs if output.key[2] in contract.quantities]
            for quantity in contract.quantities:
                if not any(output.key[2] == quantity for output in mine):
                    report.add(where, f"case {case} has no expected number for {quantity}")
            splits[split] += mine
    return splits


def frozen_problems(contract: Contract, holdout: list[Output], minimum: int) -> list[str]:
    if contract.status != "frozen":
        return []
    problems = []
    if len(contract.holdout) < minimum:
        problems.append(f"a frozen contract needs at least {minimum} holdout case(s), has "
                        f"{len(contract.holdout)}")
    weak = sorted({label(output.key) for output in holdout
                   if output.reference not in INDEPENDENT})
    problems += [f"holdout output {item} is not published or independent" for item in weak]
    return problems


def git(root: Path, *args: str) -> str | None:
    try:
        completed = subprocess.run(["git", "-C", str(root), *args], capture_output=True,
                                   text=True, encoding="utf-8", check=False)
    except OSError:
        return None
    return completed.stdout if completed.returncode == 0 else None


def newest_tag(root: Path) -> str | None:
    """The newest vX.Y.Z tag that HEAD contains, by version."""
    tags = [tag for tag in (git(root, "tag", "--merged", "HEAD") or "").split()
            if TAG.match(tag)]
    return max(tags, key=lambda tag: tuple(int(part) for part in tag[1:].split(".")),
               default=None)


def widening_problems(root: Path, registry: Registry, report: Report) -> list[str]:
    """A contract frozen at the newest tag stays frozen, and is widened only by a new decision."""
    tag = newest_tag(root)
    if tag is None:
        return []
    text = git(root, "show", f"{tag}:{REGISTRY.as_posix()}")
    tagged = read_registry(root, text, Report(), f"{REGISTRY.as_posix()} at {tag}") \
        if text is not None else None
    if tagged is None:
        return []
    current = {contract.id: contract for contract in registry.contracts}
    problems = []
    for before in tagged.contracts:
        if before.status != "frozen":
            continue
        now = current.get(before.id)
        if now is None or now.status != "frozen":
            problems.append(f"{before.id} was frozen at {tag}; a frozen contract is neither "
                            "removed nor returned to draft")
        elif now.threshold > before.threshold and now.decision == before.decision:
            problems.append(f"{before.id} threshold {now.threshold!r} is larger than "
                            f"{before.threshold!r} at {tag}; widening it needs a new decision")
    return problems


def error(metric: str, actual: float, expected: float) -> float:
    if metric == "absolute":
        return abs(actual - expected)
    if expected == 0:
        return 0.0 if actual == 0 else math.inf
    return abs(actual - expected) / abs(expected)


def read_results(path: Path, root: Path, candidate: str | None,
                 report: Report) -> dict[Key, float] | None:
    """The results of the candidate's run; None when the file is unusable or of another commit."""
    where = str(path)
    data = load_json(path, root, report)
    if data is LOAD_FAILED:
        return None
    if not isinstance(data, dict) or data.keys() != RESULTS_KEYS or \
            data["schema_version"] != 1 or data["kind"] != "test-results" or \
            not isinstance(data["results"], list):
        report.add(where, "envelope is not schema_version 1, kind 'test-results', commit "
                          "and a results list")
        return None
    commit = data["commit"]
    if not isinstance(commit, str) or not SHA.match(commit):
        report.add(where, "commit is not a full lower-case commit SHA")
        return None
    if candidate is None:
        report.add(where, "the candidate commit is unknown: pass --commit, or run in a Git "
                          "checkout")
        return None
    if commit != candidate:
        report.add(where, f"results are of commit {commit[:12]}, not of the candidate "
                          f"{candidate[:12]}")
        return None
    results: dict[Key, float] = {}
    for index, item in enumerate(data["results"]):
        if not isinstance(item, dict) or item.keys() != RESULT_KEYS or \
                not is_number(item["value"]):
            report.add(where, f"results[{index}] is not case, variant, quantity, scope and a "
                              "finite value")
            continue
        key = (str(item["case"]), item["variant"], str(item["quantity"]),
               scope_text(item["scope"]))
        if key in results:
            report.add(where, f"results[{index}] repeats {label(key)}")
        results[key] = float(item["value"])
    return results


def measure(contract: Contract, splits: Mapping[str, list[Output]], results: Mapping[Key, float],
            report: Report, where: str) -> list[str]:
    """Each output against its own tolerance and the contract; one summary line per split."""
    lines = []
    for split, outputs in splits.items():
        worst = 0.0
        for output in outputs:
            actual = results.get(output.key)
            if actual is None:
                report.add(where, f"no result for {label(output.key)}")
                continue
            bound = max(output.absolute, output.relative * abs(output.value))
            if abs(actual - output.value) > bound:
                report.add(where, f"{label(output.key)} is {actual!r}, expected "
                                  f"{output.value!r} within {bound!r}")
            worst = max(worst, error(contract.metric, actual, output.value))
        if outputs and worst > contract.threshold:
            report.add(where, f"worst {contract.metric} error {worst!r} over {split} cases "
                              f"exceeds the threshold {contract.threshold!r}")
        if outputs:
            lines.append(f"{contract.id} {split}: worst {contract.metric} error {worst!r}, "
                         f"threshold {contract.threshold!r}")
    return lines


def check(root: Path, results_path: Path | None = None,
          commit: str | None = None) -> tuple[list[str], int, list[str]]:
    """The findings, the number of contracts and the measurement lines.

    ``commit`` is the candidate the results must be of; HEAD when None."""
    report = Report()
    where = REGISTRY.as_posix()
    present = (root / REGISTRY).is_file()
    registry = read_registry(root, None, report, where)
    if registry is None:
        if not present:
            # A deleted registry still may not drop a released frozen contract.
            for problem in widening_problems(root, Registry(), report):
                report.add(where, problem)
        if results_path is not None and not present:
            report.add(where, "is missing, so there is nothing to measure")
        return report.findings, 0, []
    calculations = read_methodology(root, Report()).calculations
    candidate = commit or (git(root, "rev-parse", "HEAD") or "").strip() or None
    results = read_results(results_path, root, candidate, report) \
        if results_path is not None else None
    lines: list[str] = []
    for contract in registry.contracts:
        at = f"{where}: {contract.id}"
        splits = contract_outputs(root, contract, calculations, report, at)
        for problem in frozen_problems(contract, splits["holdout"], registry.min_holdout):
            report.add(at, problem)
        if results is not None:
            lines += measure(contract, splits, results, report, at)
    for problem in widening_problems(root, registry, report):
        report.add(where, problem)
    return report.findings, len(registry.contracts), lines


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT,
                        help="repository root (default: this checkout)")
    parser.add_argument("--results", type=Path,
                        help="results file of an Excel run, to measure the contracts")
    parser.add_argument("--commit",
                        help="full SHA of the candidate the results must be of (default: HEAD)")
    args = parser.parse_args(argv)
    findings, contracts, lines = check(args.root.resolve(), args.results, args.commit)
    for line in lines:
        print(line)
    for finding in findings:
        print(f"FAIL: {finding}")
    status = "FAIL" if findings else "PASS"
    measured = "measured against the results" if args.results else \
        "registry only, nothing measured without --results"
    print(f"accuracy: {status} ({contracts} contract(s); {measured})")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
