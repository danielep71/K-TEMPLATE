"""Positive and negative evidence fixtures for check_excel_run.py.

The checked commit is made in a temporary Git repository from the working
tree's tests/modules/, so the tests follow the suite as it is edited. Each
fixture is an evidence folder built in a temporary directory: that commit's
VBA components as the VBE exports them (CRLF), the three logs and the
certification log the harness would write for its example suite, the log and
certification log of each project suite the commit's policy lists, the results
file of its normal run, and host.txt. The commit also holds the working
tree's test cases and accuracy contracts, if any, which the results are
measured against. The logs and results are synthetic: they follow
docs/methodology/TEST_HARNESS.md and TEST_Examples.bas, not a recorded Excel
run. The tests of the measurement itself, and of a project or relation suite,
use synthetic contracts, cases and modules written into their own commit, so
they hold whatever contracts, cases and suites this repository has.
"""

from __future__ import annotations

import contextlib
import dataclasses
import io
import json
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from collections.abc import Callable, Iterator
from pathlib import Path
from typing import Any, ClassVar

SOURCE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SOURCE / "tools"))

sys.path.insert(0, str(SOURCE / "tools" / "devbuild"))

import build_workbook  # noqa: E402
import check_accuracy  # noqa: E402
import check_excel_run  # noqa: E402
import export_manifest  # noqa: E402
import vba_project  # noqa: E402
from check_excel_run import Suite, components, read_suite  # noqa: E402

# The policy of the working tree; make_repository commits it with the modules.
POLICY = check_excel_run.working_policy(SOURCE)

# The VBA test modules, with the source of each project suite of the policy.
MODULES = sorted({*(SOURCE / "tests/modules").glob("*.bas"),
                  *(path for path in (SOURCE / item.source for item in POLICY.project_suites)
                    if path.is_file())})
# A synthetic form and its binary resources, which the VBE exports as a pair.
FORM = {"src/forms/frmFixture.frm": b'VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} '
                                    b'frmFixture\n   OleObjectBlob   =   "frmFixture.frx":0000\n'
                                    b'End\nAttribute VB_Name = "frmFixture"\nOption Explicit\n',
        "src/forms/frmFixture.frx": b"\x00\x01\r\n\x02\n\xff binary"}


def make_repository(policy: bytes | None = None, *, with_policy: bool = True,
                    with_cases: bool = True, with_form: bool = True,
                    extra: dict[str, bytes] | None = None) -> tuple[Path, str]:
    """A temporary repository whose one commit holds the working tree's modules and
    evidence policy, or ``policy`` in its place, or no policy at all, and with
    ``with_cases`` the working tree's test cases and accuracy contracts. The files
    of ``extra`` are written last, so they may replace any of these."""
    root = Path(tempfile.mkdtemp(prefix="excel-run-repository-"))
    if with_policy:
        target = root / check_excel_run.POLICY
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(policy if policy is not None
                           else (SOURCE / check_excel_run.POLICY).read_bytes())
    for module in MODULES:
        target = root / module.relative_to(SOURCE)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(module.read_bytes().replace(b"\r\n", b"\n"))
    config = root / check_excel_run.BUILD_CONFIG
    config.parent.mkdir(parents=True, exist_ok=True)
    config.write_bytes((SOURCE / check_excel_run.BUILD_CONFIG).read_bytes())
    if with_cases:
        tracked = subprocess.run(["git", "-C", str(SOURCE), "ls-files", "-z", "--",
                                  *check_excel_run.CASE_PATHS],
                                 capture_output=True, check=True).stdout.decode().split("\0")
        for name in filter(None, tracked):
            target = root / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes((SOURCE / name).read_bytes())
    for name, data in [*(FORM.items() if with_form else ()), *(extra or {}).items()]:
        (root / name).parent.mkdir(parents=True, exist_ok=True)
        (root / name).write_bytes(data)
    for command in (["init", "-q"], ["add", "-A"],
                    ["-c", "user.name=Fixture", "-c", "user.email=fixture@example.invalid",
                     "-c", "commit.gpgsign=false",
                     "commit", "-qm", "Fixture"]):
        subprocess.run(["git", "-C", str(root), *command], capture_output=True, check=True)
    commit = subprocess.run(["git", "-C", str(root), "rev-parse", "HEAD"], capture_output=True,
                            check=True, text=True).stdout.strip()
    return root, commit


ROOT, COMMIT = make_repository() if MODULES and shutil.which("git") else (SOURCE, "")
ENVIRONMENT = ("ENVIRONMENT=host=Microsoft Excel; version=16.0; build=17928; "
               "os=Windows (64-bit) NT 10.00; office=64-bit; runtime=VBA7")
HOST = ("operator=Fixture Operator\n"
        "started=2026-10-24T16:59:30+02:00\n"
        "finished=2026-10-24T17:03:10+02:00\n"
        "windows=Windows 11 Pro 24H2, OS build 26100.4061\n"
        "excel=16.0 build 17928\n"
        "bitness=64-bit\n"
        "automation_security=ByUI\n"
        "vba_project_access=NOT_TRUSTED\n"
        "references=Visual Basic For Applications; Microsoft Excel 16.0 Object Library; "
        "OLE Automation; Microsoft Office 16.0 Object Library\n"
        "compile=PASS\n"
        "outcome=PASS\n")
UNCHANGED = "CLEANUP=PASS; detail=Excel state unchanged (Calculation, DisplayAlerts, " \
            "EnableEvents, ScreenUpdating, Workbooks, ActiveWindow)"
FAILED = check_excel_run.test_failed_number(components(ROOT, COMMIT)[Path(POLICY.harness).name]) \
    if COMMIT else None


def tearDownModule() -> None:
    if ROOT != SOURCE:
        shutil.rmtree(ROOT, ignore_errors=True)


def suite() -> Suite:
    return read_suite(components(ROOT, COMMIT)[Path(POLICY.suite_source).name])


def run_log(mode: str, examples: Suite | None = None, suite_name: str = "examples") -> list[str]:
    """The records the harness writes for one run of a suite, by default the example
    suite."""
    examples = examples or suite()
    skipped = set(examples.last_case) if mode == "SKIPPED_CASE" else set()
    lines = [f"RUN={suite_name}; mode={mode}; expected_cases={examples.expected_cases}; "
             f"expected_assertions={examples.expected_assertions}; started=2026-10-24 17:00:00",
             ENVIRONMENT]
    failures: list[str] = []
    cases = assertions = 0
    for case in examples.cases:
        if case in skipped:
            continue
        lines.append(f"CASE={case}")
        cases += 1
        for name in (item for item in examples.assertions if item.startswith(f"{case}/")):
            assertions += 1
            if mode == "INJECTED_FAILURE" and name in examples.injected:
                detail = f"expected={examples.injected_expected}; actual=0.75"
                lines.append(f"ASSERT={name}; status=FAIL; {detail}")
                lines.append(f"FAILURE={name}; {detail}")
                failures.append(f"{name}; {detail}")
            else:
                lines.append(f"ASSERT={name}; status=PASS")
    if mode == "SKIPPED_CASE":
        detail = (f"expected cases={examples.expected_cases}, "
                  f"assertions={examples.expected_assertions}; observed cases={cases}, "
                  f"assertions={assertions}")
        lines.append(f"FAILURE=harness.completeness; {detail}")
        failures.append(f"harness.completeness; {detail}")
    complete = "INCOMPLETE" if skipped else "COMPLETE"
    operation = "FAIL" if mode == "INJECTED_FAILURE" else "PASS"
    verdict = "PASS" if mode == "NORMAL" else "FAIL"
    lines += [f"CASES={cases}; expected={examples.expected_cases}",
              f"ASSERTIONS={assertions}; expected={examples.expected_assertions}",
              f"FAILURES={len(failures)}; errors=0"]
    lines += [f"PRIMARY_FAILURE={failures[0]}"] if failures else []
    lines += [UNCHANGED, "RECOVERY=NONE",
              f"RESULT={verdict}; suite={suite_name}; operation={operation}; completeness={complete}; "
              f"cases={cases}; assertions={assertions}; failures={len(failures)}; "
              "cleanup=PASS; recovery=NONE"]
    return lines


def certification_log(name: str = "examples",
                      modes: tuple[str, ...] = tuple(POLICY.logs)) -> list[str]:
    """The records TEST_Harness writes for a suite's certification entry point, by
    default TEST_Examples.CertifyExamples."""
    lines = [f"CERTIFY={name}; started=2026-10-24 17:00:00"]
    for mode in modes:
        normal = mode == "NORMAL"
        complete = "INCOMPLETE" if mode == "SKIPPED_CASE" else "COMPLETE"
        lines.append(f"MODE={mode}; expected={'PASS' if normal else 'FAIL'}; "
                     f"error={0 if normal else FAILED}; result={'PASS' if normal else 'FAIL'}; "
                     f"completeness={complete}; status=PASS")
    lines.append(f"CERTIFICATION=PASS; suite={name}; completeness=COMPLETE; "
                 f"modes={len(modes)}; failures=0")
    return lines


def results_document(commit: str = COMMIT, root: Path = ROOT) -> dict[str, Any]:
    """The results file a correct normal run writes for the commit checked out in
    ``root``: every output of its contracts' quantities in their cases, equal to its
    expected value, and no result when it has no accuracy registry."""
    path = root / check_accuracy.REGISTRY
    registry = json.loads(path.read_text(encoding="utf-8")) if path.is_file() \
        else {"contracts": []}
    results = []
    for contract in registry["contracts"]:
        for case in [*contract["calibration_cases"], *contract["holdout_cases"]]:
            expected = json.loads((root / "tests/expected" / f"{case}.json")
                                  .read_text(encoding="utf-8"))
            groups = [(None, expected["outputs"])] if "outputs" in expected \
                else sorted(expected["variants"].items())
            results += [{"case": case, "variant": variant, "quantity": output["quantity"],
                         "scope": output["scope"], "value": output["value"]}
                        for variant, outputs in groups for output in outputs
                        if output["quantity"] in contract["quantities"]]
    return {"schema_version": 1, "kind": "test-results", "commit": commit, "results": results}


def write_results(folder: Path, document: dict[str, Any], name: str = POLICY.results) -> None:
    (folder / name).write_text(json.dumps(document, indent=2) + "\n", encoding="utf-8")


def write_log(folder: Path, name: str, lines: list[str]) -> None:
    (folder / name).write_bytes(("\r\n".join(lines) + "\r\n").encode())


@contextlib.contextmanager
def evidence(change: Callable[[Path], None] | None = None, root: Path = ROOT,
             commit: str = COMMIT) -> Iterator[Path]:
    """A complete, valid evidence folder for the suites of ``commit``, then ``change``
    applied to it. The example suite's results file holds every result; the results
    file of a project suite that names one holds none, since the files are measured
    together."""
    policy = check_excel_run.policy_at(root, commit)
    sources = components(root, commit)
    with tempfile.TemporaryDirectory(prefix="excel-run-") as temporary:
        folder = Path(temporary)
        (folder / "exported").mkdir()
        for name, data in sources.items():
            exported = data if name.endswith(".frx") else data.replace(b"\n", b"\r\n")
            (folder / "exported" / name).write_bytes(exported)
        for mode, filename in POLICY.logs.items():
            write_log(folder, filename, run_log(mode))
        write_log(folder, POLICY.certification, certification_log())
        for item in policy.project_suites:
            project = read_suite(sources[Path(item.source).name], policy.office, item.runner)
            write_log(folder, item.log, run_log("NORMAL", project, item.name))
            write_log(folder, item.certification, certification_log(item.name, ("NORMAL",)))
            if item.results is not None:
                write_results(folder, dict(results_document(commit, root), results=[]),
                              item.results)
        (folder / "host.txt").write_text(HOST, encoding="utf-8")
        write_results(folder, results_document(commit, root))
        if change:
            change(folder)
        yield folder


def edit(path: str, old: str, new: str) -> Callable[[Path], None]:
    def apply(folder: Path) -> None:
        text = (folder / path).read_bytes().decode("ascii")
        assert old in text, (path, old)
        (folder / path).write_bytes(text.replace(old, new, 1).encode("ascii"))
    return apply


def findings(change: Callable[[Path], None] | None = None) -> list[str]:
    with evidence(change) as folder:
        return check_excel_run.check(ROOT, COMMIT, folder).findings


def workbook_bytes(commit: str, clean: bool = True) -> bytes:
    """A workbook as tools/devbuild/build_workbook.py packages it, stamped with ``commit``."""
    layout = build_workbook.load_layout(SOURCE)
    build = build_workbook.Build(commit, "2026-11-02T10:00:00+01:00", clean, "development", ())
    return build_workbook.package(layout, build, b"not a VBA project")


@unittest.skipUnless(COMMIT, "needs Git and the modules of tests/modules/")
class RebuildTest(unittest.TestCase):
    """Identity without an export: the workbook equals a rebuild of the commit."""

    # The build stops on a form, so this repository has none.
    root: ClassVar[Path]
    commit: ClassVar[str]
    temporary: ClassVar[tempfile.TemporaryDirectory[str]]
    workbook: ClassVar[Path]

    @classmethod
    def setUpClass(cls) -> None:
        cls.root, cls.commit = make_repository(with_form=False)
        cls.temporary = tempfile.TemporaryDirectory(prefix="excel-run-built-")
        cls.workbook = Path(cls.temporary.name) / "Workbook-development.xlsm"
        built = build_workbook.build(cls.root, "development", cls.workbook)
        assert built.clean and built.commit == cls.commit, built

    @classmethod
    def tearDownClass(cls) -> None:
        cls.temporary.cleanup()
        shutil.rmtree(cls.root, ignore_errors=True)

    def automated(self, folder: Path) -> None:
        """What tools/devbuild/excel_run.py leaves: no export, compile=ON_DEMAND."""
        shutil.rmtree(folder / POLICY.exported)
        write_results(folder, results_document(self.commit, self.root))
        (folder / POLICY.host).write_text(HOST.replace("compile=PASS", "compile=ON_DEMAND"),
                                          encoding="utf-8")

    def checked(self, workbook: Path, change: Callable[[Path], object] | None = None
                ) -> check_excel_run.Evidence:
        def apply(folder: Path) -> None:
            self.automated(folder)
            if change:
                change(folder)
        with evidence(apply) as folder:
            return check_excel_run.check(self.root, self.commit, folder, workbook,
                                         require_binding=True)

    def test_a_workbook_equal_to_a_rebuild_is_verified_without_an_export(self) -> None:
        result = self.checked(self.workbook)
        self.assertEqual(result.findings, [])
        self.assertEqual(result.binding, "VERIFIED")
        self.assertIn("byte-identical", " ".join(result.binding_notes))
        self.assertEqual(set(result.hashes), set(components(self.root, self.commit)))
        self.assertIn("Binding: VERIFIED", check_excel_run.report(self.commit, result))

    def repacked(self, name: str, change: Callable[[str, bytes], bytes] = lambda _, data: data,
                 **options: Any) -> Path:
        """The workbook's parts written again into a new package with ``options``."""
        target = Path(self.temporary.name) / name
        with zipfile.ZipFile(self.workbook) as source, \
                zipfile.ZipFile(target, "w", **options) as package:
            for info in source.infolist():
                package.writestr(info.filename, change(info.filename, source.read(info)))
        return target

    def test_the_parts_not_the_compression_decide(self) -> None:
        # Another platform or zlib packs the same parts into other bytes.
        other = self.repacked("other-compression.xlsm", compression=zipfile.ZIP_DEFLATED,
                              compresslevel=1)
        self.assertNotEqual(other.read_bytes(), self.workbook.read_bytes())
        result = self.checked(other)
        self.assertEqual((result.findings, result.binding), ([], "VERIFIED"))

    def test_a_changed_part_fails_and_is_named(self) -> None:
        def project(name: str, data: bytes) -> bytes:
            return data + b"\0" if name == "xl/vbaProject.bin" else data
        result = self.checked(self.repacked("changed-project.xlsm", project,
                                            compression=zipfile.ZIP_DEFLATED))
        self.assertEqual(result.binding, "FAIL")
        self.assertTrue(any("xl/vbaProject.bin (SHA-256" in item and
                            "differs from the development build of" in item
                            for item in result.findings), result.findings)

    def test_a_package_that_is_not_a_zip_fails(self) -> None:
        broken = Path(self.temporary.name) / "broken.xlsm"
        broken.write_bytes(b"not a workbook")
        result = self.checked(broken)
        self.assertEqual(result.binding, "FAIL")
        self.assertTrue(any("not a readable package" in item for item in result.findings),
                        result.findings)

    def test_a_missing_workbook_fails(self) -> None:
        result = self.checked(Path(self.temporary.name) / "absent.xlsm")
        self.assertEqual(result.binding, "FAIL")
        self.assertTrue(any("missing" in item for item in result.findings), result.findings)

    def test_compile_on_demand_is_accepted_only_with_identity_by_rebuild(self) -> None:
        def on_demand(folder: Path) -> None:
            (folder / POLICY.host).write_text(
                HOST.replace("compile=PASS", "compile=ON_DEMAND"), encoding="utf-8")
        manual = findings(on_demand)
        self.assertTrue(any("compile=ON_DEMAND, expected PASS (ON_DEMAND only with identity "
                            "by rebuild)" in item for item in manual), manual)
        failed = self.checked(self.workbook, lambda folder: (folder / POLICY.host).write_text(
            HOST.replace("compile=PASS", "compile=FAIL"), encoding="utf-8"))
        self.assertTrue(any("compile=FAIL, expected PASS or ON_DEMAND" in item
                            for item in failed.findings), failed.findings)

    def test_a_workbook_built_with_overrides_never_certifies(self) -> None:
        # Even an override identical to the commit's module marks the build.
        folder = Path(self.temporary.name) / "overrides"
        folder.mkdir()
        harness = self.root / POLICY.harness
        (folder / harness.name).write_bytes(harness.read_bytes())
        workbook = Path(self.temporary.name) / "Workbook-overridden.xlsm"
        built = build_workbook.build(self.root, "development", workbook, override=folder)
        stamp = build_workbook.commit_stamp(built)

        def overridden_results(folder: Path) -> None:
            write_results(folder, results_document(stamp, self.root))
        result = self.checked(workbook, overridden_results)
        self.assertEqual(result.binding, "FAIL")
        expected = f"built with overrides, so not evidence for {self.commit}"
        self.assertIn(f"workbook {workbook.name}: {expected}", " ".join(result.findings))
        self.assertIn(f"{POLICY.results}: written by a workbook {expected}",
                      " ".join(result.findings))
        report = check_excel_run.report(self.commit, result)
        self.assertNotIn("Binding: VERIFIED", report)
        self.assertIn("excel-run: FAIL", report)

    def test_without_an_export_or_a_workbook_the_sources_are_missing(self) -> None:
        with evidence(self.automated) as folder:
            result = check_excel_run.check(self.root, self.commit, folder)
        self.assertTrue(any("not exported from the workbook" in item
                            for item in result.findings), result.findings)


@unittest.skipUnless(COMMIT, "needs Git and the modules of tests/modules/")
class BindingTest(unittest.TestCase):
    def bound(self, change: Callable[[Path, Path], None] | None = None, *, workbook: bool = True,
              manifest: bool = True, require: bool = False,
              data: bytes | None = None) -> check_excel_run.Evidence:
        with evidence() as folder:
            path = folder.parent / f"{folder.name}-Workbook.xlsm"
            path.write_bytes(workbook_bytes(COMMIT) if data is None else data)
            try:
                if manifest:
                    export_manifest.write(path, folder)
                if change:
                    change(folder, path)
                return check_excel_run.check(ROOT, COMMIT, folder, path if workbook else None,
                                             require)
            finally:
                path.unlink(missing_ok=True)

    def test_verified(self) -> None:
        result = self.bound()
        self.assertEqual((result.binding, result.findings), ("VERIFIED", []))

    def test_unverified_without_a_manifest_or_a_workbook(self) -> None:
        for options in ({"manifest": False}, {"workbook": False}):
            with self.subTest(**options):
                result = self.bound(**options)  # type: ignore[arg-type]
                self.assertEqual((result.binding, result.findings), ("UNVERIFIED", []))
                self.assertTrue(result.binding_notes)
        required = self.bound(manifest=False, require=True)
        self.assertTrue(any("--require-binding" in item for item in required.findings))

    def test_fail_on_any_mismatch(self) -> None:
        def changed_export(folder: Path, _: Path) -> None:
            (folder / "exported/TEST_Harness.bas").write_bytes(b"changed")
        def changed_workbook(_: Path, path: Path) -> None:
            path.write_bytes(workbook_bytes(COMMIT) + b" ")
        def extra_export(folder: Path, _: Path) -> None:
            (folder / "exported/Extra.bas").write_bytes(b"x")
        def bad_manifest(folder: Path, _: Path) -> None:
            (folder / POLICY.manifest).write_text('{"format": "other"}', "utf-8")
        cases = {"changed after the export": changed_export, "but export-manifest.json records":
                 changed_workbook, "is not recorded": extra_export,
                 "not a valid export manifest": bad_manifest}
        for text, change in cases.items():
            with self.subTest(text=text):
                result = self.bound(change)
                self.assertEqual(result.binding, "FAIL")
                self.assertTrue(any(text in item for item in result.findings), result.findings)

    def test_the_stamp_must_name_the_commit_of_a_clean_build(self) -> None:
        for data, text in ((workbook_bytes(COMMIT, clean=False), "NOT evidence-grade"),
                           (workbook_bytes("0" * 40), "0" * 40)):
            with self.subTest(text=text):
                result = self.bound(data=data)
                self.assertEqual(result.binding, "FAIL")
                self.assertTrue(any("build stamp" in item and text in item
                                    for item in result.findings), result.findings)

    def test_an_overridden_workbook_never_binds(self) -> None:
        override = build_workbook.Override("TEST_Harness", Path("mutants/TEST_Harness.bas"),
                                           "0" * 64)
        build = build_workbook.Build(COMMIT, "2026-11-02T10:00:00+01:00", True, "development",
                                     (), overrides=(override,), override_folder="mutants")
        data = build_workbook.package(build_workbook.load_layout(SOURCE), build, b"project")
        result = self.bound(data=data, require=True)
        self.assertEqual(result.binding, "FAIL")
        self.assertTrue(any(f"built with overrides, so not evidence for {COMMIT}" in item and
                            "TEST_Harness SHA-256 " + "0" * 64 in item
                            for item in result.findings), result.findings)
        self.assertIn("excel-run: FAIL", check_excel_run.report(COMMIT, result))

    def test_the_stamp_is_read_from_shared_strings_after_excel_saves(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            source = zipfile.ZipFile(io.BytesIO(workbook_bytes(COMMIT)))
            parts = {name: source.read(name) for name in source.namelist()}
            sheet = parts["xl/worksheets/sheet1.xml"].decode()
            inline = f'<c r="B3" t="inlineStr"><is><t xml:space="preserve">{COMMIT}</t></is></c>'
            self.assertIn(inline, sheet)
            parts["xl/worksheets/sheet1.xml"] = sheet.replace(
                inline, '<c r="B3" t="s"><v>0</v></c>').encode()
            parts["xl/sharedStrings.xml"] = (
                '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                f"<si><t>{COMMIT}</t></si></sst>").encode()
            path = Path(temporary) / "Saved.xlsm"
            with zipfile.ZipFile(path, "w") as package:
                for name, data in parts.items():
                    package.writestr(name, data)
            self.assertEqual(check_excel_run.read_stamp(path, "build_commit"), COMMIT)
            self.assertIsNone(check_excel_run.read_stamp(path, "no_such_name"))

    def test_report_states_the_binding(self) -> None:
        with evidence() as folder, contextlib.redirect_stdout(io.StringIO()) as output:
            check_excel_run.main([str(folder), "--commit", COMMIT, "--root", str(ROOT)])
        self.assertIn("Binding: UNVERIFIED", output.getvalue())


@unittest.skipUnless(COMMIT, "needs Git and the modules of tests/modules/")
class EvidenceTest(unittest.TestCase):
    def assertFinding(self, change: Callable[[Path], None], text: str) -> None:
        found = findings(change)
        self.assertTrue(any(text in item for item in found), found)

    def test_complete_evidence_passes(self) -> None:
        self.assertEqual(findings(), [])

    def test_report_names_the_commit_and_every_component(self) -> None:
        with evidence() as folder, contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(check_excel_run.main([str(folder), "--commit", COMMIT,
                                                   "--root", str(ROOT)]), 0)
        text = output.getvalue()
        self.assertIn(f"Candidate commit: {COMMIT}", text)
        for name in components(ROOT, COMMIT):
            self.assertIn(f"  {name} ", text)
        self.assertIn("excel-run: PASS", text)

    def test_unknown_commit_fails(self) -> None:
        with evidence() as folder, contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(check_excel_run.main([str(folder), "--commit", "0" * 40,
                                                   "--root", str(ROOT)]), 1)

    def test_wrong_source_identity(self) -> None:
        self.assertFinding(edit("exported/TEST_Harness.bas", "Option Explicit",
                                "Option Explicit\r\n'local edit"), "differs from TEST_Harness.bas")
        self.assertFinding(lambda folder: (folder / "exported/TEST_Examples.bas").unlink(),
                           "TEST_Examples.bas: not exported")
        def extra(folder: Path) -> None:
            (folder / "exported/Module1.bas").write_bytes(b"x")
        self.assertFinding(extra, "Module1.bas: not a VBA component of the commit")

    def test_form_resources_are_compared_byte_for_byte(self) -> None:
        self.assertEqual(components(ROOT, COMMIT)["frmFixture.frx"], FORM["src/forms/frmFixture.frx"])
        self.assertFinding(lambda folder: (folder / "exported/frmFixture.frx").unlink(),
                           "frmFixture.frx: not exported")
        def normalized(folder: Path) -> None:
            path = folder / "exported/frmFixture.frx"
            path.write_bytes(path.read_bytes().replace(b"\r\n", b"\n"))
        self.assertFinding(normalized, "frmFixture.frx: differs from frmFixture.frx")
        self.assertFinding(edit("exported/frmFixture.frm", "Option Explicit",
                                "Option Explicit\r\n'edit"), "frmFixture.frm: differs")

    def test_incomplete_logs(self) -> None:
        def truncate(folder: Path) -> None:
            lines = (folder / "examples.log").read_bytes().splitlines()
            (folder / "examples.log").write_bytes(b"\r\n".join(lines[:-1]))
        self.assertFinding(truncate, "examples.log: incomplete log")
        self.assertFinding(lambda folder: (folder / "skipped.log").unlink(), "skipped.log: missing")
        def empty(folder: Path) -> None:
            for filename in POLICY.logs.values():
                (folder / filename).write_bytes(b"")
        found = findings(empty)
        for filename in POLICY.logs.values():
            self.assertIn(f"{filename}: incomplete log: no record", found)
        self.assertFinding(edit("examples.log", "CASE=text-exact\r\n", ""),
                           "expected CASE=text-exact")
        self.assertFinding(edit("examples.log", "RESULT=PASS", "stray line\r\nRESULT=PASS"),
                           "not a KEY=value record")

    def test_a_run_that_did_not_pass(self) -> None:
        self.assertFinding(edit("examples.log", "status=PASS", "status=FAIL"),
                           "examples.log:4: ASSERT status=FAIL, expected PASS")
        self.assertFinding(edit("examples.log", "RESULT=PASS", "RESULT=FAIL"), "RESULT=FAIL")
        self.assertFinding(edit("examples.log", "mode=NORMAL", "mode=SKIPPED_CASE"),
                           "RUN mode=SKIPPED_CASE, expected NORMAL")
        assertions = f"ASSERTIONS={suite().expected_assertions}"
        self.assertFinding(edit("examples.log", assertions, f"ERROR=run; number=5\r\n{assertions}"),
                           f"ERROR=run, expected {assertions}")
        self.assertFinding(edit("examples.log", "RECOVERY=NONE\r\n",
                                "RECOVERY=NONE\r\nRUN=examples; mode=NORMAL\r\n"),
                           "RUN=examples, expected RESULT=PASS")
        def appended(folder: Path) -> None:
            log = (folder / "examples.log").read_bytes()
            (folder / "examples.log").write_bytes(log + log)
        self.assertFinding(appended, "after RESULT (more than one run?)")

    def test_failed_cleanup(self) -> None:
        def dirty(folder: Path) -> None:
            edit("examples.log", UNCHANGED, "CLEANUP=FAIL; detail=Excel state changed during "
                 "the run: Calculation")(folder)
            edit("examples.log", "RECOVERY=NONE", "RECOVERY=REQUIRED; reason=changed")(folder)
            edit("examples.log", "cleanup=PASS; recovery=NONE",
                 "cleanup=FAIL; recovery=REQUIRED")(folder)
        self.assertFinding(dirty, "CLEANUP=FAIL, expected CLEANUP=PASS")
        for filename in POLICY.logs.values():
            with self.subTest(log=filename):
                self.assertFinding(edit(filename, "RECOVERY=NONE\r\n",
                                        "RECOVERY=REQUIRED; reason=unsafe host state\r\n"),
                                   f"{filename}:")
                self.assertFinding(edit(filename, "RECOVERY=NONE\r\n", ""),
                                   "expected RECOVERY=NONE")
                self.assertFinding(edit(filename, "cleanup=PASS; recovery=NONE",
                                        "cleanup=PASS; recovery=REQUIRED"), "RESULT=")

    def test_failure_paths_must_end_as_documented(self) -> None:
        self.assertFinding(edit("injected.log", "failures=1; cleanup", "failures=2; cleanup"),
                           "injected.log")
        self.assertFinding(edit("skipped.log", "completeness=INCOMPLETE", "completeness=COMPLETE"),
                           "skipped.log")
        self.assertFinding(edit("skipped.log", "FAILURE=harness.completeness", "FAILURE=other"),
                           "FAILURE=other, expected FAILURE=harness.completeness")
        self.assertFinding(edit("skipped.log", "status=PASS", "status=FAIL; x"),
                           "skipped.log:4: ASSERT status=FAIL, expected PASS")

    def test_the_injected_failure_must_be_the_designated_assertion(self) -> None:
        designated = suite().injected[0]
        other = next(item for item in suite().assertions if item != designated)
        detail = f"expected={suite().injected_expected}; actual=0.75"
        def moved(folder: Path) -> None:
            edit("injected.log", f"ASSERT={designated}; status=FAIL; {detail}",
                 f"ASSERT={designated}; status=PASS")(folder)
            edit("injected.log", f"ASSERT={other}; status=PASS",
                 f"ASSERT={other}; status=FAIL; {detail}")(folder)
            edit("injected.log", f"FAILURE={designated}", f"FAILURE={other}")(folder)
        self.assertFinding(moved, f"FAILURE={other}, expected FAILURE={designated}")

    def test_the_injected_failure_values_match_the_suite(self) -> None:
        designated = suite().injected[0]
        detail = f"expected={suite().injected_expected}; actual=0.75"
        def everywhere(old: str, new: str) -> Callable[[Path], None]:
            def apply(folder: Path) -> None:
                path = folder / "injected.log"
                path.write_bytes(path.read_bytes().replace(old.encode(), new.encode()))
            return apply
        self.assertFinding(everywhere(detail, "expected=1; actual=0.75"),
                           "expected=1, but the suite injects")
        self.assertFinding(everywhere(detail, "expected=0.75; actual=0.75"),
                           "expected equals actual")
        self.assertFinding(everywhere(f"{designated}; {detail}", f"{designated}; reason=x"),
                           "lacks numeric expected= and actual= values")
        self.assertFinding(edit("injected.log", f"FAILURE={designated}; {detail}",
                                f"FAILURE={designated}; expected="
                                f"{suite().injected_expected}; actual=0.7"),
                           "values differ between ASSERT, FAILURE and PRIMARY_FAILURE")

    def test_an_incomplete_normal_run_is_rejected(self) -> None:
        def skipped(folder: Path) -> None:
            (folder / "examples.log").write_bytes(
                ("\r\n".join(run_log("SKIPPED_CASE")).replace("mode=SKIPPED_CASE", "mode=NORMAL")
                 + "\r\n").encode())
        self.assertFinding(skipped, "the run is INCOMPLETE: a mandatory case or assertion "
                                    "did not run")

    def test_certification_log(self) -> None:
        log = POLICY.certification
        self.assertFinding(lambda folder: (folder / log).unlink(), f"{log}: missing")
        self.assertFinding(edit(log, "CERTIFICATION=PASS", "CERTIFICATION=FAIL"),
                           "CERTIFICATION=FAIL")
        self.assertFinding(edit(log, "CERTIFY=examples", "CERTIFY=other"),
                           "expected CERTIFY=examples")
        normal = "MODE=NORMAL; expected=PASS; error=0"
        self.assertFinding(edit(log, normal, "MODE=NORMAL; expected=PASS; error=5"),
                           f"{log}:2: MODE=NORMAL")
        # A deliberate failure that did not reach its caller as ERR_TEST_FAILED.
        self.assertFinding(edit(log, f"MODE=INJECTED_FAILURE; expected=FAIL; error={FAILED}",
                                "MODE=INJECTED_FAILURE; expected=FAIL; error=0"),
                           f"{log}:3: MODE=INJECTED_FAILURE")
        self.assertFinding(edit(log, "completeness=INCOMPLETE; status=PASS",
                                "completeness=INCOMPLETE; status=FAIL"), f"{log}:4:")
        self.assertFinding(edit(log, "MODE=SKIPPED_CASE", "MODE=NORMAL"), f"{log}:4:")
        def truncated(folder: Path) -> None:
            lines = (folder / log).read_bytes().splitlines()
            (folder / log).write_bytes(b"\r\n".join(lines[:-1]))
        self.assertFinding(truncated, f"{log}: incomplete log: it ends before CERTIFICATION=PASS")
        def repeated(folder: Path) -> None:
            text = (folder / log).read_bytes()
            (folder / log).write_bytes(text + text)
        self.assertFinding(repeated, "after CERTIFICATION (more than one certification?)")

    def test_the_harness_must_define_its_failure_number(self) -> None:
        harness = (SOURCE / POLICY.harness).read_bytes().replace(b"\r\n", b"\n")
        without = re.sub(rb"^\s*Public Const ERR_TEST_FAILED\b.*$", b"", harness, flags=re.M)
        root, commit = make_repository(extra={POLICY.harness: without})
        try:
            with evidence(root=root, commit=commit) as folder:
                found = check_excel_run.check(root, commit, folder).findings
        finally:
            shutil.rmtree(root, ignore_errors=True)
        self.assertIn(f"{POLICY.harness}: no ERR_TEST_FAILED constant", found)

    def test_environment(self) -> None:
        self.assertFinding(edit("injected.log", "build=17928", "build=18000"),
                           "different hosts")
        self.assertFinding(edit("examples.log", "build=17928; ", ""), "lacks build")
        def bitness(folder: Path) -> None:
            for name in POLICY.logs.values():
                edit(name, "office=64-bit", "office=32-bit")(folder)
        self.assertFinding(bitness, "not the supported target")

    def test_host_record(self) -> None:
        self.assertFinding(lambda folder: (folder / "host.txt").unlink(), "host.txt: missing")
        self.assertFinding(edit("host.txt", "compile=PASS", "compile=FAIL"), "compile=FAIL")
        self.assertFinding(edit("host.txt", "compile=PASS\n", ""), "no compile=")
        self.assertFinding(edit("host.txt", "OLE Automation; ", "OLE Automation; Scripting; "),
                           "references must be exactly")
        self.assertFinding(edit("host.txt", "windows=Windows 11 Pro 24H2, OS build 26100.4061\n",
                                ""), "no windows=")

    def test_the_host_record_names_the_operator_the_times_and_the_macro_policy(self) -> None:
        for key in check_excel_run.HOST_FIELDS:
            with self.subTest(key=key):
                line = re.search(rf"^{key}=.*\n", HOST, re.M)
                assert line, key
                self.assertFinding(edit("host.txt", line.group(), ""), f"host.txt: no {key}=")
        wrong = {"started=2026-10-24T16:59:30+02:00": ("started=2026-10-24 16:59:30",
                                                       "not an ISO 8601 time with its UTC offset"),
                 "finished=2026-10-24T17:03:10+02:00": ("finished=2026-10-24T16:00:00+02:00",
                                                        "finished before it started"),
                 "excel=16.0 build 17928": ("excel=Microsoft 365", "expected <version> build"),
                 "automation_security=ByUI": ("automation_security=2", "expected one of Low"),
                 "vba_project_access=NOT_TRUSTED": ("vba_project_access=no",
                                                    "expected one of TRUSTED"),
                 "bitness=64-bit": ("bitness=32-bit", "bitness is not 64-bit"),
                 "outcome=PASS": ("outcome=DONE", "outcome=DONE, expected one of PASS")}
        for old, (new, text) in wrong.items():
            with self.subTest(new=new):
                self.assertFinding(edit("host.txt", old, new), text)

    def test_the_host_record_agrees_with_the_logs(self) -> None:
        self.assertFinding(edit("host.txt", "excel=16.0 build 17928", "excel=16.0 build 18000"),
                           "but the logs' ENVIRONMENT is 16.0 build 17928")
        self.assertFinding(edit("host.txt", "windows=Windows 11 Pro 24H2, OS build 26100.4061",
                                "windows=UNKNOWN"), "windows=UNKNOWN, but only a run")
        self.assertFinding(edit("host.txt", "automation_security=ByUI",
                                "automation_security=UNKNOWN"), "may leave a value unread")
        self.assertFinding(edit("host.txt", "vba_project_access=NOT_TRUSTED",
                                "vba_project_access=UNKNOWN"), "vba_project_access=UNKNOWN, but")

    def test_a_record_that_did_not_pass_may_leave_a_host_value_unread(self) -> None:
        # FAIL is not PASS: its record may hold UNKNOWN as a stopped run's may,
        # but never for who ran it, when or how it ended.
        failed = HOST.replace("windows=Windows 11 Pro 24H2, OS build 26100.4061",
                              "windows=UNKNOWN").replace("outcome=PASS", "outcome=FAIL")
        host = dict(line.split("=", 1) for line in failed.splitlines())
        self.assertEqual(check_excel_run.host_findings(host, POLICY, {}), [])
        unread = dict(host, started="UNKNOWN")
        self.assertTrue(any("started" in item
                            for item in check_excel_run.host_findings(unread, POLICY, {})))

    def test_no_record_leaves_the_operator_unread(self) -> None:
        unread = "host.txt: operator=UNKNOWN, but every record names who ran it"
        self.assertFinding(edit("host.txt", "operator=Fixture Operator", "operator=UNKNOWN"),
                           unread)
        stopped = HOST.replace("operator=Fixture Operator", "operator=UNKNOWN") \
            .replace("outcome=PASS", "outcome=UNAVAILABLE")
        host = dict(line.split("=", 1) for line in stopped.splitlines())
        found = check_excel_run.host_findings(host, POLICY, {})
        self.assertEqual(len(found), 1, found)
        self.assertIn(unread, found[0])

    def test_an_outcome_other_than_pass_is_evidence_of_itself_only(self) -> None:
        for outcome in check_excel_run.OUTCOMES[1:]:
            with self.subTest(outcome=outcome):
                self.assertFinding(edit("host.txt", "outcome=PASS", f"outcome={outcome}"),
                                   f"evidence of a run that ended {outcome}, never of a pass")
        # A run that stopped before Excel started cannot read its host values.
        stopped = re.sub(r"^(excel|bitness|automation_security|vba_project_access)=.*$",
                         r"\1=UNKNOWN", HOST, flags=re.M).replace("outcome=PASS",
                                                                  "outcome=UNAVAILABLE")
        with tempfile.TemporaryDirectory() as temporary:
            (Path(temporary) / POLICY.host).write_text(stopped, encoding="utf-8")
            result = check_excel_run.Evidence()
            check_excel_run.check_host(Path(temporary), result, POLICY)
        self.assertEqual(result.findings, ["host.txt: outcome=UNAVAILABLE: the record is "
                                           "evidence of a run that ended UNAVAILABLE, never "
                                           "of a pass"])
        with evidence() as folder, contextlib.redirect_stdout(io.StringIO()) as output:
            (folder / POLICY.host).write_text(stopped, encoding="utf-8")
            self.assertEqual(check_excel_run.main([str(folder), "--commit", COMMIT,
                                                   "--root", str(ROOT)]), 1)
        self.assertIn("Outcome: UNAVAILABLE", output.getvalue())
        self.assertIn("Operator: Fixture Operator", output.getvalue())
        self.assertIn("excel-run: FAIL", output.getvalue())


# A synthetic calculation, declared after the working tree's calculation contracts,
# and the draft accuracy contracts on synthetic cases of it that the measurement
# tests use in place of this repository's contracts.
CALCULATION = "CALC-90"
QUANTITY = "synthetic_total_amount"
CALCULATIONS = "docs/methodology/CALCULATION_CONTRACTS.md"
# By contract ID: its calibration cases and its holdout cases, each with its
# one expected value.
Contracts = dict[str, tuple[dict[str, float], dict[str, float]]]
EXAMPLE_CONTRACTS: Contracts = {
    "ACC-01": ({"synthetic-calibration": 1000000}, {"synthetic-holdout": 2500})}


def accuracy_files(contracts: Contracts) -> dict[str, bytes]:
    """The files of a commit whose registry holds only ``contracts``, with their
    cases and the declaration of their calculation."""
    files = {CALCULATIONS: (SOURCE / CALCULATIONS).read_bytes()
             + f"\n### *Illustrative:* `{CALCULATION}` Synthetic total\n".encode()}
    registry: dict[str, Any] = {"schema_version": 1, "kind": "accuracy-contracts",
                                "min_holdout_cases": 1, "contracts": []}
    for identifier, (calibration, holdout) in contracts.items():
        registry["contracts"].append({
            "id": identifier, "calculation": CALCULATION, "regime": "illustrative",
            "quantities": [QUANTITY], "metric": "absolute", "threshold": 0.000001,
            "provenance": "Synthetic fixture", "status": "draft", "decision": None,
            "calibration_cases": list(calibration), "holdout_cases": list(holdout)})
        for case, value in {**calibration, **holdout}.items():
            fixture = {"schema_version": 1, "kind": "test-fixture", "id": case,
                       "calculation": CALCULATION}
            expected = {"schema_version": 1, "kind": "test-expected", "fixture": case,
                        "outputs": [{"quantity": QUANTITY, "scope": None, "value": value,
                                     "tolerance": {"absolute": 0.000001, "relative": 0},
                                     "reference": {"class": "illustrative"}}]}
            files[f"tests/fixtures/{case}.json"] = json.dumps(fixture).encode()
            files[f"tests/expected/{case}.json"] = json.dumps(expected).encode()
    files[check_accuracy.REGISTRY.as_posix()] = json.dumps(registry).encode()
    return files


@unittest.skipUnless(COMMIT, "needs Git and the modules of tests/modules/")
class ResultsTest(unittest.TestCase):
    """The results of a commit whose registry holds the synthetic EXAMPLE_CONTRACTS."""

    root: ClassVar[Path]
    commit: ClassVar[str]

    @classmethod
    def setUpClass(cls) -> None:
        cls.root, cls.commit = make_repository(extra=accuracy_files(EXAMPLE_CONTRACTS))

    @classmethod
    def tearDownClass(cls) -> None:
        shutil.rmtree(cls.root, ignore_errors=True)

    def findings(self, change: Callable[[Path], None] | None = None) -> list[str]:
        with evidence(change, self.root, self.commit) as folder:
            return check_excel_run.check(self.root, self.commit, folder).findings

    def changed(self, change: Callable[[dict[str, Any]], object]) -> list[str]:
        def apply(folder: Path) -> None:
            document = results_document(self.commit, self.root)
            change(document)
            write_results(folder, document)
        return self.findings(apply)

    def test_the_contracts_are_measured_and_reported(self) -> None:
        with evidence(None, self.root, self.commit) as folder:
            result = check_excel_run.check(self.root, self.commit, folder)
        self.assertEqual(result.findings, [])
        self.assertEqual([line.split(":")[0] for line in result.accuracy],
                         ["ACC-01 calibration", "ACC-01 holdout"])
        self.assertIn("Accuracy: ACC-01 calibration: worst absolute error 0.0",
                      check_excel_run.report(self.commit, result))

    def test_the_repository_contracts_are_measured(self) -> None:
        # One line for each split of this repository's contracts that lists a case.
        registry = SOURCE / check_accuracy.REGISTRY
        if registry.is_file():
            contracts = json.loads(registry.read_text(encoding="utf-8"))["contracts"]
            expected = [f"{contract['id']} {split}" for contract in contracts
                        for split in ("calibration", "holdout") if contract[f"{split}_cases"]]
        else:
            expected = ["not measured"]
        with evidence() as folder:
            result = check_excel_run.check(ROOT, COMMIT, folder)
        self.assertEqual(result.findings, [])
        self.assertEqual([line.split(":")[0] for line in result.accuracy], expected)

    def test_a_missing_results_file_fails_when_contracts_exist(self) -> None:
        found = self.findings(lambda folder: (folder / POLICY.results).unlink())
        self.assertEqual(len(found), 1, found)
        self.assertIn(f"{POLICY.results}: missing", found[0])

    def test_results_of_another_commit_fail(self) -> None:
        found = self.changed(lambda document: document.update(commit="f" * 40))
        self.assertIn("results are of commit ffffffffffff", " ".join(found))

    def test_a_malformed_results_file_is_a_finding(self) -> None:
        def broken(folder: Path) -> None:
            (folder / POLICY.results).write_text("{broken", encoding="utf-8")
        with evidence(broken, self.root, self.commit) as folder:
            result = check_excel_run.check(self.root, self.commit, folder)
        self.assertEqual(len(result.findings), 1, result.findings)
        self.assertIn(f"{POLICY.results}: {folder / POLICY.results}: not valid UTF-8 JSON",
                      result.findings[0])

    def test_results_of_an_overridden_build_are_named(self) -> None:
        stamp = f"{self.commit}{build_workbook.OVERRIDDEN}mutants: CORE_A SHA-256 " \
            f"{'0' * 64}: NOT evidence-grade)"
        found = self.changed(lambda document: document.update(commit=stamp))
        self.assertIn(f"{POLICY.results}: written by a workbook built with overrides, so not "
                      f"evidence for {self.commit}", " ".join(found))

    def test_a_marked_or_missing_stamp_fails(self) -> None:
        for commit in (self.commit + "+local-changes", ""):
            with self.subTest(commit=commit):
                found = self.changed(lambda document: document.update(commit=commit))
                self.assertIn("commit is not a full lower-case commit SHA", " ".join(found))

    def test_a_missing_or_wrong_result_fails(self) -> None:
        def drop(document: dict[str, Any]) -> None:
            document["results"].pop()
        self.assertIn("no result for", " ".join(self.changed(drop)))

        def wrong(document: dict[str, Any]) -> None:
            document["results"][0]["value"] += 1
        found = " ".join(self.changed(wrong))
        self.assertIn("synthetic-calibration synthetic_total_amount is 1000001", found)
        self.assertIn("expected 1000000", found)
        self.assertIn("exceeds the threshold", found)

        def repeated(document: dict[str, Any]) -> None:
            document["results"].append(dict(document["results"][0]))
        self.assertIn("repeats", " ".join(self.changed(repeated)))

    def test_a_commit_without_contracts_needs_no_results(self) -> None:
        root, commit = make_repository(with_cases=False)
        try:
            with evidence(lambda folder: (folder / POLICY.results).unlink()) as folder:
                result = check_excel_run.check(root, commit, folder)
        finally:
            shutil.rmtree(root, ignore_errors=True)
        self.assertEqual(result.findings, [])
        self.assertEqual(result.accuracy, [f"not measured: commit {commit[:12]} has no "
                                           f"{check_accuracy.REGISTRY.as_posix()}"])

    def test_the_contracts_of_the_commit_judge_the_results(self) -> None:
        # A working-tree change to the registry does not change what the commit holds.
        with evidence(None, self.root, self.commit) as folder:
            before = check_excel_run.check(self.root, self.commit, folder).accuracy
            (self.root / check_accuracy.REGISTRY).write_text("{}", encoding="utf-8")
            try:
                after = check_excel_run.check(self.root, self.commit, folder)
            finally:
                subprocess.run(["git", "-C", str(self.root), "checkout", "--",
                                check_accuracy.REGISTRY.as_posix()], check=True)
        self.assertEqual(after.findings, [])
        self.assertEqual(after.accuracy, before)


@unittest.skipUnless(COMMIT, "needs Git and the modules of tests/modules/")
class SuiteTest(unittest.TestCase):
    def test_suite_matches_the_documented_counts(self) -> None:
        examples = suite()
        self.assertEqual(len(examples.cases), examples.expected_cases)
        self.assertEqual(len(examples.assertions), examples.expected_assertions)
        self.assertEqual(examples.last_case[0], examples.cases[-1])
        self.assertEqual(len(examples.injected), 1)
        self.assertEqual(examples.injected_expected, "0.76")

    def test_the_entry_point_saves_the_logs_the_checker_reads(self) -> None:
        source = (ROOT / POLICY.suite_source).read_text(encoding="ascii")
        saved = re.findall(r'CertifyMode "(\w+)", (True|False), Folder & "\\([\w.]+)"', source)
        self.assertEqual({mode: log for mode, _, log in saved}, POLICY.logs)
        self.assertEqual({mode for mode, must, _ in saved if must == "True"},
                         set(POLICY.must_pass))
        self.assertIn(f'EndCertification Folder & "\\{POLICY.certification}"', source)
        # The results are those of the normal run, saved before the next run resets them.
        normal = re.search(r'CertifyMode "NORMAL".*\n\s+TEST_Harness\.SaveResults Folder & '
                           r'"\\([\w.]+)"\n\s+CertifyMode', source)
        self.assertIsNotNone(normal)
        assert normal is not None
        self.assertEqual(normal.group(1), POLICY.results)

    def test_failed_number_is_the_harness_constant(self) -> None:
        self.assertEqual(FAILED, -2147221504 + 4065)

    def test_the_policy_matches_the_suite_and_the_host_policy(self) -> None:
        examples = suite()
        self.assertEqual((POLICY.expected_cases, POLICY.expected_assertions),
                         (examples.expected_cases, examples.expected_assertions))
        source = (SOURCE / POLICY.suite_source).read_text(encoding="ascii")
        self.assertIn(f"Public Sub {POLICY.entry_point}(", source)
        self.assertIn(f'CertifyMode "{next(iter(POLICY.must_pass))}"', source)
        self.assertRegex(source, rf'Const SUITE_NAME\s+As String = "{POLICY.suite}"')
        installation = (SOURCE / "INSTALLATION.md").read_text(encoding="utf-8")
        self.assertIn(f"TEST_Examples.{POLICY.entry_point}", installation)
        for reference in POLICY.references:
            self.assertIn(reference, installation)
        # The installation guide tells the reader which files to leave in the
        # evidence folder, so it must name each one the policy names.
        for name in (*POLICY.logs.values(), POLICY.certification, POLICY.host, POLICY.manifest,
                     POLICY.results, f"{POLICY.exported}\\"):
            with self.subTest(file=name):
                self.assertIn(f"`{name}", installation)
        self.assertIn(POLICY.office, installation)


Change = Callable[[dict[str, Any]], object]

# A project suite: two cases, and a third that only a 64-bit host compiles.
PROJECT = "tests/modules/TEST_Project.bas"
PROJECT_SOURCE = b"""Attribute VB_Name = "TEST_Project"
Option Explicit
Option Private Module

#If Win64 Then
    Private Const EXPECTED_CASES As Long = 3
    Private Const EXPECTED_ASSERTIONS As Long = 4
#Else
    Private Const EXPECTED_CASES As Long = 2
    Private Const EXPECTED_ASSERTIONS As Long = 3
#End If

Public Sub CertifyProject( _
    ByVal Folder As String, _
    ByVal Fixtures As String)
End Sub

Private Sub RunSuite()
        CaseFirst
        CaseSecond
#If Win64 Then
        CaseWide
#End If
End Sub

Private Sub CaseFirst()
        TEST_Harness.BeginCase "project-first"
        TEST_Harness.AssertTrue "holds", True
End Sub

Private Sub CaseSecond()
        TEST_Harness.BeginCase "project-second"
        TEST_Harness.AssertTrue "left", True
        TEST_Harness.AssertTrue "right", True
End Sub

Private Sub CaseWide()
        TEST_Harness.BeginCase "project-wide"
        TEST_Harness.AssertTrue "pointer", True
End Sub
"""
PROJECT_ENTRY: dict[str, Any] = {
    "name": "project", "source": PROJECT, "certification_entry_point": "CertifyProject",
    "expected_cases": 3, "expected_assertions": 4, "log": "project.log",
    "certification_log": "project-certification.log"}


# The same suite when it records the output of an accuracy contract of its own.
PROJECT_RESULTS_ENTRY: dict[str, Any] = dict(PROJECT_ENTRY, results="project-results.json")
PROJECT_CASE = "project-total"
# The example contract, and a second contract, ACC-02, whose one case only the
# project suite computes.
PROJECT_CONTRACTS: Contracts = {**EXAMPLE_CONTRACTS, "ACC-02": ({PROJECT_CASE: 250}, {})}


# A relation suite, whose public runner is the entry point that certifies it.
RELATIONS = "tests/modules/TEST_Relations.bas"
RELATIONS_SOURCE = b"""Attribute VB_Name = "TEST_Relations"
Option Explicit
Option Private Module

Private Const EXPECTED_CASES As Long = 2
Private Const EXPECTED_ASSERTIONS As Long = 4

Public Sub RunRelationTests()
        CaseAdditivity
        CaseSymmetry
End Sub

Private Sub CaseAdditivity()
        TEST_Harness.BeginCase "additivity"
        TEST_Harness.AssertTrue "holds", True
        TEST_Harness.AssertTrue "detects-corruption", Not False
End Sub

Private Sub CaseSymmetry()
        TEST_Harness.BeginCase "symmetry"
        TEST_Harness.AssertTrue "holds", True
        TEST_Harness.AssertTrue "detects-corruption", Not False
End Sub
"""


def policy_text(change: Change) -> str:
    data = json.loads((SOURCE / check_excel_run.POLICY).read_text(encoding="utf-8"))
    change(data)
    return json.dumps(data)


class PolicyTest(unittest.TestCase):
    def test_the_working_policy_is_valid(self) -> None:
        self.assertEqual(POLICY.must_pass, frozenset({"NORMAL"}))
        self.assertEqual(list(POLICY.logs), list(check_excel_run.MODE_KINDS))

    def test_a_malformed_policy_is_rejected(self) -> None:
        def mode(index: int, key: str, value: object) -> Change:
            def apply(data: dict[str, Any]) -> None:
                data["modes"][index][key] = value
            return apply
        def section(name: str, key: str, value: object) -> Change:
            def apply(data: dict[str, Any]) -> None:
                data[name][key] = value
            return apply
        cases: dict[str, Change] = {
            "schema_version": lambda data: data.update(schema_version=3),
            "or 1 without suites": lambda data: data.update(schema_version=1),
            "its own name and source": lambda data: data["suites"].append(
                dict(PROJECT_ENTRY, source=data["suite"]["source"])),
            "modes must be": mode(1, "mode", "OTHER"),
            "only NORMAL must pass": mode(1, "must_pass", True),
            "must_pass is true or false": mode(0, "must_pass", "yes"),
            "distinct and named": mode(1, "log", "examples.log"),
            "plain names": section("evidence", "host_record", "../host.txt"),
            "inside the repository": section("suite", "source", "/tmp/TEST_Examples.bas"),
            "never certified evidence": lambda data: data["suites"].append(
                dict(PROJECT_ENTRY, source="tests/modules/TEST_BenchmarkTotals.bas")),
            "positive integer": section("suite", "expected_cases", 0),
            "listed once each": section("host", "references", []),
            "host office": section("host", "office", "ARM"),
            "KeyError": lambda data: data.pop("harness"),
        }
        for text, change in cases.items():
            with self.subTest(text=text), self.assertRaises(check_excel_run.PolicyError) as raised:
                check_excel_run.parse_policy(policy_text(change))
            self.assertIn(text, str(raised.exception))
        with self.assertRaises(check_excel_run.PolicyError):
            check_excel_run.parse_policy("not json")

    def test_a_project_suite_is_validated_like_the_example_suite(self) -> None:
        cases: dict[str, Change] = {
            "its own name and source": lambda data: data["suites"].extend(
                [PROJECT_ENTRY, dict(PROJECT_ENTRY, log="other.log",
                                     certification_log="other-certification.log")]),
            "positive integer": lambda data: data["suites"].append(
                dict(PROJECT_ENTRY, expected_cases=0)),
            "distinct and named": lambda data: data["suites"].append(
                dict(PROJECT_ENTRY, log="examples.log")),
            "evidence files must be distinct": lambda data: data["suites"].append(
                dict(PROJECT_ENTRY, results="results.json")),
            "must be plain names": lambda data: data["suites"].append(
                dict(PROJECT_ENTRY, results="../project-results.json")),
            "evidence files must be distinct and named": lambda data: data["suites"].append(
                dict(PROJECT_ENTRY, results="")),
            "KeyError": lambda data: data["suites"].append(
                {key: value for key, value in PROJECT_ENTRY.items() if key != "log"}),
            "name of a VBA procedure": lambda data: data["suites"].append(
                dict(PROJECT_ENTRY, runner="Run Suite")),
        }
        for text, change in cases.items():
            with self.subTest(text=text), self.assertRaises(check_excel_run.PolicyError) as raised:
                check_excel_run.parse_policy(policy_text(change))
            self.assertIn(text, str(raised.exception))
        accepted = check_excel_run.parse_policy(policy_text(
            lambda data: data["suites"].append(PROJECT_ENTRY)))
        # The suites the working policy lists come first.
        self.assertEqual([item.name for item in accepted.project_suites],
                         [*(item.name for item in POLICY.project_suites), "project"])
        self.assertIsNone(accepted.project_suites[-1].results)
        self.assertEqual(accepted.project_suites[-1].runner, "RunSuite")
        named = check_excel_run.parse_policy(policy_text(
            lambda data: data["suites"].append(dict(PROJECT_ENTRY, runner="RunRelationTests"))))
        self.assertEqual(named.project_suites[-1].runner, "RunRelationTests")

    def test_a_suite_is_read_through_the_runner_its_entry_names(self) -> None:
        # Relation tests, as the harness guide says to certify them: their runner is
        # RunRelationTests, not RunSuite.
        suite = read_suite(RELATIONS_SOURCE, "64-bit", "RunRelationTests")
        self.assertEqual(suite.cases, ["additivity", "symmetry"])
        self.assertEqual((len(suite.assertions), suite.expected_cases,
                          suite.expected_assertions), (4, 2, 4))
        self.assertIn("symmetry/detects-corruption", suite.assertions)
        entry = check_excel_run.ProjectSuite(
            "relations", RELATIONS, "RunRelationTests", 2, 4,
            "relations.log", "relations-certification.log", runner="RunRelationTests")
        self.assertEqual(check_excel_run.suite_findings(entry, suite, RELATIONS_SOURCE, "x"), [])
        unread = check_excel_run.suite_findings(
            dataclasses.replace(entry, runner="RunSuite"), read_suite(RELATIONS_SOURCE, "64-bit"),
            RELATIONS_SOURCE, "x")
        self.assertEqual(len(unread), 1)
        self.assertIn(f"no suite in {RELATIONS} at x: RunSuite calls no Case procedure",
                      unread[0])
        saving = check_excel_run.parse_policy(policy_text(
            lambda data: data["suites"].append(PROJECT_RESULTS_ENTRY)))
        self.assertEqual(check_excel_run.results_files(saving),
                         [*check_excel_run.results_files(POLICY), "project-results.json"])
        # A version 1 policy, as older commits hold it, is still read.
        old = check_excel_run.parse_policy(policy_text(
            lambda data: (data.update(schema_version=1), data.pop("suites"))))
        self.assertEqual(old.project_suites, ())

    @unittest.skipUnless(COMMIT, "needs Git and the modules of tests/modules/")
    def test_the_commit_policy_judges_the_evidence(self) -> None:
        def run(root: Path, commit: str) -> list[str]:
            try:
                with evidence() as folder:
                    return check_excel_run.check(root, commit, folder).findings
            finally:
                shutil.rmtree(root, ignore_errors=True)
        missing = run(*make_repository(with_policy=False))
        self.assertEqual(len(missing), 1)
        self.assertIn(f"{check_excel_run.POLICY}: missing at", missing[0])
        malformed = run(*make_repository(b"{}"))
        self.assertEqual(len(malformed), 1)
        self.assertIn(check_excel_run.POLICY, malformed[0])
        stated = suite().expected_cases
        counts = run(*make_repository(policy_text(
            lambda data: data["suite"].update(expected_cases=stated + 1)).encode()))
        self.assertTrue(any(f"EXPECTED_CASES is {stated}, but {check_excel_run.POLICY} "
                            f"expects {stated + 1}" in item for item in counts), counts)
        entry = run(*make_repository(policy_text(
            lambda data: data["suite"].update(
                certification_entry_point="CertifyOther")).encode()))
        self.assertTrue(any("no Public Sub CertifyOther" in item for item in entry), entry)
        for key, name in (("harness", "TEST_Harness.bas"), ("suite", "TEST_Examples.bas")):
            with self.subTest(moved=key):
                path = f"does-not-exist/{name}"
                def moved(data: dict[str, Any], key: str = key, path: str = path) -> None:
                    if key == "harness":
                        data["harness"] = path
                    else:
                        data["suite"]["source"] = path
                found = run(*make_repository(policy_text(moved).encode()))
                self.assertIn(f"{path}: not in commit", " ".join(found))
        logs = run(*make_repository(policy_text(
            lambda data: data["modes"][2].update(log="other.log")).encode()))
        self.assertIn("other.log: missing", logs)



class ProjectSuiteTest(unittest.TestCase):
    """A project suite beside the example suite: its normal run and its certification."""

    root: ClassVar[Path]
    commit: ClassVar[str]

    @classmethod
    def setUpClass(cls) -> None:
        if not shutil.which("git"):
            raise unittest.SkipTest("needs Git")
        policy = policy_text(lambda data: data["suites"].append(PROJECT_ENTRY)).encode()
        cls.root, cls.commit = make_repository(policy, extra={PROJECT: PROJECT_SOURCE})

    @classmethod
    def tearDownClass(cls) -> None:
        shutil.rmtree(cls.root, ignore_errors=True)

    def findings(self, change: Callable[[Path], None] | None = None) -> list[str]:
        with evidence(change, self.root, self.commit) as folder:
            return check_excel_run.check(self.root, self.commit, folder).findings

    def test_both_suites_certified_pass(self) -> None:
        # evidence() writes the two logs the project suite's certification entry
        # point saves, as it does for each suite a working policy lists.
        with evidence(None, self.root, self.commit) as folder:
            self.assertTrue((folder / "project.log").read_bytes().endswith(
                b"RESULT=PASS; suite=project; operation=PASS; completeness=COMPLETE; cases=3; "
                b"assertions=4; failures=0; cleanup=PASS; recovery=NONE\r\n"))
            self.assertIn(b"CERTIFICATION=PASS; suite=project; completeness=COMPLETE; modes=1",
                          (folder / "project-certification.log").read_bytes())
        self.assertEqual(self.findings(), [])

    def test_a_missing_or_wrong_project_log_fails(self) -> None:
        self.assertIn("project.log: missing",
                      self.findings(lambda folder: (folder / "project.log").unlink()))
        self.assertIn("project-certification.log: missing",
                      self.findings(lambda folder: (folder / "project-certification.log")
                                    .unlink()))
        dropped = self.findings(edit("project.log", "CASE=project-wide\r\n", ""))
        self.assertTrue(any(item.startswith("project.log:") for item in dropped), dropped)
        wrong = self.findings(edit("project-certification.log", "CERTIFICATION=PASS",
                                   "CERTIFICATION=FAIL"))
        self.assertTrue(any(item.startswith("project-certification.log:") for item in wrong),
                        wrong)

    def test_the_project_logs_come_from_the_same_host(self) -> None:
        other = self.findings(edit("project.log", "build=17928", "build=17929"))
        self.assertIn("the logs come from different hosts (ENVIRONMENT differs)", other)

    def test_a_suite_that_names_no_results_file_saves_none(self) -> None:
        self.assertEqual(check_excel_run.results_files(check_excel_run.policy_at(
            self.root, self.commit)), check_excel_run.results_files(POLICY))

    def test_counts_follow_the_bitness_the_policy_supports(self) -> None:
        wide, narrow = read_suite(PROJECT_SOURCE, "64-bit"), read_suite(PROJECT_SOURCE, "32-bit")
        self.assertEqual((len(wide.cases), len(wide.assertions), wide.expected_cases),
                         (3, 4, 3))
        self.assertEqual((len(narrow.cases), len(narrow.assertions), narrow.expected_cases),
                         (2, 3, 2))
        stated = check_excel_run.ProjectSuite("project", PROJECT, "CertifyProject", 2, 3,
                                              "project.log", "project-certification.log")
        self.assertTrue(any("EXPECTED_CASES is 3, but" in item for item in
                            check_excel_run.suite_findings(stated, wide, PROJECT_SOURCE, "x")))
        self.assertEqual(check_excel_run.suite_findings(stated, narrow, PROJECT_SOURCE, "x"), [])


class ProjectResultsTest(unittest.TestCase):
    """A project suite that computes the output of its own accuracy contract saves it
    in its own results file, which is measured together with the example suite's."""

    root: ClassVar[Path]
    commit: ClassVar[str]

    @classmethod
    def setUpClass(cls) -> None:
        if not shutil.which("git"):
            raise unittest.SkipTest("needs Git")
        policy = policy_text(lambda data: data["suites"].append(PROJECT_RESULTS_ENTRY)).encode()
        cls.root, cls.commit = make_repository(policy, extra={
            PROJECT: PROJECT_SOURCE, **accuracy_files(PROJECT_CONTRACTS)})

    @classmethod
    def tearDownClass(cls) -> None:
        shutil.rmtree(cls.root, ignore_errors=True)

    def both_suites(self, folder: Path) -> None:
        """The example suite's results file with the outputs of every contract but
        the project suite's."""
        document = results_document(self.commit, self.root)
        document["results"] = [item for item in document["results"]
                               if item["case"] != PROJECT_CASE]
        write_results(folder, document)

    def checked(self, results: list[dict[str, Any]] | None = None
                ) -> check_excel_run.Evidence:
        """The evidence of both suites; the project suite's results file holds
        ``results``, by default the one output of its case."""
        project = [{"case": PROJECT_CASE, "variant": None, "quantity": QUANTITY,
                    "scope": None, "value": 250}] if results is None else results

        def apply(folder: Path) -> None:
            self.both_suites(folder)
            (folder / "project-results.json").write_text(json.dumps(
                {"schema_version": 1, "kind": "test-results", "commit": self.commit,
                 "results": project}), encoding="utf-8")
        with evidence(apply, self.root, self.commit) as folder:
            return check_excel_run.check(self.root, self.commit, folder)

    def test_the_project_contract_is_measured_alongside_the_example_contract(self) -> None:
        result = self.checked()
        self.assertEqual(result.findings, [])
        self.assertEqual([line.split(":")[0] for line in result.accuracy],
                         ["ACC-01 calibration", "ACC-01 holdout", "ACC-02 calibration"])

    def test_the_project_results_are_needed(self) -> None:
        found = self.checked([]).findings
        self.assertTrue(any(f"no result for {PROJECT_CASE}" in item for item in found), found)

        def without_project_results(folder: Path) -> None:
            self.both_suites(folder)
            (folder / "project-results.json").unlink()
        with evidence(without_project_results, self.root, self.commit) as folder:
            missing = check_excel_run.check(self.root, self.commit, folder).findings
        self.assertEqual(len(missing), 1, missing)
        self.assertIn("project-results.json: missing, but", missing[0])

    def test_a_result_held_by_two_files_is_rejected(self) -> None:
        repeated = [*results_document(self.commit, self.root)["results"][:1],
                    {"case": PROJECT_CASE, "variant": None, "quantity": QUANTITY,
                     "scope": None, "value": 250}]
        found = self.checked(repeated).findings
        names = check_excel_run.results_files(check_excel_run.policy_at(self.root, self.commit))
        self.assertEqual(names[-1], "project-results.json")
        self.assertEqual(len(found), 1, found)
        self.assertIn(f"{', '.join(names)}: ", found[0])
        self.assertIn(f"project-results.json: synthetic-calibration {QUANTITY} is also in ",
                      found[0])
        self.assertTrue(found[0].endswith("results.json"), found[0])


# A document module as Excel exports it, and a class with the same attributes.
DOCUMENT_HEADER = "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\n"
DOCUMENT_ATTRIBUTE_LINES = "\n".join(
    vba_project.document_module("ThisWorkbook", True).source.split("\r\n")[1:-2])
DOCUMENT = (f'{DOCUMENT_HEADER}Attribute VB_Name = "ThisWorkbook"\n{DOCUMENT_ATTRIBUTE_LINES}\n'
            "Option Explicit\n\nPrivate Sub Workbook_Open()\n"
            '    Application.StatusBar = "Opened"   \'Shown once\nEnd Sub\n').encode()
CLASS = DOCUMENT.replace(b'"ThisWorkbook"', b'"CFixture"')


class SourceIdentityTest(unittest.TestCase):
    """The export is compared with the commit without the letter case of code, as
    the VBE re-cases a name (HB-01), and without the attributes Excel writes for a
    document module; any other difference fails."""

    root: ClassVar[Path]
    commit: ClassVar[str]

    @classmethod
    def setUpClass(cls) -> None:
        if not shutil.which("git"):
            raise unittest.SkipTest("needs Git")
        cls.root, cls.commit = make_repository(extra={"src/workbook/ThisWorkbook.cls": DOCUMENT,
                                                      "src/classes/CFixture.cls": CLASS})

    @classmethod
    def tearDownClass(cls) -> None:
        shutil.rmtree(cls.root, ignore_errors=True)

    def checked(self, *changes: Callable[[Path], None]) -> check_excel_run.Evidence:
        def apply(folder: Path) -> None:
            for change in changes:
                change(folder)
        with evidence(apply, self.root, self.commit) as folder:
            return check_excel_run.check(self.root, self.commit, folder)

    def test_a_recased_identifier_passes_with_a_counted_note(self) -> None:
        result = self.checked(edit("exported/TEST_Examples.bas",
                                   'If mMode = "INJECTED_FAILURE" Then',
                                   'If MMODE = "INJECTED_FAILURE" Then'),
                              edit("exported/TEST_Harness.bas", "ERR_TEST_FAILED        As Long",
                                   "Err_Test_Failed        As Long"))
        self.assertEqual(result.findings, [])
        self.assertIn("exported/TEST_Examples.bas: 1 identifier(s) re-cased outside strings and "
                      "comments, as the VBE re-cases a name (HB-01): mMode -> MMODE",
                      result.notes)
        self.assertTrue(any("TEST_Harness.bas: 1 identifier(s)" in note and
                            "ERR_TEST_FAILED -> Err_Test_Failed" in note
                            for note in result.notes), result.notes)
        self.assertIn("Note: exported/TEST_Examples.bas: 1 identifier(s)",
                      check_excel_run.report(self.commit, result))

    def test_a_change_of_case_in_a_string_or_a_comment_fails(self) -> None:
        for change in (edit("exported/TEST_Examples.bas", '"INJECTED_FAILURE" Then',
                            '"injected_failure" Then'),
                       edit("exported/TEST_Examples.bas", "'Expected sum;", "'expected sum;"),
                       edit("exported/TEST_Harness.bas", "'vbObjectError + 4065",
                            "'VbObjectError + 4065"),
                       edit("exported/ThisWorkbook.cls", '"Opened"', '"opened"'),
                       edit("exported/ThisWorkbook.cls", "'Shown once", "'shown once")):
            with self.subTest(change=change):
                found = self.checked(change).findings
                self.assertTrue(any("differs from" in item and "first difference at line" in item
                                    for item in found), found)

    def test_any_other_difference_still_fails(self) -> None:
        found = self.checked(edit("exported/TEST_Examples.bas", "If mMode =", "If mMode2 =")
                             ).findings
        self.assertTrue(any("TEST_Examples.bas: differs" in item for item in found), found)

    def test_a_comment_continued_on_the_next_line_keeps_its_case(self) -> None:
        commit = b"x = 1 ' note _\nContinued Text\ny = 2\n"
        self.assertEqual(check_excel_run.source_difference(
            b"x = 1 ' note _\ncontinued text\ny = 2\n", commit, False)[0], 2)
        self.assertEqual(check_excel_run.source_difference(
            b"X = 1 ' note _\nContinued Text\nY = 2\n", commit, False), (0, ["x -> X", "y -> Y"], []))
        self.assertEqual(check_excel_run.source_difference(
            b"x = [Sheet1!A1]\n", b"x = [sheet1!A1]\n", False)[0], 1)
        self.assertEqual(check_excel_run.source_difference(b"x = 1\n", b"x = 1\ny = 2\n",
                                                           False)[0], 2)

    def test_the_tolerated_attributes_are_those_excel_writes_for_a_document(self) -> None:
        # The build writes the same attributes for a document module without a source file.
        written = re.findall(r"^Attribute (\w+) =", DOCUMENT_ATTRIBUTE_LINES, re.M)
        self.assertEqual(tuple(written), check_excel_run.DOCUMENT_ATTRIBUTES)
        self.assertEqual(check_excel_run.document_modules(self.root, self.commit),
                         {"ThisWorkbook.cls"})

    def test_document_module_attributes_are_not_compared(self) -> None:
        def attributes(name: str) -> Callable[[Path], None]:
            def apply(folder: Path) -> None:
                path = folder / "exported" / name
                text = path.read_bytes().decode("ascii")
                for line in ("Attribute VB_TemplateDerived = False\r\n",
                             "Attribute VB_Customizable = True\r\n"):
                    text = text.replace(line, "")
                text = text.replace("VB_PredeclaredId = True", "VB_PredeclaredId = False")
                path.write_bytes(text.encode("ascii"))
            return apply
        result = self.checked(attributes("ThisWorkbook.cls"))
        self.assertEqual(result.findings, [])
        self.assertIn("exported/ThisWorkbook.cls: document-module attributes Excel writes, not "
                      "compared: VB_Customizable, VB_PredeclaredId, VB_TemplateDerived",
                      result.notes)
        # A class module's attributes are its source, and so is a document's code name.
        found = self.checked(attributes("CFixture.cls")).findings
        self.assertTrue(any("CFixture.cls: differs" in item for item in found), found)
        found = self.checked(edit("exported/ThisWorkbook.cls", 'VB_Name = "ThisWorkbook"',
                                  'VB_Name = "thisworkbook"')).findings
        self.assertTrue(any("ThisWorkbook.cls: differs" in item for item in found), found)


if __name__ == "__main__":
    unittest.main()
