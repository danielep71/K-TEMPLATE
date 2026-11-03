"""What tools/devbuild/excel_run.py does without Excel.

The command runs only on Windows with Excel. These tests check the rest: the
host record it writes is accepted by check_excel_run.py and names the
project's references, the Windows description, the evidence folder it
empties, and the PowerShell script that drives Excel.
"""

from __future__ import annotations

import base64
import io
import shutil
import subprocess
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools" / "devbuild"))
sys.path.insert(0, str(ROOT / "tools"))

import check_excel_run  # noqa: E402
import excel_run  # noqa: E402

POLICY = check_excel_run.working_policy(ROOT)


class HostRecordTest(unittest.TestCase):
    def test_the_built_references_are_the_policy_references(self) -> None:
        self.assertEqual(sorted(excel_run.built_references()), sorted(POLICY.references))

    def test_the_record_passes_the_check_only_with_identity_by_rebuild(self) -> None:
        record = excel_run.host_record("Windows 11 Pro, version 24H2, OS build 26100.4061",
                                       excel_run.built_references())
        self.assertIn("compile=ON_DEMAND", record)
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            (folder / POLICY.host).write_text(record, encoding="utf-8")
            for rebuilt, expected in ((True, []), (False, ["compile=ON_DEMAND"])):
                with self.subTest(rebuilt=rebuilt):
                    evidence = check_excel_run.Evidence()
                    check_excel_run.check_host(folder, evidence, POLICY, rebuilt)
                    self.assertEqual(len(evidence.findings), len(expected), evidence.findings)
                    for text, finding in zip(expected, evidence.findings):
                        self.assertIn(text, finding)

    def test_windows_is_described_as_winver_shows_it(self) -> None:
        values = {"ProductName": "Windows 10 Pro", "DisplayVersion": "24H2",
                  "CurrentBuild": "26100", "UBR": 4061}
        self.assertEqual(excel_run.describe_windows(values),
                         "Windows 11 Pro, version 24H2, OS build 26100.4061")
        older = {"ProductName": "Windows 10 Enterprise", "DisplayVersion": "22H2",
                 "CurrentBuild": "19045", "UBR": 5247}
        self.assertEqual(excel_run.describe_windows(older),
                         "Windows 10 Enterprise, version 22H2, OS build 19045.5247")
        self.assertEqual(excel_run.describe_windows({"ProductName": "Windows 10 Pro"}),
                         "Windows 10 Pro")


class EvidenceFolderTest(unittest.TestCase):
    def test_a_folder_under_test_results_is_emptied(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            folder = root / "test-results" / "excel-run"
            (folder / "exported").mkdir(parents=True)
            (folder / "examples.log").write_text("old", encoding="utf-8")
            excel_run.prepare(folder, root)
            self.assertEqual(list(folder.iterdir()), [])

    def test_another_non_empty_folder_is_refused_and_kept(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            folder = root / "elsewhere"
            folder.mkdir()
            (folder / "keep.txt").write_text("mine", encoding="utf-8")
            with self.assertRaisesRegex(excel_run.RunError, "not empty"):
                excel_run.prepare(folder, root)
            self.assertTrue((folder / "keep.txt").is_file())
            new = root / "new"
            excel_run.prepare(new, root)
            self.assertTrue(new.is_dir())


class PowerShellTest(unittest.TestCase):
    script = excel_run.powershell_script(Path(r"C:\Dev\it's\Workbook-development.xlsm"),
                                         Path(r"C:\Dev\run"), Path(r"C:\Dev\tests\fixtures"),
                                         "TEST_Examples", POLICY.entry_point,
                                         "TEST_Harness.LastResult", 300)

    def test_the_workbook_is_opened_read_only_and_never_saved(self) -> None:
        self.assertIn("$excel.Workbooks.Open($workbookPath, 0, $true)", self.script)
        self.assertIn("$book.Close($false)", self.script)
        self.assertIn("$excel.Quit()", self.script)
        self.assertNotIn(".Save", self.script)
        # Nothing reaches the VBA project, so no Trust access is needed.
        self.assertNotIn("VBProject", self.script)

    def test_no_macro_setting_changes_and_a_person_enables_the_macros(self) -> None:
        # A policy may open the workbook with macros disabled (HB-02); the script never
        # overrides it, and waits for Enable Content instead.
        self.assertNotIn("AutomationSecurity", self.script)
        probe = self.script.index("[void]$excel.Run($prefix + 'TEST_Harness.LastResult')")
        opened = self.script.index("$excel.Workbooks.Open(")
        run = self.script.index(f'$excel.Run($prefix + "TEST_Examples.{POLICY.entry_point}"')
        self.assertLess(opened, probe)
        self.assertLess(probe, run)
        self.assertIn("$deadline = (Get-Date).AddSeconds(300)", self.script)
        self.assertIn("Start-Sleep -Seconds 2", self.script)
        self.assertIn('$failure = "NOT-ENABLED"', self.script)
        self.assertIn("exit 4", self.script)

    def test_progress_records_do_not_reach_the_output(self) -> None:
        self.assertIn("$ProgressPreference = 'SilentlyContinue'", self.script)

    def test_the_entry_point_runs_with_both_folders(self) -> None:
        self.assertIn(f'$excel.Run($prefix + "TEST_Examples.{POLICY.entry_point}", $evidence, '
                      "$fixtures)", self.script)
        self.assertIn(r"$evidence = 'C:\Dev\run'", self.script)
        self.assertIn(r"$fixtures = 'C:\Dev\tests\fixtures'", self.script)
        # A quote in a path is doubled, so a path cannot end the literal.
        self.assertIn(r"$workbookPath = 'C:\Dev\it''s\Workbook-development.xlsm'", self.script)

    def test_the_command_is_encoded_and_skips_the_profile(self) -> None:
        text = excel_run.encoded(self.script)
        self.assertEqual(base64.b64decode(text).decode("utf-16-le"), self.script)
        with mock.patch.object(subprocess, "run") as run:
            run.return_value = subprocess.CompletedProcess([], 0, "EXCEL-RUN=DONE", "")
            self.assertEqual(excel_run.run_excel(self.script, 5), (0, "EXCEL-RUN=DONE"))
        command = run.call_args.args[0]
        self.assertEqual(command[:4], ["powershell.exe", "-NoProfile", "-NonInteractive",
                                       "-EncodedCommand"])
        self.assertNotIn("Bypass", command)


class CommandTest(unittest.TestCase):
    def test_macros_never_enabled_stop_without_a_record(self) -> None:
        built = mock.Mock(commit="a" * 40, clean=True)
        folder = ROOT / "test-results" / "excel-run-unit-test"
        try:
            with mock.patch.object(excel_run.build_workbook, "build", return_value=built), \
                    mock.patch.object(excel_run, "run_excel",
                                      return_value=(excel_run.NOT_ENABLED,
                                                    "EXCEL-RUN=NOT-ENABLED")), \
                    mock.patch.object(excel_run, "windows_version") as windows, \
                    redirect_stdout(io.StringIO()) as output:
                with self.assertRaisesRegex(excel_run.RunError, "not enabled within 300 s"):
                    excel_run.record(ROOT, folder, 600, 300)
            windows.assert_not_called()
            self.assertFalse((folder / POLICY.host).exists())
            self.assertIn("click Enable Content", output.getvalue())
        finally:
            shutil.rmtree(folder, ignore_errors=True)

    def test_the_wait_for_enable_content_is_shorter_than_the_run(self) -> None:
        with redirect_stdout(io.StringIO()), mock.patch("sys.stderr", io.StringIO()), \
                self.assertRaises(SystemExit):
            excel_run.main(["--timeout", "300", "--enable-timeout", "300"])

    @unittest.skipIf(sys.platform == "win32", "checks the refusal off Windows")
    def test_off_windows_it_records_nothing(self) -> None:
        output = io.StringIO()
        with redirect_stdout(output):
            self.assertEqual(excel_run.main([]), 2)
        self.assertIn("SKIP", output.getvalue())

    def test_a_build_that_is_not_evidence_grade_stops_before_excel(self) -> None:
        built = mock.Mock(commit="a" * 40, clean=False)
        with mock.patch.object(excel_run.build_workbook, "build", return_value=built), \
                mock.patch.object(excel_run, "run_excel") as run, \
                redirect_stdout(io.StringIO()):
            with self.assertRaisesRegex(excel_run.RunError, "local changes"):
                excel_run.record(ROOT, ROOT / "test-results" / "excel-run-test", 5)
        run.assert_not_called()


if __name__ == "__main__":
    unittest.main()
