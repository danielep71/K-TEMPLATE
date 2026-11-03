#!/usr/bin/env python3
"""Record an Excel run of the harness certification with one command, on Windows.

Usage: python tools/devbuild/excel_run.py [--evidence FOLDER] [--timeout SECONDS]
                                         [--enable-timeout SECONDS]

Run it in a clean checkout of the candidate commit, on Windows with Excel
installed. It replaces the manual steps of INSTALLATION.md, Recording an
Excel run, and never needs Trust access to the VBA project object model:

1. builds the development workbook with build_workbook.py, and stops if the
   build is not evidence-grade (a change in the checkout);
2. empties the evidence folder (default test-results/excel-run, which Git
   ignores; another folder must be empty or absent);
3. opens the workbook read-only in a visible Excel through COM, from
   PowerShell started with -NoProfile and -EncodedCommand. It changes no
   macro setting: where Excel opens it with macros disabled, as under a
   policy that makes programs follow the user's macro setting
   (VALIDATION_PLAN.md, HB-02), you click Enable Content once, and the
   command waits for that, up to --enable-timeout seconds. Then it runs the
   certification entry point of the example suite and of each project suite
   the policy lists, in order, with the evidence folder and tests/fixtures,
   closes the workbook without saving and quits Excel;
4. writes host.txt: Windows from the registry, the references the build
   gives the project, and compile=ON_DEMAND, because nothing can compile the
   whole project without Trust access, so VBA compiled the modules the run
   called;
5. runs tools/check_excel_run.py with --workbook and --require-binding: with
   no exported folder it proves identity by rebuilding the commit and
   comparing the workbook's parts. It prints the report, and saves it
   beside the evidence folder as <folder>-report.txt.

It exits 0 when the check passes, 1 when it fails or the run could not
complete, and 2 when it is not run on Windows. Standard library only.
"""

from __future__ import annotations

import argparse
import base64
import os
import shutil
import subprocess
import sys
from collections.abc import Mapping, Sequence
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(ROOT / "tools"))

import build_workbook  # noqa: E402
import check_excel_run  # noqa: E402
import vba_project  # noqa: E402

DEFAULT_EVIDENCE = Path("test-results") / "excel-run"
WORKBOOK = Path("test-results") / "build" / "Workbook-development.xlsm"
FIXTURES = Path("tests") / "fixtures"
# References every VBA project in Excel has, besides those the build writes.
IMPLICIT_REFERENCES = ("Visual Basic For Applications", "Microsoft Excel 16.0 Object Library")
WINDOWS_KEY = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion"
# Windows 11 keeps "Windows 10" in ProductName; its builds start at 22000.
WINDOWS_11_BUILD = 22000


class RunError(Exception):
    """The run could not be recorded."""


def built_references() -> list[str]:
    """The references of the project build_workbook.py writes, by their names in Excel."""
    return [*IMPLICIT_REFERENCES,
            *(libid.rsplit("#", 1)[1] for _, libid in vba_project.REFERENCES)]


def describe_windows(values: Mapping[str, object]) -> str:
    """The text winver shows, from the CurrentVersion registry values."""
    build = str(values.get("CurrentBuild", "")).strip()
    product = str(values.get("ProductName", "Windows")).strip()
    if build.isdigit() and int(build) >= WINDOWS_11_BUILD:
        product = product.replace("Windows 10", "Windows 11")
    version = str(values.get("DisplayVersion", "")).strip()
    patch = values.get("UBR")
    full = f"{build}.{patch}" if build and patch is not None else build
    return ", ".join(part for part in (product, f"version {version}" if version else "",
                                       f"OS build {full}" if full else "") if part)


def windows_version() -> str:
    import winreg  # Windows only

    values: dict[str, object] = {}
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, WINDOWS_KEY) as key:  # type: ignore[attr-defined]
        for name in ("ProductName", "DisplayVersion", "CurrentBuild", "UBR"):
            try:
                values[name] = winreg.QueryValueEx(key, name)[0]  # type: ignore[attr-defined]
            except OSError:
                continue
    return describe_windows(values)


def host_record(windows: str, references: Sequence[str]) -> str:
    return (f"windows={windows}\n"
            f"references={'; '.join(references)}\n"
            f"compile={check_excel_run.COMPILE_ON_DEMAND}\n")


def prepare(folder: Path, root: Path) -> None:
    """An empty evidence folder. Only a folder under test-results/ is emptied."""
    scratch = (root / "test-results").resolve()
    folder = folder.resolve()
    if folder.exists() and any(folder.iterdir()):
        if scratch not in folder.parents:
            raise RunError(f"{folder} is not empty, and only a folder under {scratch} is "
                           "emptied: choose an empty or new folder")
        shutil.rmtree(folder)
    folder.mkdir(parents=True, exist_ok=True)


def quoted(text: str) -> str:
    """A PowerShell single-quoted literal."""
    return "'" + text.replace("'", "''") + "'"


def powershell_script(workbook: Path, folder: Path, fixtures: Path,
                      entry_points: Sequence[str], probe: str, enable_timeout: int) -> str:
    """Open the workbook read-only, wait until its macros run, run each certification
    entry point (``Module.Procedure``) in order, close without saving, quit. Every
    entry point runs even when an earlier one fails, so each suite leaves its logs; the
    first failure is reported. It never changes a macro setting: when the workbook opens
    with macros disabled, a person enables them with Enable Content (HB-02)."""
    calls = "\n".join(
        f"""        try {{
            [void]$excel.Run($prefix + {quoted(entry)}, $evidence, $fixtures)
        }} catch {{
            if (-not $failure) {{ $failure = {quoted(entry + ": ")} + $_.Exception.Message }}
        }}""" for entry in entry_points)
    return f"""$ErrorActionPreference = 'Stop'
# Progress records would reach the caller as CLIXML on standard error.
$ProgressPreference = 'SilentlyContinue'
$workbookPath = {quoted(str(workbook))}
$evidence = {quoted(str(folder))}
$fixtures = {quoted(str(fixtures))}
$excel = $null
$book = $null
$failure = $null
try {{
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    # UpdateLinks 0, ReadOnly: the workbook on disk stays the build.
    $book = $excel.Workbooks.Open($workbookPath, 0, $true)
    [void]$book.Activate()
    $prefix = "'" + $book.Name.Replace("'", "''") + "'!"
    # Wait until the workbook's macros run: at once where Excel enables them, or
    # after a person clicks Enable Content. The probe only reads harness state.
    $deadline = (Get-Date).AddSeconds({enable_timeout})
    $enabled = $false
    while (-not $enabled) {{
        try {{
            [void]$excel.Run($prefix + {quoted(probe)})
            $enabled = $true
        }} catch {{
            if ((Get-Date) -gt $deadline) {{ break }}
            Start-Sleep -Seconds 2
        }}
    }}
    if (-not $enabled) {{
        $failure = "NOT-ENABLED"
    }} else {{
{calls}
    }}
}} catch {{
    $failure = $_.Exception.Message
}} finally {{
    if ($null -ne $book) {{ $book.Close($false) }}
    if ($null -ne $excel) {{
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}}
if ($failure -eq "NOT-ENABLED") {{
    Write-Output "EXCEL-RUN=NOT-ENABLED"
    exit 4
}}
if ($failure) {{
    Write-Output "EXCEL-RUN=ERROR; $failure"
    exit 3
}}
Write-Output "EXCEL-RUN=DONE"
"""


def entry_points(policy: check_excel_run.Policy) -> list[str]:
    """The certification entry point of the example suite, then of each project suite,
    as ``Module.Procedure``."""
    return [f"{Path(policy.suite_source).stem}.{policy.entry_point}",
            *(f"{Path(item.source).stem}.{item.entry_point}" for item in policy.project_suites)]


def encoded(script: str) -> str:
    """The script as -EncodedCommand takes it: Base64 of UTF-16LE."""
    return base64.b64encode(script.encode("utf-16-le")).decode("ascii")


def run_excel(script: str, timeout: int) -> tuple[int, str]:
    command = ["powershell.exe", "-NoProfile", "-NonInteractive", "-EncodedCommand",
               encoded(script)]
    try:
        completed = subprocess.run(command, capture_output=True, text=True, timeout=timeout,
                                   check=False)
    except subprocess.TimeoutExpired as error:
        raise RunError(f"Excel did not finish within {timeout} s; close it, check for a "
                       "dialog, then run TEST_Harness.ResetTests if it reports recovery") \
            from error
    return completed.returncode, (completed.stdout + completed.stderr).strip()


# The exit code of the PowerShell script when the macros were never enabled.
NOT_ENABLED = 4
ENABLE_HINT = ("Excel opens the workbook now. If it shows a yellow bar SECURITY WARNING, "
               "Macros have been disabled, click Enable Content (or File > Info > Enable "
               "Content > Enable All Content). Change nothing else, and leave Excel alone "
               "until it closes.")


def record(root: Path, folder: Path, timeout: int, enable_timeout: int = 300) -> int:
    policy = check_excel_run.working_policy(root)
    workbook = root / WORKBOOK
    built = build_workbook.build(root, check_excel_run.REBUILT_CONFIGURATION, workbook)
    print(f"built {workbook}: commit {built.commit}")
    if not built.clean:
        raise RunError("the checkout has local changes, so the build is not evidence-grade: "
                       "commit or remove them, or use a clean clone")
    prepare(folder, root)
    probe = f"{Path(policy.harness).stem}.LastResult"
    script = powershell_script(workbook, folder.resolve(), (root / FIXTURES).resolve(),
                               entry_points(policy), probe, enable_timeout)
    print(f"running {', '.join(entry_points(policy))} in Excel ...")
    print(ENABLE_HINT)
    code, output = run_excel(script, timeout)
    if output:
        print(output)
    if code == NOT_ENABLED:
        raise RunError(f"the workbook's macros were not enabled within {enable_timeout} s, so "
                       "nothing ran: run the command again and click Enable Content in Excel")
    if code not in (0, 3):
        raise RunError(f"PowerShell or Excel failed (exit {code}); see the output above")
    (folder / policy.host).write_text(host_record(windows_version(), built_references()),
                                      encoding="utf-8")
    evidence = check_excel_run.check(root, built.commit, folder, workbook, require_binding=True)
    report = check_excel_run.report(built.commit, evidence)
    print(report)
    saved = folder.parent / f"{folder.name}-report.txt"
    saved.write_text(report + "\n", encoding="utf-8")
    print(f"report saved to {saved}")
    return 1 if evidence.findings else 0


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--evidence", type=Path, default=None,
                        help=f"the evidence folder (default: {DEFAULT_EVIDENCE.as_posix()})")
    parser.add_argument("--timeout", type=int, default=600,
                        help="seconds to wait for Excel in all (default: 600)")
    parser.add_argument("--enable-timeout", type=int, default=300,
                        help="seconds to wait for Enable Content (default: 300)")
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    if args.enable_timeout >= args.timeout:
        parser.error("--enable-timeout must be shorter than --timeout")
    if os.name != "nt":
        print("excel-run: SKIP (needs Windows with Excel; this records no evidence)")
        return 2
    root = args.root.resolve()
    folder = args.evidence if args.evidence is not None else root / DEFAULT_EVIDENCE
    try:
        return record(root, folder, args.timeout, args.enable_timeout)
    except (RunError, build_workbook.BuildError, check_excel_run.PolicyError, OSError) as error:
        print(f"FAIL: {error}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
