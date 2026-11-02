#!/usr/bin/env python3
"""Run the harness certification in headless LibreOffice as an early warning.

Usage: python3 tools/devbuild/lo_smoke.py [WORKBOOK]
Without WORKBOOK it builds the development workbook first.

It opens the workbook in LibreOffice with VBA compatibility and runs
TEST_Examples.CertifyExamples into a temporary folder. Then it prints a JSON
summary: the modules LibreOffice loaded, the error the certification raised,
each run log's RESULT and failures, the certification log, and the failures
that LibreOffice is not known to cause. It exits 0 when there is none, 1 when
there is one or the certification wrote no log, and 2 when LibreOffice or
its Python UNO bridge (python3-uno) is missing. Run it with the Python that
provides the bridge.

LibreOffice is not a supported host, and this is never Excel evidence. It
only catches some runtime defects before a run on the supported host. It runs
valid VBA differently in known ways:
- vbObjectError in a constant expression stops the whole project from
  compiling, so the harness writes its error numbers as literals;
- Debug.Print raises when it runs headless, so the harness prints through
  TEST_Harness.Echo;
- Application.Build does not exist, so ENVIRONMENT has an empty build;
- after a call that raised under On Error Resume Next, Err.Number reads 0,
  so the assertions on raised errors fail (KNOWN_FAILURES below).
"""

from __future__ import annotations

import importlib
import importlib.util
import json
import shutil
import socket
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))
LOGS = ("examples.log", "injected.log", "skipped.log", "certification.log")
# Assertions of the example suite that fail in LibreOffice only, because
# Err.Number reads 0 after a call that raised under On Error Resume Next.
KNOWN_FAILURES = frozenset({
    "raised-error/negative-input",
    "cleanup-keeps-primary-error/primary-error",
    "cleanup-partial-capture/capture-error",
    "cleanup-change-failure/change-error",
    "cleanup-each-restoration-failure/calculation-cleanup-error",
    "cleanup-each-restoration-failure/screen-cleanup-error",
    "cleanup-only-failure/cleanup-error",
})
# The failure each deliberate failure mode is designed to produce.
DESIGNED = {"injected.log": {"exact-double/half-plus-quarter"},
            "skipped.log": {"harness.completeness"}}


def free_port() -> int:
    with socket.socket() as probe:
        probe.bind(("127.0.0.1", 0))
        return int(probe.getsockname()[1])


def property_value(name: str, value: object) -> Any:
    item = importlib.import_module("uno").createUnoStruct("com.sun.star.beans.PropertyValue")
    item.Name = name
    item.Value = value
    return item


def connect(port: int) -> Any:
    local = importlib.import_module("uno").getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(120):
        try:
            return resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;"
                                    "StarOffice.ComponentContext")
        except Exception:  # noqa: BLE001 - the office is still starting
            time.sleep(0.5)
    raise RuntimeError("LibreOffice did not accept a connection")


def enable_vba(context: Any) -> None:
    provider = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.configuration.ConfigurationProvider", context)
    access = provider.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationUpdateAccess",
        (property_value("nodepath", "/org.openoffice.Office.Calc/Filter/Import/VBA"),))
    access.setPropertyValue("Load", True)
    access.setPropertyValue("Executable", True)
    access.commitChanges()


def certify(context: Any, workbook: Path, folder: Path) -> dict[str, object]:
    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",
                                                              context)
    document = desktop.loadComponentFromURL(
        workbook.resolve().as_uri(), "_blank", 0,
        (property_value("Hidden", True), property_value("MacroExecutionMode", 4)))
    summary: dict[str, object] = {"workbook": str(workbook)}
    try:
        summary["modules"] = sorted(document.BasicLibraries.getByName("VBAProject")
                                    .getElementNames())
        script = document.getScriptProvider().getScript(
            "vnd.sun.star.script:VBAProject.TEST_Examples.CertifyExamples"
            "?language=Basic&location=document")
        try:
            script.invoke((str(folder),), (), ())
            summary["error"] = None
        except Exception as error:  # noqa: BLE001 - reported, not raised
            summary["error"] = str(error)
    finally:
        document.close(True)
        try:
            desktop.terminate()
        except Exception:  # noqa: BLE001 - the office may already be gone
            pass
    return summary


def read_logs(folder: Path) -> dict[str, object]:
    found: dict[str, object] = {}
    for name in LOGS:
        matches = sorted(folder.rglob(name)) + sorted(folder.parent.glob(f"*{name}"))
        if not matches:
            found[name] = None
            continue
        lines = matches[0].read_text(encoding="utf-8", errors="replace").splitlines()
        if name == "certification.log":
            found[name] = lines
        else:
            found[name] = {"result": lines[-1] if lines else "",
                           "failures": [line for line in lines
                                        if line.startswith(("FAILURE=", "ERROR="))]}
    return found


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    soffice = shutil.which("soffice")
    if soffice is None or importlib.util.find_spec("uno") is None:
        print("lo-smoke: SKIP (needs LibreOffice and python3-uno; this is never Excel evidence)")
        return 2
    with tempfile.TemporaryDirectory(prefix="lo-smoke-") as temporary:
        base = Path(temporary)
        if args:
            workbook = Path(args[0])
        else:
            import build_workbook
            workbook = base / "Workbook-development.xlsm"
            build_workbook.build(ROOT, "development", workbook)
        folder = base / "evidence"
        folder.mkdir()
        port = free_port()
        office = subprocess.Popen(
            [soffice, "--headless", "--invisible", "--norestore", "--nologo",
             f"-env:UserInstallation={(base / 'profile').as_uri()}",
             f"--accept=socket,host=127.0.0.1,port={port};urp;"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        try:
            context = connect(port)
            enable_vba(context)
            summary = certify(context, workbook, folder)
        finally:
            try:
                office.wait(timeout=60)
            except subprocess.TimeoutExpired:
                office.kill()
        summary["logs"] = read_logs(folder)
    summary["unexpected"] = unexpected(summary["logs"])  # type: ignore[arg-type]
    summary["note"] = "LibreOffice is not a supported host; this is never Excel evidence"
    print(json.dumps(summary, indent=1, default=str))
    return 1 if summary["unexpected"] else 0


def unexpected(logs: dict[str, Any]) -> list[str]:
    """Failures and errors that LibreOffice is not known to cause, and missing logs."""
    found = []
    for name, content in logs.items():
        if content is None:
            found.append(f"{name}: not written")
        elif isinstance(content, dict):
            for line in content["failures"]:
                key, _, rest = line.partition("=")
                identifier = rest.split(";", 1)[0]
                if key == "ERROR" or identifier not in KNOWN_FAILURES | DESIGNED.get(name, set()):
                    found.append(f"{name}: {line}")
    return found


if __name__ == "__main__":
    sys.exit(main())
