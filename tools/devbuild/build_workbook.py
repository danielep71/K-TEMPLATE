#!/usr/bin/env python3
"""Build a macro-enabled workbook from the source of one commit.

Usage: python tools/devbuild/build_workbook.py [--configuration NAME] [OUTPUT]
Default: the development configuration, written to
test-results/build/Workbook-development.xlsm, which Git ignores.

The workbook has the sheets, code names and defined names of
tools/devbuild/workbook.json, which tests/test_devbuild.py keeps equal to
docs/methodology/WORKBOOK_CONTRACT.md. Sheets with the Test role, and their
names, are built only in the development configuration. It carries the build stamp on the
reference sheet. Its VBA project is written from source by vba_project.py:
the standard modules (.bas) and class modules (.cls) of the configuration's
directories, plus a document module for ThisWorkbook and for each sheet. A
document module comes from src/workbook/<code name>.cls when that file
exists. Excel compiles the project from source on open, so the build needs
no Trust access to the VBA project object model.

Forms are not built: a .frm in a source directory stops the build, and forms
are imported by hand. For a given commit and working tree the output is
byte-identical, because the stamp carries the commit's committer time, not
the clock. A tracked change, or an untracked file in a source directory,
stamps the build as not evidence-grade. Standard library only.
"""

from __future__ import annotations

import argparse
import io
import json
import re
import subprocess
import sys
import uuid
import zipfile
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape, quoteattr

sys.path.insert(0, str(Path(__file__).resolve().parent))
from vba_project import (  # noqa: E402
    CLASS,
    DOCUMENT,
    STANDARD,
    Module,
    build_vba_project,
    document_module,
)

ROOT = Path(__file__).resolve().parents[2]
CONFIG = "tools/devbuild/workbook.json"
DOCUMENT_HOME = "src/workbook"
CLASS_HOME = "src/classes"
DEVELOPMENT = "development"
# Sheet roles of the workbook contract; Test sheets exist only in a development build.
ROLES = ("Reference", "Input", "Parameter", "Output", "Diagnostics", "Test")
TEST_ROLE = "Test"
FORM_HOME = "src/forms"
VBA_ENCODING = "cp1252"
NOT_EVIDENCE = " (working tree with changes: NOT evidence-grade)"
ZIP_TIME = (1980, 1, 1, 0, 0, 0)
CELL = re.compile(r"^([A-Z]{1,3})([1-9][0-9]{0,6})$")
NAME = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
CODE_NAME = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,30}$")
HEADER = re.compile(r"^(VERSION \d+\.\d+( CLASS)?|(?i:begin)(\s.*)?|(?i:end)|\s+\S.*)$")
TITLE = "Built from source by tools/devbuild/build_workbook.py. Never commit this workbook."

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships"
WORKSHEET_TYPE = f"{RELATIONSHIPS}/worksheet"
STYLES_TYPE = f"{RELATIONSHIPS}/styles"
VBA_TYPE = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


class BuildError(Exception):
    """The configuration or the source cannot be built."""


@dataclass(frozen=True)
class Sheet:
    tab: str
    code_name: str
    role: str


@dataclass(frozen=True)
class DefinedName:
    name: str
    sheet: str
    cell: str
    label: str


@dataclass(frozen=True)
class Layout:
    workbook_code_name: str
    sheets: tuple[Sheet, ...]
    names: tuple[DefinedName, ...]
    stamp: dict[str, str]
    configurations: dict[str, tuple[str, ...]]

    def sheet(self, code_name: str) -> Sheet:
        return next(sheet for sheet in self.sheets if sheet.code_name == code_name)


@dataclass(frozen=True)
class Build:
    commit: str
    time: str
    clean: bool
    configuration: str
    components: tuple[Path, ...]


def load_layout(root: Path) -> Layout:
    """The build configuration, validated; BuildError names the first problem."""
    path = root / CONFIG
    try:
        data: dict[str, Any] = json.loads(path.read_text(encoding="utf-8"))
        layout = Layout(
            data["workbook_code_name"],
            tuple(Sheet(item["tab"], item["code_name"], item["role"]) for item in data["sheets"]),
            tuple(DefinedName(item["name"], item["sheet"], item["cell"], item.get("label", ""))
                  for item in data["names"]),
            dict(data["stamp"]),
            {key: tuple(value) for key, value in data["configurations"].items()})
    except (OSError, ValueError, KeyError, TypeError) as error:
        raise BuildError(f"{CONFIG}: {error!r}") from error
    problems = layout_problems(layout)
    if problems:
        raise BuildError(f"{CONFIG}: {problems[0]}")
    return layout


def layout_problems(layout: Layout) -> list[str]:
    problems = []
    codes = [layout.workbook_code_name, *(sheet.code_name for sheet in layout.sheets)]
    tabs = [sheet.tab.casefold() for sheet in layout.sheets]
    if not layout.sheets:
        problems.append("no sheet")
    if len({code.casefold() for code in codes}) != len(codes):
        problems.append("a code name is used twice")
    if len(set(tabs)) != len(tabs):
        problems.append("a tab name is used twice")
    problems += [f"code name {code!r} is not a VBA identifier of at most 31 characters"
                 for code in codes if not CODE_NAME.match(code)]
    problems += [f"tab name {sheet.tab!r} is empty, longer than 31 characters or uses []:*?/\\"
                 for sheet in layout.sheets
                 if not sheet.tab or len(sheet.tab) > 31 or re.search(r"[\[\]:*?/\\]", sheet.tab)]
    names = [item.name.casefold() for item in layout.names]
    if len(set(names)) != len(names):
        problems.append("a defined name is used twice")
    for item in layout.names:
        if not NAME.match(item.name) or CELL.match(item.name.upper()):
            problems.append(f"defined name {item.name!r} is not a valid name")
        if item.sheet not in codes[1:]:
            problems.append(f"defined name {item.name} refers to unknown sheet {item.sheet}")
        if not CELL.match(item.cell):
            problems.append(f"defined name {item.name} refers to {item.cell!r}, not one cell")
    problems += [f"sheet {sheet.code_name} has role {sheet.role!r}, not one of "
                 f"{', '.join(ROLES)}" for sheet in layout.sheets if sheet.role not in ROLES]
    tests = {sheet.code_name for sheet in layout.sheets if sheet.role == TEST_ROLE}
    for role in ("commit", "time", "configuration"):
        stamp = next((item for item in layout.names if item.name == layout.stamp.get(role)), None)
        if stamp is None:
            problems.append(f"stamp {role} is not a defined name")
        elif stamp.sheet in tests:
            problems.append(f"stamp {role} is on a {TEST_ROLE} sheet, which a production "
                            "build leaves out")
    if DEVELOPMENT not in layout.configurations:
        problems.append("no development configuration")
    return problems


def git(root: Path, *args: str) -> str:
    completed = subprocess.run(["git", "-C", str(root), *args], capture_output=True, text=True,
                               check=False)
    if completed.returncode:
        raise BuildError(f"git {' '.join(args)}: {completed.stderr.strip()}")
    return completed.stdout.strip()


def source_directories(layout: Layout, configuration: str) -> tuple[str, ...]:
    if configuration not in layout.configurations:
        raise BuildError(f"unknown configuration {configuration!r}; {CONFIG} defines "
                         f"{', '.join(sorted(layout.configurations))}")
    return layout.configurations[configuration]


def components(root: Path, layout: Layout, configuration: str) -> list[Path]:
    """The standard and class modules the configuration embeds, in file name order."""
    found: list[Path] = []
    for directory in (*source_directories(layout, configuration), FORM_HOME):
        forms = sorted(path.name for path in (root / directory).glob("*.frm")) \
            if (root / directory).is_dir() else []
        if forms:
            raise BuildError(f"{directory}: forms are not built ({', '.join(forms)}); import "
                             "them by hand as INSTALLATION.md describes")
    for directory in source_directories(layout, configuration):
        folder = root / directory
        if not folder.is_dir():
            continue
        found += sorted(folder.glob("*.bas"))
        if directory == CLASS_HOME:
            found += sorted(folder.glob("*.cls"))
    documents = {code.casefold(): code for code in
                 [layout.workbook_code_name, *(sheet.code_name for sheet in layout.sheets)]}
    exported = sorted((root / DOCUMENT_HOME).glob("*.cls")) \
        if (root / DOCUMENT_HOME).is_dir() else []
    strays = [path.name for path in exported if path.stem.casefold() not in documents]
    if strays:
        raise BuildError(f"{DOCUMENT_HOME}: {', '.join(strays)} matches no sheet of {CONFIG}")
    # The file is found by its exact name, so a name that differs only in case
    # would be missed on a case-sensitive file system and replaced by an empty
    # document module.
    miscased = [f"{path.name} (code name {documents[path.stem.casefold()]})"
                for path in exported if path.stem != documents[path.stem.casefold()]]
    if miscased:
        raise BuildError(f"{DOCUMENT_HOME}: {', '.join(miscased)} differs from its code name "
                         "in case; rename it to the code name")
    return found


def module_text(path: Path) -> str:
    """A component's source as the VBE stores it: CRLF, without the VERSION and
    BEGIN ... END lines an exported class writes before its attributes."""
    lines = path.read_bytes().decode(VBA_ENCODING).replace("\r\n", "\n").split("\n")
    start = next((index for index, line in enumerate(lines) if line.startswith("Attribute ")), 0)
    if any(not HEADER.match(line) for line in lines[:start]):
        start = 0
    text = "\r\n".join(lines[start:])
    return text if text.endswith("\r\n") else text + "\r\n"


def vba_modules(root: Path, layout: Layout, paths: Sequence[Path]) -> list[Module]:
    modules = []
    for code in [layout.workbook_code_name, *(sheet.code_name for sheet in layout.sheets)]:
        path = root / DOCUMENT_HOME / f"{code}.cls"
        modules.append(Module(code, module_text(path), kind=DOCUMENT) if path.is_file()
                       else document_module(code, code == layout.workbook_code_name))
    for path in paths:
        modules.append(Module(path.stem, module_text(path),
                              kind=CLASS if path.suffix.lower() == ".cls" else STANDARD))
    return modules


def is_clean(root: Path, layout: Layout, configuration: str) -> bool:
    """No tracked change, and no untracked file where the build reads source."""
    directories = [d for d in {*source_directories(layout, configuration), DOCUMENT_HOME}
                   if (root / d).is_dir()]
    tracked = git(root, "status", "--porcelain", "--untracked-files=no")
    untracked = git(root, "status", "--porcelain", "--untracked-files=all", "--",
                    *sorted(directories)) if directories else ""
    return not tracked and not untracked


def column_number(letters: str) -> int:
    number = 0
    for letter in letters:
        number = number * 26 + ord(letter) - 64
    return number


def column_letters(number: int) -> str:
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def quoted_tab(tab: str) -> str:
    return "'" + tab.replace("'", "''") + "'"


def sheet_cells(layout: Layout, build: Build, code_name: str) -> dict[tuple[int, int], str]:
    """The text cells of one sheet, by (row, column): stamp values and name labels."""
    stamp = {layout.stamp["commit"]: build.commit + ("" if build.clean else NOT_EVIDENCE),
             layout.stamp["time"]: build.time,
             layout.stamp["configuration"]: build.configuration}
    cells: dict[tuple[int, int], str] = {}
    stamp_sheet = next(item.sheet for item in layout.names
                       if item.name == layout.stamp["commit"])
    if code_name == stamp_sheet:
        cells[(1, 1)] = TITLE
    for item in layout.names:
        if item.sheet != code_name:
            continue
        match = CELL.match(item.cell)
        assert match is not None  # layout_problems() checked every cell
        row, column = int(match.group(2)), column_number(match.group(1))
        if item.label and column > 1:
            cells[(row, column - 1)] = item.label
        if item.name in stamp:
            cells[(row, column)] = stamp[item.name]
    return cells


def worksheet_xml(code_name: str, cells: dict[tuple[int, int], str]) -> str:
    rows: dict[int, list[tuple[int, str]]] = {}
    for (row, column), text in cells.items():
        rows.setdefault(row, []).append((column, text))
    body = "".join(
        f'<row r="{row}">' + "".join(
            f'<c r="{column_letters(column)}{row}" t="inlineStr"><is>'
            f'<t xml:space="preserve">{escape(text)}</t></is></c>'
            for column, text in sorted(rows[row])) + "</row>"
        for row in sorted(rows))
    return (f'{XML_HEADER}<worksheet xmlns="{MAIN}" xmlns:r="{RELATIONSHIPS}">'
            f"<sheetPr codeName={quoteattr(code_name)}/>"
            f"<sheetData>{body}</sheetData></worksheet>")


def workbook_xml(layout: Layout) -> str:
    sheets = "".join(f"<sheet name={quoteattr(sheet.tab)} sheetId=\"{index}\" "
                     f'r:id="rIdSheet{index}"/>'
                     for index, sheet in enumerate(layout.sheets, 1))
    names = ""
    for item in sorted(layout.names, key=lambda entry: entry.name.casefold()):
        match = CELL.match(item.cell)
        assert match is not None
        reference = f"{quoted_tab(layout.sheet(item.sheet).tab)}!${match.group(1)}${match.group(2)}"
        names += f"<definedName name={quoteattr(item.name)}>{escape(reference)}</definedName>"
    return (f'{XML_HEADER}<workbook xmlns="{MAIN}" xmlns:r="{RELATIONSHIPS}">'
            f"<workbookPr codeName={quoteattr(layout.workbook_code_name)}/>"
            f"<bookViews><workbookView/></bookViews><sheets>{sheets}</sheets>"
            f"<definedNames>{names}</definedNames></workbook>")


STYLES = (f'{XML_HEADER}<styleSheet xmlns="{MAIN}">'
          '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
          '<fills count="2"><fill><patternFill patternType="none"/></fill>'
          '<fill><patternFill patternType="gray125"/></fill></fills>'
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
          '</cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" '
          'borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" '
          'xfId="0" builtinId="0"/></cellStyles></styleSheet>')


def package(layout: Layout, build: Build, vba_bin: bytes) -> bytes:
    """The .xlsm package: the same bytes for the same input."""
    count = len(layout.sheets)
    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for index in range(1, count + 1))
    parts = {
        "[Content_Types].xml": (
            f'{XML_HEADER}<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
            'content-types"><Default Extension="rels" ContentType="application/'
            'vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" '
            'ContentType="application/xml"/><Default Extension="bin" '
            'ContentType="application/vnd.ms-office.vbaProject"/><Override '
            'PartName="/xl/workbook.xml" ContentType="application/'
            'vnd.ms-excel.sheet.macroEnabled.main+xml"/><Override PartName="/xl/styles.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.'
            f'styles+xml"/>{overrides}</Types>'),
        "_rels/.rels": (
            f'{XML_HEADER}<Relationships xmlns="{PACKAGE_RELATIONSHIPS}"><Relationship '
            f'Id="rId1" Type="{RELATIONSHIPS}/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>"),
        "xl/workbook.xml": workbook_xml(layout),
        "xl/_rels/workbook.xml.rels": (
            f'{XML_HEADER}<Relationships xmlns="{PACKAGE_RELATIONSHIPS}">' + "".join(
                f'<Relationship Id="rIdSheet{index}" Type="{WORKSHEET_TYPE}" '
                f'Target="worksheets/sheet{index}.xml"/>' for index in range(1, count + 1))
            + f'<Relationship Id="rIdStyles" Type="{STYLES_TYPE}" Target="styles.xml"/>'
            f'<Relationship Id="rIdVBA" Type="{VBA_TYPE}" Target="vbaProject.bin"/>'
            "</Relationships>"),
        "xl/styles.xml": STYLES,
    }
    for index, sheet in enumerate(layout.sheets, 1):
        parts[f"xl/worksheets/sheet{index}.xml"] = worksheet_xml(
            sheet.code_name, sheet_cells(layout, build, sheet.code_name))
    blobs: dict[str, bytes] = {name: text.encode("utf-8") for name, text in parts.items()}
    blobs["xl/vbaProject.bin"] = vba_bin
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, data in blobs.items():
            info = zipfile.ZipInfo(name, date_time=ZIP_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            # ZipInfo takes the creating system from sys.platform; fix it to Unix, the
            # system the permission bits below belong to, so every platform agrees.
            info.create_system = 3
            info.external_attr = 0o644 << 16
            archive.writestr(info, data)
    return buffer.getvalue()


def for_configuration(layout: Layout, configuration: str) -> Layout:
    """The layout a configuration builds: Test sheets and their names only in
    the development configuration, as the workbook contract requires."""
    if configuration == DEVELOPMENT:
        return layout
    sheets = tuple(sheet for sheet in layout.sheets if sheet.role != TEST_ROLE)
    kept = {sheet.code_name for sheet in sheets}
    return Layout(layout.workbook_code_name, sheets,
                  tuple(item for item in layout.names if item.sheet in kept),
                  layout.stamp, layout.configurations)


def build(root: Path, configuration: str, target: Path) -> Build:
    """Write the workbook to ``target`` and describe what it holds."""
    # Document modules are checked against every sheet, Test sheets included,
    # before the configuration drops what it does not build.
    complete = load_layout(root)
    paths = components(root, complete, configuration)
    layout = for_configuration(complete, configuration)
    commit = git(root, "rev-parse", "HEAD")
    result = Build(commit, git(root, "show", "-s", "--format=%cI", "HEAD"),
                   is_clean(root, layout, configuration), configuration, tuple(paths))
    project_id = "{" + str(uuid.uuid5(uuid.NAMESPACE_URL,
                                      f"workbook-build:{commit}:{configuration}")).upper() + "}"
    vba_bin = build_vba_project(vba_modules(root, layout, paths), project_id)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(package(layout, result, vba_bin))
    return result


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("output", nargs="?", type=Path,
                        help="the .xlsm to write (default: test-results/build/"
                             "Workbook-<configuration>.xlsm)")
    parser.add_argument("--configuration", default="development",
                        help=f"a configuration of {CONFIG} (default: development)")
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    root = args.root.resolve()
    target = args.output or root / "test-results" / "build" / \
        f"Workbook-{args.configuration}.xlsm"
    try:
        result = build(root, args.configuration, target)
    except (BuildError, ValueError) as error:
        print(f"FAIL: {error}")
        return 1
    names = ", ".join(path.stem for path in result.components) or \
        "no VBA component: the workbook holds only its document modules"
    grade = "" if result.clean else NOT_EVIDENCE
    print(f"built {target}: {result.configuration}, commit {result.commit}{grade}")
    print(f"components: {names}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
