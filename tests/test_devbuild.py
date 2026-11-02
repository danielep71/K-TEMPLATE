"""Read back the built workbook and its VBA project with an independent reader.

tools/devbuild/vba_project.py writes vbaProject.bin from the [MS-OVBA] and
[MS-CFB] specifications. These tests decompress and parse its output with a
separate, minimal implementation of the same specifications, so a writer
defect fails here without Excel. They also check that
tools/devbuild/build_workbook.py builds the sheets, code names, defined names
and stamp of tools/devbuild/workbook.json, that this configuration matches
docs/methodology/WORKBOOK_CONTRACT.md, and that a rebuild is byte-identical.
"""

from __future__ import annotations

import contextlib
import io
import random
import shutil
import struct
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools" / "devbuild"))
sys.path.insert(0, str(ROOT / "tools"))
import build_workbook  # noqa: E402
import check_excel_run  # noqa: E402
import export_manifest  # noqa: E402
import vba_project  # noqa: E402

ENDOFCHAIN = 0xFFFFFFFE


def decompress(data: bytes) -> bytes:
    """[MS-OVBA] 2.4.1.3.1: decompress a CompressedContainer."""
    if not data or data[0] != 0x01:
        raise ValueError("missing signature byte")
    out = bytearray()
    pos = 1
    while pos < len(data):
        start = pos
        header = struct.unpack_from("<H", data, pos)[0]
        pos += 2
        if (header >> 12) & 0x7 != 0b011:
            raise ValueError("bad chunk signature")
        end = min(start + (header & 0x0FFF) + 3, len(data))
        if not header & 0x8000:
            out += data[pos:pos + 4096]
            pos += 4096
            continue
        chunk_start = len(out)
        while pos < end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= end:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[pos])
                    pos += 1
                    continue
                token = struct.unpack_from("<H", data, pos)[0]
                pos += 2
                difference = len(out) - chunk_start
                bit_count = 4
                while (1 << bit_count) < difference:
                    bit_count += 1
                length_mask = 0xFFFF >> bit_count
                length = (token & length_mask) + 3
                offset = (token >> (16 - bit_count)) + 1
                source = len(out) - offset
                for i in range(length):
                    out.append(out[source + i])
    return bytes(out)


class CompoundFile:
    """[MS-CFB] version-3 reader: FAT, directory tree and mini stream."""

    def __init__(self, blob: bytes) -> None:
        if blob[:8] != bytes.fromhex("D0CF11E0A1B11AE1"):
            raise ValueError("not a compound file")
        self.blob = blob
        self.sector = 1 << struct.unpack_from("<H", blob, 30)[0]
        self.mini = 1 << struct.unpack_from("<H", blob, 32)[0]
        n_fat, dir_start = struct.unpack_from("<II", blob, 44)
        self.cutoff, minifat_start = struct.unpack_from("<II", blob, 56)
        difat = struct.unpack_from("<109I", blob, 76)[:n_fat]
        fat_bytes = b"".join(self._sector(s) for s in difat)
        self.fat = list(struct.unpack(f"<{len(fat_bytes) // 4}I", fat_bytes))
        directory = self._chain_bytes(dir_start)
        self.entries = [self._entry(directory[i:i + 128]) for i in range(0, len(directory), 128)]
        root = self.entries[0]
        self.mini_stream = self._chain_bytes(root["start"])[:root["size"]]
        minifat_bytes = self._chain_bytes(minifat_start) if minifat_start != ENDOFCHAIN else b""
        self.minifat = list(struct.unpack(f"<{len(minifat_bytes) // 4}I", minifat_bytes))

    def _sector(self, index: int) -> bytes:
        offset = self.sector * (index + 1)
        return self.blob[offset:offset + self.sector]

    def _chain(self, start: int, table: list[int]) -> list[int]:
        chain, seen = [], set()
        while start != ENDOFCHAIN:
            if start in seen:
                raise ValueError("cyclic chain")
            seen.add(start)
            chain.append(start)
            start = table[start]
        return chain

    def _chain_bytes(self, start: int) -> bytes:
        return b"".join(self._sector(s) for s in self._chain(start, self.fat))

    @staticmethod
    def _entry(raw: bytes) -> dict[str, Any]:
        name_length = struct.unpack_from("<H", raw, 64)[0]
        return {"name": raw[:max(0, name_length - 2)].decode("utf-16-le"), "kind": raw[66],
                "left": struct.unpack_from("<I", raw, 68)[0], "right": struct.unpack_from("<I", raw, 72)[0],
                "child": struct.unpack_from("<I", raw, 76)[0], "start": struct.unpack_from("<I", raw, 116)[0],
                "size": struct.unpack_from("<Q", raw, 120)[0]}

    def children(self, sid: int) -> dict[str, int]:
        found: dict[str, int] = {}
        stack = [self.entries[sid]["child"]]
        while stack:
            node = stack.pop()
            if node in (0xFFFFFFFF, None):
                continue
            entry = self.entries[node]
            found[entry["name"]] = node
            stack += [entry["left"], entry["right"]]
        return found

    def read(self, path: str) -> bytes:
        sid = 0
        for part in path.split("/"):
            sid = self.children(sid)[part]
        entry = self.entries[sid]
        if entry["size"] >= self.cutoff:
            return self._chain_bytes(entry["start"])[:entry["size"]]
        if entry["size"] == 0:
            return b""
        chunks = [self.mini_stream[s * self.mini:(s + 1) * self.mini] for s in self._chain(entry["start"], self.minifat)]
        return b"".join(chunks)[:entry["size"]]


def sample_modules() -> list[vba_project.Module]:
    body = "\r\n".join(f"' line {i}: some repeated text, repeated text, repeated text" for i in range(200))
    return [vba_project.document_module("ThisWorkbook", True),
            vba_project.document_module("shChecks", False),
            vba_project.Module("CORE_Small", 'Attribute VB_Name = "CORE_Small"\r\nOption Explicit\r\n'),
            vba_project.Module("TEST_Large", 'Attribute VB_Name = "TEST_Large"\r\n' + body + "\r\n")]


class Compression(unittest.TestCase):
    def test_specification_example_decompresses(self) -> None:
        # [MS-OVBA] 3.2.3, "Normal Compression" example.
        compressed = bytes.fromhex(
            "012FB000236161616263646582660070616768696A01380861"
            "6B6C00306D6E6F7006710270041072737475761077"
            "78797A00" "3C")
        self.assertEqual(decompress(compressed), b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa")

    def test_round_trip(self) -> None:
        rng = random.Random(8)
        samples = [b"", b"a", b"abc" * 3000, bytes(rng.randrange(256) for _ in range(9000)),
                   b"Option Explicit\r\n" * 700, bytes(range(256)) * 40]
        for data in samples:
            with self.subTest(length=len(data)):
                self.assertEqual(decompress(vba_project.compress(data)), data)


    def test_incompressible_last_chunk_names_the_module(self) -> None:
        # About 4 KB of random printable text in the last, short chunk does not fit
        # one compressed chunk; the build stops and names the module.
        rng = random.Random(97)
        noise = "".join(chr(rng.randrange(33, 127)) for _ in range(4000))
        module = vba_project.Module("CORE_Noise", f'Attribute VB_Name = "CORE_Noise"\r\n\' {noise}\r\n')
        with self.assertRaisesRegex(ValueError, r"module CORE_Noise: .*INSTALLATION.md"):
            vba_project.build_vba_project([module], ProjectStorage.project_id)


class Encryption(unittest.TestCase):
    def test_project_protection_fields_round_trip(self) -> None:
        rng = random.Random("{ID}")
        for data in (b"\x00", b"\xff", struct.pack("<I", 0), b"abcdefgh"):
            self.assertEqual(vba_project.decrypt(vba_project._encrypt(data, "{ID}", rng)), data)


class ProjectStorage(unittest.TestCase):
    project_id = "{0B9A3C1D-1111-4222-8333-444455556666}"

    def build(self) -> CompoundFile:
        return CompoundFile(vba_project.build_vba_project(sample_modules(), self.project_id))

    def test_every_module_source_reads_back_unchanged(self) -> None:
        cfb = self.build()
        for module in sample_modules():
            with self.subTest(module=module.name):
                self.assertEqual(decompress(cfb.read(f"VBA/{module.name}")).decode("cp1252"), module.source)

    def test_project_and_dir_streams_list_every_module(self) -> None:
        cfb = self.build()
        project = cfb.read("PROJECT").decode("cp1252")
        self.assertIn(f'ID="{self.project_id}"', project)
        self.assertIn("Document=ThisWorkbook/&H00000000", project)
        self.assertIn("Module=CORE_Small", project)
        directory = decompress(cfb.read("VBA/dir"))
        for module in sample_modules():
            self.assertIn(module.name.encode("cp1252"), directory)
        self.assertEqual(cfb.read("VBA/_VBA_PROJECT")[:2], b"\xcc\x61")

    def test_output_is_deterministic_for_a_project_id(self) -> None:
        first = vba_project.build_vba_project(sample_modules(), self.project_id)
        self.assertEqual(first, vba_project.build_vba_project(sample_modules(), self.project_id))
        self.assertNotEqual(first, vba_project.build_vba_project(sample_modules(), "{" + "1" * 8 + "-0000-0000-0000-" + "0" * 12 + "}"))


    def test_class_modules_are_declared_as_classes(self) -> None:
        module = vba_project.Module("CPair", 'Attribute VB_Name = "CPair"\r\nOption Explicit\r\n',
                                    kind=vba_project.CLASS)
        cfb = CompoundFile(vba_project.build_vba_project([module], self.project_id))
        self.assertIn("Class=CPair", cfb.read("PROJECT").decode("cp1252"))
        directory = decompress(cfb.read("VBA/dir"))
        # MODULETYPE 0x0022: a document, class or designer module.
        self.assertIn(struct.pack("<HI", 0x0022, 0), directory)
        with self.assertRaisesRegex(ValueError, "unknown kind"):
            vba_project.Module("X", "", kind="form")


def contract_rows(heading: str) -> list[list[str]]:
    """The body rows of the first table under a heading of the workbook contract,
    each cell without its Illustrative marker and code formatting."""
    text = (ROOT / "docs/methodology/WORKBOOK_CONTRACT.md").read_text(encoding="utf-8")
    section = text.split(f"## {heading}\n", 1)[1].split("\n## ", 1)[0]
    lines = [line for line in section.splitlines() if line.startswith("|")][2:]
    return [[cell.strip().replace("*Illustrative:*", "").strip().strip("`")
             for cell in line.strip("|").split("|")] for line in lines]


class Layout(unittest.TestCase):
    layout = build_workbook.load_layout(ROOT)

    def test_sheets_and_names_match_the_workbook_contract(self) -> None:
        sheets = [(row[1].strip("`"), row[2].strip("`")) for row in contract_rows("Sheets")]
        self.assertEqual([(sheet.tab, sheet.code_name) for sheet in self.layout.sheets], sheets)
        names = [(row[0], row[1]) for row in contract_rows("Defined names")]
        self.assertEqual([(item.name, item.sheet) for item in self.layout.names], names)

    def test_configurations_follow_the_structure_guide(self) -> None:
        structure = (ROOT / "docs/REPOSITORY_STRUCTURE.md").read_text(encoding="utf-8")
        for directories in self.layout.configurations.values():
            for directory in directories:
                with self.subTest(directory=directory):
                    self.assertIn(f"| `{directory}/` |", structure)
        production = set(self.layout.configurations["production"])
        self.assertFalse({"tests/modules", "examples/modules"} & production)
        self.assertLess(production, set(self.layout.configurations["development"]))

    def test_invalid_layouts_are_refused(self) -> None:
        base = self.layout
        cases = {
            "no sheet": build_workbook.Layout(base.workbook_code_name, (), base.names,
                                              base.stamp, base.configurations),
            "code name is used twice": build_workbook.Layout(
                "shReadme", base.sheets, base.names, base.stamp, base.configurations),
            "not one cell": build_workbook.Layout(
                base.workbook_code_name, base.sheets,
                (*base.names, build_workbook.DefinedName("cfg_x", "shInputs", "C5:C6", "")),
                base.stamp, base.configurations),
            "unknown sheet": build_workbook.Layout(
                base.workbook_code_name, base.sheets,
                (*base.names, build_workbook.DefinedName("cfg_x", "shNone", "A1", "")),
                base.stamp, base.configurations),
            "stamp commit": build_workbook.Layout(
                base.workbook_code_name, base.sheets, base.names,
                {**base.stamp, "commit": "missing"}, base.configurations),
            "not a valid name": build_workbook.Layout(
                base.workbook_code_name, base.sheets,
                (*base.names, build_workbook.DefinedName("AB12", "shInputs", "A1", "")),
                base.stamp, base.configurations),
        }
        for expected, layout in cases.items():
            with self.subTest(expected=expected):
                problems = build_workbook.layout_problems(layout)
                self.assertTrue(any(expected in problem for problem in problems), problems)
        self.assertEqual(build_workbook.layout_problems(base), [])


@contextlib.contextmanager
def source_tree(files: dict[str, bytes]) -> Any:
    """A directory with the build configuration and the given files."""
    with tempfile.TemporaryDirectory(prefix="devbuild-") as temporary:
        root = Path(temporary)
        (root / "tools/devbuild").mkdir(parents=True)
        shutil.copy(ROOT / build_workbook.CONFIG, root / build_workbook.CONFIG)
        for name, data in files.items():
            (root / name).parent.mkdir(parents=True, exist_ok=True)
            (root / name).write_bytes(data)
        yield root


class Components(unittest.TestCase):
    layout = build_workbook.load_layout(ROOT)

    def test_standard_and_class_modules_by_configuration(self) -> None:
        files = {"src/core/CORE_A.bas": b"x", "src/classes/CPair.cls": b"x",
                 "src/workbook/M_Run.bas": b"x", "tests/modules/TEST_A.bas": b"x",
                 "src/workbook/shInputs.cls": b"x", "docs/notes.bas": b"x"}
        with source_tree(files) as root:
            names = {configuration: [path.relative_to(root).as_posix() for path in
                                     build_workbook.components(root, self.layout, configuration)]
                     for configuration in ("production", "development")}
        self.assertEqual(names["production"],
                         ["src/core/CORE_A.bas", "src/classes/CPair.cls", "src/workbook/M_Run.bas"])
        self.assertEqual(names["development"], [*names["production"], "tests/modules/TEST_A.bas"])

    def test_forms_and_unknown_document_modules_stop_the_build(self) -> None:
        with source_tree({"src/forms/frmA.frm": b"x"}) as root, \
                self.assertRaisesRegex(build_workbook.BuildError, "forms are not built"):
            build_workbook.components(root, self.layout, "production")
        with source_tree({"src/workbook/shOther.cls": b"x"}) as root, \
                self.assertRaisesRegex(build_workbook.BuildError, "matches no sheet"):
            build_workbook.components(root, self.layout, "production")
        with self.assertRaisesRegex(build_workbook.BuildError, "unknown configuration"):
            build_workbook.components(ROOT, self.layout, "staging")

    def test_module_text_drops_the_class_header_only(self) -> None:
        exported = (b'VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\n'
                    b'Attribute VB_Name = "CPair"\r\nOption Explicit\r\n')
        standard = b'Attribute VB_Name = "CORE_A"\nOption Explicit'
        with source_tree({"src/classes/CPair.cls": exported,
                          "src/core/CORE_A.bas": standard}) as root:
            self.assertEqual(build_workbook.module_text(root / "src/classes/CPair.cls"),
                             'Attribute VB_Name = "CPair"\r\nOption Explicit\r\n')
            self.assertEqual(build_workbook.module_text(root / "src/core/CORE_A.bas"),
                             'Attribute VB_Name = "CORE_A"\r\nOption Explicit\r\n')


class Workbook(unittest.TestCase):
    layout = build_workbook.load_layout(ROOT)

    def build(self, configuration: str = "development") -> tuple[bytes, build_workbook.Build]:
        with tempfile.TemporaryDirectory(prefix="devbuild-") as temporary:
            target = Path(temporary) / "Workbook.xlsm"
            result = build_workbook.build(ROOT, configuration, target)
            return target.read_bytes(), result

    def test_a_rebuild_is_byte_identical(self) -> None:
        first, result = self.build()
        second, _ = self.build()
        self.assertEqual(first, second)
        self.assertEqual(result.commit, build_workbook.git(ROOT, "rev-parse", "HEAD"))

    def test_the_package_holds_the_contract_layout_and_the_stamp(self) -> None:
        data, result = self.build()
        with zipfile.ZipFile(io.BytesIO(data)) as package:
            self.assertIn("macroEnabled", package.read("[Content_Types].xml").decode())
            book = package.read("xl/workbook.xml").decode()
            self.assertIn(f'codeName="{self.layout.workbook_code_name}"', book)
            for index, sheet in enumerate(self.layout.sheets, 1):
                self.assertIn(f'<sheet name="{sheet.tab}" sheetId="{index}"', book)
                self.assertIn(f'<sheetPr codeName="{sheet.code_name}"/>',
                              package.read(f"xl/worksheets/sheet{index}.xml").decode())
            for item in self.layout.names:
                self.assertRegex(book, rf'<definedName name="{item.name}">[^<]*'
                                       rf'!\$[A-Z]+\$[0-9]+</definedName>')
            for info in package.infolist():
                self.assertEqual(info.date_time, build_workbook.ZIP_TIME)
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "Workbook.xlsm"
            path.write_bytes(data)
            stamp = check_excel_run.read_stamp(path, self.layout.stamp["commit"])
            self.assertEqual(stamp, result.commit + ("" if result.clean
                                                      else build_workbook.NOT_EVIDENCE))
            self.assertEqual(check_excel_run.read_stamp(path, self.layout.stamp["configuration"]),
                             "development")
            self.assertEqual(check_excel_run.read_stamp(path, self.layout.stamp["time"]),
                             result.time)

    def test_the_project_embeds_every_component_and_document_module(self) -> None:
        data, result = self.build()
        with zipfile.ZipFile(io.BytesIO(data)) as package:
            cfb = CompoundFile(package.read("xl/vbaProject.bin"))
        project = cfb.read("PROJECT").decode("cp1252")
        for code in [self.layout.workbook_code_name, *(s.code_name for s in self.layout.sheets)]:
            with self.subTest(document=code):
                self.assertIn(f"Document={code}/&H00000000", project)
                self.assertIn(f'Attribute VB_Name = "{code}"',
                              decompress(cfb.read(f"VBA/{code}")).decode("cp1252"))
        self.assertTrue(result.components)
        for path in result.components:
            with self.subTest(component=path.stem):
                self.assertEqual(decompress(cfb.read(f"VBA/{path.stem}")).decode("cp1252"),
                                 build_workbook.module_text(path))
                record = "Class" if path.suffix.lower() == ".cls" else "Module"
                self.assertIn(f"{record}={path.stem}", project)

    def test_production_holds_no_test_module(self) -> None:
        _, result = self.build("production")
        self.assertFalse([path for path in result.components if "tests" in path.parts])

    def test_main_reports_the_build(self) -> None:
        with tempfile.TemporaryDirectory() as temporary, \
                contextlib.redirect_stdout(io.StringIO()) as output:
            code = build_workbook.main([str(Path(temporary) / "W.xlsm"), "--root", str(ROOT),
                                        "--configuration", "production"])
        self.assertEqual(code, 0)
        selected = build_workbook.components(ROOT, self.layout, "production")
        expected = ", ".join(path.stem for path in selected) or "no VBA component"
        self.assertIn(f"components: {expected}", output.getvalue())
        with git_tree({}) as root, tempfile.TemporaryDirectory() as temporary, \
                contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(build_workbook.main([str(Path(temporary) / "W.xlsm"), "--root",
                                                  str(root), "--configuration", "production"]), 0)
        self.assertIn("no VBA component: the workbook holds only its document modules",
                      output.getvalue())
        with contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(build_workbook.main(["--configuration", "staging"]), 1)
        self.assertIn("unknown configuration", output.getvalue())


CLASS_SOURCE = (b'VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\n'
                b'Attribute VB_Name = "CPair"\r\nOption Explicit\r\n')
SHEET_SOURCE = (b'VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\n'
                b'Attribute VB_Name = "shInputs"\r\nOption Explicit\r\n'
                b'Private Sub Worksheet_Change(ByVal Target As Range)\r\nEnd Sub\r\n')


@contextlib.contextmanager
def git_tree(files: dict[str, bytes]) -> Any:
    """A committed Git repository with the build configuration and the given files."""
    with source_tree(files) as root:
        for command in (["init", "-q"], ["add", "-A"],
                        ["-c", "user.name=Fixture", "-c", "user.email=fixture@example.invalid",
                         "-c", "commit.gpgsign=false", "commit", "-qm", "Fixture"]):
            build_workbook.git(root, *command)
        yield root


class ProjectSource(unittest.TestCase):
    """Builds of a project with its own standard, class and document modules."""

    files = {"src/core/CORE_Demo.bas": b'Attribute VB_Name = "CORE_Demo"\r\nOption Explicit\r\n',
             "src/classes/CPair.cls": CLASS_SOURCE,
             "src/workbook/shInputs.cls": SHEET_SOURCE}

    def test_every_kind_of_module_is_embedded(self) -> None:
        with git_tree(self.files) as root, tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary) / "W.xlsm"
            result = build_workbook.build(root, "production", target)
            with zipfile.ZipFile(target) as package:
                cfb = CompoundFile(package.read("xl/vbaProject.bin"))
        self.assertTrue(result.clean)
        self.assertEqual([path.name for path in result.components], ["CORE_Demo.bas", "CPair.cls"])
        project = cfb.read("PROJECT").decode("cp1252")
        for record in ("Module=CORE_Demo", "Class=CPair", "Document=shInputs/&H00000000"):
            self.assertIn(record, project)
        sheet = decompress(cfb.read("VBA/shInputs")).decode("cp1252")
        self.assertIn("Private Sub Worksheet_Change", sheet)
        self.assertTrue(sheet.startswith('Attribute VB_Name = "shInputs"'))

    def test_a_document_module_in_the_wrong_case_stops_the_build(self) -> None:
        files = {"src/workbook/shINPUTS.cls": SHEET_SOURCE.replace(b"shInputs", b"shINPUTS")}
        with git_tree(files) as root, tempfile.TemporaryDirectory() as temporary, \
                self.assertRaisesRegex(build_workbook.BuildError,
                                       r"shINPUTS\.cls \(code name shInputs\) differs"):
            build_workbook.build(root, "production", Path(temporary) / "W.xlsm")


class Manifest(unittest.TestCase):
    def test_manifest_records_the_workbook_and_every_export(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            (folder / "exported").mkdir()
            (folder / "exported/A.bas").write_bytes(b"a")
            workbook = folder / "W.xlsm"
            workbook.write_bytes(b"workbook")
            path = export_manifest.write(workbook, folder)
            digest, files = check_excel_run.load_manifest(path)
            self.assertEqual(digest, export_manifest.sha256(workbook))
            self.assertEqual(set(files), {"A.bas"})
            (folder / "exported/A.bas").unlink()
            with contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(export_manifest.main([str(workbook), str(folder)]), 1)
            self.assertIn("holds no exported file", output.getvalue())


if __name__ == "__main__":
    unittest.main()
