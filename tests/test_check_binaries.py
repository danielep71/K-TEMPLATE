"""Tests for tools/check_binaries.py, in temporary Git repositories."""

from __future__ import annotations

import contextlib
import hashlib
import io
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_binaries  # noqa: E402

PNG = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR"


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class BinaryCheckTest(unittest.TestCase):
    def repository(self, files: dict[str, bytes], allowlist: object = None) -> list[str]:
        """Stage files in a new repository with this repository's .gitattributes; check it."""
        with tempfile.TemporaryDirectory(prefix="binaries-") as directory:
            root = Path(directory)
            subprocess.run(["git", "init", "-q"], cwd=root, check=True)
            files = {".gitattributes": (ROOT / ".gitattributes").read_bytes(), **files}
            if allowlist is not None:
                files[".github/binary-allowlist.json"] = json.dumps(allowlist).encode()
            for name, content in files.items():
                (root / name).parent.mkdir(parents=True, exist_ok=True)
                (root / name).write_bytes(content)
            subprocess.run(["git", "add", "-f", "."], cwd=root, check=True)
            return check_binaries.check(root)[0]

    def assertFinding(self, findings: list[str], text: str) -> None:
        self.assertTrue(any(text in finding for finding in findings), findings)

    def test_repository_passes(self) -> None:
        self.assertEqual(check_binaries.check(ROOT)[0], [])

    def test_text_files_pass(self) -> None:
        self.assertEqual(self.repository({"README.md": b"# Title\n", "data.csv": b"a,b\n1,2\n"}),
                         [])

    def test_binary_types_fail(self) -> None:
        for name in ("sample.xlsm", "book.xlsx", "docs/report.pdf", "archive.zip", "tool.exe",
                     "docs/assets/logo.png", "src/forms/frmRun.frx"):
            with self.subTest(name=name):
                self.assertFinding(self.repository({name: b"PK\x03\x04\0binary"}),
                                   f"{name}: binary file outside the approved paths")

    def test_approved_paths_pass(self) -> None:
        self.assertEqual(self.repository({
            "docs/methodology/sources/SRC-01_rule.pdf": b"%PDF-1.7\0",
            "docs/methodology/sources/templates/SRC-02_template.xlsx": b"PK\x03\x04\0",
            "src/workbook/Template.xlsx": b"PK\x03\x04\0",
            "src/forms/frmRun.frm": b"VERSION 5.00\r\n",
            "src/forms/frmRun.frx": b"\0\1\2",
            "tests/fixtures/a-case/crlf/positions.csv": b"id\r\nP1\r\n",
            "tests/fixtures/a-case/utf-16/positions.csv": "id\n".encode("utf-16"),
        }), [])

    def test_file_case_manifest_is_not_a_variant_file(self) -> None:
        self.assertFinding(self.repository({"tests/fixtures/a-case/case.json": b"{\0}"}),
                           "tests/fixtures/a-case/case.json: binary content")
        self.assertFinding(self.repository({"tests/fixtures/data.bin": b"\0"}),
                           "tests/fixtures/data.bin: binary content")

    def test_nul_byte_in_a_text_file_fails(self) -> None:
        self.assertFinding(self.repository({"notes.txt": b"text\0more"}),
                           "notes.txt: binary content in a file treated as text")
        self.assertFinding(self.repository({"payload.bin": b"\0\1\2"}),
                           "payload.bin: binary content")

    def test_allowlist(self) -> None:
        digest = hashlib.sha256(PNG).hexdigest()
        entry = {"sha256": digest, "reason": "Social preview image"}
        listed = {"schema_version": 1, "files": {"docs/assets/preview.png": entry}}
        self.assertEqual(self.repository({"docs/assets/preview.png": PNG}, listed), [])
        self.assertFinding(self.repository({"docs/assets/preview.png": PNG + b"x"}, listed),
                           "content changed since it was allowlisted")
        self.assertFinding(self.repository({}, listed),
                           "docs/assets/preview.png is not a tracked file")

    def test_allowlisted_file_replaced_by_text_still_fails(self) -> None:
        # Not marked binary in .gitattributes, so only its NUL bytes made it binary.
        original = b"\0\1\2payload"
        entry = {"sha256": hashlib.sha256(original).hexdigest(), "reason": "Reviewed payload"}
        listed = {"schema_version": 1, "files": {"payload.bin": entry}}
        self.assertEqual(self.repository({"payload.bin": original}, listed), [])
        self.assertFinding(self.repository({"payload.bin": b"plain text now\n"}, listed),
                           "payload.bin: content changed since it was allowlisted")

    def test_malformed_allowlist(self) -> None:
        for allowlist, text in (([], "needs schema_version 1"),
                                ({"schema_version": 2, "files": {}}, "needs schema_version 1"),
                                ({"schema_version": 1, "files": {"a.png": {"sha256": "x",
                                                                           "reason": "r"}}},
                                 "a.png needs a 64-character sha256 and a reason")):
            with self.subTest(allowlist=allowlist):
                self.assertFinding(self.repository({}, allowlist), text)

    def test_main_reports_the_verdict(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()) as printed:
            self.assertEqual(check_binaries.main(["--root", str(ROOT)]), 0)
        self.assertIn("binaries: PASS", printed.getvalue())


if __name__ == "__main__":
    unittest.main()
