"""Tests for tools/format_vba.py, with synthetic VBA fragments.

Each rule case states one line or a few lines as the formatter must leave
them, so that a failure names the rule. The literal and apostrophe cases are
the ones a line-based formatter is most likely to get wrong.
"""

from __future__ import annotations

import contextlib
import io
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import format_vba  # noqa: E402

HEAD = ['Attribute VB_Name = "CORE_Probe"', "Option Explicit", "Option Private Module", ""]


def module(*body: str) -> str:
    return "".join(f"{line}\r\n" for line in [*HEAD, *body])


def body_of(text: str) -> list[str]:
    return text.split("\r\n")[len(HEAD):-1]


class RuleTest(unittest.TestCase):
    def assertFormats(self, before: list[str], after: list[str]) -> None:
        formatted = format_vba.format_text(module(*before))
        self.assertEqual(body_of(formatted), after)
        self.assertEqual(format_vba.format_text(formatted), formatted, "not idempotent")

    def test_formatted_module_is_unchanged(self) -> None:
        text = module("Private Sub Run()", "    Dim Count As Long", "End Sub")
        self.assertEqual(format_vba.format_text(text), text)

    def test_line_endings_and_end_of_file(self) -> None:
        lf = "\n".join([*HEAD, "Private Sub Run()", "End Sub", "", "", ""])
        self.assertEqual(format_vba.format_text(lf),
                         module("Private Sub Run()", "End Sub"))
        no_final = module("Private Sub Run()", "End Sub").removesuffix("\r\n")
        self.assertEqual(format_vba.format_text(no_final),
                         module("Private Sub Run()", "End Sub"))

    def test_trailing_white_space(self) -> None:
        self.assertFormats(["Private Sub Run()  ", "    Count = 1\t ", "End Sub"],
                           ["Private Sub Run()", "    Count = 1", "End Sub"])

    def test_indentation_tabs(self) -> None:
        self.assertFormats(["Private Sub Run()", "\tIf Ready Then", "\t\tCount = 1",
                            "\tEnd If", "End Sub"],
                           ["Private Sub Run()", "    If Ready Then", "        Count = 1",
                            "    End If", "End Sub"])

    def test_outer_level_lines(self) -> None:
        self.assertFormats(["    Public Function Total() As Double", "Clean_Exit:",
                            "    Exit Function", "  Fail:  ' handler", "    End Function"],
                           ["Public Function Total() As Double", "Clean_Exit:",
                            "    Exit Function", "Fail:  ' handler", "End Function"])
        self.assertFormats(["  Friend Property Get Name() As String", "  End Property"],
                           ["Friend Property Get Name() As String", "End Property"])
        self.assertFormats(["   Private Static Sub Tick()", " End Sub"],
                           ["Private Static Sub Tick()", "End Sub"])

    def test_continuation_lines_keep_their_indentation(self) -> None:
        self.assertFormats(["Public Function Total( _", "    ByVal Amount As Double, _",
                            "    ByVal Rate As Double) As Double", "End Function"],
                           ["Public Function Total( _", "    ByVal Amount As Double, _",
                            "    ByVal Rate As Double) As Double", "End Function"])
        # A continued statement whose next line looks like a label is left alone.
        self.assertFormats(["Private Sub Run()", "    Debug.Print _",
                            "        Done:", "End Sub"],
                           ["Private Sub Run()", "    Debug.Print _",
                            "        Done:", "End Sub"])

    def test_two_blank_lines_between_procedures(self) -> None:
        self.assertFormats(["Private Sub One()", "End Sub", "'Header of Two",
                            "Private Sub Two()", "End Sub"],
                           ["Private Sub One()", "End Sub", "", "", "'Header of Two",
                            "Private Sub Two()", "End Sub"])
        self.assertFormats(["Private Sub One()", "End Sub", "", "", "", "",
                            "Private Sub Two()", "End Sub"],
                           ["Private Sub One()", "End Sub", "", "",
                            "Private Sub Two()", "End Sub"])

    def test_other_lines_are_not_moved(self) -> None:
        lines = ["Private Sub Run()", "    If Ready Then Exit Sub", "    Select Case Mode",
                 "        Case Else: Count = 0", "    End Select", "    End",
                 "    Debug.Print Count: Count = 1", "End Sub"]
        self.assertFormats(lines, lines)


class LiteralTest(unittest.TestCase):
    """No rule changes a string literal, or mistakes one for code."""

    CASES = [
        ("a label name in a literal is data", '    Message = "Fail:"'),
        ("a procedure keyword in a literal is data", '    Message = "End Sub"'),
        ("an apostrophe in a literal does not start a comment",
         "    Const Quote As String = \"a 'b' c\""),
        ("a doubled quote does not unbalance the scan",
         "    Const Said As String = \"he said \"\"go\"\" 'x' now\""),
        ("a tab inside a literal is kept", '    Message = "a\tb"'),
        ("spaces inside a literal are kept", '    Message = "a     b"'),
    ]

    def test_literals_are_unchanged(self) -> None:
        for label, line in self.CASES:
            with self.subTest(label):
                text = module("Private Sub Run()", line, "End Sub")
                self.assertIn(line, body_of(format_vba.format_text(text)))

    def test_comment_after_a_literal_with_an_apostrophe(self) -> None:
        line = "Done:  ' after \"a 'quoted' word\""
        self.assertEqual(format_vba.split_code_comment(line),
                         ("Done:  ", "' after \"a 'quoted' word\""))
        self.assertEqual(format_vba.split_code_comment("    X = \"it's\" ' note"),
                         ('    X = "it\'s" ', "' note"))


class HeaderTest(unittest.TestCase):
    def test_vbe_header_is_kept_byte_for_byte(self) -> None:
        header = ["VERSION 5.00", "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRun ",
                  "   Caption         =   \"Run\"   ", "\tClientHeight    =   3015",
                  "   Begin VB.CommandButton cmdOk ", "      Caption = \"OK\" ", "   End",
                  "End", "Attribute VB_Name = \"frmRun\"", "Attribute VB_Exposed = False"]
        text = "".join(f"{line}\r\n" for line in [*header, "Option Explicit", "",
                                                    "  Private Sub cmdOk_Click()  ",
                                                    "  End Sub"])
        lines = format_vba.format_text(text).split("\r\n")
        self.assertEqual(lines[:len(header)], header)
        self.assertEqual(lines[len(header):-1],
                         ["Option Explicit", "", "Private Sub cmdOk_Click()", "End Sub"])

    def test_class_header(self) -> None:
        header = ["VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1  'True", "END",
                  'Attribute VB_Name = "CResult"', "Attribute VB_PredeclaredId = False"]
        lines = [*header, "Option Explicit"]
        self.assertEqual(format_vba.format_body([]), [])
        self.assertEqual(format_vba.header_length(lines), len(header))


class MainTest(unittest.TestCase):
    def run_main(self, files: dict[str, str], *args: str) -> tuple[int, str, Path]:
        directory = tempfile.TemporaryDirectory(prefix="format-vba-")
        self.addCleanup(directory.cleanup)
        root = Path(directory.name)
        for name, text in files.items():
            (root / name).parent.mkdir(parents=True, exist_ok=True)
            (root / name).write_bytes(text.encode("latin-1"))
        with contextlib.redirect_stdout(io.StringIO()) as printed:
            code = format_vba.main([*args, "--root", str(root)])
        return code, printed.getvalue(), root

    def test_no_source_passes(self) -> None:
        code, printed, _ = self.run_main({}, "--check")
        self.assertEqual(code, 0)
        self.assertIn("vba-format: PASS (no VBA source found)", printed)

    def test_repository_passes(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(format_vba.main(["--check", "--root", str(ROOT)]), 0)

    def test_check_and_write(self) -> None:
        drifted = module("  Private Sub Run()  ", "End Sub")
        files = {"src/core/CORE_Probe.bas": drifted,
                 "src/workbook/frmRun.frx": "\x00\x01",
                 "docs/notes.bas": drifted}
        code, printed, root = self.run_main(files, "--check")
        self.assertEqual(code, 1)
        self.assertIn("FAIL: src/core/CORE_Probe.bas: not in house-style layout", printed)
        self.assertNotIn("docs/notes.bas", printed)
        self.assertIn("vba-format: FAIL (1 file(s)", printed)
        code, printed, root = self.run_main(files, "--write")
        self.assertEqual(code, 0)
        self.assertEqual((root / "src/core/CORE_Probe.bas").read_bytes().decode("latin-1"),
                         module("Private Sub Run()", "End Sub"))
        self.assertEqual((root / "docs/notes.bas").read_bytes().decode("latin-1"), drifted)

    def test_bytes_outside_ascii_round_trip(self) -> None:
        text = module("Private Sub Run()", "    ' caf\xe9", "End Sub")
        code, _, root = self.run_main({"src/core/CORE_Probe.bas": text}, "--write")
        self.assertEqual(code, 0)
        self.assertEqual((root / "src/core/CORE_Probe.bas").read_bytes(),
                         text.encode("latin-1"))

    def test_explicit_paths(self) -> None:
        files = {"elsewhere/CORE_Probe.bas": module("  Private Sub Run()", "End Sub")}
        code, printed, _ = self.run_main(files, "--check", "elsewhere/CORE_Probe.bas")
        self.assertEqual(code, 1)
        self.assertIn("FAIL: elsewhere/CORE_Probe.bas", printed)


if __name__ == "__main__":
    unittest.main()
