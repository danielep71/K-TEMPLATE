"""Tests for tools/render_docs.py: generated documentation blocks."""

from __future__ import annotations

import contextlib
import io
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_repository  # noqa: E402
import render_docs  # noqa: E402
from render_docs import Block  # noqa: E402


def begin(name: str) -> str:
    return f"<!-- BEGIN generated: {name} -->"


def end(name: str) -> str:
    return f"<!-- END generated: {name} -->"


class RenderTest(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="render-docs-")
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        blocks = {"facts": Block("GUIDE.md", lambda root: "| fact |\n| --- |\n| one |")}
        patcher = mock.patch.dict(render_docs.BLOCKS, blocks, clear=True)
        patcher.start()
        self.addCleanup(patcher.stop)

    def write(self, name: str, text: str) -> None:
        (self.root / name).write_text(text, encoding="utf-8")

    def problems(self, write: bool = False) -> list[str]:
        with contextlib.redirect_stdout(io.StringIO()):
            return render_docs.run(self.root, write)

    def test_a_fresh_block_passes(self) -> None:
        self.write("GUIDE.md", f"Intro\n{begin('facts')}\n| fact |\n| --- |\n| one |\n"
                               f"{end('facts')}\nOutro\n")
        self.assertEqual(self.problems(), [])

    def test_a_stale_block_fails_and_write_renders_it(self) -> None:
        self.write("GUIDE.md", f"Intro\n{begin('facts')}\nold\n{end('facts')}\nOutro\n")
        self.assertEqual(self.problems(),
                         ["GUIDE.md: block facts differs from a fresh render; run "
                          "python tools/render_docs.py --write"])
        self.assertEqual(self.problems(write=True), [])
        self.assertEqual((self.root / "GUIDE.md").read_text(encoding="utf-8"),
                         f"Intro\n{begin('facts')}\n| fact |\n| --- |\n| one |\n{end('facts')}"
                         "\nOutro\n")
        self.assertEqual(self.problems(), [])

    def test_marker_problems(self) -> None:
        cases = {
            f"{begin('facts')}\n": "block facts is not closed",
            f"{end('facts')}\n": "END facts without its BEGIN",
            f"{begin('facts')}\n{begin('other')}\n{end('other')}\n{end('facts')}\n":
                "BEGIN other inside block facts",
            f"{begin('facts')}\nx\n{end('facts')}\n{begin('facts')}\nx\n{end('facts')}\n":
                "block facts appears twice",
            "no markers\n": "block facts is missing",
        }
        for text, expected in cases.items():
            with self.subTest(expected=expected):
                self.write("GUIDE.md", text)
                found = self.problems()
                self.assertTrue(any(expected in problem for problem in found), found)

    def test_an_unknown_block_anywhere_fails(self) -> None:
        self.write("GUIDE.md", f"{begin('facts')}\n| fact |\n| --- |\n| one |\n{end('facts')}\n")
        (self.root / "docs").mkdir()
        self.write("docs/OTHER.md", f"{begin('nothing')}\nx\n{end('nothing')}\n")
        self.assertEqual(self.problems(), ["docs/OTHER.md: block nothing has no renderer"])

    def test_a_missing_file_fails(self) -> None:
        self.assertEqual(self.problems(), ["GUIDE.md: missing"])

    def test_main_exit_codes(self) -> None:
        self.write("GUIDE.md", f"{begin('facts')}\nold\n{end('facts')}\n")
        with contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(render_docs.main(["--check", "--root", str(self.root)]), 1)
        self.assertIn("generated-docs: FAIL", output.getvalue())
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(render_docs.main(["--write", "--root", str(self.root)]), 0)
            self.assertEqual(render_docs.main(["--check", "--root", str(self.root)]), 0)


class RepositoryTest(unittest.TestCase):
    def test_the_repository_blocks_are_fresh(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(render_docs.run(ROOT, write=False), [])

    def test_the_suite_results_follow_the_suite(self) -> None:
        text = render_docs.example_suite_results(ROOT)
        self.assertIn("RESULT=PASS; suite=examples; operation=PASS; completeness=COMPLETE; "
                      "cases=13; assertions=30; failures=0; cleanup=PASS; recovery=NONE", text)
        self.assertIn("`TEST_Examples.RunExampleTestsWithSkippedCase` | `RESULT=FAIL`, "
                      "`operation=PASS`, `completeness=INCOMPLETE`, `cases=12`, "
                      "`assertions=28`, `failures=1`", text)
        self.assertIn("Two more entry points", text)

    def test_every_check_is_listed_once_and_described(self) -> None:
        text = render_docs.integrity_checks(ROOT)
        context = check_repository.Context(ROOT, repository=None, default_branch=None,
                                           is_template=None)
        for check in check_repository.checks(context):
            with self.subTest(check=check.name):
                self.assertEqual(text.count(f"| `{check.name}` |"), 1)
        self.assertNotIn(sys.executable, text)
        self.assertNotIn("Built in: |", text)

    def test_a_check_without_a_description_is_refused(self) -> None:
        def undocumented(context: check_repository.Context) -> list[tuple[str, str]]:
            return []

        listed = [check_repository.Check("silent", function=undocumented)]
        with mock.patch.object(check_repository, "checks", return_value=listed), \
                self.assertRaisesRegex(ValueError, "silent: undocumented has no docstring"):
            render_docs.integrity_checks(ROOT)

    def test_commands_are_shown_as_typed(self) -> None:
        self.assertEqual(render_docs.shown_command([sys.executable, "-m", "x"]), "python -m x")
        self.assertEqual(render_docs.shown_command(["/usr/bin/node", "a.mjs"]), "node a.mjs")
        # The same table on Windows, whose executables carry a suffix.
        self.assertEqual(render_docs.shown_command([r"C:\Program Files\nodejs\node.exe", "a.mjs"]),
                         "node a.mjs")
        self.assertEqual(render_docs.shown_command([r"C:\tools\actionlint.EXE", "-no-color"]),
                         "actionlint -no-color")
        self.assertEqual(render_docs.shown_command(["git", "diff"]), "git diff")


if __name__ == "__main__":
    unittest.main()
