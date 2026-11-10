"""Tests for tools/render_docs.py: generated documentation blocks."""

from __future__ import annotations

import contextlib
import io
import json
import re
import shutil
import sys
import tempfile
import unittest
from collections.abc import Callable
from pathlib import Path
from typing import Any
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_repository  # noqa: E402
import render_docs  # noqa: E402
from render_docs import Block  # noqa: E402


def suite_counts() -> tuple[int, int, int]:
    """The example suite's expected cases and assertions, read from its source as
    tests/test_vba_harness.py reads them, and the assertions of the last case its
    runner calls."""
    policy = json.loads((ROOT / ".github/excel-evidence-policy.json").read_text("utf-8"))
    source = (ROOT / policy["suite"]["source"]).read_text("ascii").replace("\r\n", "\n")

    def constant(name: str) -> int:
        match = re.search(rf"^\s*Private Const {name}\s+As Long = (\d+)\b", source, re.M)
        assert match, name
        return int(match.group(1))

    def body(name: str) -> str:
        match = re.search(rf"^Private Sub {name}\(.*?^End Sub$", source, re.M | re.S)
        assert match, name
        return match.group(0)

    last = re.findall(r"^\s+(Case\w+)$", body("RunSuite"), re.M)[-1]
    return (constant("EXPECTED_CASES"), constant("EXPECTED_ASSERTIONS"),
            len(re.findall(r"TEST_Harness\.Assert\w+ ", body(last))))


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


class DataContractTest(unittest.TestCase):
    """The tables of the data contract follow docs/methodology/data-contract.json."""

    GUIDE = "docs/methodology/DATA_CONTRACT.md"
    CONTRACT = "docs/methodology/data-contract.json"

    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="render-contract-")
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        (self.root / "docs/methodology").mkdir(parents=True)
        for path in (self.GUIDE, self.CONTRACT):
            shutil.copyfile(ROOT / path, self.root / path)
        blocks = {name: block for name, block in render_docs.BLOCKS.items()
                  if block.path == self.GUIDE}
        patcher = mock.patch.dict(render_docs.BLOCKS, blocks, clear=True)
        patcher.start()
        self.addCleanup(patcher.stop)

    def edit_contract(self, edit: Callable[[dict[str, Any]], None]) -> None:
        path = self.root / self.CONTRACT
        contract = json.loads(path.read_text(encoding="utf-8"))
        edit(contract)
        path.write_text(json.dumps(contract), encoding="utf-8")

    def problems(self, write: bool = False) -> list[str]:
        with contextlib.redirect_stdout(io.StringIO()):
            return render_docs.run(self.root, write)

    def test_the_guide_has_the_three_blocks(self) -> None:
        self.assertEqual(sorted(render_docs.BLOCKS), ["data-contract-codes",
                                                      "data-contract-datasets",
                                                      "data-contract-fields"])
        self.assertEqual(self.problems(), [])

    def test_a_changed_contract_makes_the_tables_stale(self) -> None:
        def optional_first_field(contract: dict[str, Any]) -> None:
            contract["datasets"][0]["fields"][0]["required"] = False
        self.edit_contract(optional_first_field)
        self.assertEqual(self.problems(), [
            f"{self.GUIDE}: block data-contract-fields differs from a fresh render; run "
            "python tools/render_docs.py --write"])
        self.edit_contract(lambda contract: contract["datasets"][0].update(provider="Other"))
        self.assertIn(f"{self.GUIDE}: block data-contract-datasets differs from a fresh "
                      "render; run python tools/render_docs.py --write", self.problems())
        self.assertEqual(self.problems(write=True), [])
        self.assertIn("| Other |", (self.root / self.GUIDE).read_text(encoding="utf-8"))

    def test_an_edited_table_fails(self) -> None:
        path = self.root / self.GUIDE
        text = path.read_text(encoding="utf-8")
        path.write_text(text.replace("| Yes |", "| No |", 1), encoding="utf-8")
        self.assertEqual(self.problems(), [
            f"{self.GUIDE}: block data-contract-fields differs from a fresh render; run "
            "python tools/render_docs.py --write"])

    def test_an_invalid_contract_cannot_be_rendered(self) -> None:
        self.edit_contract(lambda contract: contract.update(kind="other"))
        reason = (f"{self.CONTRACT}: envelope does not match: schema_version 1, kind "
                  "'data-contract'; python tools/check_data_contract.py lists every problem")
        # The codes come from the validator, not the contract, and still render.
        self.assertEqual(sorted(self.problems()), [
            f"{self.GUIDE}: block data-contract-datasets cannot be rendered: {reason}",
            f"{self.GUIDE}: block data-contract-fields cannot be rendered: {reason}"])

    def test_how_each_kind_of_dataset_and_field_is_shown(self) -> None:
        base = {"illustrative": False, "description": "D", "provider": "P",
                "frequency": "F", "granularity": "G", "fields": [
                    {"name": "kind", "type": "text", "unit": None, "required": True,
                     "allowed": {"values": ["A", "B"]}, "notes": "Two kinds"},
                    {"name": "weight", "type": "decimal", "unit": "Decimal fraction",
                     "required": False, "allowed": {"min": 0, "max": 1}, "notes": ""},
                    {"name": "cap", "type": "integer", "unit": None, "required": False,
                     "allowed": {"max": 9}, "notes": ""}]}
        self.edit_contract(lambda contract: contract.update(datasets=[
            {**base, "id": "DS-02", "source": "a.tsv", "encoding": None, "delimiter": "\t",
             "decimal_separator": ","},
            {**base, "id": "DS-03", "source": "b.xlsx", "encoding": None, "delimiter": None,
             "decimal_separator": None},
            {**base, "id": "DS-04", "source": "Inputs", "encoding": None, "delimiter": None,
             "decimal_separator": None}]))
        datasets = render_docs.data_contract_datasets(self.root)
        self.assertIn("| `DS-02` | D | P | F | G | `a.tsv` |  |", datasets)
        fields = render_docs.data_contract_fields(self.root)
        self.assertIn("### `DS-02` D\n\nRead from `a.tsv`, delimited text in `UTF-8`, with a "
                      "tab between cells and `,` as the decimal separator.", fields)
        self.assertIn("Read from `b.xlsx`, which is not delimited text", fields)
        self.assertIn("Read from the sheet `Inputs`.", fields)
        self.assertIn("| `kind` | Text | — | Yes | `A`, `B` | Two kinds |", fields)
        self.assertIn("| `weight` | Decimal | Decimal fraction | No | 0 to 1 |  |", fields)
        self.assertIn("| `cap` | Integer | — | No | ≤ 9 |  |", fields)
        self.assertNotIn("Illustrative", fields)


class RepositoryTest(unittest.TestCase):
    def test_the_repository_blocks_are_fresh(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(render_docs.run(ROOT, write=False), [])

    def test_the_suite_results_follow_the_suite(self) -> None:
        cases, assertions, last = suite_counts()
        text = render_docs.example_suite_results(ROOT)
        self.assertIn("RESULT=PASS; suite=examples; operation=PASS; completeness=COMPLETE; "
                      f"cases={cases}; assertions={assertions}; failures=0; cleanup=PASS; "
                      "recovery=NONE", text)
        # The skipped-case run leaves out the last case and its assertions.
        self.assertIn("`TEST_Examples.RunExampleTestsWithSkippedCase` | `RESULT=FAIL`, "
                      "`operation=PASS`, `completeness=INCOMPLETE`, "
                      f"`cases={cases - 1}`, `assertions={assertions - last}`, `failures=1`",
                      text)
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
