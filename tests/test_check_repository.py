"""Tests for tools/check_repository.py, in temporary Git repositories.

The run itself is tested with small checks injected in place of the real
list, so that the test suite never runs itself.
"""

from __future__ import annotations

import contextlib
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_repository  # noqa: E402
from check_repository import Check, Context  # noqa: E402


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class RepositoryTest(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="check-repository-")
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        self.git("init", "-q", "-b", "main")
        self.git("config", "user.name", "Fixture")
        self.git("config", "user.email", "fixture@example.invalid")
        self.git("config", "core.autocrlf", "false")

    def git(self, *args: str) -> str:
        return subprocess.run(["git", *args], cwd=self.root, capture_output=True, text=True,
                              check=True).stdout

    def write(self, files: dict[str, str | bytes]) -> None:
        for name, content in files.items():
            path = self.root / name
            path.parent.mkdir(parents=True, exist_ok=True)
            if isinstance(content, bytes):
                path.write_bytes(content)
            else:
                path.write_text(content, encoding="utf-8")

    def commit(self, files: dict[str, str | bytes] | None = None) -> None:
        self.write(files or {})
        self.git("add", "-A")
        self.git("commit", "-qm", "Fixture", "--allow-empty")

    def context(self, **values: object) -> Context:
        base: dict[str, object] = {"repository": "owner/project", "default_branch": "main",
                                   "is_template": False}
        base.update(values)
        return Context(self.root, **base)  # type: ignore[arg-type]


class FileCheckTest(RepositoryTest):
    def test_template_only_content(self) -> None:
        self.write({"AGENTS.md": "<!-- template-only:start -->\n"})
        self.assertEqual(len(check_repository.template_only(self.context())), 1)
        self.assertEqual(check_repository.template_only(self.context(is_template=True)), [])
        self.write({"AGENTS.md": "# Instructions\n"})
        self.assertEqual(check_repository.template_only(self.context()), [])

    def test_contact_links(self) -> None:
        self.write({".github/ISSUE_TEMPLATE/config.yml":
                    'contact_links:\n  - name: Security\n'
                    '    url: "https://github.com/Owner/Project/security"\n'
                    "  - name: Other\n    url: https://github.com/other/template/discussions\n"})
        findings = check_repository.contact_links(self.context())
        self.assertEqual(len(findings), 1)
        self.assertIn("https://github.com/other/template/discussions", findings[0][1])

    def test_traffic_badges(self) -> None:
        badge = "https://raw.githubusercontent.com/{}/traffic-history/badges/views.svg"
        self.write({"README.md": f"![a]({badge.format('owner/project')})\n"})
        self.assertEqual(check_repository.traffic_badges(self.context()), [])
        self.write({"README.md": f"![a]({badge.format('other/template')})\n"})
        self.assertIn("reads other/template",
                      check_repository.traffic_badges(self.context())[0][1])

    def test_dependabot_target_branches(self) -> None:
        self.commit({"README.md": "x\n"})
        self.git("update-ref", "refs/remotes/origin/release/1.0.0", "HEAD")
        self.write({".github/dependabot.yml":
                    "updates:\n  - package-ecosystem: pip\n    target-branch: release/1.0.0\n"
                    '  - package-ecosystem: npm\n    target-branch: "release/0.9.0"\n'})
        findings = check_repository.dependabot_branches(self.context())
        self.assertEqual([message for _, message in findings],
                         ["target-branch release/0.9.0 does not exist; set it to the active "
                          "release branch or remove it."])

    def test_codeql_default_branch(self) -> None:
        self.write({"tests/strict_yaml.py": (ROOT / "tests/strict_yaml.py").read_text("utf-8"),
                    ".github/workflows/codeql.yml":
                    "on:\n  push:\n    branches: [main]\n  pull_request:\n"
                    "    branches: [main, 'release/**']\n"})
        self.assertEqual(check_repository.codeql_default_branch(self.context()), [])
        findings = check_repository.codeql_default_branch(self.context(default_branch="trunk"))
        self.assertEqual(len(findings), 2)
        self.assertIn("omit the default branch trunk", findings[0][1])

    def test_required_documents(self) -> None:
        self.write({"README.md": "x\n", "LICENSE": "",
                    "docs/repo_governance/SETTINGS.md": "x\n"})
        self.assertEqual(check_repository.required_documents(self.context()),
                         [("LICENSE", "is missing or empty")])


class WhitespaceTest(RepositoryTest):
    def outcomes(self) -> dict[str, int]:
        return {name: subprocess.run(command, cwd=self.root, capture_output=True,
                                     check=False).returncode
                for name, command in check_repository.whitespace_commands(self.root)}

    def test_committed_trailing_white_space_fails_in_a_clean_checkout(self) -> None:
        self.commit({"notes.txt": "trailing \n"})
        self.assertEqual(self.git("status", "--porcelain"), "")
        self.assertNotEqual(self.outcomes()["whitespace-committed"], 0)
        self.assertEqual(self.outcomes()["whitespace-local"], 0)

    def test_clean_text_crlf_and_binary_pass(self) -> None:
        self.commit({".gitattributes": "*.cmd text eol=crlf\n*.bin binary\n",
                     "notes.txt": "clean\n", "run.cmd": b"echo clean\r\n",
                     "data.bin": b"\x00binary \n"})
        self.assertEqual(set(self.outcomes().values()), {0})

    def test_local_white_space_fails(self) -> None:
        self.commit({"notes.txt": "clean\n"})
        self.write({"notes.txt": "bad\t\n"})
        self.assertNotEqual(self.outcomes()["whitespace-local"], 0)


class RunTest(RepositoryTest):
    def run_checks(self, listed: list[Check], strict: bool = False,
                   summary: Path | None = None) -> tuple[int, str, dict[str, object]]:
        with contextlib.redirect_stdout(io.StringIO()) as printed, \
                mock.patch.dict(os.environ, {"PR_HEAD_SHA": "f" * 40, "GITHUB_ACTIONS": ""}):
            code = check_repository.run(self.context(), listed,
                                        self.root / "test-results/integrity", strict, summary)
        report = json.loads((self.root / "test-results/integrity/report.json")
                            .read_text(encoding="utf-8"))
        return code, printed.getvalue(), report

    def test_report_logs_and_summary(self) -> None:
        self.commit({"README.md": "x\n"})
        self.write({"untracked.txt": "y\n"})
        listed = [Check("ok", [sys.executable, "-c", "print('fine')"]),
                  Check("broken", [sys.executable, "-c", "import sys; sys.exit('bad input')"]),
                  Check("finding", function=lambda context: [("README.md", "is wrong")])]
        step_summary = self.root / "step-summary.md"
        step_summary.write_text("before\n", encoding="utf-8")
        code, printed, report = self.run_checks(listed, summary=step_summary)
        self.assertEqual(code, 1)
        self.assertIn("broken: FAIL (exit code 1)", printed)
        self.assertIn("bad input", printed)
        self.assertEqual(report["status"], "fail")
        self.assertEqual(report["commit"], self.git("rev-parse", "HEAD").strip())
        self.assertIs(report["dirty"], True)
        self.assertEqual(report["pull_request_head"], "f" * 40)
        self.assertEqual(set(report["files_sha256"]), {"README.md"})  # type: ignore[arg-type]
        statuses = {check["name"]: (check["status"], check["exit_code"])
                    for check in report["checks"]}  # type: ignore[attr-defined]
        self.assertEqual(statuses, {"ok": ("pass", 0), "broken": ("fail", 1),
                                    "finding": ("fail", 1)})
        out = self.root / "test-results/integrity"
        self.assertIn("fine", (out / "ok.log").read_text(encoding="utf-8"))
        self.assertIn("FAIL: README.md: is wrong", (out / "finding.log").read_text("utf-8"))
        text = step_summary.read_text(encoding="utf-8")
        self.assertTrue(text.startswith("before\n## Repository integrity: FAIL"))
        self.assertIn("| `broken` | ❌ fail | exit code 1 |", text)
        self.assertIn("no Excel run", text)
        self.assertEqual((out / "summary.md").read_text(encoding="utf-8"),
                         text.removeprefix("before\n"))

    def test_a_check_that_raises_fails_and_the_run_continues(self) -> None:
        self.commit({"tests/strict_yaml.py": (ROOT / "tests/strict_yaml.py").read_text("utf-8"),
                     ".github/workflows/codeql.yml": "on:\n  push: {}\n  push: {}\n"})
        listed = [Check("codeql-default-branch", function=check_repository.codeql_default_branch),
                  Check("after", [sys.executable, "-c", "print('ran')"])]
        code, printed, report = self.run_checks(listed)
        self.assertEqual(code, 1)
        statuses = {check["name"]: check["status"]
                    for check in report["checks"]}  # type: ignore[attr-defined]
        self.assertEqual(statuses, {"codeql-default-branch": "fail", "after": "pass"})
        log = (self.root / "test-results/integrity/codeql-default-branch.log").read_text("utf-8")
        self.assertIn("duplicate key", log)
        self.assertIn("codeql-default-branch: FAIL (raised an error)", printed)

    def test_missing_context_is_skipped_unless_strict(self) -> None:
        self.commit()
        listed = [Check("needs-node", ["node"], missing="node is not on PATH")]
        code, printed, report = self.run_checks(listed)
        self.assertEqual(code, 0)
        self.assertIn("needs-node: SKIP (node is not on PATH)", printed)
        self.assertEqual(report["status"], "pass")
        code, _, report = self.run_checks(listed, strict=True)
        self.assertEqual(code, 1)
        self.assertEqual(report["status"], "fail")

    def test_github_annotations(self) -> None:
        with mock.patch.dict(os.environ, {"GITHUB_ACTIONS": "true"}):
            self.assertEqual(check_repository.annotate(("a.md", "bad")), "::error file=a.md::bad")

    def test_main_detects_the_repository_and_takes_arguments(self) -> None:
        self.commit()
        self.git("remote", "add", "origin", "https://github.com/Owner/Project.git")
        seen: list[Context] = []

        def listed(context: Context) -> list[Check]:
            seen.append(context)
            return [Check("ok", [sys.executable, "-c", "pass"])]

        with contextlib.redirect_stdout(io.StringIO()):
            code = check_repository.main(["--root", str(self.root), "--template", "false"],
                                         listed)
        self.assertEqual(code, 0)
        self.assertEqual((seen[0].repository, seen[0].default_branch, seen[0].is_template),
                         ("Owner/Project", None, False))
        with contextlib.redirect_stdout(io.StringIO()):
            check_repository.main(["--root", str(self.root), "--repository", "a/b",
                                   "--default-branch", "trunk", "--template", ""], listed)
        self.assertEqual((seen[1].repository, seen[1].default_branch, seen[1].is_template),
                         ("a/b", "trunk", None))

    def test_real_check_list_names_each_check_once(self) -> None:
        self.commit()
        names = [check.name for check in check_repository.checks(self.context())]
        self.assertEqual(len(names), len(set(names)))
        for expected in ("tests", "ruff", "mypy", "actionlint", "binaries", "vba-format",
                         "test-cases", "derivations", "changelog", "generated-docs",
                         "markdown-links",
                         "codeql-default-branch", "whitespace-committed"):
            self.assertIn(expected, names)
        without = check_repository.checks(self.context(repository=None, is_template=None,
                                                       default_branch=None))
        missing = {check.name for check in without if check.missing}
        self.assertTrue({"issue-contact-links", "readme-traffic-badges", "template-only-content",
                         "codeql-default-branch"} <= missing)


if __name__ == "__main__":
    unittest.main()
