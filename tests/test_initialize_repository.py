"""Tests for tools/initialize_repository.py.

Each adaptation is tested on a small text. One test copies the tracked tree
to a temporary Git repository, initializes it, and runs the Repository
integrity checks there with --template false, as a generated repository's CI
does. That run leaves out three checks: the test suite, which would run this
test again inside itself, and Ruff and mypy. They read only the Python code
of tools/ and tests/, which the initializer never changes, so the template's
own run of them already covers the copy. In place of the suite it runs the
test modules that compare the adapted files with each other; the integrity
checks of the copy cover the rest. The same run is
repeated after deleting the initializer and this test, which a generated
repository may do once it is initialized; there, this copy test is skipped.
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

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_changelog  # noqa: E402
import check_repository  # noqa: E402
import initialize_repository as initialize  # noqa: E402
from check_repository import Check, Context  # noqa: E402
from initialize_repository import InitializeError, Options  # noqa: E402

REPOSITORY = "owner/new-workbook"
OPTIONS = Options(REPOSITORY, "LCR", "release/0.1.0")
# The checks a generated repository fails until it is initialized.
CLEANUP_CHECKS = ("template-only-content", "issue-contact-links", "readme-traffic-badges",
                  "dependabot-target-branches", "changelog")
# Read only code the initializer never changes; see the module docstring.
NOT_RUN = {"tests", "ruff", "mypy"}
ADAPTED_FILE_TESTS = ("test_repository_consistency", "test_workflow_controls",
                      "test_check_issues")
INITIALIZER = ("tools/initialize_repository.py", "tests/test_initialize_repository.py")
# Only the template still holds the template-only section; once it is
# initialized, a copy of this tree has nothing left to adapt.
IS_TEMPLATE = (ROOT / "AGENTS.md").is_file() and \
    initialize.SECTION_START in (ROOT / "AGENTS.md").read_text(encoding="utf-8")

AGENTS = """# Instructions

- A rule.

<!-- template-only:start -->
## Template maintenance

- A template rule.
<!-- template-only:end -->
"""

CONFIG = """blank_issues_enabled: false
contact_links:
  - name: "Security"
    url: "https://github.com/template-owner/template/security/policy"
  - name: "Labels"
    url: "https://github.com/template-owner/template/blob/HEAD/docs/LABELS.md"
"""

BADGES = "https://raw.githubusercontent.com/template-owner/template/traffic-history/data/"
README = f"""# Template

[![Clones]({BADGES}clones.json)](docs/TRAFFIC.md)
[![Visitors]({BADGES}visitors.json)](docs/TRAFFIC.md)

I use this template.
"""

DEPENDABOT = """version: 2
updates:
  - package-ecosystem: pip
    directory: /tools
    target-branch: release/1.0.0
    schedule:
      interval: weekly
  - package-ecosystem: github-actions
    directory: /
    target-branch: "release/1.0.0"
"""

# The template after its first release: a dated section, VERSION, and a
# Current status that names the release.
CHANGELOG = """# Changelog

Notable changes. It is neither a commit log nor release evidence.

## Maintenance policy

- Add a material change under `[Unreleased]`.

## [Unreleased]

### Added

- Something new (#9).

## [1.0.0] - 2026-12-01

### Added

- The first release (#1).

[Unreleased]: https://github.com/template-owner/template/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/template-owner/template/releases/tag/v1.0.0
"""

RELEASING = """# Releasing

## Current status

| Item | State |
| --- | --- |
| Released versions | 1.0.0 |
| `VERSION` | 1.0.0 |
"""

HOUSE_STYLE = """# House style

Replace `<REGIME>` with the project's short name, once, in this table.

| Element | Convention | Example |
| --- | --- | --- |
| Public facade module | `<REGIME>_` plus subject | `<REGIME>_Functions` |
| Enum and member | `<REGIME>` plus noun | `<REGIME>RunStatus`, `rsFailed` |
"""

# The template's evidence: a tested environment and a host behaviour row that
# cite its own commit and issue.
INSTALLATION = """# Installation

Record each host on which a commit has been compiled and run:

| Tested environment | Commit | Evidence |
| --- | --- | --- |
| Excel 16.0, 64-bit | `0123456789abcdef0123456789abcdef01234567` | #204: `excel-run: PASS` |
| Excel 16.0, 32-bit | `fedcba9876543210fedcba9876543210fedcba98` | #300: `excel-run: PASS` |

## Next section
"""

VALIDATION_PLAN = """# Validation plan

| ID | Behaviour | Status | Probe | Host | Date | Result | Issue |
| --- | --- | --- | --- | --- | --- | --- | --- |
| `HB-01` | A name is re-cased | `assumed` | By hand | \u2014 | \u2014 | Not measured | #245 |
| `HB-02` | Macros start disabled | `measured` | `Run` it | Excel 16.0 | 2026-11-03 | It failed | #204 |

| Other | Issue |
| --- | --- |
| Not host behaviour | #9 |
"""


def write(root: Path, files: dict[str, str]) -> None:
    for name, text in files.items():
        path = root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")


class AdaptationTests(unittest.TestCase):
    def assertIdempotent(self, adapt: initialize.Adapter, text: str,
                         options: Options = OPTIONS) -> str:
        adapted, notes = adapt(text, options)
        self.assertTrue(notes)
        self.assertEqual(adapt(adapted, options), (adapted, []))
        return adapted

    def test_agents_loses_the_template_only_section_and_its_markers(self) -> None:
        adapted = self.assertIdempotent(initialize.adapt_agents, AGENTS)
        self.assertEqual(adapted, "# Instructions\n\n- A rule.\n")
        middle = AGENTS.replace("<!-- template-only:end -->\n",
                                "<!-- template-only:end -->\n\n## Later\n\n- Kept.\n")
        self.assertEqual(initialize.adapt_agents(middle, OPTIONS)[0],
                         "# Instructions\n\n- A rule.\n\n## Later\n\n- Kept.\n")
        with self.assertRaisesRegex(InitializeError, "has no <!-- template-only:end -->"):
            initialize.adapt_agents("<!-- template-only:start -->\n", OPTIONS)

    def test_contact_links_point_at_the_new_repository(self) -> None:
        adapted = self.assertIdempotent(initialize.adapt_contact_links, CONFIG)
        self.assertIn('url: "https://github.com/owner/new-workbook/security/policy"', adapted)
        self.assertIn("https://github.com/owner/new-workbook/blob/HEAD/docs/LABELS.md", adapted)
        self.assertNotIn("template-owner", adapted)
        with self.assertRaisesRegex(InitializeError, "not a GitHub repository link"):
            initialize.adapt_contact_links('    url: "https://example.com/policy"\n', OPTIONS)

    def test_traffic_badges_are_repointed_or_removed(self) -> None:
        adapted = self.assertIdempotent(initialize.adapt_badges, README)
        self.assertEqual(adapted, README.replace("template-owner/template", REPOSITORY))
        removing = Options(REPOSITORY, "LCR", remove_traffic_badges=True)
        removed = self.assertIdempotent(initialize.adapt_badges, README, removing)
        self.assertEqual(removed, "# Template\n\nI use this template.\n")

    def test_dependabot_target_branch_is_set_added_or_removed(self) -> None:
        adapted = self.assertIdempotent(initialize.adapt_dependabot, DEPENDABOT)
        self.assertEqual(adapted.count("    target-branch: release/0.1.0\n"), 2)
        self.assertNotIn("release/1.0.0", adapted)
        default = Options(REPOSITORY, "LCR")
        removed = self.assertIdempotent(initialize.adapt_dependabot, DEPENDABOT, default)
        self.assertNotIn("target-branch", removed)
        added = self.assertIdempotent(initialize.adapt_dependabot, removed)
        self.assertEqual(added, adapted.replace('"release/0.1.0"', "release/0.1.0"))
        self.assertIn("    directory: /tools\n    target-branch: release/0.1.0\n", added)

    def test_house_style_takes_the_regime_in_its_table_only(self) -> None:
        adapted = self.assertIdempotent(initialize.adapt_house_style, HOUSE_STYLE)
        self.assertIn("| `LCR_` plus subject | `LCR_Functions` |", adapted)
        self.assertIn("`LCRRunStatus`", adapted)
        self.assertIn("Replace `<REGIME>` with the project's short name", adapted)
        with self.assertRaisesRegex(InitializeError, "already LCR_, not FRTB_"):
            initialize.adapt_house_style(adapted, Options(REPOSITORY, "FRTB"))

    def test_tested_environments_lose_their_rows_but_keep_the_header(self) -> None:
        adapted = self.assertIdempotent(initialize.adapt_tested_environments, INSTALLATION)
        self.assertEqual(adapted, INSTALLATION.split("| Excel 16.0, 64-bit")[0]
                         + "\n## Next section\n")
        self.assertIn("| Tested environment | Commit | Evidence |\n| --- | --- | --- |\n", adapted)
        self.assertNotIn("#204", adapted)

    def test_host_behaviour_rows_cite_no_issue_and_other_tables_stay(self) -> None:
        adapted = self.assertIdempotent(initialize.adapt_host_behaviour, VALIDATION_PLAN)
        self.assertEqual(adapted, VALIDATION_PLAN.replace("| #245 |", "| \u2014 |")
                         .replace("| #204 |", "| \u2014 |"))
        self.assertIn("| Not host behaviour | #9 |", adapted)

    def test_the_template_evidence_goes_with_the_changelog_entries(self) -> None:
        with tempfile.TemporaryDirectory(prefix="initialize-evidence-") as directory:
            root = Path(directory)
            files = {"INSTALLATION.md": INSTALLATION,
                     "docs/methodology/VALIDATION_PLAN.md": VALIDATION_PLAN}
            # A repository whose changelog is its own keeps its evidence.
            write(root, files)
            self.assertEqual(initialize.plan(root, OPTIONS), [])
            write(root, {"CHANGELOG.md": CHANGELOG})
            changes = initialize.plan(root, OPTIONS)
            self.assertEqual([change.path for change in changes],
                             ["CHANGELOG.md", *files])
            initialize.apply(root, changes)
            for name in files:
                self.assertNotIn("#2", (root / name).read_text(encoding="utf-8"))
            self.assertEqual(initialize.plan(root, OPTIONS), [])

    def test_options_are_validated(self) -> None:
        for options in (Options("no-slash", "LCR"), Options(REPOSITORY, "CORE"),
                        Options(REPOSITORY, "lcr"), Options(REPOSITORY, "LCR", "bad branch")):
            with self.subTest(options=options), self.assertRaises(InitializeError):
                initialize.validate(options)
        initialize.validate(OPTIONS)

    def test_a_released_template_changelog_becomes_a_fresh_one(self) -> None:
        with tempfile.TemporaryDirectory(prefix="initialize-changelog-") as directory:
            root = Path(directory)
            write(root, {"CHANGELOG.md": CHANGELOG, "VERSION": "1.0.0\n",
                         "RELEASING.md": RELEASING})
            before, _ = check_changelog.check(root, REPOSITORY)
            self.assertTrue(any("does not point to" in finding for finding in before), before)
            changes = initialize.plan(root, OPTIONS)
            self.assertEqual([change.path for change in changes],
                             ["CHANGELOG.md", "VERSION", "RELEASING.md"])
            initialize.apply(root, changes)
            self.assertFalse((root / "VERSION").exists())
            changelog = (root / "CHANGELOG.md").read_text(encoding="utf-8")
            self.assertEqual(changelog, (
                "# Changelog\n\nNotable changes. It is neither a commit log nor release "
                "evidence. **No version has been released.**\n\n## Maintenance policy\n\n"
                "- Add a material change under `[Unreleased]`.\n\n## [Unreleased]\n\n"
                "[Unreleased]: https://github.com/owner/new-workbook/commits/main\n"))
            self.assertEqual(check_changelog.check(root, REPOSITORY), ([], 0))
            self.assertEqual(initialize.plan(root, OPTIONS), [])


def copy_tracked_tree(destination: Path) -> None:
    names = subprocess.run(["git", "ls-files", "-z"], cwd=ROOT, capture_output=True,
                           check=True).stdout.decode("utf-8").split("\0")
    for name in filter(None, names):
        if (ROOT / name).is_file():
            (destination / name).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(ROOT / name, destination / name)


def adapted_checks(context: Context) -> list[Check]:
    listed = [check for check in check_repository.checks(context) if check.name not in NOT_RUN]
    tests = [sys.executable, "-m", "unittest", *[f"tests.{name}" for name in ADAPTED_FILE_TESTS]]
    return [Check("tests-of-adapted-files", tests), *listed]


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class ReleaseBranchWarningTests(unittest.TestCase):
    """--release-branch is checked against origin as dependabot-target-branches checks it."""

    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="initialize-warning-")
        self.addCleanup(directory.cleanup)
        self.base = Path(directory.name)
        self.origin = self.base / "origin.git"
        work = self.base / "work"
        work.mkdir()
        self.git(self.base, "init", "-q", "--bare", str(self.origin))
        for command in (["init", "-q", "-b", "main"], ["config", "user.name", "Fixture"],
                        ["config", "user.email", "fixture@example.invalid"],
                        ["config", "commit.gpgsign", "false"],
                        ["commit", "-q", "--allow-empty", "-m", "Created"],
                        ["push", "-q", str(self.origin), "HEAD:refs/heads/main",
                         "HEAD:refs/heads/release/0.1.0"]):
            self.git(work, *command)

    @staticmethod
    def git(root: Path, *args: str) -> None:
        subprocess.run(["git", *args], cwd=root, capture_output=True, text=True, check=True)

    def clone(self, *options: str) -> Path:
        target = self.base / "clone"
        self.git(self.base, "clone", "-q", *options, str(self.origin), str(target))
        return target

    def test_a_branch_only_on_origin_gives_no_warning(self) -> None:
        clone = self.clone("--single-branch", "--branch", "main")
        self.assertEqual(initialize.warnings(clone, OPTIONS), [])

    def test_a_branch_origin_lacks_is_a_warning(self) -> None:
        clone = self.clone()
        found = initialize.warnings(clone, Options(REPOSITORY, "LCR", "release/0.2.0"))
        self.assertEqual(len(found), 1)
        self.assertIn("origin has no branch release/0.2.0", found[0])
        self.assertIn("dependabot-target-branches check fails", found[0])

    def test_no_answer_from_origin_is_a_warning_that_says_so(self) -> None:
        clone = self.clone()
        self.git(clone, "remote", "remove", "origin")
        found = initialize.warnings(clone, OPTIONS)
        self.assertEqual(len(found), 1)
        self.assertIn("origin cannot be asked whether branch release/0.1.0 exists", found[0])

    def test_no_release_branch_needs_no_check(self) -> None:
        self.assertEqual(initialize.warnings(self.base, Options(REPOSITORY, "LCR")), [])


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
@unittest.skipUnless(IS_TEMPLATE, "this repository is already initialized")
class GeneratedRepositoryTest(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="initialize-repository-")
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        copy_tracked_tree(self.root)
        for command in (["init", "-q", "-b", "main"], ["config", "user.name", "Fixture"],
                        ["config", "user.email", "fixture@example.invalid"],
                        ["config", "core.autocrlf", "false"],
                        ["config", "commit.gpgsign", "false"]):
            self.git(*command)
        self.commit("Created from the template")
        # An origin with the release branch that Dependabot targets; the push
        # also makes it a fetched remote branch.
        origin = tempfile.TemporaryDirectory(prefix="initialize-repository-origin-")
        self.addCleanup(origin.cleanup)
        self.git("init", "-q", "--bare", origin.name)
        self.git("remote", "add", "origin", origin.name)
        self.git("push", "-q", "origin", "HEAD:refs/heads/main", "HEAD:refs/heads/release/0.1.0")

    def git(self, *args: str) -> str:
        return subprocess.run(["git", *args], cwd=self.root, capture_output=True, text=True,
                              check=True).stdout

    def commit(self, message: str) -> None:
        self.git("add", "-A")
        self.git("commit", "-qm", message)

    def initialize(self, *extra: str) -> subprocess.CompletedProcess[str]:
        return subprocess.run([sys.executable, "tools/initialize_repository.py", "--repository",
                               REPOSITORY, "--regime", "LCR", "--release-branch",
                               "release/0.1.0", *extra], cwd=self.root, capture_output=True,
                              text=True, check=False)

    def integrity(self, listed: list[str] | None = None) -> tuple[int, str, dict[str, str]]:
        """Run the checks with --template false; return the code, output and statuses."""
        def chosen(context: Context) -> list[Check]:
            return [check for check in adapted_checks(context)
                    if listed is None or check.name in listed]
        printed = io.StringIO()
        with contextlib.redirect_stdout(printed):
            code = check_repository.main(["--root", str(self.root), "--template", "false",
                                          "--repository", REPOSITORY, "--default-branch",
                                          "main"], chosen)
        report = json.loads((self.root / check_repository.REPORT_DIR / "report.json")
                            .read_text(encoding="utf-8"))
        return code, printed.getvalue(), {check["name"]: check["status"]
                                          for check in report["checks"]}

    def test_the_dry_run_writes_no_file_not_even_an_ignored_one(self) -> None:
        def every_path() -> list[str]:
            return sorted(path.relative_to(self.root).as_posix() for path in self.root.rglob("*")
                          if ".git" not in path.relative_to(self.root).parts)
        before = every_path()
        # Without this variable Python writes the bytecode of each module it imports.
        environment = {key: value for key, value in os.environ.items()
                       if key != "PYTHONDONTWRITEBYTECODE"}
        dry_run = subprocess.run([sys.executable, "tools/initialize_repository.py", "--repository",
                                  REPOSITORY, "--regime", "LCR"], cwd=self.root, env=environment,
                                 capture_output=True, text=True, check=False)
        self.assertEqual(dry_run.returncode, 0, dry_run.stderr)
        self.assertIn("dry run, nothing written", dry_run.stdout)
        self.assertEqual(every_path(), before)

    def test_an_initialized_copy_passes_repository_integrity(self) -> None:
        code, _, statuses = self.integrity(list(CLEANUP_CHECKS))
        self.assertEqual((code, statuses), (1, dict.fromkeys(CLEANUP_CHECKS, "fail")))

        dry_run = self.initialize()
        self.assertEqual(dry_run.returncode, 0, dry_run.stderr)
        self.assertIn("- AGENTS.md: delete the template-only section", dry_run.stdout)
        self.assertIn("dry run, nothing written", dry_run.stdout)
        self.assertEqual(self.git("status", "--porcelain", "--untracked-files=all"), "")

        (self.root / "notes.txt").write_text("unrelated\n", encoding="utf-8")
        refused = self.initialize("--apply")
        self.assertEqual(refused.returncode, 2)
        self.assertIn("the working tree differs from the commit", refused.stderr)
        self.assertEqual(self.git("status", "--porcelain"), "?? notes.txt\n")
        (self.root / "notes.txt").unlink()

        applied = self.initialize("--apply")
        self.assertEqual(applied.returncode, 0, applied.stderr)
        self.assertIn("Nothing on GitHub was changed", applied.stdout)
        # No evidence row cites a commit or an issue of the template.
        installation = (self.root / initialize.INSTALLATION).read_text(encoding="utf-8")
        rows = installation.split(initialize.TESTED_HEADER, 1)[1].split("\n")[2]
        self.assertFalse(rows.startswith("|"), rows)
        plan = (self.root / initialize.VALIDATION_PLAN).read_text(encoding="utf-8")
        host = [line for line in plan.split("\n") if line.startswith("| `HB-")]
        self.assertTrue(host)
        self.assertTrue(all(line.endswith(f"| {initialize.NO_EVIDENCE} |") for line in host), host)
        self.commit("Initialize the repository")
        again = self.initialize("--apply")
        self.assertEqual(again.returncode, 0, again.stderr)
        self.assertIn("nothing to change", again.stdout)
        self.assertEqual(self.git("status", "--porcelain", "--untracked-files=all"), "")

        code, printed, statuses = self.integrity()
        self.assertEqual(code, 0, printed)
        self.assertNotIn("fail", statuses.values(), printed)

        # Once run, the initializer can be deleted.
        for name in INITIALIZER:
            (self.root / name).unlink()
        self.commit("Delete the initializer")
        code, printed, _ = self.integrity()
        self.assertEqual(code, 0, printed)


if __name__ == "__main__":
    unittest.main()
