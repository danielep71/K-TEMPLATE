"""Valid and invalid CHANGELOG.md and VERSION fixtures for tools/check_changelog.py.

Each fixture is a temporary directory holding only the two files. The
versions and dates are synthetic; no release of this repository is implied.
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

import check_changelog  # noqa: E402

REPOSITORY = "example/repository"
TAGS = f"https://github.com/{REPOSITORY}/releases/tag"
UNRELEASED_ONLY = (
    "# Changelog\n\nNo version has been released.\n\n## Policy\n\nText.\n\n"
    "## [Unreleased]\n\n### Added\n\n- A change.\n\n"
    f"[Unreleased]: https://github.com/{REPOSITORY}/commits/main\n"
)
RELEASED = (
    "# Changelog\n\n## [Unreleased]\n\n## [0.2.0] - 2026-11-02\n\n### Fixed\n\n- A fix.\n\n"
    "## [0.1.0] - 2026-10-30\n\n### Added\n\n- A start.\n\n"
    f"[Unreleased]: https://github.com/{REPOSITORY}/compare/v0.2.0...HEAD\n"
    f"[0.2.0]: {TAGS}/v0.2.0\n[0.1.0]: {TAGS}/v0.1.0\n"
)


STATUS_BEFORE = ("# Releasing\n\n## 🧭 Current status\n\n| Item | State |\n| --- | --- |\n"
                 "| Released versions | None |\n"
                 "| `VERSION` | Does not exist yet; it is created when the first release is "
                 "prepared |\n\n## 🏷️ Versioning\n\n| Item | Convention |\n| --- | --- |\n"
                 "| `VERSION` | One line |\n")
STATUS_AFTER = (STATUS_BEFORE.replace("| Released versions | None |",
                                      "| Released versions | 0.2.0, 0.1.0 |")
                .replace("Does not exist yet; it is created when the first release is prepared",
                         "0.2.0"))


def findings(changelog: str | None, version: str | None = None,
             repository: str | None = None, **files: str) -> list[str]:
    """Check a directory holding the given CHANGELOG.md, VERSION and other files.

    files maps RELEASING and README to the text of RELEASING.md and README.md.
    """
    with tempfile.TemporaryDirectory(prefix="changelog-") as temporary:
        root = Path(temporary)
        if changelog is not None:
            (root / "CHANGELOG.md").write_text(changelog, encoding="utf-8")
        if version is not None:
            (root / "VERSION").write_text(version, encoding="utf-8")
        for name, text in files.items():
            (root / f"{name}.md").write_text(text, encoding="utf-8")
        return check_changelog.check(root, repository)[0]


class ChangelogTest(unittest.TestCase):
    def assertFinding(self, changelog: str | None, text: str, version: str | None = None,
                      repository: str | None = None) -> None:
        found = findings(changelog, version, repository)
        self.assertTrue(any(text in item for item in found), found)

    def test_valid_files_pass(self) -> None:
        self.assertEqual(findings(UNRELEASED_ONLY), [])
        self.assertEqual(findings(RELEASED, "0.2.0\n"), [])
        self.assertEqual(findings(RELEASED, "0.2.0\n", "Example/Repository"), [])

    def test_repository_files_pass(self) -> None:
        found, _ = check_changelog.check(ROOT)
        self.assertEqual(found, [])

    def test_missing_changelog(self) -> None:
        self.assertFinding(None, "CHANGELOG.md is missing")

    def test_unreleased_comes_first_without_a_date(self) -> None:
        self.assertFinding(RELEASED.replace("## [Unreleased]\n\n", ""),
                           "first version heading must be '## [Unreleased]'", "0.2.0\n")
        self.assertFinding(UNRELEASED_ONLY.replace("## [Unreleased]", "## [Unreleased] - 2026-10-30"),
                           "takes no date or suffix")
        self.assertFinding(UNRELEASED_ONLY.replace("[Unreleased]: ", "[Other]: "),
                           "missing link reference for [Unreleased]")

    def test_release_headings(self) -> None:
        for heading, text in (("## [0.2.0]", "must be '## [X.Y.Z] - YYYY-MM-DD'"),
                              ("## [v0.2.0] - 2026-11-02", "must be '## [X.Y.Z] - YYYY-MM-DD'"),
                              ("## [0.2] - 2026-11-02", "must be '## [X.Y.Z] - YYYY-MM-DD'"),
                              ("## [0.2.0] - 2026-02-30", "is not a calendar date")):
            with self.subTest(heading=heading):
                self.assertFinding(RELEASED.replace("## [0.2.0] - 2026-11-02", heading), text,
                                   "0.2.0\n")

    def test_unbracketed_version_headings(self) -> None:
        for heading in ("## 0.2.0 - 2026-11-02", "## v0.2.0 - 2026-11-02", "## Unreleased",
                        "##  [0.2.0] - 2026-11-02"):
            with self.subTest(heading=heading):
                self.assertFinding(RELEASED.replace("## [0.2.0] - 2026-11-02", heading),
                                   "version heading must be", "0.1.0\n")

    def test_order_and_duplicates(self) -> None:
        swapped = RELEASED.replace("0.2.0", "X").replace("0.1.0", "0.2.0").replace("X", "0.1.0")
        self.assertFinding(swapped, "newest first", "0.1.0\n")
        self.assertFinding(RELEASED.replace("2026-10-30", "2026-11-05"), "newest first", "0.2.0\n")
        self.assertFinding(RELEASED + "\n## [Unreleased]\n", "duplicate heading [Unreleased]",
                           "0.2.0\n")

    def test_release_link_references(self) -> None:
        self.assertFinding(RELEASED.replace(f"[0.1.0]: {TAGS}/v0.1.0\n", ""),
                           "missing link reference for [0.1.0]", "0.2.0\n")
        self.assertFinding(RELEASED.replace(f"{TAGS}/v0.1.0", f"{TAGS}/v0.0.9"),
                           "must end with /releases/tag/v0.1.0", "0.2.0\n")

    def test_link_references_point_to_the_repository(self) -> None:
        # Without --repository only the format is checked.
        self.assertEqual(findings(RELEASED, "0.2.0\n"), [])
        for label in ("Unreleased", "0.2.0", "0.1.0"):
            with self.subTest(label=label):
                self.assertFinding(RELEASED, f"the link reference for [{label}] does not point "
                                   "to other/project", "0.2.0\n", "other/project")
        copied = UNRELEASED_ONLY.replace(REPOSITORY, "example/repository-template")
        self.assertFinding(copied, "does not point to example/repository", None, REPOSITORY)


class ReleaseStateTest(unittest.TestCase):
    def assertFinding(self, found: list[str], text: str) -> None:
        self.assertTrue(any(text in item for item in found), found)

    def test_consistent_state_passes(self) -> None:
        readme = "# Project\n\n- [x] Done\n"
        self.assertEqual(findings(UNRELEASED_ONLY, RELEASING=STATUS_BEFORE, README=readme), [])
        self.assertEqual(findings(RELEASED, "0.2.0\n", RELEASING=STATUS_AFTER, README=readme),
                         [])

    def test_changelog_release_statement(self) -> None:
        self.assertFinding(findings(RELEASED.replace("# Changelog\n", "# Changelog\n\n"
                                                     "No version has been released.\n"),
                                    "0.2.0\n"),
                           "says 'No version has been released' but releases [0.2.0]")
        self.assertFinding(findings(UNRELEASED_ONLY.replace("No version has been released.", "")),
                           "does not say 'No version has been released'")

    def test_placeholder_lines(self) -> None:
        for line in ("- TBD", "_Nothing yet._", "- TODO", "..."):
            with self.subTest(line=line):
                self.assertFinding(findings(UNRELEASED_ONLY.replace("- A change.", line)),
                                   f"placeholder line in a version section: {line}")
        # A placeholder word inside a sentence is not a placeholder line.
        self.assertEqual(findings(UNRELEASED_ONLY.replace("- A change.", "- Remove a TODO.")), [])

    def test_releasing_current_status(self) -> None:
        self.assertFinding(findings(RELEASED, "0.2.0\n", RELEASING=STATUS_BEFORE),
                           "'Released versions' does not name the newest release [0.2.0]: None")
        self.assertFinding(findings(RELEASED, "0.2.0\n", RELEASING=STATUS_BEFORE),
                           "'`VERSION`' does not name the newest release [0.2.0]")
        self.assertFinding(findings(UNRELEASED_ONLY, RELEASING=STATUS_AFTER),
                           "'Released versions' must be None before the first release")
        self.assertFinding(findings(UNRELEASED_ONLY, RELEASING=STATUS_AFTER),
                           "'`VERSION`' must say it does not exist before the first release")
        self.assertFinding(findings(UNRELEASED_ONLY, RELEASING=STATUS_BEFORE.replace(
            "| Released versions | None |\n", "")), "Current status has no 'Released versions' row")
        self.assertFinding(findings(UNRELEASED_ONLY, RELEASING="# Releasing\n"),
                           "RELEASING.md: has no Current status section")

    def test_readme_release_state(self) -> None:
        self.assertFinding(findings(UNRELEASED_ONLY, README="Badge: Release_Candidate\n"),
                           "README.md: mentions a Release Candidate")
        self.assertFinding(findings(UNRELEASED_ONLY, README="## Release\n\n- [ ] Tag it\n"),
                           "README.md: holds an unchecked task-list item")


class VersionTest(unittest.TestCase):
    def test_version_only_from_the_first_release(self) -> None:
        self.assertIn("VERSION exists but CHANGELOG.md has no dated release; it is created "
                      "when a release is prepared (RELEASING.md)",
                      findings(UNRELEASED_ONLY, "0.1.0\n"))
        self.assertIn("VERSION is missing; CHANGELOG.md releases [0.2.0]", findings(RELEASED))

    def test_version_names_the_newest_release(self) -> None:
        self.assertIn("VERSION 0.1.0 differs from the newest CHANGELOG.md release [0.2.0]",
                      findings(RELEASED, "0.1.0\n"))
        for raw in ("0.2.0", "v0.2.0\n", "0.2.0\n\n", "0.2\n"):
            with self.subTest(raw=raw):
                self.assertIn("VERSION must hold one X.Y.Z line ending in a newline",
                              findings(RELEASED, raw))

    def test_main_reports_the_verdict(self) -> None:
        for version, code, verdict in (("0.2.0\n", 0, "changelog: PASS (2 dated release(s)"),
                                       ("0.1.0\n", 1, "changelog: FAIL (2 dated release(s)")):
            with self.subTest(version=version), \
                    tempfile.TemporaryDirectory(prefix="changelog-") as temporary, \
                    contextlib.redirect_stdout(io.StringIO()) as output:
                (Path(temporary) / "CHANGELOG.md").write_text(RELEASED, encoding="utf-8")
                (Path(temporary) / "VERSION").write_text(version, encoding="utf-8")
                self.assertEqual(check_changelog.main(["--root", temporary,
                                                       "--repository", REPOSITORY]), code)
                self.assertIn(verdict, output.getvalue())


if __name__ == "__main__":
    unittest.main()
