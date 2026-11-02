"""Tests for tools/release.py, in temporary Git repositories with an origin."""

from __future__ import annotations

import contextlib
import gzip
import hashlib
import io
import json
import shutil
import subprocess
import sys
import tarfile
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import release  # noqa: E402

try:
    import yaml
except ImportError:  # pragma: no cover - installed from .github/requirements.txt
    yaml = None

CHANGELOG = """# Changelog

## [Unreleased]

## [1.2.0] - 2026-11-02

### Added

- A feature.

## [1.1.0] - 2026-10-01

### Fixed

- A bug.

[Unreleased]: https://github.com/owner/project/compare/v1.2.0...HEAD
[1.2.0]: https://github.com/owner/project/releases/tag/v1.2.0
[1.1.0]: https://github.com/owner/project/releases/tag/v1.1.0
"""


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class ReleaseTest(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="release-")
        self.addCleanup(directory.cleanup)
        base = Path(directory.name)
        self.origin = base / "origin.git"
        self.root = base / "work"
        subprocess.run(["git", "init", "-q", "--bare", "-b", "main", str(self.origin)],
                       check=True)
        self.root.mkdir()
        for command in (["init", "-q", "-b", "main"], ["config", "user.name", "Fixture"],
                        ["config", "user.email", "fixture@example.invalid"],
                        ["config", "commit.gpgsign", "false"],
                        ["config", "tag.gpgsign", "false"],
                        ["remote", "add", "origin", str(self.origin)]):
            self.git(*command)
        self.commit({"CHANGELOG.md": CHANGELOG, "VERSION": "1.2.0\n", "README.md": "x\n"})
        self.git("push", "-q", "origin", "main")
        self.sha = self.git("rev-parse", "HEAD").strip()

    def git(self, *args: str) -> str:
        return subprocess.run(["git", "-C", str(self.root), *args], capture_output=True,
                              text=True, check=True).stdout

    def commit(self, files: dict[str, str]) -> None:
        for name, text in files.items():
            (self.root / name).write_text(text, encoding="utf-8")
        self.git("add", "-A")
        self.git("commit", "-qm", "Fixture")

    def tag(self, name: str = "v1.2.0", message: str | None = None) -> None:
        text = message if message is not None else (
            f"{name} - Release\n\nCandidate {self.git('rev-parse', 'HEAD').strip()}.\n"
            "Deviations: none.\n")
        self.git("tag", "-a", name, "-m", text)

    def problems(self, tag: str = "v1.2.0") -> list[str]:
        return release.verify(self.root, tag, "main")

    def test_a_consistent_tag_passes(self) -> None:
        self.tag()
        self.assertEqual(self.problems(), [])

    def test_the_tag_must_be_annotated_and_well_named(self) -> None:
        self.git("tag", "v1.2.0")
        self.assertTrue(any("not an annotated tag" in p for p in self.problems()))
        self.assertEqual(self.problems("1.2.0"), ["tag '1.2.0' is not vX.Y.Z"])

    def test_the_version_must_match_the_changelog_and_version(self) -> None:
        self.tag("v1.3.0", f"v1.3.0\n\nCandidate {self.sha}.\nDeviations: none.\n")
        found = self.problems("v1.3.0")
        self.assertTrue(any("releases 1.2.0 as its newest version, not 1.3.0" in p
                            for p in found), found)
        self.assertTrue(any("VERSION at v1.3.0 is 1.2.0, not 1.3.0" in p for p in found), found)

    def test_version_must_exist(self) -> None:
        self.git("rm", "-q", "VERSION")
        self.git("commit", "-qm", "Drop VERSION")
        self.git("push", "-q", "origin", "main")
        self.tag()
        self.assertEqual(self.problems(), ["VERSION at v1.2.0 is missing, not 1.2.0"])

    def test_the_commit_must_be_on_the_default_branch(self) -> None:
        self.git("switch", "-q", "-c", "side")
        self.commit({"README.md": "side\n"})
        self.tag()
        self.assertTrue(any("is not on origin/main" in p for p in self.problems()))

    def test_the_message_must_name_the_candidate_and_its_deviations(self) -> None:
        self.tag(message="v1.2.0\n\nCandidate " + "0" * 40 + ".\n")
        found = self.problems()
        self.assertTrue(any(f"'Candidate {self.sha}.'" in p for p in found), found)
        self.assertTrue(any("no 'Deviations:' line" in p for p in found), found)

    def test_notes_hold_the_message_and_the_changelog_section(self) -> None:
        self.tag()
        text = release.notes(self.root, "v1.2.0")
        self.assertTrue(text.startswith("v1.2.0 - Release\n\nCandidate "))
        self.assertIn("## Changes\n\n### Added\n\n- A feature.\n", text)
        self.assertNotIn("A bug.", text)
        self.assertNotIn("releases/tag", text)
        with self.assertRaisesRegex(release.ReleaseError, "no section for 1.3.0"):
            self.git("tag", "-a", "v1.3.0", "-m", "x")
            release.notes(self.root, "v1.3.0")

    def test_bundle_is_reproducible(self) -> None:
        self.tag()
        first = release.bundle(self.root, "v1.2.0", "project", self.root.parent / "one")
        second = release.bundle(self.root, "v1.2.0", "project", self.root.parent / "two")
        self.assertEqual([path.name for path in first],
                         ["project-1.2.0-source.tar.gz", "SHA256SUMS"])
        self.assertEqual(first[0].read_bytes(), second[0].read_bytes())
        digest = hashlib.sha256(first[0].read_bytes()).hexdigest()
        self.assertEqual(first[1].read_text(encoding="utf-8"),
                         f"{digest}  project-1.2.0-source.tar.gz\n")
        with gzip.open(first[0]) as packed, tarfile.open(fileobj=io.BytesIO(packed.read())) as tar:
            self.assertIn("project-1.2.0/CHANGELOG.md", tar.getnames())
        other = release.bundle(self.root, "HEAD", "project", self.root.parent / "three")
        self.assertEqual(other[0].name, f"project-{self.sha[:12]}-source.tar.gz")

    def test_main_reports_and_exits(self) -> None:
        self.tag()
        with contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(release.main(["--root", str(self.root), "verify", "v1.2.0",
                                           "--default-branch", "main"]), 0)
            self.assertEqual(release.main(["--root", str(self.root), "verify", "v9.9.9",
                                           "--default-branch", "main"]), 1)
        self.assertIn("release-verify: PASS (v1.2.0)", output.getvalue())
        self.assertIn("release-verify: FAIL (v9.9.9)", output.getvalue())


class IntegrityTest(unittest.TestCase):
    def runs(self, payload: object, returncode: int = 0) -> list[str]:
        completed = subprocess.CompletedProcess([], returncode, stdout=json.dumps(payload),
                                                stderr="denied")
        with mock.patch.object(release.subprocess, "run", return_value=completed):
            return release.integrity("owner/project", "a" * 40)

    def test_a_successful_run_passes(self) -> None:
        self.assertEqual(self.runs([{"name": "Repository integrity", "conclusion": "success"},
                                    {"name": "CodeQL", "conclusion": "failure"}]), [])

    def test_missing_failed_or_unreadable_runs_fail(self) -> None:
        self.assertEqual(self.runs([{"name": "CodeQL", "conclusion": "success"}]),
                         ["Repository integrity has not run on aaaaaaaaaaaa"])
        self.assertEqual(self.runs([{"name": "Repository integrity", "conclusion": "failure"}]),
                         ["Repository integrity on aaaaaaaaaaaa did not succeed: failure"])
        self.assertEqual(self.runs([], returncode=1),
                         ["cannot read the check runs of aaaaaaaaaaaa: denied"])


class WorkflowTest(unittest.TestCase):
    def test_the_workflow_calls_the_tool_it_is_tested_with(self) -> None:
        text = (ROOT / ".github/workflows/release.yml").read_text(encoding="utf-8")
        for command in ("python tools/release.py verify", "python tools/release.py integrity",
                        "python tools/release.py bundle", "python tools/release.py notes",
                        "gh release create \"$TAG\" --draft --verify-tag"):
            with self.subTest(command=command):
                self.assertIn(command, text)

    @unittest.skipIf(yaml is None, "needs PyYAML from .github/requirements.txt")
    def test_a_repository_without_attestations_still_gets_its_draft(self) -> None:
        assert yaml is not None
        jobs = yaml.safe_load((ROOT / ".github/workflows/release.yml")
                              .read_text(encoding="utf-8"))["jobs"]
        attest, draft = jobs["attest"]["if"], jobs["draft"]["if"]
        self.assertIn("github.event.repository.private == false", attest)
        self.assertIn("vars.RELEASE_ATTESTATIONS == 'enabled'", attest)
        self.assertEqual(jobs["draft"]["needs"], ["build", "attest"])
        for condition in ("!cancelled()", "needs.build.result == 'success'",
                          "needs.attest.result == 'success'",
                          "needs.attest.result == 'skipped'"):
            with self.subTest(condition=condition):
                self.assertIn(condition, draft)
        script = jobs["draft"]["steps"][-1]["run"]
        self.assertIn('if [ "$ATTESTED" != "success" ]', script)
        self.assertLess(script.index("no build provenance attestation"),
                        script.index("gh release create"))


if __name__ == "__main__":
    unittest.main()
