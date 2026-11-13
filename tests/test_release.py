"""Tests for tools/release.py, in temporary Git repositories with an origin."""

from __future__ import annotations

import argparse
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
from collections.abc import Sequence
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

    def test_deviations_may_be_a_list_on_the_next_lines(self) -> None:
        head = self.git("rev-parse", "HEAD").strip()
        self.tag(message=f"v1.2.0 - Release\n\nCandidate {head}.\nDeviations:\n"
                         "- Invariant 8: the published workbook was not checked.\n"
                         "* Invariant 4: the run names another build.\n")
        self.assertEqual(self.problems(), [])

    def test_deviations_without_text_or_a_list_fail(self) -> None:
        head = self.git("rev-parse", "HEAD").strip()
        self.tag(message=f"v1.2.0 - Release\n\nCandidate {head}.\nDeviations:\n\nSigned.\n")
        self.assertEqual(self.problems(),
                         ["the message of v1.2.0 has no 'Deviations:' line followed by text, "
                          "such as 'none', or by a list"])

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

    def test_verify_published_reads_the_bundle_and_checksums_that_bundle_writes(
            self) -> None:
        self.tag()
        archive, checksums = release.bundle(self.root, "v1.2.0", "project",
                                            self.root.parent / "dist")
        gh = FakeGh()
        gh.files = {"1": archive.read_bytes(), "2": checksums.read_bytes()}
        self.assertEqual([path.name for path in (archive, checksums)],
                         [BUNDLE_NAME, release.SUMS])
        facts = release.verify_published(REPOSITORY, "v1.2.0", runner=gh)
        self.assertEqual([status for status, _ in facts], ["PASS"] * 7, facts)

    def test_main_reports_and_exits(self) -> None:
        self.tag()
        with contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(release.main(["--root", str(self.root), "verify", "v1.2.0",
                                           "--default-branch", "main"]), 0)
            self.assertEqual(release.main(["--root", str(self.root), "verify", "v9.9.9",
                                           "--default-branch", "main"]), 1)
        self.assertIn("release-verify: PASS (v1.2.0)", output.getvalue())
        self.assertIn("FAIL: tag v9.9.9 is missing", output.getvalue())
        self.assertIn("release-verify: FAIL (v9.9.9)", output.getvalue())

    def test_notes_and_bundle_report_a_failure_without_a_traceback(self) -> None:
        # Neither subcommand has --sha: the summary line names its own argument.
        with contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(release.main(["--root", str(self.root), "notes", "v9.9.9"]), 1)
            self.assertEqual(release.main(["--root", str(self.root), "bundle", "no-such-ref",
                                           "--name", "Project", "--out", str(self.root / "out")]),
                             1)
        self.assertIn("release-notes: FAIL (v9.9.9)", output.getvalue())
        self.assertIn("release-bundle: FAIL (no-such-ref)", output.getvalue())


class IntegrityTest(unittest.TestCase):
    def runs(self, payload: object, returncode: int = 0) -> list[str]:
        completed = subprocess.CompletedProcess(
            [], returncode, stdout=json.dumps({"check_runs": payload}).encode(),
            stderr=b"denied")
        with mock.patch("release.subprocess.run", return_value=completed) as run:
            problems = release.integrity("owner/project", "a" * 40)
        self.assertEqual(run.call_args.args[0][:2], ["gh", "api"])
        return problems

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


REPOSITORY = "owner/project"
COMMIT = "c" * 40
BUNDLE_NAME = "project-1.2.0-source.tar.gz"
BUNDLE = b"the source bundle"
WRITES = {"-X", "--method", "-f", "-F", "--field", "--raw-field", "--input"}


def sums(content: bytes, name: str = BUNDLE_NAME) -> bytes:
    return f"{hashlib.sha256(content).hexdigest()}  {name}\n".encode()


def asset(name: str, number: int) -> dict[str, object]:
    return {"name": name,
            "apiUrl": f"https://api.github.com/repos/{REPOSITORY}/releases/assets/{number}",
            "url": f"https://github.com/{REPOSITORY}/releases/download/v1.2.0/{name}"}


class FakeGh:
    """gh for one published release, as JSON fixtures that a test changes."""

    def __init__(self) -> None:
        self.tag_ref: dict[str, object] = {"object": {"type": "tag", "sha": "d" * 40}}
        self.tag: dict[str, object] = {
            "message": f"v1.2.0 - Release\n\nCandidate {COMMIT}.\nDeviations: none.\n",
            "object": {"type": "commit", "sha": COMMIT}}
        self.check_runs: list[dict[str, object]] = [
            {"name": "Repository integrity", "conclusion": "success"}]
        self.release: dict[str, object] | None = {
            "tagName": "v1.2.0", "isDraft": False, "publishedAt": "2026-11-04T10:00:00Z",
            "url": f"https://github.com/{REPOSITORY}/releases/tag/v1.2.0",
            "body": f"v1.2.0 - Release\n\nCandidate {COMMIT}.\nDeviations: none.\n",
            "assets": [asset(BUNDLE_NAME, 1), asset("SHA256SUMS", 2)]}
        self.files: dict[str, bytes] = {"1": BUNDLE, "2": sums(BUNDLE)}
        self.attested = True
        self.milestone: dict[str, object] = {"title": "v1.2.0 - Release", "state": "closed",
                                             "open_issues": 0}
        self.calls: list[list[str]] = []
        self.attested_files: list[tuple[Path, bytes]] = []

    @staticmethod
    def answer(stdout: bytes = b"", error: str = "") -> subprocess.CompletedProcess[bytes]:
        return subprocess.CompletedProcess([], 1 if error else 0, stdout=stdout,
                                           stderr=error.encode())

    def json(self, payload: object) -> subprocess.CompletedProcess[bytes]:
        return self.answer(json.dumps(payload).encode())

    def __call__(self, arguments: Sequence[str]) -> subprocess.CompletedProcess[bytes]:
        args = list(arguments)
        self.calls.append(args)
        base = f"repos/{REPOSITORY}/"
        if args[0] == "api":
            path = args[-1]
            # GitHub resolves an owner and a name in any letter case.
            if path.lower().startswith(base.lower()):
                path = base + path[len(base):]
            if path.lower() == base.rstrip("/").lower():
                return self.json({"full_name": REPOSITORY, "name": "project"})
            if path == base + "git/ref/tags/v1.2.0":
                return self.json(self.tag_ref)
            if path == base + "git/tags/" + "d" * 40:
                return self.json(self.tag)
            if path.startswith(base + "commits/"):
                return self.json({"check_runs": self.check_runs})
            if path.startswith(base + "releases/assets/"):
                number = path.rsplit("/", 1)[-1]
                if number in self.files:
                    return self.answer(self.files[number])
            if path == base + "milestones/7":
                return self.json(self.milestone)
            return self.answer(error="gh: Not Found (HTTP 404)")
        if args[:2] == ["release", "view"]:
            if self.release is None:
                return self.answer(error="release not found")
            return self.json(self.release)
        if args[:2] == ["attestation", "verify"]:
            self.attested_files.append((Path(args[2]), Path(args[2]).read_bytes()))
            if not self.attested:
                return self.answer(error="Error: no matching attestations found")
            return self.json([{"verificationResult": {}}])
        return self.answer(error=f"unexpected gh {' '.join(args)}")


class VerifyPublishedTest(unittest.TestCase):
    def setUp(self) -> None:
        self.gh = FakeGh()

    def facts(self, milestone: int | None = 7,
              extra: dict[str, str] | None = None) -> list[tuple[str, str]]:
        found = release.verify_published(REPOSITORY, "v1.2.0", milestone, extra, self.gh)
        for args in self.gh.calls:
            with self.subTest(call=args):
                self.assertIn(args[0] if args[0] == "api" else " ".join(args[:2]),
                              {"api", "release view", "attestation verify"})
                self.assertFalse(WRITES & set(args), args)
        return found

    def failures(self, milestone: int | None = 7,
                 extra: dict[str, str] | None = None) -> list[str]:
        return [text for status, text in self.facts(milestone, extra) if status == "FAIL"]

    def test_a_published_release_passes_with_one_line_per_fact(self) -> None:
        facts = self.facts()
        self.assertEqual([status for status, _ in facts], ["PASS"] * 8, facts)
        text = "\n".join(line for _, line in facts)
        for expected in (f"tag v1.2.0 is annotated and points at {COMMIT}",
                         f"Repository integrity succeeded on {COMMIT}",
                         "the release of v1.2.0 is published, not a draft",
                         f"the release notes name the candidate {COMMIT}",
                         f"exactly the expected assets: {BUNDLE_NAME}, SHA256SUMS",
                         f"{BUNDLE_NAME} has SHA-256 {hashlib.sha256(BUNDLE).hexdigest()}, "
                         "as SHA256SUMS records",
                         f"gh attestation verify passed on {BUNDLE_NAME}",
                         "milestone 7 (v1.2.0 - Release) is closed and has no open issue"):
            with self.subTest(expected=expected):
                self.assertIn(expected, text)

    def test_the_attestation_is_verified_on_a_removed_copy_of_the_published_bundle(
            self) -> None:
        self.facts()
        [(copy, content)] = self.gh.attested_files
        self.assertEqual(content, BUNDLE)
        self.assertFalse(copy.exists())
        call = next(args for args in self.gh.calls if args[0] == "attestation")
        for option in (["--repo", REPOSITORY],
                       ["--signer-workflow", f"{REPOSITORY}/.github/workflows/release.yml"],
                       ["--source-ref", "refs/tags/v1.2.0"], ["--format", "json"]):
            with self.subTest(option=option):
                index = call.index(option[0])
                self.assertEqual(call[index:index + 2], option)

    def test_a_draft_release_fails(self) -> None:
        assert self.gh.release is not None
        self.gh.release["isDraft"] = True
        self.assertEqual(self.failures(), ["the release of v1.2.0 is a draft, not published"])

    def test_a_missing_release_fails(self) -> None:
        self.gh.release = None
        self.assertEqual(self.failures(),
                         ["cannot read the release of v1.2.0: release not found"])

    def test_a_tag_that_points_elsewhere_than_its_candidate_fails(self) -> None:
        self.gh.tag["object"] = {"type": "commit", "sha": "e" * 40}
        self.assertIn(f"tag v1.2.0 points at {'e' * 40}, but its Candidate line names "
                      f"{COMMIT}", self.failures())

    def test_a_lightweight_or_missing_tag_fails(self) -> None:
        self.gh.tag_ref = {"object": {"type": "commit", "sha": COMMIT}}
        self.assertIn("tag v1.2.0 is commit, not an annotated tag", self.failures())
        self.gh.tag_ref = {}
        self.assertIn("tag v1.2.0 is missing, not an annotated tag", self.failures())
        self.assertEqual(release.verify_published(REPOSITORY, "1.2.0", runner=self.gh),
                         [("FAIL", "tag '1.2.0' is not vX.Y.Z")])

    def test_a_failed_integrity_run_fails(self) -> None:
        self.gh.check_runs = [{"name": "Repository integrity", "conclusion": "failure"}]
        self.assertEqual(self.failures(),
                         [f"Repository integrity on {COMMIT[:12]} did not succeed: failure"])

    def test_notes_that_do_not_name_the_candidate_fail(self) -> None:
        assert self.gh.release is not None
        self.gh.release["body"] = "Notes."
        self.assertEqual(self.failures(),
                         [f"the release notes do not name the candidate {COMMIT}"])

    def test_a_missing_or_unexpected_asset_fails(self) -> None:
        assert self.gh.release is not None
        self.gh.release["assets"] = [asset(BUNDLE_NAME, 1), asset("notes.txt", 3)]
        self.assertEqual(self.failures(), ["asset SHA256SUMS is missing from the release",
                                           "asset notes.txt is not an expected asset"])

    def test_an_altered_asset_fails(self) -> None:
        self.gh.files["1"] = b"an altered bundle"
        digest = hashlib.sha256(b"an altered bundle").hexdigest()
        self.assertIn(f"{BUNDLE_NAME} has SHA-256 {digest}, but SHA256SUMS records "
                      f"{hashlib.sha256(BUNDLE).hexdigest()}", self.failures())

    def test_checksums_must_record_exactly_the_bundle(self) -> None:
        self.gh.files["2"] = sums(BUNDLE) + sums(b"x", "other.zip")
        self.assertEqual(self.failures(), [f"SHA256SUMS records other.zip, {BUNDLE_NAME}, not "
                                           f"exactly {BUNDLE_NAME}"])

    def test_an_unreadable_asset_fails(self) -> None:
        del self.gh.files["2"]
        self.assertEqual(self.failures(),
                         ["cannot read asset SHA256SUMS: gh: Not Found (HTTP 404)"])

    def test_other_assets_are_expected_with_their_recorded_digests(self) -> None:
        assert self.gh.release is not None
        workbook = b"the production workbook"
        self.gh.release["assets"] = [asset(BUNDLE_NAME, 1), asset("SHA256SUMS", 2),
                                     asset("Workbook.xlsm", 4)]
        self.gh.files["4"] = workbook
        digest = hashlib.sha256(workbook).hexdigest()
        self.assertEqual(self.failures(extra={"Workbook.xlsm": digest}), [])
        self.assertEqual(self.failures(extra={"Workbook.xlsm": "0" * 64}),
                         [f"Workbook.xlsm has SHA-256 {digest}, not the recorded {'0' * 64}"])
        self.assertEqual(self.failures(), ["asset Workbook.xlsm is not an expected asset"])
        self.assertIn("asset Manual.pdf is missing from the release",
                      self.failures(extra={"Workbook.xlsm": digest, "Manual.pdf": "1" * 64}))

    def test_a_failed_attestation_fails(self) -> None:
        self.gh.attested = False
        self.assertEqual(self.failures(), [f"gh attestation verify failed on {BUNDLE_NAME}: "
                                           "Error: no matching attestations found"])

    def test_an_unattested_bundle_stated_in_the_notes_is_a_note(self) -> None:
        assert self.gh.release is not None
        self.gh.release["body"] = f"{self.gh.release['body']}\n{release.UNATTESTED}: this ...\n"
        self.gh.attested = False
        facts = self.facts()
        self.assertIn(("NOTE", f"{BUNDLE_NAME} is not attested, as the release notes state: "
                               "record its absence as a deviation"), facts)
        self.assertNotIn("FAIL", [status for status, _ in facts])
        self.assertEqual(self.gh.attested_files, [])

    def test_an_open_milestone_fails(self) -> None:
        self.gh.milestone = {"title": "v1.2.0 - Release", "state": "open", "open_issues": 0}
        self.assertEqual(self.failures(),
                         ["milestone 7 (v1.2.0 - Release) is open, with 0 open issue(s)"])
        self.gh.milestone["state"] = "closed"
        self.gh.milestone["open_issues"] = 2
        self.assertEqual(self.failures(),
                         ["milestone 7 (v1.2.0 - Release) is closed, with 2 open issue(s)"])
        self.assertEqual(self.failures(milestone=8),
                         ["cannot read milestone 8: gh: Not Found (HTTP 404)"])

    def test_the_repository_is_checked_under_its_canonical_name(self) -> None:
        found = release.verify_published("Owner/Project", "v1.2.0", 7, None, self.gh)
        self.assertEqual([status for status, _ in found], ["PASS"] * 8, found)
        self.assertIn(f"exactly the expected assets: {BUNDLE_NAME}, SHA256SUMS",
                      "\n".join(text for _, text in found))
        call = next(args for args in self.gh.calls if args[0] == "attestation")
        self.assertEqual(call[call.index("--repo") + 1], REPOSITORY)
        self.assertEqual(call[call.index("--signer-workflow") + 1],
                         f"{REPOSITORY}/.github/workflows/release.yml")
        self.gh.calls.clear()
        self.assertEqual(release.verify_published("owner/missing", "v1.2.0", runner=self.gh),
                         [("FAIL", "cannot read repository owner/missing: "
                                   "gh: Not Found (HTTP 404)")])

    def test_without_a_milestone_none_is_read(self) -> None:
        self.assertEqual(self.failures(milestone=None), [])
        self.assertFalse(any("milestones" in arg for args in self.gh.calls for arg in args))

    def test_main_prints_each_fact_and_exits(self) -> None:
        arguments = ["verify-published", "v1.2.0", "--repository", REPOSITORY,
                     "--milestone", "7"]
        with mock.patch("release.run_gh", self.gh), \
                contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(release.main(arguments), 0)
            self.gh.attested = False
            self.assertEqual(release.main(arguments), 1)
        lines = output.getvalue().splitlines()
        self.assertEqual(lines[8], "release-verify-published: PASS (v1.2.0)")
        self.assertTrue(all(line.startswith("PASS: ") for line in lines[:8]), lines)
        self.assertIn(f"FAIL: gh attestation verify failed on {BUNDLE_NAME}: "
                      "Error: no matching attestations found", lines)
        self.assertEqual(lines[-1], "release-verify-published: FAIL (v1.2.0)")

    def test_an_asset_record_is_file_equals_sha256(self) -> None:
        self.assertEqual(release.asset_record("a=b.xlsm=" + "A" * 64), ("a=b.xlsm", "a" * 64))
        for value in ("Workbook.xlsm", "=" + "a" * 64, "Workbook.xlsm=abc"):
            with self.subTest(value=value), self.assertRaises(argparse.ArgumentTypeError):
                release.asset_record(value)


class WorkflowTest(unittest.TestCase):
    def test_the_workflow_states_an_unattested_bundle_as_verify_published_reads_it(
            self) -> None:
        text = (ROOT / ".github/workflows/release.yml").read_text(encoding="utf-8")
        self.assertIn(release.UNATTESTED, text)
        self.assertIn(f"dist/{release.SUMS}\n", text)

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
