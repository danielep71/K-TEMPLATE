#!/usr/bin/env python3
"""Verify, describe and package a release tag for .github/workflows/release.yml.

Subcommands, all read-only apart from the files that bundle writes:

- verify TAG --default-branch BRANCH
  Fails unless:
  - TAG is vX.Y.Z and annotated;
  - X.Y.Z is the newest dated release of CHANGELOG.md and the content of
    VERSION at the tag;
  - the tag's commit is on origin/BRANCH;
  - the tag message names that commit on a "Candidate <sha>." line and has
    a "Deviations:" line, as RELEASING.md requires.
- integrity --repository OWNER/NAME --sha SHA
  Fails unless the check run Repository integrity succeeded on SHA. It reads
  the check runs through the GitHub CLI.
- notes TAG
  Prints the release notes: the tag message, then the changelog section of
  the version.
- bundle REF --name NAME --out DIR
  Writes DIR/NAME-VERSION-source.tar.gz, a gzip of `git archive` of REF with
  no timestamp in the gzip header, and DIR/SHA256SUMS. The same REF gives the
  same bytes. VERSION is the tag's version, or the short commit for another
  ref, as in a rehearsal.

Standard library only, apart from git, and gh for integrity.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import re
import subprocess
import sys
from collections.abc import Sequence
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from check_changelog import CHANGELOG, VERSION, released_versions  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
TAG = re.compile(r"^v(\d+\.\d+\.\d+)$")
CANDIDATE = re.compile(r"^Candidate ([0-9a-f]{40})\.\s*$", re.M)
DEVIATIONS = re.compile(r"^Deviations: \S", re.M)
SECTION = re.compile(r"^## \[(?P<label>[^\]]+)\][^\n]*\n(?P<body>.*?)(?=^## \[|^\[[^\]]+\]: |\Z)",
                     re.M | re.S)
REQUIRED_CHECK = "Repository integrity"


class ReleaseError(Exception):
    """The release cannot proceed; the message says why."""


def git(root: Path, *args: str) -> str:
    completed = subprocess.run(["git", "-C", str(root), *args], capture_output=True,
                               check=False)
    if completed.returncode:
        raise ReleaseError(f"git {' '.join(args)}: "
                           f"{completed.stderr.decode('utf-8', 'replace').strip()}")
    return completed.stdout.decode("utf-8")


def version_of(tag: str) -> str:
    match = TAG.match(tag)
    if not match:
        raise ReleaseError(f"tag {tag!r} is not vX.Y.Z")
    return match.group(1)


def verify(root: Path, tag: str, default_branch: str) -> list[str]:
    """Every reason the tag cannot be released; empty when it can."""
    try:
        version = version_of(tag)
    except ReleaseError as error:
        return [str(error)]
    problems = []
    kind = git(root, "cat-file", "-t", f"refs/tags/{tag}").strip()
    if kind != "tag":
        problems.append(f"tag {tag} is {kind or 'missing'}, not an annotated tag")
    commit = git(root, "rev-parse", f"{tag}^{{commit}}").strip()
    try:
        changelog = git(root, "show", f"{commit}:{CHANGELOG}")
        released = released_versions(changelog)
    except ReleaseError:
        released = []
    if not released or released[0] != version:
        newest = released[0] if released else "none"
        problems.append(f"{CHANGELOG} at {tag} releases {newest} as its newest version, "
                        f"not {version}")
    try:
        recorded = git(root, "show", f"{commit}:{VERSION}").strip()
    except ReleaseError:
        recorded = ""
    if recorded != version:
        problems.append(f"{VERSION} at {tag} is {recorded or 'missing'}, not {version}")
    reachable = subprocess.run(["git", "-C", str(root), "merge-base", "--is-ancestor", commit,
                                f"origin/{default_branch}"], capture_output=True, check=False)
    if reachable.returncode:
        problems.append(f"{tag} ({commit[:12]}) is not on origin/{default_branch}")
    if kind == "tag":
        message = git(root, "tag", "-l", "--format=%(contents)", tag)
        candidates = CANDIDATE.findall(message)
        if candidates != [commit]:
            problems.append(f"the message of {tag} must name its commit on one line "
                            f"'Candidate {commit}.'")
        if not DEVIATIONS.search(message):
            problems.append(f"the message of {tag} has no 'Deviations:' line")
    return problems


def integrity(repository: str, sha: str) -> list[str]:
    """Why Repository integrity did not succeed on sha, if it did not."""
    completed = subprocess.run(
        ["gh", "api", f"repos/{repository}/commits/{sha}/check-runs?per_page=100",
         "--jq", ".check_runs"], capture_output=True, text=True, check=False)
    if completed.returncode:
        return [f"cannot read the check runs of {sha[:12]}: {completed.stderr.strip()}"]
    runs = [run for run in json.loads(completed.stdout or "[]")
            if run.get("name") == REQUIRED_CHECK]
    if not runs:
        return [f"{REQUIRED_CHECK} has not run on {sha[:12]}"]
    if not any(run.get("conclusion") == "success" for run in runs):
        conclusions = sorted({str(run.get("conclusion")) for run in runs})
        return [f"{REQUIRED_CHECK} on {sha[:12]} did not succeed: {', '.join(conclusions)}"]
    return []


def notes(root: Path, tag: str) -> str:
    version = version_of(tag)
    commit = git(root, "rev-parse", f"{tag}^{{commit}}").strip()
    changelog = git(root, "show", f"{commit}:{CHANGELOG}")
    section = next((match.group("body").strip() for match in SECTION.finditer(changelog)
                    if match.group("label") == version), None)
    if section is None:
        raise ReleaseError(f"{CHANGELOG} at {tag} has no section for {version}")
    message = git(root, "tag", "-l", "--format=%(contents)", tag).strip()
    return f"{message}\n\n## Changes\n\n{section}\n"


def bundle(root: Path, ref: str, name: str, out: Path) -> list[Path]:
    """The source archive of ref and its checksums, byte for byte reproducible."""
    commit = git(root, "rev-parse", f"{ref}^{{commit}}").strip()
    match = TAG.match(ref)
    version = match.group(1) if match else commit[:12]
    stem = f"{name}-{version}"
    archive = subprocess.run(["git", "-C", str(root), "archive", "--format=tar",
                              f"--prefix={stem}/", commit], capture_output=True, check=False)
    if archive.returncode:
        raise ReleaseError(f"git archive {ref}: {archive.stderr.decode('utf-8', 'replace')}")
    out.mkdir(parents=True, exist_ok=True)
    target = out / f"{stem}-source.tar.gz"
    with open(target, "wb") as raw, gzip.GzipFile(filename="", mode="wb", fileobj=raw,
                                                  mtime=0, compresslevel=9) as packed:
        packed.write(archive.stdout)
    sums = out / "SHA256SUMS"
    digest = hashlib.sha256(target.read_bytes()).hexdigest()
    sums.write_text(f"{digest}  {target.name}\n", encoding="utf-8")
    return [target, sums]


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    commands = parser.add_subparsers(dest="command", required=True)
    check = commands.add_parser("verify", help="verify a release tag")
    check.add_argument("tag")
    check.add_argument("--default-branch", required=True)
    runs = commands.add_parser("integrity", help="require a successful Repository integrity")
    runs.add_argument("--repository", required=True)
    runs.add_argument("--sha", required=True)
    describe = commands.add_parser("notes", help="print the release notes of a tag")
    describe.add_argument("tag")
    package = commands.add_parser("bundle", help="write the source archive and checksums")
    package.add_argument("ref")
    package.add_argument("--name", required=True)
    package.add_argument("--out", type=Path, required=True)
    args = parser.parse_args(argv)
    root = args.root.resolve()
    try:
        if args.command == "notes":
            sys.stdout.write(notes(root, args.tag))
            return 0
        if args.command == "bundle":
            for path in bundle(root, args.ref, args.name, args.out):
                print(f"wrote {path}")
            return 0
        problems = (verify(root, args.tag, args.default_branch) if args.command == "verify"
                    else integrity(args.repository, args.sha))
    except ReleaseError as error:
        problems = [str(error)]
    for problem in problems:
        print(f"FAIL: {problem}")
    label = args.tag if args.command == "verify" else args.sha[:12]
    print(f"release-{args.command}: {'FAIL' if problems else 'PASS'} ({label})")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
