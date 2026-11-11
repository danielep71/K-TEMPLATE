#!/usr/bin/env python3
"""Adapt the files of a repository created from the template, once, in one diff.

It makes the edits that README.md, *Start a new repository*, step 2, lists
and that **Repository integrity** requires with --template false:

- AGENTS.md: delete the template-only section, with its markers;
- .github/ISSUE_TEMPLATE/config.yml: point the contact links at the new
  repository;
- README.md: point the traffic badges at the new repository, or remove them
  with --remove-traffic-badges;
- .github/dependabot.yml: set target-branch to --release-branch, or remove it;
- CHANGELOG.md: keep the introduction and the policy sections, and replace the
  entries and link references with an empty [Unreleased] section of the new
  repository; with them go VERSION and the release rows of the *Current
  status* of RELEASING.md, which describe the template's releases;
- docs/VBA_HOUSE_STYLE.md: replace <REGIME> in the naming table with
  --regime.

Without --apply it only lists the changes. With --apply it writes them, and
refuses a working tree that differs from the commit, so that the result is
one reviewable diff. It changes nothing on GitHub and commits nothing. A file
already adapted is left as it is, so a second run changes nothing.
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))

from check_changelog import NO_RELEASE, VERSION_HEADING, foreign_references  # noqa: E402
from check_repository import (  # noqa: E402
    BADGE,
    CONTACT_URL,
    TARGET_BRANCH,
    MissingContext,
    origin_has_branch,
)
from vba_lexer import FACADE_ROW, HOUSE_STYLE, RESERVED_PREFIXES  # noqa: E402

REPOSITORY = re.compile(r"[\w.-]+/[\w.-]+")
REGIME = re.compile(r"[A-Z][A-Z0-9]*")
BRANCH = re.compile(r"[\w.-]+(?:/[\w.-]+)*")
SECTION_START = "<!-- template-only:start -->"
SECTION_END = "<!-- template-only:end -->"
GITHUB_URL = re.compile(r"https://github\.com/[\w.-]+/[\w.-]+/")
DIRECTORY = re.compile(r"^(\s*)directory:.*$", re.M)
STATUS_ROWS = {
    "Released versions": "None",
    "`VERSION`": "Does not exist yet; it is created when the first release is prepared",
}


class InitializeError(Exception):
    pass


@dataclass(frozen=True)
class Options:
    repository: str
    regime: str
    release_branch: str | None = None
    remove_traffic_badges: bool = False
    default_branch: str = "main"


@dataclass
class Change:
    """A file's new content (None deletes it) and what changes in it."""

    path: str
    content: str | None
    notes: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# One function per file: the adapted text and what changed, or no notes.

Adapted = tuple[str, list[str]]
Adapter = Callable[[str, Options], Adapted]


def adapt_agents(text: str, options: Options) -> Adapted:
    if SECTION_START not in text:
        return text, []
    before, rest = text.split(SECTION_START, 1)
    if SECTION_END not in rest:
        raise InitializeError(f"AGENTS.md: {SECTION_START} has no {SECTION_END}")
    after = rest.split(SECTION_END, 1)[1].lstrip("\n")
    adapted = before.rstrip("\n") + "\n" + ("\n" + after if after.strip() else "")
    return adapted, ["delete the template-only section, with its markers"]


def adapt_contact_links(text: str, options: Options) -> Adapted:
    own = f"https://github.com/{options.repository}/"
    notes = []
    for url in CONTACT_URL.findall(text):
        if url.lower().startswith(own.lower()):
            continue
        if not GITHUB_URL.match(url):
            raise InitializeError(f".github/ISSUE_TEMPLATE/config.yml: contact link {url} is "
                                  f"not a GitHub repository link; point it at "
                                  f"{options.repository} by hand")
        new = GITHUB_URL.sub(own, url, count=1)
        text = text.replace(url, new)
        notes.append(f"point contact link {url} at {new}")
    return text, notes


def adapt_badges(text: str, options: Options) -> Adapted:
    foreign = sorted({name for name in BADGE.findall(text)
                      if name.lower() != options.repository.lower()})
    if options.remove_traffic_badges:
        lines = text.split("\n")
        badges = sum(1 for line in lines if BADGE.search(line))
        if not badges:
            return text, []
        kept: list[str] = []
        removed = False
        for line in lines:
            if BADGE.search(line):
                removed = True
            # Of the blank lines around a removed badge block, keep one.
            elif not (removed and not line.strip() and kept and not kept[-1].strip()):
                removed = removed and not line.strip()
                kept.append(line)
        return "\n".join(kept), [f"remove the {badges} traffic badge line(s)"]
    if not foreign:
        return text, []
    adapted = BADGE.sub(f"raw.githubusercontent.com/{options.repository}/traffic-history/",
                        text)
    return adapted, [f"point the traffic badges at {options.repository}, not "
                     f"{', '.join(foreign)}"]


def adapt_dependabot(text: str, options: Options) -> Adapted:
    branch = options.release_branch
    current = TARGET_BRANCH.findall(text)
    if branch is None:
        if not current:
            return text, []
        adapted = re.sub(r"^\s*target-branch:.*\n", "", text, flags=re.M)
        return adapted, [f"remove target-branch ({', '.join(sorted(set(current)))}), so that "
                         "updates target the default branch"]
    if current:
        if set(current) == {branch}:
            return text, []
        adapted = re.sub(r"^(\s*target-branch:\s*).*$", rf"\g<1>{branch}", text, flags=re.M)
        return adapted, [f"set target-branch to {branch}, not "
                         f"{', '.join(sorted(set(current)))}"]
    adapted = DIRECTORY.sub(lambda match: f"{match.group(0)}\n{match.group(1)}target-branch: "
                            f"{branch}", text)
    return adapted, [f"add target-branch {branch} to each update"]


def adapt_changelog(text: str, options: Options) -> Adapted:
    # A link reference to another repository marks the template's changelog.
    foreign = foreign_references(text, options.repository, set())
    if not foreign:
        return text, []
    first = VERSION_HEADING.search(text)
    kept = text[:first.start()] if first else text
    introduction, heading, rest = kept.partition("\n## ")
    if NO_RELEASE not in introduction:
        introduction = f"{introduction.rstrip()} **{NO_RELEASE}.**\n"
    kept = (introduction + heading + rest).rstrip()
    adapted = (f"{kept}\n\n## [Unreleased]\n\n[Unreleased]: "
               f"https://github.com/{options.repository}/commits/{options.default_branch}\n")
    entries = len(VERSION_HEADING.findall(text))
    return adapted, [f"replace its {entries} version section(s) and {len(foreign)} link "
                     "reference(s) to another repository with an empty [Unreleased] section; "
                     "the introduction and the policy stay"]


def adapt_releasing(text: str, options: Options) -> Adapted:
    notes = []
    for name, state in STATUS_ROWS.items():
        row = re.compile(rf"^(\|\s*{re.escape(name)}\s*\|)(.*)\|\s*$", re.M)
        match = row.search(text)
        if match and match.group(2).strip() != state:
            text = row.sub(lambda found: f"{found.group(1)} {state} |", text, count=1)
            notes.append(f"set the Current status row {name} to: {state}")
    return text, notes


def adapt_house_style(text: str, options: Options) -> Adapted:
    lines = text.split("\n")
    rows = [index for index, line in enumerate(lines)
            if line.startswith("|") and "<REGIME>" in line]
    if not rows:
        match = FACADE_ROW.search(text)
        if match and match.group(1) != options.regime:
            raise InitializeError(f"{HOUSE_STYLE}: the facade prefix is already "
                                  f"{match.group(1)}_, not {options.regime}_; change it by hand")
        return text, []
    for index in rows:
        lines[index] = lines[index].replace("<REGIME>", options.regime)
    return "\n".join(lines), [f"replace <REGIME> with {options.regime} in {len(rows)} row(s) of "
                              "the naming table"]


ADAPTERS: tuple[tuple[str, Adapter], ...] = (
    ("AGENTS.md", adapt_agents),
    (".github/ISSUE_TEMPLATE/config.yml", adapt_contact_links),
    ("README.md", adapt_badges),
    (".github/dependabot.yml", adapt_dependabot),
    ("CHANGELOG.md", adapt_changelog),
    (HOUSE_STYLE, adapt_house_style),
)


# ---------------------------------------------------------------------------
# Planning and applying.

def read(root: Path, path: str) -> str | None:
    file = root / path
    # Bytes, so that line endings are kept as they are on every platform.
    return file.read_bytes().decode("utf-8") if file.is_file() else None


def plan(root: Path, options: Options) -> list[Change]:
    changes = []
    for path, adapt in ADAPTERS:
        text = read(root, path)
        if text is None:
            continue
        adapted, notes = adapt(text, options)
        if notes:
            changes.append(Change(path, adapted, notes))
        if path == "CHANGELOG.md" and notes:
            # The template's release state goes with its entries.
            if (root / "VERSION").is_file():
                changes.append(Change("VERSION", None, ["delete it: no version of this "
                                                        "repository has been released"]))
            releasing = read(root, "RELEASING.md")
            if releasing is not None:
                adapted, notes = adapt_releasing(releasing, options)
                if notes:
                    changes.append(Change("RELEASING.md", adapted, notes))
    return changes


def git(root: Path, *args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(["git", *args], cwd=root, capture_output=True, text=True,
                          check=False)


def dirty(root: Path) -> list[str]:
    status = git(root, "status", "--porcelain", "--untracked-files=all")
    if status.returncode:
        raise InitializeError(f"{root} is not a Git working tree")
    return [line for line in status.stdout.splitlines() if line]


def warnings(root: Path, options: Options) -> list[str]:
    """Whether --release-branch exists on origin, asked as dependabot-target-branches asks it.

    The check looks for the remote-tracking branch, and else asks origin, so a
    single-branch or shallow clone of a repository that has the branch needs
    no warning.
    """
    branch = options.release_branch
    if not branch:
        return []
    try:
        if origin_has_branch(root, branch):
            return []
    except MissingContext as missing:
        return [f"{missing}: make sure origin has it before Repository integrity runs, or "
                "the dependabot-target-branches check fails there"]
    return [f"origin has no branch {branch}: push it before Repository integrity runs, or the "
            "dependabot-target-branches check fails"]


def apply(root: Path, changes: Sequence[Change]) -> None:
    for change in changes:
        path = root / change.path
        if change.content is None:
            path.unlink()
        else:
            path.write_bytes(change.content.encode("utf-8"))


def validate(options: Options) -> None:
    if not REPOSITORY.fullmatch(options.repository):
        raise InitializeError(f"--repository must be OWNER/NAME, not {options.repository!r}")
    if not REGIME.fullmatch(options.regime) or options.regime in RESERVED_PREFIXES:
        raise InitializeError(f"--regime must be upper-case letters and digits, starting with "
                              f"a letter, and not one of {', '.join(RESERVED_PREFIXES)}: "
                              f"{options.regime!r}")
    for name, branch in (("--release-branch", options.release_branch),
                         ("--default-branch", options.default_branch)):
        if branch is not None and not BRANCH.fullmatch(branch):
            raise InitializeError(f"{name} is not a branch name: {branch!r}")


def run(root: Path, options: Options, write: bool) -> int:
    validate(options)
    if write and (changed := dirty(root)):
        raise InitializeError("the working tree differs from the commit; commit or discard "
                              "these first, so that the initialization is one diff:\n  "
                              + "\n  ".join(changed))
    changes = plan(root, options)
    for warning in warnings(root, options):
        print(f"WARNING: {warning}")
    if not changes:
        print(f"initialize-repository: nothing to change; the files are already initialized "
              f"for {options.repository}")
        return 0
    for change in changes:
        for note in change.notes:
            print(f"- {change.path}: {note}")
    count = f"{len(changes)} file(s)"
    if write:
        apply(root, changes)
        print(f"initialize-repository: changed {count} for {options.repository}; review them "
              "with git diff and commit them. Nothing on GitHub was changed.")
    else:
        print(f"initialize-repository: dry run, nothing written; --apply changes {count}.")
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--repository", required=True, metavar="OWNER/NAME",
                        help="the new repository on GitHub")
    parser.add_argument("--regime", required=True, metavar="PREFIX",
                        help="the project's short name for the VBA facade prefix, such as LCR")
    parser.add_argument("--release-branch", metavar="BRANCH",
                        help="Dependabot's target-branch (default: none, so the default branch)")
    parser.add_argument("--default-branch", default="main",
                        help="the branch whose commits the changelog's [Unreleased] link shows "
                             "(default: main)")
    parser.add_argument("--remove-traffic-badges", action="store_true",
                        help="remove the README's traffic badges instead of repointing them")
    parser.add_argument("--apply", action="store_true", help="write the changes")
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    options = Options(args.repository, args.regime, args.release_branch,
                      args.remove_traffic_badges, args.default_branch)
    try:
        return run(args.root.resolve(), options, args.apply)
    except InitializeError as error:
        print(f"initialize-repository: FAIL: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
