#!/usr/bin/env python3
"""Run every Repository integrity check, the same way locally and in CI.

Each check is a command or a small function. The run writes, under
test-results/integrity/:

- one <check>.log per check, with its command and output;
- report.json: the commit, whether the tree is dirty, the pull request head,
  the Python version and platform, the SHA-256 of every tracked file, and each
  check's command, status and exit code;
- summary.md, which --summary also appends to a file such as
  $GITHUB_STEP_SUMMARY.

A check whose tool or GitHub context is missing is skipped and reported as
such, as is a check that cannot reach the origin remote; --strict, which CI
passes, makes it fail instead. An advisory check reports its findings as a
warning and does not fail the run. --ci, which CI also passes, adds a last
check that fails when the working tree differs from the commit, so that the
result describes the commit and nothing a step changed. The run covers the
repository only: it compiles no VBA, runs no Excel and certifies no figure.
"""

from __future__ import annotations

import argparse
import datetime
import fnmatch
import hashlib
import importlib
import importlib.util
import json
import os
import platform
import re
import shutil
import subprocess
import sys
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
REPORT_DIR = Path("test-results/integrity")
SCOPE = "Repository checks only: no VBA compilation, no Excel run, no numerical certification"
TIMEOUT_SECONDS = 600
# How long to wait for the origin remote to list a branch.
REMOTE_TIMEOUT_SECONDS = 60
REQUIRED_DOCUMENTS = ("README.md", "LICENSE", "docs/repo_governance/SETTINGS.md")
BADGE = re.compile(r"raw\.githubusercontent\.com/([^/\s]+/[^/\s]+)/traffic-history/")
CONTACT_URL = re.compile(r'^\s*url:\s*"?([^"\s]+)"?\s*$', re.M)
TARGET_BRANCH = re.compile(r'^\s*target-branch:\s*"?([^"\s]+)"?\s*$', re.M)
GITHUB_REMOTE = re.compile(r"github\.com[:/]([\w.-]+/[\w.-]+?)(?:\.git)?/?$")
# File names that mark a secret or a private key, matched case-insensitively
# against the last path component. A reviewed list: .gitignore ignores the same
# names, but git add -f bypasses it, and secret scanning matches token
# patterns, not file types.
SECRET_FILES = (".env", ".env.*", "*.pem", "*.key", "*.p12", "*.pfx", "*.pvk", "*.jks",
                "*.keystore", "id_rsa*", "id_ed25519*", "*.kdbx")
# A sanitized example and public SSH keys stay trackable.
SECRET_FILE_EXCEPTIONS = (".env.example", "id_rsa*.pub", "id_ed25519*.pub")

Finding = tuple[str, str]


class MissingContext(Exception):
    """Raised by a check function that cannot run here; the message says why.

    The run reports it as it reports a Check's missing reason: skipped, or
    failed under --strict.
    """


@dataclass
class Context:
    """What the checks may need to know about the repository and the run."""

    root: Path
    repository: str | None = None
    default_branch: str | None = None
    is_template: bool | None = None
    # Whether to fail a working tree that differs from the commit (--ci).
    ci: bool = False


@dataclass
class Check:
    """A command to run, or a function returning findings as (file, message)."""

    name: str
    command: list[str] | None = None
    function: Callable[[Context], list[Finding]] | None = None
    # Why the check cannot run here, if it cannot.
    missing: str | None = None
    # Findings are reported as a warning and do not fail the run.
    advisory: bool = False


@dataclass
class Result:
    name: str
    status: str
    command: list[str] | None
    exit_code: int | None
    detail: str = ""
    log: str = field(default="", repr=False)


def git(root: Path, *args: str) -> str:
    return subprocess.run(["git", *args], cwd=root, capture_output=True, text=True,
                          check=True).stdout


# ---------------------------------------------------------------------------
# Checks that read files or Git, without a separate tool.

def template_only(context: Context) -> list[Finding]:
    """A generated repository has deleted the template-only section of AGENTS.md."""
    text = (context.root / "AGENTS.md").read_text(encoding="utf-8") \
        if (context.root / "AGENTS.md").is_file() else ""
    if context.is_template is False and "template-only:start" in text:
        return [("AGENTS.md", "Delete the template-only section of AGENTS.md, including its "
                              "markers, in a generated repository.")]
    return []


def contact_links(context: Context) -> list[Finding]:
    """The issue forms' contact links point to this repository."""
    path = ".github/ISSUE_TEMPLATE/config.yml"
    if not (context.root / path).is_file():
        return []
    expected = f"https://github.com/{str(context.repository).lower()}/"
    return [(path, f"Contact link {url} does not point to {context.repository}; update it "
                   "after creating a repository from the template.")
            for url in CONTACT_URL.findall((context.root / path).read_text(encoding="utf-8"))
            if not url.lower().startswith(expected)]


def traffic_badges(context: Context) -> list[Finding]:
    """The README's traffic badges read this repository's traffic."""
    path = context.root / "README.md"
    if not path.is_file():
        return []
    read = sorted(set(BADGE.findall(path.read_text(encoding="utf-8"))))
    return [("README.md", f"A traffic badge reads {name}; point it at {context.repository} or "
                          "remove it.")
            for name in read if name.lower() != str(context.repository).lower()]


def secret_pattern(path: str) -> str | None:
    """The SECRET_FILES pattern a tracked path matches, unless an exception covers it."""
    name = path.rsplit("/", 1)[-1].lower()
    if any(fnmatch.fnmatchcase(name, pattern) for pattern in SECRET_FILE_EXCEPTIONS):
        return None
    return next((pattern for pattern in SECRET_FILES if fnmatch.fnmatchcase(name, pattern)), None)


def secret_files(context: Context) -> list[Finding]:
    """No tracked file is named as a secret or private key file."""
    findings = []
    for path in git(context.root, "ls-files", "-z").split("\0"):
        pattern = secret_pattern(path) if path else None
        if pattern:
            findings.append((path, f"is tracked and named as a secret or key file ({pattern}); "
                                   "remove it from the repository, rotate the secret it holds, "
                                   "and follow Handling a finding in "
                                   "docs/repo_governance/DISTRIBUTION.md."))
    return findings


def dependabot_branches(context: Context) -> list[Finding]:
    """Every Dependabot target-branch exists."""
    path = ".github/dependabot.yml"
    if not (context.root / path).is_file():
        return []
    findings = []
    for branch in sorted(set(TARGET_BRANCH.findall((context.root / path).read_text("utf-8")))):
        if not origin_has_branch(context.root, branch):
            findings.append((path, f"target-branch {branch} does not exist; set it to the active "
                                   "release branch or remove it."))
    return findings


def origin_has_branch(root: Path, branch: str) -> bool:
    """Whether origin has the branch, by its remote-tracking branch or else by asking it.

    A single-branch or shallow clone lacks the remote-tracking branch of a
    branch that exists. Raises MissingContext when origin cannot answer: no
    such remote, no network or no access.
    """
    tracked = subprocess.run(["git", "rev-parse", "--verify", "--quiet",
                              f"refs/remotes/origin/{branch}"], cwd=root, capture_output=True,
                             check=False)
    if not tracked.returncode:
        return True
    ref = f"refs/heads/{branch}"
    try:
        # --exit-code exits with 2 when origin answers that it has no such branch.
        asked = subprocess.run(["git", "ls-remote", "--exit-code", "--heads", "origin", ref],
                               cwd=root, stdin=subprocess.DEVNULL, capture_output=True,
                               text=True, timeout=REMOTE_TIMEOUT_SECONDS, check=False,
                               env={**os.environ, "GIT_TERMINAL_PROMPT": "0"})
    except subprocess.TimeoutExpired:
        raise MissingContext(f"origin did not answer within {REMOTE_TIMEOUT_SECONDS} s whether "
                             f"branch {branch} exists") from None
    if asked.returncode == 2:
        return False
    if asked.returncode:
        error = next((line for line in asked.stderr.splitlines() if line.strip()),
                     f"exit code {asked.returncode}")
        raise MissingContext(f"origin cannot be asked whether branch {branch} exists "
                             f"(git ls-remote: {error.strip()})")
    # A pattern also matches the end of a longer name: keep only the exact branch.
    return any(line.split("\t")[-1] == ref for line in asked.stdout.splitlines())


def codeql_default_branch(context: Context) -> list[Finding]:
    """codeql.yml lists the default branch, which branch filters cannot name."""
    path = ".github/workflows/codeql.yml"
    if not (context.root / path).is_file():
        return []
    sys.path.insert(0, str(context.root / "tests"))
    try:
        strict_yaml: Any = importlib.import_module("strict_yaml")
    finally:
        sys.path.pop(0)
    document = strict_yaml.load((context.root / path).read_text(encoding="utf-8")) or {}
    triggers = document.get("on", document.get(True)) or {}
    findings = []
    for event in ("push", "pull_request"):
        branches = (triggers.get(event) or {}).get("branches")
        if branches is not None and context.default_branch not in branches:
            findings.append((path, f"{event} branches {branches} omit the default branch "
                                   f"{context.default_branch}; add it."))
    return findings


def required_documents(context: Context) -> list[Finding]:
    """The documents the repository requires exist and are not empty."""
    return [(name, "is missing or empty") for name in REQUIRED_DOCUMENTS
            if not (context.root / name).is_file() or not (context.root / name).stat().st_size]


def clean_working_tree(context: Context) -> list[Finding]:
    """No file differs from the commit, after every other check has run."""
    changes = git(context.root, "status", "--porcelain", "--untracked-files=all")
    return [(line[3:], f"differs from the commit (git status {line[:2].strip()}); the checks "
                       "must describe the commit as it is")
            for line in changes.splitlines() if line]


def whitespace_commands(root: Path) -> list[tuple[str, list[str]]]:
    """Check every committed file, then local changes, without needing history."""
    empty_tree = subprocess.run(["git", "hash-object", "-t", "tree", "--stdin"], cwd=root,
                                input="", capture_output=True, text=True,
                                check=True).stdout.strip()
    return [("whitespace-committed", ["git", "diff", "--check", empty_tree, "HEAD"]),
            ("whitespace-local", ["git", "diff", "--check", "HEAD"])]


# ---------------------------------------------------------------------------
# The list of checks.

def tool_command(name: str, *args: str, module: bool = False) -> tuple[list[str], str | None]:
    """A command for an installed tool, and why it is missing if it is."""
    if module:
        found = importlib.util.find_spec(name) is not None
        return [sys.executable, "-m", name, *args], None if found else f"{name} is not installed"
    executable = shutil.which(name)
    return [executable or name, *args], None if executable else f"{name} is not on PATH"


def checks(context: Context) -> list[Check]:
    python = sys.executable
    changelog = [python, "tools/check_changelog.py"]
    if context.repository:
        changelog += ["--repository", context.repository]
    ruff, ruff_missing = tool_command("ruff", "check", "--no-cache", "tools", "tests",
                                      module=True)
    mypy, mypy_missing = tool_command("mypy", "--no-incremental", module=True)
    actionlint, actionlint_missing = tool_command("actionlint", "-no-color")
    node = shutil.which("node")
    node_missing = None if node else "node is not on PATH"
    no_repository = None if context.repository else "the repository is unknown (--repository)"
    # The CodeQL check reads its workflow through tests/strict_yaml.py, which needs PyYAML.
    codeql_missing = "; ".join(reason for found, reason in (
        (context.default_branch, "the default branch is unknown (--default-branch)"),
        (importlib.util.find_spec("yaml"), "PyYAML is not installed (.github/requirements.txt)"),
    ) if not found) or None
    listed: list[Check] = [
        Check("tests", [python, "-m", "unittest", "discover", "-s", "tests", "-v"]),
        Check("ruff", ruff, missing=ruff_missing),
        Check("mypy", mypy, missing=mypy_missing),
        Check("actionlint", actionlint, missing=actionlint_missing),
        Check("labels-sync-self-test", [node or "node", ".github/scripts/labels-sync.mjs",
                                        "--self-test"], missing=node_missing),
        Check("labels-catalogue", [node or "node", ".github/scripts/labels-sync.mjs", "--mode",
                                   "validate"], missing=node_missing),
        Check("labels-drift-self-test", [node or "node", ".github/scripts/labels-drift.mjs",
                                         "--self-test"], missing=node_missing),
        # check_issues.py validates the label catalogue with the label script, under Node.js.
        Check("issue-configuration", [python, "tools/check_issues.py", "--check-config"],
              missing=node_missing),
        Check("template-only-content", function=template_only,
              missing=None if context.is_template is not None
              else "whether this is the template is unknown (--template)"),
        Check("issue-contact-links", function=contact_links, missing=no_repository),
        Check("readme-traffic-badges", function=traffic_badges, missing=no_repository),
        Check("binaries", [python, "tools/check_binaries.py"]),
        Check("secret-files", function=secret_files),
        Check("vba-format", [python, "tools/format_vba.py", "--check"]),
        *[Check(f"vba-{name.replace('_', '-')}", [python, f"tools/check_vba_{name}.py"])
          for name in ("source", "conditionals", "jumps", "public_api", "dependencies",
                       "identifiers", "errors")],
        Check("vba-calls", [python, "tools/check_vba_calls.py"], advisory=True),
        Check("data-contract", [python, "tools/check_data_contract.py"]),
        Check("test-cases", [python, "tools/check_test_cases.py"]),
        Check("derivations", [python, "tools/check_derivations.py"]),
        Check("independent-implementations", [python, "tools/check_independent.py"]),
        Check("accuracy", [python, "tools/check_accuracy.py"]),
        Check("workbook-protection", [python, "tools/check_workbook_protection.py"]),
        Check("changelog", changelog),
        Check("generated-docs", [python, "tools/render_docs.py", "--check"]),
        Check("markdown-links", [python, "tools/check_links.py"]),
        Check("documented-commands", [python, "tools/check_doc_commands.py"]),
        Check("dependabot-target-branches", function=dependabot_branches),
        Check("codeql-default-branch", function=codeql_default_branch, missing=codeql_missing),
        Check("required-documents", function=required_documents),
    ]
    listed += [Check(name, command) for name, command in whitespace_commands(context.root)]
    if context.ci:
        # Last, so that it also sees what the other checks changed.
        listed.append(Check("clean-working-tree", function=clean_working_tree))
    return listed


# ---------------------------------------------------------------------------
# Running, reporting.

def annotate(finding: Finding) -> str:
    path, message = finding
    if os.environ.get("GITHUB_ACTIONS") == "true":
        return f"::error file={path}::{message}"
    return f"FAIL: {path}: {message}"


def run_check(check: Check, context: Context, strict: bool) -> Result:
    result = run_one(check, context, strict)
    if check.advisory and result.status == "fail":
        result.status = "warn"
        result.detail = f"advisory, not a failure: {result.detail}"
    return result


def run_one(check: Check, context: Context, strict: bool) -> Result:
    if check.missing:
        return missing_result(check, check.missing, strict)
    if check.function is not None:
        try:
            findings = check.function(context)
        except MissingContext as missing:
            return missing_result(check, str(missing), strict)
        except Exception as error:  # noqa: BLE001 - reported as a failed check, not raised
            # A malformed file must fail its check, not stop the run before the
            # other checks, the report and the summary are written.
            detail = f"{type(error).__name__}: {error}"
            return Result(check.name, "fail", None, 1, "raised an error", detail)
        log = "\n".join(annotate(finding) for finding in findings)
        return Result(check.name, "fail" if findings else "pass", None, 1 if findings else 0,
                      f"{len(findings)} finding(s)" if findings else "", log)
    assert check.command is not None
    try:
        completed = subprocess.run(check.command, cwd=context.root, capture_output=True,
                                   text=True, timeout=TIMEOUT_SECONDS, check=False)
        output, code = completed.stdout + completed.stderr, completed.returncode
    except (OSError, subprocess.TimeoutExpired) as error:
        output, code = str(error), 1
    return Result(check.name, "pass" if code == 0 else "fail", check.command, code,
                  "" if code == 0 else f"exit code {code}", output)


def missing_result(check: Check, reason: str, strict: bool) -> Result:
    """A check that cannot run here: skipped, or failed under --strict."""
    return Result(check.name, "fail" if strict else "skip", check.command, None, reason, reason)


def tracked_hashes(root: Path) -> dict[str, str]:
    names = [name for name in git(root, "ls-files", "-z").split("\0") if name]
    return {name: hashlib.sha256((root / name).read_bytes()).hexdigest()
            for name in names if (root / name).is_file()}


def build_report(context: Context, results: Sequence[Result]) -> dict[str, Any]:
    root = context.root
    return {
        "schema_version": 1,
        "commit": git(root, "rev-parse", "HEAD").strip(),
        "dirty": bool(git(root, "status", "--porcelain")),
        "pull_request_head": os.environ.get("PR_HEAD_SHA") or None,
        "repository": context.repository,
        "generated_at": datetime.datetime.now(datetime.UTC).isoformat(timespec="seconds"),
        "python": sys.version.split()[0],
        "platform": platform.platform(),
        "scope": SCOPE,
        "status": "fail" if any(r.status == "fail" for r in results) else "pass",
        "checks": [{"name": r.name, "status": r.status, "command": r.command,
                    "exit_code": r.exit_code, "detail": r.detail, "log": f"{r.name}.log"}
                   for r in results],
        "files_sha256": tracked_hashes(root),
    }


def summary(report: dict[str, Any]) -> str:
    marks = {"pass": "✅ pass", "fail": "❌ fail", "warn": "⚠️ warn", "skip": "⏭️ skip"}
    lines = [f"## Repository integrity: {report['status'].upper()}", "",
             f"- Commit: `{report['commit']}`" + (" (dirty working tree)" if report["dirty"]
                                                   else ""),
             f"- Pull request head: `{report['pull_request_head']}`"
             if report["pull_request_head"] else "- Pull request head: none",
             f"- Scope: {report['scope']}.", "",
             "| Check | Result | Detail |", "| --- | --- | --- |"]
    lines += [f"| `{check['name']}` | {marks[check['status']]} | "
              f"{check['detail'].replace('|', '/')} |" for check in report["checks"]]
    lines += ["", "Logs and `report.json`, with the SHA-256 of every tracked file, are in "
              f"`{REPORT_DIR.as_posix()}/`, kept as the run's artifact in CI."]
    return "\n".join(lines) + "\n"


def run(context: Context, listed: Sequence[Check], out: Path, strict: bool,
        summary_path: Path | None = None) -> int:
    out.mkdir(parents=True, exist_ok=True)
    results = []
    for check in listed:
        result = run_check(check, context, strict)
        command = " ".join(result.command) if result.command else check.name
        (out / f"{check.name}.log").write_text(f"$ {command}\n{result.log}", encoding="utf-8")
        results.append(result)
        print(f"{check.name}: {result.status.upper()}"
              + (f" ({result.detail})" if result.detail else ""), flush=True)
        if result.status in ("fail", "warn") and result.log:
            print(result.log.rstrip(), flush=True)
    report = build_report(context, results)
    (out / "report.json").write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    text = summary(report)
    (out / "summary.md").write_text(text, encoding="utf-8")
    if summary_path is not None:
        with summary_path.open("a", encoding="utf-8") as handle:
            handle.write(text)
    skipped = sum(result.status == "skip" for result in results)
    print(f"repository: {report['status'].upper()} ({len(results)} check(s), {skipped} "
          f"skipped; {SCOPE})")
    return 1 if report["status"] == "fail" else 0


def detect_repository(root: Path) -> str | None:
    try:
        url = git(root, "remote", "get-url", "origin").strip()
    except subprocess.CalledProcessError:
        return None
    match = GITHUB_REMOTE.search(url)
    return match.group(1) if match else None


def detect_default_branch(root: Path) -> str | None:
    try:
        ref = git(root, "symbolic-ref", "--quiet", "refs/remotes/origin/HEAD").strip()
    except subprocess.CalledProcessError:
        return None
    return ref.removeprefix("refs/remotes/origin/") or None


def parse_flag(text: str | None) -> bool | None:
    if text is None or text == "":
        return None
    if text.lower() in ("true", "false"):
        return text.lower() == "true"
    raise argparse.ArgumentTypeError(f"expected true or false, not {text!r}")


def main(argv: Sequence[str] | None = None,
         listed: Callable[[Context], list[Check]] = checks) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--repository", metavar="OWNER/REPOSITORY",
                        help="default: the GitHub repository of the origin remote")
    parser.add_argument("--default-branch", help="default: origin's HEAD branch")
    parser.add_argument("--template", type=parse_flag, metavar="true|false",
                        help="whether this repository is the template (default: unknown)")
    parser.add_argument("--strict", action="store_true",
                        help="fail a check whose tool or context is missing, as CI does")
    parser.add_argument("--ci", action="store_true",
                        help="fail when the working tree differs from the commit, as CI does")
    parser.add_argument("--summary", type=Path, help="also append the summary to this file")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    context = Context(root, args.repository or detect_repository(root),
                      args.default_branch or detect_default_branch(root), args.template,
                      args.ci)
    return run(context, listed(context), root / REPORT_DIR, args.strict, args.summary)


if __name__ == "__main__":
    sys.exit(main())
