#!/usr/bin/env python3
"""Check the links in the repository's Markdown files.

Offline, every relative link must name an existing directory or non-empty file
inside the repository, and every fragment must name a heading or an explicit anchor of
the Markdown file it points to. An external link must use HTTPS: a plain http://
link is an error, whether it is an inline link, a reference definition, an
autolink or an HTML attribute; text in code spans and fenced blocks is not a link.

With --external, each https link is also requested once, over HTTPS only and
without following a redirect to another protocol, except absolute links into
this repository itself: a private repository answers 404 to anonymous requests,
so they are not checked at all. The repository is --repository, else
GITHUB_REPOSITORY, else the GitHub repository of the origin remote, so a local
run skips the same links as the workflow. Write links to this repository's files as
relative links, which the offline check covers. A 403, 429 or 5xx answer, or no
answer at all, is a warning, not an error, because the site may only be busy or
refuse automated requests. To record that such a link is known and still
wanted, list it in .github/link-exceptions.json with a reason and an expiry
date; after that date the warning becomes an error, so the exception is
reviewed again. The file is optional; when present, its form is

    {"schema_version": 1,
     "exceptions": [{"url": "https://example.org/page",
                     "reason": "Answers 403 to automated requests",
                     "expires": "2027-06-30"}]}

and the offline check fails while it is malformed or lists a URL that no
Markdown file links.
"""

from __future__ import annotations

import argparse
import html
import json
import os
import re
import subprocess
import sys
import unicodedata
from collections.abc import Iterable, Sequence
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from datetime import UTC, date, datetime
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlsplit

ROOT = Path(__file__).resolve().parents[1]
GITHUB_REMOTE = re.compile(r"github\.com[:/]([\w.-]+/[\w.-]+?)(?:\.git)?/?$")
FENCE = re.compile(r" {0,3}(`{3,}|~{3,})")
ATTRIBUTE = r"""\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+))"""
EXCEPTIONS = Path(".github/link-exceptions.json")
EXCEPTION_KEYS = {"url", "reason", "expires"}
# curl's exit status when a redirect leads to a protocol --proto-redir refuses.
CURL_UNSUPPORTED_PROTOCOL = 1


@dataclass(frozen=True)
class LinkException:
    """Why a URL's unverifiable answer is tolerated, and until when (inclusive)."""

    reason: str
    expires: date


def github_slug(heading: str) -> str:
    """Anchor GitHub generates for a Markdown heading, before duplicate suffixes."""
    text = re.sub(r"!?\[([^\]]*)\]\([^)]*\)", r"\1", heading)
    # Decode only complete character references (ending in ';'), as Markdown does.
    text = re.sub(r"&(?:#[0-9]{1,7}|#[xX][0-9a-fA-F]{1,6}|[A-Za-z][A-Za-z0-9]{1,31});",
                  lambda match: html.unescape(match.group(0)), re.sub(r"<[^>]+>", "", text))
    return "".join(char for char in text.lower() if slug_character(char)).replace(" ", "-")


def slug_character(char: str) -> bool:
    """GitHub keeps word characters (letters, marks, decimal digits and connector
    punctuation, such as the variation selector U+FE0F after many emoji), "-" and spaces."""
    return char in "- " or unicodedata.category(char)[:1] in ("L", "M") \
        or unicodedata.category(char) in ("Nd", "Pc")


def heading_anchor(slug: str, used: dict[str, int]) -> str:
    """Make a heading slug unique the way GitHub does: -1, -2... until unused."""
    anchor = slug
    while anchor in used:
        used[slug] += 1
        anchor = f"{slug}-{used[slug]}"
    used[anchor] = 0
    return anchor


def inline_destinations(line: str) -> list[str]:
    """Destinations of [text](destination "title"), allowing <...> and balanced parentheses."""
    found = []
    for match in re.finditer(r"\]\(", line):
        rest = line[match.end():].lstrip()
        if rest.startswith("<"):
            end = rest.find(">")
            if end > 1:
                found.append(rest[1:end])
            continue
        depth, destination = 0, ""
        for char in rest:
            if char.isspace() or (char == ")" and depth == 0):
                break
            depth += {"(": 1, ")": -1}.get(char, 0)
            destination += char
        if destination:
            found.append(destination)
    return found


def ends_paragraph(line: str, heading: bool) -> bool:
    return (heading or not line.strip()
            or bool(re.match(r" {0,3}([-*_])(\s*\1){2,}\s*$", line))
            or bool(re.match(r" {4}| {0,3}([-*+]\s|\d+[.)]\s|[>|<])", line)))


def line_links(line: str) -> tuple[list[str], set[str]]:
    """Links and explicit anchors on one line outside code."""
    line = re.sub(r"(`+).*?\1", "", line)
    anchors = {"".join(value) for value in
               re.findall(r"<a\s[^>]*?(?<![\w-])(?:id|name)" + ATTRIBUTE, line)}
    links = inline_destinations(line)
    links += ["".join(value) for value in
              re.findall(r"^ {0,3}\[(?!\^)[^\]]+\]:\s*(?:<([^>]*)>|(\S+))", line)]
    links += re.findall(r"<(https?://[^>\s]+)>", line)
    links += ["".join(value) for value in re.findall(r"(?<![\w-])(?:href|src)" + ATTRIBUTE, line)]
    return links, anchors


def markdown_links(text: str) -> tuple[list[str], set[str]]:
    """Return (links, anchors) of a Markdown document, ignoring fenced and inline code."""
    links: list[str] = []
    anchors: set[str] = set()
    used: dict[str, int] = {}
    fence: str | None = None
    paragraph: list[str] = []
    for line in text.splitlines():
        match = FENCE.match(line)
        if fence:
            if match and match.group(1)[0] == fence[0] and len(match.group(1)) >= len(fence):
                fence = None
            continue
        if match:
            fence, paragraph = match.group(1), []
            continue
        if paragraph and re.match(r" {0,3}(=+|-+)\s*$", line):
            # Setext heading: the paragraph above an === or --- underline.
            anchors.add(heading_anchor(github_slug(" ".join(paragraph)), used))
            paragraph = []
            continue
        heading = re.match(r" {0,3}#{1,6}\s+(.*?)(?:\s+#+)?\s*$", line)
        if heading:
            anchors.add(heading_anchor(github_slug(heading.group(1)), used))
        paragraph = [] if ends_paragraph(line, bool(heading)) else [*paragraph, line.strip()]
        found, explicit = line_links(line)
        links += found
        anchors |= explicit
    return links, anchors


def tracked_markdown(root: Path) -> list[Path]:
    try:
        names = subprocess.run(["git", "-C", str(root), "ls-files", "-z", "--", "*.md"],
                               capture_output=True, text=True, encoding="utf-8",
                               check=True).stdout.split("\0")
        return [root / name for name in names if name]
    except (OSError, subprocess.CalledProcessError):
        return [path for path in root.rglob("*.md") if ".git" not in path.relative_to(root).parts]


def local_error(root: Path, source: Path, link: str, anchors: dict[Path, set[str]]) -> str | None:
    url = urlsplit(link)
    if url.path.startswith("/"):
        return f"root-relative link: {link}"
    target = (source.parent / unquote(url.path)).resolve() if url.path else source.resolve()
    if not target.is_relative_to(root.resolve()):
        return f"path outside the repository: {link}"
    if not target.exists():
        return f"missing target: {link}"
    if target.is_file() and target.stat().st_size == 0:
        return f"empty target: {link}"
    if url.fragment and target in anchors and unquote(url.fragment) not in anchors[target]:
        return f"missing anchor: {link}"
    return None


def exception_problem(entry: Any, seen: dict[str, LinkException]) -> str | None:
    """Why an entry of the exceptions file is malformed, or None."""
    if not isinstance(entry, dict) or set(entry) != EXCEPTION_KEYS:
        return "needs exactly url, reason and expires"
    url, reason, expires = entry["url"], entry["reason"], entry["expires"]
    if not isinstance(url, str) or not url.startswith("https://") or not urlsplit(url).netloc \
            or "#" in url or url != url.strip():
        return f"url must be an https URL without a fragment: {url!r}"
    if url in seen:
        return f"{url} is listed twice"
    if not isinstance(reason, str) or not reason.strip():
        return f"{url} needs a reason"
    if not isinstance(expires, str) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", expires):
        return f"{url} needs expires as YYYY-MM-DD"
    try:
        date.fromisoformat(expires)
    except ValueError:
        return f"{url} has an expires date that does not exist: {expires}"
    return None


def load_exceptions(root: Path) -> tuple[dict[str, LinkException], list[str]]:
    """Return (exceptions by URL, problems) of the optional exceptions file."""
    path, name = root / EXCEPTIONS, EXCEPTIONS.as_posix()
    if not path.is_file():
        return {}, []
    try:
        data: Any = json.loads(path.read_text(encoding="utf-8"))
    except (UnicodeDecodeError, ValueError) as error:
        return {}, [f"{name}: not valid UTF-8 JSON ({error})"]
    if not isinstance(data, dict) or set(data) != {"schema_version", "exceptions"} \
            or type(data["schema_version"]) is not int or data["schema_version"] != 1 \
            or not isinstance(data["exceptions"], list):
        return {}, [f"{name}: needs exactly schema_version 1 and an exceptions list"]
    exceptions: dict[str, LinkException] = {}
    problems = []
    for index, entry in enumerate(data["exceptions"]):
        problem = exception_problem(entry, exceptions)
        if problem:
            problems.append(f"{name}: exception {index + 1}: {problem}")
        else:
            exceptions[entry["url"]] = LinkException(entry["reason"].strip(),
                                                     date.fromisoformat(entry["expires"]))
    return exceptions, problems


def check(root: Path = ROOT) -> tuple[list[str], set[str], int]:
    """Return (errors, external URLs, number of links checked) for the Markdown files."""
    errors: list[str] = []
    documents = {}
    for path in tracked_markdown(root):
        if not path.is_file():
            # Tracked by Git but deleted from the working tree without git rm.
            errors.append(f"{path.relative_to(root).as_posix()}: tracked but missing from the working tree")
            continue
        documents[path.resolve()] = markdown_links(path.read_text(encoding="utf-8"))
    anchors = {path: found for path, (_, found) in documents.items()}
    external: set[str] = set()
    count = 0
    for source, (links, _) in sorted(documents.items()):
        for link in links:
            count += 1
            url = urlsplit(link)
            if url.scheme == "http":
                errors.append(f"{source.relative_to(root.resolve()).as_posix()}: "
                              f"plain HTTP link, use https: {link}")
                continue
            if url.scheme == "https" or (not url.scheme and url.netloc):
                external.add(("https:" + link if link.startswith("//") else link).split("#")[0])
                continue
            if url.scheme in ("mailto", "tel", "data"):
                continue
            problem = (f"unsupported URL scheme: {link}" if url.scheme
                       else local_error(root, source, link, anchors))
            if problem:
                errors.append(f"{source.relative_to(root.resolve()).as_posix()}: {problem}")
    exceptions, problems = load_exceptions(root)
    errors += problems
    errors += [f"{EXCEPTIONS.as_posix()}: {url} is not linked from any Markdown file; "
               "remove its exception" for url in sorted(set(exceptions) - external)]
    return errors, external, count


def detect_repository(root: Path) -> str:
    """owner/name of the origin remote on GitHub, or "" when there is none."""
    result = subprocess.run(["git", "-C", str(root), "remote", "get-url", "origin"],
                            capture_output=True, text=True, check=False)
    match = GITHUB_REMOTE.search(result.stdout.strip()) if result.returncode == 0 else None
    return match.group(1) if match else ""


def own_repository(url: str, repository: str) -> bool:
    """Whether a URL points into this repository; such URLs are not requested."""
    if not repository:
        return False
    prefixes = (f"https://github.com/{repository}", f"https://raw.githubusercontent.com/{repository}/")
    lowered = url.lower()
    return any(lowered == prefix.lower() or lowered.startswith(prefix.lower().rstrip("/") + "/")
               for prefix in prefixes)


def check_external(url: str) -> tuple[str, str]:
    # GET follows redirects; HEAD is rejected by some otherwise working sites.
    result = subprocess.run([
        "curl", "--silent", "--show-error", "--location", "--output", os.devnull,
        "--write-out", "%{http_code}", "--connect-timeout", "10", "--max-time", "25",
        "--retry", "2", "--retry-delay", "2", "--proto", "=https",
        "--proto-redir", "=https", url,
    ], capture_output=True, text=True, timeout=95, check=False)
    code = result.stdout.strip()
    if result.returncode == CURL_UNSUPPORTED_PROTOCOL:
        return "ERROR", f"{url}: redirects to a URL that is not HTTPS"
    if result.returncode or code in ("403", "429") or code.startswith("5"):
        return "WARN", f"{url}: unavailable for verification (HTTP {code or 'unknown'})"
    if not code.startswith("2"):
        return "ERROR", f"{url}: HTTP {code}"
    return "PASS", url


def apply_exception(status: str, message: str, exception: LinkException | None,
                    today: date) -> tuple[str, str]:
    """Keep a tolerated answer a warning until its exception expires, then fail it."""
    if status != "WARN" or exception is None:
        return status, message
    if today > exception.expires:
        return "ERROR", (f"{message}; its exception in {EXCEPTIONS.as_posix()} expired on "
                         f"{exception.expires.isoformat()} ({exception.reason})")
    return status, (f"{message}; tolerated until {exception.expires.isoformat()} "
                    f"({exception.reason})")


def external_errors(urls: Iterable[str], repository: str,
                    exceptions: dict[str, LinkException] | None = None,
                    today: date | None = None) -> list[str]:
    checked = sorted(url for url in urls if not own_repository(url, repository))
    listed = exceptions or {}
    day = today or datetime.now(UTC).date()
    errors = []
    with ThreadPoolExecutor(max_workers=3) as pool:
        for url, result in zip(checked, pool.map(check_external, checked), strict=True):
            status, message = apply_exception(*result, listed.get(url), day)
            print(f"{status}: {message}")
            if status == "ERROR":
                errors.append(message)
    return errors


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--external", action="store_true",
                        help="also request every external https link")
    parser.add_argument("--repository", default=os.environ.get("GITHUB_REPOSITORY", ""),
                        help="owner/name whose own URLs are skipped (default: GITHUB_REPOSITORY, "
                             "else the origin remote)")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    errors, external, count = check(root)
    if args.external:
        repository = args.repository or detect_repository(root)
        errors += external_errors(external, repository, load_exceptions(root)[0])
    for error in errors:
        print(f"ERROR: {error}")
    print(f"Checked {count} link(s) in Markdown; {len(errors)} error(s).")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
