#!/usr/bin/env python3
"""Bring exported VBA source into the layout of docs/VBA_HOUSE_STYLE.md.

The formatter applies only mechanical, behaviour-neutral rules that the house
style states (*Source representation*):

1. CRLF line endings, one final line ending, no blank lines at the end;
2. no trailing white space;
3. four spaces for each tab of a line's indentation;
4. procedure declarations, their End statements and labels at the outer
   level;
5. two blank lines after each procedure's End statement.

It never changes a string literal: every rule works on indentation, on the
end of a line, or on whole lines it recognises outside literals, and a
comment starts only at an apostrophe outside a literal. A statement keyword
followed by a colon, such as Else: or Loop:, is not a label. A line that ends
with a space or tab and an underscore continues on the next one, a comment
included. No blank line is added to or removed from one logical line: a
physical line that a continuation carries over is kept as it stands, blank or
not, the two blank lines after an End statement wait until its whole logical
line has been emitted, and a blank line that ends a continuation at the end
of the file stays. The header that the VBE writes (VERSION, BEGIN ... END,
Object and Attribute lines) is kept byte for byte.

--check reports each module that differs from its formatted form; --write
rewrites it. Without paths it reads the .bas, .cls and .frm files that Git
tracks under the source directories of docs/REPOSITORY_STRUCTURE.md, and
passes when there is none; a path may name a module or a folder, whose
modules are formatted.

An optional baseline, tools/vba_layout_baseline.json by default, lets a
project adopt the check one module at a time. It holds schema_version 1 and
a sorted, unique "files" list of root-relative POSIX paths, and --check
ratchets it:

- a module out of layout that the baseline does not list fails;
- a listed module that is in layout fails until its entry is removed;
- a listed path that is not a tracked module under the source directories
  fails (checked when no path is named).

Without the file every module must be in layout; --no-baseline forces that.
Exit status: 0 when everything passes, 1 for a layout finding, 2 for an
error: a missing path, an unreadable Git index, or a missing or malformed
baseline (a missing default baseline is not an error). It neither compiles
nor runs VBA.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from collections.abc import Sequence
from pathlib import Path, PurePosixPath

ROOT = Path(__file__).resolve().parents[1]
SOURCE_DIRECTORIES = ("src", "tests/modules", "examples/modules")
EXTENSIONS = frozenset({".bas", ".cls", ".frm"})
BASELINE = "tools/vba_layout_baseline.json"
BASELINE_SCHEMA = 1
INDENT = "    "

PROCEDURE = re.compile(r"^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
                       r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+[A-Za-z_]\w*",
                       re.I)
PROCEDURE_END = re.compile(r"^End\s+(?:Sub|Function|Property)\b", re.I)
LABEL = re.compile(r"^([A-Za-z_]\w*):$")
# Statement keywords that a colon separator can follow; none of them is a label.
NOT_LABELS = frozenset({"do", "else", "end", "loop", "next", "rem", "resume", "return",
                        "stop", "wend"})
HEADER_LINE = re.compile(r"^(?:VERSION\s|Object\s*=|Attribute\s)", re.I)
# A space or tab and an underscore at the end of a line continue it, in code
# and in a comment alike.
CONTINUED = re.compile(r"[ \t]_$")


class BaselineError(ValueError):
    """The baseline file is missing, malformed or lists an invalid path."""


def split_code_comment(line: str) -> tuple[str, str]:
    """Split a line at the apostrophe that starts its comment, if any.

    An apostrophe inside a string literal is data. A doubled quote inside a
    literal toggles the scan twice, which leaves it inside the literal.
    """
    in_literal = False
    for index, char in enumerate(line):
        if char == '"':
            in_literal = not in_literal
        elif char == "'" and not in_literal:
            return line[:index], line[index:]
    return line, ""


def header_length(lines: Sequence[str]) -> int:
    """The number of leading lines the VBE writes and the formatter keeps."""
    depth = 0
    for index, line in enumerate(lines):
        word = line.strip().split(" ", 1)[0].upper()
        if word == "BEGIN":
            depth += 1
        elif word == "END" and depth:
            depth -= 1
        elif not depth and not HEADER_LINE.match(line):
            return index
    return len(lines)


def expand_indentation(line: str) -> str:
    """Replace each tab of the leading white space with four spaces."""
    body = line.lstrip(" \t")
    return line[:len(line) - len(body)].replace("\t", INDENT) + body


def is_label(code: str) -> bool:
    """Whether a line's code is a label: a word and a colon, not a keyword."""
    match = LABEL.match(code)
    return match is not None and match.group(1).casefold() not in NOT_LABELS


def outer_level(line: str) -> bool:
    """Whether a line is a procedure declaration, its End statement or a label."""
    code, _ = split_code_comment(line.strip())
    code = code.strip()
    return bool(PROCEDURE.match(code) or PROCEDURE_END.match(code) or is_label(code))


def continues(line: str) -> bool:
    """Whether the next physical line belongs to this line's logical line.

    A continuation also continues a comment, so the whole line is tested.
    """
    return bool(CONTINUED.search(line))


def format_body(lines: Sequence[str]) -> list[str]:
    """The formatted body lines.

    No blank line is inserted into, or removed from, one logical line: a
    physical line that a continuation carries over is emitted as it stands,
    blank or not, and the blank lines after a procedure's End wait until its
    whole logical line has been emitted.
    """
    out: list[str] = []
    continued = False
    pending_end = False
    for raw in lines:
        line = expand_indentation(raw.rstrip())
        if continued:
            out.append(line)
            continued = continues(line)
            continue
        if not line:
            if not pending_end:
                out.append(line)
            continue
        if pending_end:
            # Two blank lines between a procedure's End and what follows it.
            out.extend(["", ""])
            pending_end = False
        if outer_level(line):
            line = line.lstrip()
            pending_end = bool(PROCEDURE_END.match(split_code_comment(line)[0].strip()))
        out.append(line)
        continued = continues(line)
    # A blank line that ends a continuation belongs to its logical line and stays.
    while out and not out[-1] and not (len(out) > 1 and continues(out[-2])):
        out.pop()
    return out


def format_text(text: str) -> str:
    """The formatted form of one module's text; idempotent."""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if lines and lines[-1] == "":
        lines.pop()
    split = header_length(lines)
    body = format_body(lines[split:])
    return "".join(f"{line}\r\n" for line in [*lines[:split], *body])


def is_module(path: Path) -> bool:
    return path.suffix.lower() in EXTENSIONS


def modules_under(folder: Path) -> list[Path]:
    return [path for path in sorted(folder.rglob("*")) if path.is_file() and is_module(path)]


def tracked_modules(root: Path) -> list[str]:
    """The tracked modules under the source directories, as sorted root-relative paths.

    Git's index decides, staged files included; an untracked file is not read.
    """
    completed = subprocess.run(["git", "-C", str(root), "ls-files", "-z", "--",
                                *SOURCE_DIRECTORIES],
                               capture_output=True, check=False)
    if completed.returncode:
        message = completed.stderr.decode("utf-8", errors="replace").strip()
        raise RuntimeError(f"cannot list the tracked files: {message or 'git ls-files failed'}")
    names = (item.decode("utf-8", errors="surrogateescape")
             for item in completed.stdout.split(b"\0") if item)
    return sorted({name for name in names if is_module(Path(name))})


def named_files(paths: Sequence[Path]) -> list[Path]:
    """The modules that the named paths stand for: a file as itself, a folder's modules."""
    files: list[Path] = []
    for path in paths:
        if path.is_dir():
            files.extend(modules_under(path))
        elif path.is_file():
            files.append(path)
        else:
            raise FileNotFoundError(path)
    return files


def load_baseline(path: Path, shown_as: str) -> list[str]:
    """The sorted, unique root-relative POSIX paths that a baseline file lists."""
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        reason = error.strerror if isinstance(error, OSError) and error.strerror else error
        raise BaselineError(f"cannot read baseline {shown_as}: {reason}") from error
    if not isinstance(data, dict) or data.get("schema_version") != BASELINE_SCHEMA:
        raise BaselineError(f"{shown_as}: expected an object with schema_version "
                            f"{BASELINE_SCHEMA}")
    files = data.get("files")
    if not isinstance(files, list) or not all(isinstance(name, str) for name in files):
        raise BaselineError(f"{shown_as}: 'files' must be a list of paths")
    for name in files:
        parts = PurePosixPath(name).parts
        if (not name or name.startswith("/") or "\\" in name or ":" in name
                or ".." in parts or "." in name.split("/") or name.endswith("/")):
            raise BaselineError(f"{shown_as}: {name!r} is not a root-relative POSIX path")
    if files != sorted(set(files)):
        raise BaselineError(f"{shown_as}: 'files' must be sorted and unique")
    return files


def shown(path: Path, root: Path) -> str:
    """A path relative to the root when it is inside it, else absolute."""
    resolved = path.resolve()
    if resolved.is_relative_to(root):
        return resolved.relative_to(root).as_posix()
    return resolved.as_posix()


def read(path: Path) -> str:
    # Latin-1 maps every byte to one character, so any byte round-trips.
    return path.read_bytes().decode("latin-1")


def check(root: Path, paths: Sequence[Path], baseline: Sequence[str], baseline_name: str,
          tracked: Sequence[str] | None) -> int:
    """Report each module out of layout and each stale baseline entry; 1 if any.

    ``tracked`` is the full list of tracked modules when no path was named:
    then a listed path that is not one of them is reported too.
    """
    listed = set(baseline)
    failures = 0
    for path in paths:
        name = shown(path, root)
        text = read(path)
        in_layout = format_text(text) == text
        if name in listed:
            if in_layout:
                failures += 1
                print(f"FAIL: {name}: in layout but still listed in {baseline_name}; "
                      "remove its entry")
            else:
                print(f"baselined: {name}: not in house-style layout, listed in "
                      f"{baseline_name}")
        elif not in_layout:
            failures += 1
            print(f"FAIL: {name}: not in house-style layout; run python tools/format_vba.py "
                  f"--write {name} and re-import it into the VBE")
    if tracked is not None:
        known = set(tracked)
        for name in baseline:
            if name not in known:
                failures += 1
                print(f"FAIL: {name}: listed in {baseline_name} but not a tracked VBA module "
                      "under the source directories")
    status = "FAIL" if failures else "PASS"
    print(f"vba-format: {status} ({len(paths)} file(s), {len(baseline)} baselined; "
          "layout only, nothing compiled)")
    return 1 if failures else 0


def write(root: Path, paths: Sequence[Path], baseline: Sequence[str],
          baseline_name: str) -> int:
    """Rewrite each module out of layout and say what each one now needs."""
    listed = set(baseline)
    written: list[str] = []
    for path in paths:
        text = read(path)
        formatted = format_text(text)
        if formatted != text:
            path.write_bytes(formatted.encode("latin-1"))
            written.append(shown(path, root))
            print(f"formatted: {written[-1]}")
    print(f"vba-format: {len(written)} of {len(paths)} file(s) rewritten")
    if written:
        print("Re-import each rewritten module into the VBE before committing.")
    for name in written:
        if name in listed:
            print(f"Remove {name} from {baseline_name}: it is now in layout.")
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--check", action="store_true",
                      help="report modules out of layout, ratcheting the baseline")
    mode.add_argument("--write", action="store_true", help="rewrite modules in place")
    parser.add_argument("--root", type=Path, default=ROOT)
    source = parser.add_mutually_exclusive_group()
    source.add_argument("--baseline", type=Path,
                        help=f"baseline file, relative to the root (default: {BASELINE}, "
                             "when it exists)")
    source.add_argument("--no-baseline", action="store_true",
                        help="ignore any baseline: every module must be in layout")
    parser.add_argument("paths", nargs="*", type=Path,
                        help="modules or folders to format (default: the tracked modules "
                             "under the VBA source directories)")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    # Relative paths are relative to the root, wherever the tool runs from.
    named = [path if path.is_absolute() else root / path for path in args.paths]
    baseline_path = args.baseline or Path(BASELINE)
    baseline_path = baseline_path if baseline_path.is_absolute() else root / baseline_path
    baseline_name = shown(baseline_path, root)
    baseline: list[str] = []
    tracked: list[str] | None = None
    try:
        if named:
            paths = named_files(named)
        else:
            tracked = tracked_modules(root)
            paths = [root / name for name in tracked if (root / name).is_file()]
        # The default baseline is optional; one named with --baseline must exist.
        if not args.no_baseline and (args.baseline or baseline_path.exists()):
            baseline = load_baseline(baseline_path, baseline_name)
    except FileNotFoundError as missing:
        print(f"vba-format: ERROR: {shown(Path(str(missing)), root)} is not a file or folder",
              file=sys.stderr)
        return 2
    except (BaselineError, RuntimeError) as error:
        print(f"vba-format: ERROR: {error}", file=sys.stderr)
        return 2
    if not paths and not baseline:
        print("vba-format: PASS (no VBA source found)")
        return 0
    if args.write:
        return write(root, paths, baseline, baseline_name)
    return check(root, paths, baseline, baseline_name, tracked)


if __name__ == "__main__":
    sys.exit(main())
