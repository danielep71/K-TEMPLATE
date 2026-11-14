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
comment starts only at an apostrophe outside a literal. A line that ends with
a space or tab and an underscore continues on the next one, a comment
included: the lines that continue it are never moved, and no blank line is
inserted before them. The header that the
VBE writes (VERSION, BEGIN ... END, Object and Attribute lines) is kept byte
for byte.

--check reports each file that differs from its formatted form; --write
rewrites it. A path may name a module or a folder, whose .bas, .cls and .frm
files are formatted. Without paths it formats those files under the source
directories of docs/REPOSITORY_STRUCTURE.md, and passes when there is none.
It neither compiles nor runs VBA.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Sequence
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SOURCE_DIRECTORIES = ("src", "tests/modules", "examples/modules")
EXTENSIONS = frozenset({".bas", ".cls", ".frm"})
INDENT = "    "

PROCEDURE = re.compile(r"^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
                       r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+[A-Za-z_]\w*",
                       re.I)
PROCEDURE_END = re.compile(r"^End\s+(?:Sub|Function|Property)\b", re.I)
LABEL = re.compile(r"^[A-Za-z_]\w*:$")
HEADER_LINE = re.compile(r"^(?:VERSION\s|Object\s*=|Attribute\s)", re.I)
# A space or tab and an underscore at the end of a line continue it, in code
# and in a comment alike.
CONTINUED = re.compile(r"[ \t]_$")


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


def outer_level(line: str) -> bool:
    """Whether a line is a procedure declaration, its End statement or a label."""
    code, _ = split_code_comment(line.strip())
    code = code.strip()
    return bool(PROCEDURE.match(code) or PROCEDURE_END.match(code) or LABEL.match(code))


def format_body(lines: Sequence[str]) -> list[str]:
    out: list[str] = []
    continued = False
    pending_end = False
    for raw in lines:
        line = expand_indentation(raw.rstrip())
        if not line:
            # A blank line ends a continued statement or comment.
            continued = False
            if not pending_end:
                out.append(line)
            continue
        if pending_end and not continued:
            # Two blank lines between a procedure's End and what follows it,
            # once the End statement and any comment it continues are over.
            out.extend(["", ""])
            pending_end = False
        if not continued and outer_level(line):
            line = line.lstrip()
            pending_end = bool(PROCEDURE_END.match(split_code_comment(line)[0].strip()))
        out.append(line)
        continued = bool(CONTINUED.search(line))
    while out and not out[-1]:
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


def modules_under(folder: Path) -> list[Path]:
    return [path for path in sorted(folder.rglob("*"))
            if path.is_file() and path.suffix.lower() in EXTENSIONS]


def source_files(root: Path) -> list[Path]:
    files = [path for directory in SOURCE_DIRECTORIES for path in modules_under(root / directory)]
    return sorted(set(files))


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


def shown(path: Path, root: Path) -> str:
    """A path relative to the root when it is inside it, else absolute."""
    resolved = path.resolve()
    if resolved.is_relative_to(root):
        return resolved.relative_to(root).as_posix()
    return resolved.as_posix()


def read(path: Path) -> str:
    # Latin-1 maps every byte to one character, so any byte round-trips.
    return path.read_bytes().decode("latin-1")


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--check", action="store_true", help="report files that differ")
    mode.add_argument("--write", action="store_true", help="rewrite files in place")
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("paths", nargs="*", type=Path,
                        help="files to format (default: the VBA source directories)")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    named = [path if path.is_absolute() else root / path for path in args.paths]
    try:
        paths = named_files(named) if named else source_files(root)
    except FileNotFoundError as missing:
        print(f"vba-format: ERROR: {shown(Path(str(missing)), root)} is not a file or folder",
              file=sys.stderr)
        return 2
    if not paths:
        print("vba-format: PASS (no VBA source found)")
        return 0
    drifted = []
    for path in paths:
        text = read(path)
        formatted = format_text(text)
        if formatted == text:
            continue
        if args.write:
            path.write_bytes(formatted.encode("latin-1"))
            print(f"formatted: {shown(path, root)}")
        else:
            drifted.append(path)
            print(f"FAIL: {shown(path, root)}: not in house-style layout; "
                  "run python tools/format_vba.py --write and re-import it into the VBE")
    status = "FAIL" if drifted else "PASS"
    print(f"vba-format: {status} ({len(paths)} file(s); layout only, nothing compiled)")
    return 1 if drifted else 0


if __name__ == "__main__":
    sys.exit(main())
