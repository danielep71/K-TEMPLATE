#!/usr/bin/env python3
"""Check that every VBA jump target is a label of the procedure that jumps.

``GoTo``, ``GoSub``, ``Resume <label>`` and ``On Error GoTo <label>`` must
name a label defined in the same procedure, labels must be unique within a
procedure, and every procedure must be closed. Each compilation environment
of check_vba_conditionals.py is checked separately, so that mutually
exclusive ``#If`` branches are never combined.

An ``Err.Raise`` reached after a ``Resume <label>`` while one of the
procedure's own ``On Error GoTo <label>`` handlers is enabled is reported:
that handler traps the raise again, so a handler that resumes to the same
code loops, and the error never reaches the caller. Raise a saved error
after ``On Error GoTo 0``. The model of the handler state:

- ``On Error GoTo <label>`` enables that local handler, ``On Error GoTo 0``
  disables it and ``On Error Resume Next`` replaces it. ``On Error GoTo -1``
  only ends the error-handling state, so the enabled handler stays enabled.
- A ``Resume <label>`` belongs to the handler whose label precedes it most
  closely, and re-enables that handler, which trapped the error. A
  ``Resume`` before any handler label is not followed.
- From the ``Resume`` target, the check follows the text of the procedure.
  ``If``, ``Select Case`` and loop blocks merge the states of their
  branches. A one-line ``If ... Then a: b Else c: d`` is read as a block
  ``If`` whose branches hold the statements in order, up to the end of the
  line. Where the branches disagree, the state is unknown and nothing is
  reported.
- The path ends at an ``Exit Sub``, ``Exit Function``, ``Exit Property``,
  ``End``, ``GoTo``, ``Resume`` or ``Return`` outside any block, and at the
  end of the procedure. Jumps are not followed, and a label on the path does
  not merge the states of other jumps to it.
- ``Resume`` and ``Resume Next`` continue at a statement that cannot be known
  statically, so no path starts from them. A raise inside a handler before
  its ``Resume`` is not trapped by that handler, so it is not reported.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import Path

from check_vba_conditionals import reachable_sources
from vba_lexer import (
    Finding,
    logical_units,
    read_component,
    report,
    split_statements,
    tracked_components,
)

ROOT = Path(__file__).resolve().parents[1]
PROC_OPEN = re.compile(
    r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\b", re.IGNORECASE)
PROC_CLOSE = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
NAMED_LABEL = re.compile(r"^\s*([A-Za-z_]\w*|\d+):")
# A line number, followed by a statement or by nothing once a comment such
# as "100 Rem target" or "100 ' target" has been stripped.
NUMBERED_LABEL = re.compile(r"^\s*(\d+)(?=\s|$)")
# A keyword after a dot is a member call, such as Application.Goto or .Goto
# in a With block, not a jump.
JUMP = re.compile(r"(?<![\w.])(GoTo|GoSub|Resume)\b(?:\s+([A-Za-z_]\w*|\d+|-1))?", re.IGNORECASE)
ON_ERROR = re.compile(r"\bon\s+error\s*$", re.IGNORECASE)
# On <expression> GoTo|GoSub <label>, <label>, ...: every listed target.
LABEL = r"(?:[A-Za-z_]\w*|\d+)"
COMPUTED_JUMP = re.compile(rf"\bOn\b(?!\s+Error\b).*?(?<![\w.])(GoTo|GoSub)\s+"
                           rf"({LABEL}(?:\s*,\s*{LABEL})*)", re.IGNORECASE)
# Statements that can stand alone before a colon, as in "Else: x = 1"; a
# keyword is never a label.
KEYWORDS = frozenset({"beep", "do", "else", "end", "loop", "next", "resume", "return", "stop",
                      "wend"})
# The handler-state model of the re-raise rule; see the module docstring.
ON_ERROR_SET = re.compile(r"(?<![\w.])On\s+Error\s+(?:GoTo\s+(-1|[A-Za-z_]\w*|\d+)|Resume\s+Next)\b",
                          re.IGNORECASE)
ERR_RAISE = re.compile(r"(?<![\w.])Err\s*\.\s*Raise\b", re.IGNORECASE)
RESUME_TO_LABEL = re.compile(r"^Resume\s+([A-Za-z_]\w*|\d+)$", re.IGNORECASE)
ONE_LINE_IF = re.compile(r"^(If\b.*?\bThen)\s+(\S.*)$", re.IGNORECASE)
ONE_LINE_ELSE = re.compile(r"(?<![\w.])Else\b", re.IGNORECASE)
CASE_ELSE = re.compile(r"^Case\s+Else$", re.IGNORECASE)
BLOCK_IF = re.compile(r"^If\b.*\bThen$", re.IGNORECASE)
SELECT = re.compile(r"^Select\s+Case\b", re.IGNORECASE)
BRANCH = re.compile(r"^(?:ElseIf\b|Else$|Case\b)", re.IGNORECASE)
BRANCH_ELSE = re.compile(r"^(?:Else|Case\s+Else)$", re.IGNORECASE)
BLOCK_END = re.compile(r"^End\s+(?:If|Select)$", re.IGNORECASE)
LOOP_OPEN = re.compile(r"^(?:Do|For|While)\b", re.IGNORECASE)
LOOP_CLOSE = re.compile(r"^(?:Loop|Next|Wend)\b", re.IGNORECASE)
LOOP_EXIT = re.compile(r"^Exit\s+(?:Do|For)$", re.IGNORECASE)
PATH_END = re.compile(r"^(?:Exit\s+(?:Sub|Function|Property)|End|GoTo\b.*|Resume\b.*|Return)$",
                      re.IGNORECASE)
# Handler states other than an enabled local handler, which is its label.
OFF, RESUME_NEXT, UNKNOWN, DEAD = "<off>", "<resume next>", "<unknown>", "<dead>"


@dataclass
class Step:
    """A label or a statement of a procedure, in source order."""
    line: int
    label: str | None = None
    code: str = ""


@dataclass
class Block:
    """An If, Select Case or loop block open on a re-raise path."""
    loop: bool
    entry: str
    ends: list[str] = field(default_factory=list)
    has_else: bool = False
    started: bool = False


@dataclass
class Procedure:
    name: str
    line: int
    labels: dict[str, int] = field(default_factory=dict)
    jumps: list[tuple[int, str, str]] = field(default_factory=list)
    steps: list[Step] = field(default_factory=list)


def is_procedure(code: str) -> re.Match[str] | None:
    """A procedure header; a Declare statement never matches, since its
    keyword stands where PROC_OPEN expects Sub, Function or Property."""
    return PROC_OPEN.match(code)


def label_at_start(code: str) -> str | None:
    match = NAMED_LABEL.match(code) or NUMBERED_LABEL.match(code)
    return match.group(1) if match and match.group(1).casefold() not in KEYWORDS else None


def statement_jumps(code: str) -> list[tuple[str, str]]:
    """``(operation, target)`` for each jump to a label in one statement.

    ``Resume``, ``Resume 0``, ``Resume Next`` and ``On Error GoTo 0`` or ``-1`` name no
    label;
    ``On <expression> GoTo`` and ``GoSub`` name every target of their list.
    """
    jumps: list[tuple[str, str]] = []
    computed = COMPUTED_JUMP.search(code)
    if computed:
        jumps.extend((computed.group(1), target.strip())
                     for target in computed.group(2).split(","))
        code = code[:computed.start()] + " " * (computed.end() - computed.start()) + \
            code[computed.end():]
    for match in JUMP.finditer(code):
        operation, target = match.group(1), match.group(2)
        if target is None or (operation.casefold() == "resume"
                              and target.casefold() in {"next", "0"}):
            continue
        if operation.casefold() == "goto" and target in {"0", "-1"} \
                and ON_ERROR.search(code[:match.start()]):
            continue
        jumps.append((operation, target))
    return jumps


def scan(path: str, text: str) -> tuple[list[Procedure], list[Finding]]:
    """Collect each procedure's labels and jumps; report structural defects."""
    procedures: list[Procedure] = []
    findings: list[Finding] = []
    current: Procedure | None = None
    for line, _, _, code in logical_units(text.split("\n"), blank_strings=True):
        # A label starts the logical line; the statements after it, split at
        # ":" separators, may open or close a procedure and may jump.
        label = label_at_start(code) if current is not None else None
        if current is not None and label is not None:
            if label.casefold() in current.labels:
                findings.append(Finding(path, line, f"duplicate label {label!r} in procedure "
                                                    f"{current.name}; first defined at line "
                                                    f"{current.labels[label.casefold()]}"))
            else:
                current.labels[label.casefold()] = line
            current.steps.append(Step(line, label=label))
        # The statements of each procedure on this line, recorded once the
        # line, or the procedure, ends, so that a one-line If spans them all.
        owner, texts = current, []
        for index, statement in enumerate(split_statements(code)):
            # The label opens the first statement, as in "Cleanup" or "100 Exit Sub".
            text = statement[len(label):].strip() if index == 0 and label else statement
            before = current
            current = scan_statement(path, line, statement, current, procedures, findings)
            if before is not None and current is before:
                texts.append(text)
            else:
                record_steps(owner, line, texts)
                owner, texts = current, []
        record_steps(owner, line, texts)
    if current is not None:
        findings.append(Finding(path, current.line, f"procedure {current.name} is not closed"))
    return procedures, findings


def scan_statement(path: str, line: int, statement: str, current: Procedure | None,
                   procedures: list[Procedure], findings: list[Finding]) -> Procedure | None:
    """Apply one statement; return the procedure open after it."""
    opener = is_procedure(statement)
    if opener:
        if current is not None:
            findings.append(Finding(path, line, f"procedure {current.name} opened at line "
                                                f"{current.line} is not closed before "
                                                f"{opener.group(2)}"))
        current = Procedure(opener.group(2), line)
        procedures.append(current)
    elif PROC_CLOSE.match(statement):
        if current is None:
            findings.append(Finding(path, line, "End statement outside any procedure"))
        current = None
    elif current is not None:
        current.jumps.extend((line, operation, target)
                             for operation, target in statement_jumps(statement))
    return current


def record_steps(procedure: Procedure | None, line: int, statements: list[str]) -> None:
    """Append the statements of one line to ``procedure``, in block form."""
    if procedure is not None:
        procedure.steps.extend(Step(line, code=code) for code in block_form(statements)
                               if code)


def one_line_parts(statement: str) -> list[tuple[str, str]]:
    """``(kind, code)`` parts of a statement: ``"if"`` for the head of a
    one-line If, ``"else"`` for its Else, ``"code"`` for anything else."""
    match = ONE_LINE_IF.match(statement)
    if match:
        return [("if", match.group(1)), *one_line_parts(match.group(2))]
    pieces = ONE_LINE_ELSE.split(statement)
    if len(pieces) == 1 or CASE_ELSE.match(statement):
        return [("code", statement)]
    parts: list[tuple[str, str]] = []
    for index, piece in enumerate(pieces):
        if index:
            parts.append(("else", "Else"))
        if piece.strip():
            parts.extend(one_line_parts(piece.strip()))
    return parts


def block_form(statements: list[str]) -> list[str]:
    """The statements of one line, with each one-line ``If ... Then a: b
    Else c`` written as a block ``If`` that ends with the line. An Else
    belongs to the nearest open one-line If, as in VBA."""
    parts = [part for statement in statements for part in one_line_parts(statement)]

    def branch(index: int, nested: bool) -> tuple[list[str], int]:
        codes: list[str] = []
        while index < len(parts):
            kind, code = parts[index]
            if kind == "else" and nested:
                return codes, index
            if kind == "if":
                then, index = branch(index + 1, True)
                codes += [code, *then]
                if index < len(parts):
                    other, index = branch(index + 1, True)
                    codes += ["Else", *other]
                codes.append("End If")
                continue
            # A block If's Else, as in "Else: x = 1", stays a statement.
            codes.append(code)
            index += 1
        return codes, index

    return branch(0, False)[0]


def enabled(state: str) -> bool:
    """Whether ``state`` is a local handler, named by its label."""
    return not state.startswith("<")


def handler_after(code: str, state: str) -> str:
    """The handler state after the ``On Error`` statements in ``code``."""
    for match in ON_ERROR_SET.finditer(code):
        target = match.group(1)
        if target is None:
            state = RESUME_NEXT
        elif target == "0":
            state = OFF
        elif target != "-1":
            state = target
    return state


def merge(states: list[str]) -> str:
    """The state where branches meet: DEAD if none reaches it, UNKNOWN if they disagree."""
    live = [state for state in states if state != DEAD]
    if not live:
        return DEAD
    return live[0] if len({state.casefold() for state in live}) == 1 else UNKNOWN


def block_statement(code: str, state: str, blocks: list[Block]) -> str | None:
    """The state after a statement that opens, divides, closes or leaves a
    block, or None for any other statement."""
    if BLOCK_IF.match(code) or SELECT.match(code):
        blocks.append(Block(loop=False, entry=state, started=bool(BLOCK_IF.match(code))))
        return state
    if LOOP_OPEN.match(code):
        blocks.append(Block(loop=True, entry=state))
        return state
    if LOOP_EXIT.match(code) and any(block.loop for block in blocks):
        next(block for block in reversed(blocks) if block.loop).ends.append(state)
        return DEAD
    if not blocks:
        return None
    block = blocks[-1]
    if BRANCH.match(code) and not block.loop:
        if block.started:
            block.ends.append(state)
        block.started = True
        block.has_else = block.has_else or bool(BRANCH_ELSE.match(code))
        return block.entry
    if BLOCK_END.match(code) and not block.loop:
        blocks.pop()
        return merge([*block.ends, state, *([] if block.has_else else [block.entry])])
    if LOOP_CLOSE.match(code) and block.loop:
        blocks.pop()
        return merge([*block.ends, state, block.entry])
    return None


def trapped_raises(steps: list[Step], handler: str) -> list[tuple[int, str]]:
    """``(line, handler)`` for each ``Err.Raise`` that a local handler traps
    on the path through ``steps``, entered with ``handler`` enabled."""
    found: list[tuple[int, str]] = []
    state = handler
    blocks: list[Block] = []
    for step in steps:
        if step.label is not None:
            # Another jump may reach a label in a branch this path left.
            state = UNKNOWN if state == DEAD else state
            continue
        after = block_statement(step.code, state, blocks)
        if after is not None:
            state = after
            continue
        if state == DEAD:
            continue
        if ERR_RAISE.search(step.code) and enabled(state):
            found.append((step.line, state))
        if PATH_END.match(step.code):
            if not blocks:
                break
            state = DEAD
        else:
            state = handler_after(step.code, state)
    return found


def reraise_findings(path: str, procedure: Procedure) -> list[Finding]:
    """Each ``Err.Raise`` reached after a ``Resume <label>`` while a local
    handler of ``procedure`` is enabled."""
    handlers = {match.group(1).casefold() for step in procedure.steps
                for match in ON_ERROR_SET.finditer(step.code)
                if match.group(1) not in {None, "0", "-1"}}
    targets = {step.label.casefold(): index for index, step in enumerate(procedure.steps)
               if step.label is not None}
    messages: dict[tuple[int, str], str] = {}
    trapping: str | None = None
    for step in procedure.steps:
        if step.label is not None and step.label.casefold() in handlers:
            trapping = step.label
        for operation, target in statement_jumps(step.code):
            if trapping is None or operation.casefold() != "resume" \
                    or target.casefold() not in targets:
                continue
            for line, handler in trapped_raises(procedure.steps[targets[target.casefold()]:],
                                                trapping):
                messages.setdefault((line, handler.casefold()), (
                    f"Err.Raise after Resume {target} is trapped again by handler "
                    f"{handler!r} of procedure {procedure.name}; put On Error GoTo 0 "
                    f"before the raise"))
    return [Finding(path, line, message) for (line, _), message in sorted(messages.items())]


def analyze_source(path: str, text: str) -> list[Finding]:
    procedures, findings = scan(path, text)
    for procedure in procedures:
        findings.extend(Finding(path, line, f"{operation} target {target!r} is not a label of "
                                            f"procedure {procedure.name}")
                        for line, operation, target in procedure.jumps
                        if target.casefold() not in procedure.labels)
        findings.extend(reraise_findings(path, procedure))
    return findings


def analyze(path: str, text: str) -> list[Finding]:
    """Findings of every environment, each listed once with the environments it occurs in."""
    sources, findings = reachable_sources(path, text)
    environments: dict[Finding, list[str]] = {}
    for environment, source in sources.items():
        for finding in analyze_source(path, source):
            environments.setdefault(finding, []).append(environment)
    findings.extend(Finding(item.path, item.line, f"{item.message} [{', '.join(names)}]")
                    for item, names in environments.items())
    return findings


def check(root: Path) -> tuple[list[Finding], int, int]:
    components = tracked_components(root)
    findings: list[Finding] = []
    procedures = 0
    for path in components:
        text = read_component(root, path)
        findings.extend(analyze(path, text))
        procedures += sum(bool(is_procedure(statement)) for _, _, _, code
                          in logical_units(text.split("\n"), blank_strings=True)
                          for statement in split_statements(code))
    return findings, len(components), procedures


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components, procedures = check(args.root.resolve())
    return report("vba-jumps", findings,
                  f"{components} VBA component(s), {procedures} procedure(s)")


if __name__ == "__main__":
    sys.exit(main())
