#!/usr/bin/env python3
"""Check VBA error handling: Err is copied before any other statement can change it.

A read of ``Err.Number``, ``Err.Source`` or ``Err.Description`` is reported
when, on the straight-line path that leads to it, a statement has already
cleared or may have replaced the error it means to read:

- an ``On Error`` statement, which clears ``Err``;
- in an error handler, a call or a member access (``Helper``, ``x = Helper(1)``,
  ``x = Helper`` without parentheses, ``obj.Method``, ``Debug.Print``,
  ``Application.StatusBar = ...``, ``Err.Clear``, ``New``), which can run code
  that clears or replaces it. A name without parentheses is a call unless it
  is a variable, a constant, an Enum member or a component: a function,
  property or Declare of the project, or a host global such as ActiveCell.
  Names that start with ``vb`` or ``xl`` are the constants of VBA and Excel.
  A name is resolved against the project first: a procedure, property or
  Declare of the module, or a public one of a standard module, shadows a VBA
  function, statement, statement target or constant of its name, as in VBA,
  while ``VBA.Format`` still names the VBA function. ``Err`` and ``Erl``
  always mean VBA's.
  In the same way, an assignment (``=``, ``Let``, ``Set``, ``LSet``,
  ``RSet``) to anything but a variable, an array element or a statement
  target (``Date``, ``Mid(Text, 1, 1)``) sets a property, a call made once
  the value is read: a Property Let or Set of the project, a member
  (``.Value = x`` in a ``With`` block) or a property of the host.

After ``On Error Resume Next``, and outside a handler, ``Err`` belongs to the
most recent statement that can fail, so a read after such a statement checks
that statement and is not reported. A plain copy into a variable
(``n = Err.Number``) and a literal change nothing. The VBA functions of
PURE_FUNCTIONS (``CLng``, ``CStr``, ``Len``, ..., also written ``VBA.CLng``),
arithmetic, concatenation and an array element can fail but run no code that
clears ``Err``: in a handler they change nothing, and after ``On Error Resume
Next`` they are a statement that can fail. So are the VBA statements of
BUILTIN_STATEMENT (``ReDim``, ``Erase``, ``Open``, ``Print #``, ...), whose
operands alone are read as expressions; a statement that starts with any
other name is a call.

Within a statement, the check follows VBA's evaluation order: a read inside
the arguments of a call, as in ``CLng(Err.Number)``, runs before that call.
Only what runs before a read counts for it, under the same rules as the
statements before: after ``On Error Resume Next`` it can refresh ``Err``, and
in a handler a call or member access makes the read stale, as in
``Report Helper(), Err.Number``. An operator written before a read may run
after it, as ``&`` in ``a & b / Err.Number``; it counts as running before.
A name qualified by the VBA library or one of its modules (``VBA.CLng``,
``VBA.Strings.Len``, ``VBA.vbCrLf``, ``VBA.Beep``, ``VBA.Mid``) is VBA's, as
unqualified, whatever the project declares.
Assigning a property of ``Err`` (``Err.Description = x``) writes it and
reads nothing, and ``.Number`` inside ``With Err`` reads ``Err.Number``.
``Err`` read through its default member (``n = Err``), ``Err`` passed to a
procedure and ``Erl`` are not followed.

A segment starts at each label and after each branch or loop statement:
``If``, ``Select Case``, ``Do``, ``For``, ``While``, ``Exit``, ``Resume``,
``GoTo``, ``Return`` and the statements that close them end the segment
being tracked, and a single-line ``If`` ends it with its whole line. What
follows a branch is never claimed, so a finding is a read that is stale on
every path to it. Each compilation environment of check_vba_conditionals.py
is analyzed separately, and a finding names the environments it occurs in.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import Path

from check_vba_calls import dates_as_values
from check_vba_conditionals import reachable_sources
from check_vba_identifiers import Module, closing_parenthesis, scan
from check_vba_jumps import PROC_CLOSE, is_procedure, label_at_start
from vba_lexer import (
    Finding,
    logical_units,
    read_component,
    report,
    split_statements,
    tracked_components,
)

ROOT = Path(__file__).resolve().parents[1]
ERR_READ = re.compile(r"(?<![\w.!])(?:VBA\s*\.\s*)?Err\s*\.\s*(Number|Source|Description)\b",
                      re.IGNORECASE)
ON_ERROR = re.compile(r"^On\s+Error\b", re.IGNORECASE)
RESUME_NEXT = re.compile(r"^On\s+Error\s+Resume\s+Next\b", re.IGNORECASE)
HANDLER = re.compile(r"\bOn\s+Error\s+GoTo\s+([A-Za-z_]\w*|\d+)", re.IGNORECASE)
IF_START = re.compile(r"^If\b", re.IGNORECASE)
# The branches of an If, as in If x Then a Else b: separators, not values.
IF_KEYWORDS = re.compile(r"^If\b|\b(?:Then|Else)\b", re.IGNORECASE)
# Any property of Err as an assignment target, as in Err.Description = x.
ERR_TARGET = re.compile(r"(?:VBA\s*\.\s*)?Err\s*\.\s*[A-Za-z_]\w*", re.IGNORECASE)
WITH_ERR = re.compile(r"^With\s+(?:VBA\s*\.\s*)?Err\s*$", re.IGNORECASE)
# A member of the object of a With block, as in .Number; not x.y, f(1).y or 1.5.
WITH_MEMBER = re.compile(r"(?<![\w)\].!])\.(?=\s*[A-Za-z_])")
# What an operand is an argument or operand of, before it: the name of an
# enclosing call (the last name only, as in .Method of obj.Method(), since
# obj is read first) and an operator, which both run after the operand.
# The VBA library and its modules, as in VBA.CLng, VBA.Strings.Len or
# Strings.Len: what they qualify is VBA's, whatever the project declares.
VBA_MODULE = (r"(?:Constants|Conversion|DateTime|FileSystem|Financial|Information|Interaction"
              r"|Math|Strings)\s*\.\s*")
VBA_QUALIFIER = rf"(?:VBA\s*\.\s*(?:{VBA_MODULE})?|{VBA_MODULE})"
VBA_HEAD = re.compile(rf"^{VBA_QUALIFIER}", re.IGNORECASE)
ENCLOSING_CALL = re.compile(rf"(?:(?<![\w.!]){VBA_QUALIFIER}|\.\s*)?[A-Za-z_]\w*[%&!#@$]?\s*$",
                            re.IGNORECASE)
TRAILING_OPERATOR = re.compile(r"(?:[-&+*/\\^<>=]|\b(?:Mod|Not|And|Or|Xor|Eqv|Imp|Like|Is)\b)\s*$",
                               re.IGNORECASE)
LABEL_PREFIX = re.compile(r"^\s*(?:[A-Za-z_]\w*:(?!=)|\d+:?)")
# Statements that end the tracked segment once their own expression is read:
# that expression runs straight after the statement before.
BRANCH = re.compile(r"^(?:If|Select|Do|For|While|Exit|Resume|GoTo|GoSub|Return|Stop|End"
                    r"(?!\s+With\b))\b|^On\s+(?!Error\b)", re.IGNORECASE)
# Statements reached from elsewhere, which end the segment before anything is read.
JOIN = re.compile(r"^(?:ElseIf|Else|Case|Loop|Next|Wend|End\s+If|End\s+Select)\b", re.IGNORECASE)
DECLARATION = re.compile(r"^(?:Dim|Static|Const)\b", re.IGNORECASE)
WITH = re.compile(r"^With\b", re.IGNORECASE)
END_WITH = re.compile(r"^End\s+With\b", re.IGNORECASE)
# Statements of the VBA language that can fail but run no project code: only
# their operands are expressions, and the words of their syntax are not names.
BUILTIN_STATEMENT = re.compile(
    r"^(?:ReDim(?:\s+Preserve)?|Erase|Open|Close|Print|Write|Input|Line\s+Input|Get|Put|Seek"
    r"|Lock|Unlock|Kill|MkDir|RmDir|ChDir|ChDrive|FileCopy|Name|SetAttr|Randomize|Beep|Reset"
    r"|Width|Error|AppActivate|SaveSetting|DeleteSetting|SendKeys)\b", re.IGNORECASE)
STATEMENT_SYNTAX = re.compile(r"\b(?:Access|Append|Binary|Input|Lock|Output|Random|Read|Shared"
                              r"|Write)\b", re.IGNORECASE)
# Assignment targets that are VBA statements: Date = x sets the system date,
# Mid(Text, 1, 1) = x replaces characters.
BUILTIN_TARGETS = frozenset({"date", "time", "mid", "midb"})
ASSIGNMENT_KEYWORD = re.compile(r"^(?:Let|Set|LSet|RSet)\s+", re.IGNORECASE)
TARGET = re.compile(r"\.?[A-Za-z_]\w*[%&!#@$]?(?:\(\))?(?:[.!][A-Za-z_]\w*[%&!#@$]?(?:\(\))?)*")
PLAIN_TARGET = re.compile(r"[A-Za-z_]\w*[%&!#@$]?")
HEAD = re.compile(r"^(?:Call\s+)?\.?[A-Za-z_]\w*[%&!#@$]?(?:\s*[.!]\s*[A-Za-z_]\w*[%&!#@$]?)*",
                  re.IGNORECASE)
CALL = re.compile(r"(?<![\w.!#])([A-Za-z_]\w*)[%&!#@$]?\s*\(")
# A name not followed by an argument list or by := (a named argument).
NAME = re.compile(r"(?<![\w.!#])([A-Za-z_]\w*)\b(?![%&!#@$]?\s*\()(?!\s*:=)")
# A ! before a name is a member access, as in Me!txtAmount; x! is a Single.
MEMBER = re.compile(r"\.\s*[A-Za-z_]|![A-Za-z_]")
NEW = re.compile(r"\bNew\s+[A-Za-z_]", re.IGNORECASE)
NUMBER_LITERAL = re.compile(r"&[HO][0-9A-F]+&?", re.IGNORECASE)
OPERATOR = re.compile(r"[*/\\^]|(?<![\w&])&|\bMod\b", re.IGNORECASE)
# VBA functions that run no project code, so they leave Err as it is unless
# they fail; extend the set when a copy of Err goes through another one.
PURE_FUNCTIONS = frozenset("""
    abs array asc ascw cbool cbyte ccur cdate cdbl cdec chr chrw cint clng clnglng clngptr csng
    cstr cvar cverr dateserial day fix format hex iif instr instrrev int isarray isdate isempty
    iserror ismissing isnull isnumeric isobject join lbound lcase left len lenb ltrim mid midb
    month now oct replace right round rtrim sgn space spc split str strcomp string tab time timer
    trim typename ubound ucase val vartype year
""".split())
# The same functions qualified by their library, as in VBA.CLng(x), and the
# constants of VBA so qualified, as in VBA.vbCrLf: neither a member access of
# an object nor a procedure of the project of that name.
VBA_PURE = re.compile(rf"(?<![\w.!]){VBA_QUALIFIER}(?:{'|'.join(sorted(PURE_FUNCTIONS))})\b"
                      r"[%&!#@$]?", re.IGNORECASE)
VBA_CONSTANT = re.compile(rf"(?<![\w.!]){VBA_QUALIFIER}vb\w*", re.IGNORECASE)
# Prefixes of the constants of VBA (vbCrLf) and Excel (xlUp), which are values.
CONSTANT_PREFIXES = ("vb", "xl")
# Words that may stand in an expression without being a value that can fail.
WORDS = frozenset("""
    and as byref byval call case do each else elseif empty end eqv erl err error exit false for
    function gosub goto if imp in is let like loop lset me mod new next not nothing null on or
    property resume return rset select set step stop sub then to true typeof until wend while
    with xor
""".split())


@dataclass
class Segment:
    """What is known of Err at one point of a procedure's straight-line code."""

    in_handler: bool = False
    resume_next: bool = False
    # The line and the reason that made Err stale, or None while it is current.
    stale: tuple[int, str] | None = None

    @property
    def follows_last_failure(self) -> bool:
        """Whether Err belongs to the most recent statement that can fail."""
        return self.resume_next or not self.in_handler

    def restart(self, in_handler: bool | None = None) -> None:
        if in_handler is not None:
            self.in_handler = in_handler
            self.resume_next = False
        self.stale = None


@dataclass
class Procedure:
    name: str
    line: int
    units: list[tuple[int, str | None, list[str]]] = field(default_factory=list)


@dataclass(frozen=True)
class Names:
    """The names a procedure resolves before VBA's own: ``values``, read
    without a call, and ``callables``, the procedures, properties and Declares
    of the project it can call unqualified. A project declaration shadows a
    VBA function, statement or constant of the same name, as in VBA."""

    values: frozenset[str] = frozenset()
    callables: frozenset[str] = frozenset()


@dataclass(frozen=True)
class Effect:
    """Whether code can fail, and whether it can run code that changes Err."""

    fails: bool = False
    changes: bool = False

    def __or__(self, other: Effect) -> Effect:
        return Effect(self.fails or other.fails, self.changes or other.changes)


def arithmetic(text: str) -> bool:
    """Whether ``text`` computes: an operator other than a sign before a number."""
    if OPERATOR.search(text):
        return True
    for match in re.finditer(r"[-+]", text):
        before = text[:match.start()].rstrip()
        after = text[match.end():].lstrip()
        if (before and (before[-1].isalnum() or before[-1] in "_)]")) or \
                not after[:1].isdigit():
            return True
    return False


def effect(text: str, names: Names) -> Effect:
    """What evaluating the expression ``text`` can do to Err; strings are blanked.

    A name followed by parentheses is an array element when it is a value,
    and a call otherwise; a bare name that is not a value is a call too: a
    function or property of the project, a Declare, or a member of the host's
    global object, such as ActiveSheet, called without parentheses. Only a
    name the project does not declare can be a VBA function or constant."""
    text = ERR_READ.sub(lambda match: "0" * len(match.group()), text)
    # A qualified VBA function can fail, as CLng("x") does, but runs no project code.
    qualified = VBA_PURE.search(text) is not None
    text = VBA_PURE.sub(lambda match: "0" * len(match.group()), text)
    text = VBA_CONSTANT.sub(lambda match: "0" * len(match.group()), text)
    text = NUMBER_LITERAL.sub(lambda match: "0" * len(match.group()), dates_as_values(text))
    result = Effect(fails=arithmetic(text) or qualified)
    if MEMBER.search(text) or NEW.search(text):
        result = Effect(True, True)
    for match in CALL.finditer(text):
        name = match.group(1).casefold()
        if name in names.values:
            result = result | Effect(fails=True)
        elif name in names.callables:
            result = Effect(True, True)
        elif name not in WORDS:
            result = result | Effect(True, name not in PURE_FUNCTIONS)
    for match in NAME.finditer(text):
        name = match.group(1).casefold()
        if name in names.callables and name not in names.values:
            result = Effect(True, True)
        elif not (name in names.values or name in WORDS or name.startswith(CONSTANT_PREFIXES)):
            result = result | Effect(True, name not in PURE_FUNCTIONS)
    return result


def assignment(statement: str) -> tuple[str, str] | None:
    """``(target, value)`` of an assignment, or None for any other statement."""
    keyword = ASSIGNMENT_KEYWORD.match(statement)
    body = statement[keyword.end():] if keyword else statement
    depth = 0
    for index, char in enumerate(body):
        depth += {"(": 1, ")": -1}.get(char, 0)
        if char == "=" and not depth:
            target = body[:index].strip()
            collapsed, position = "", 0
            while position < len(target):
                end = closing_parenthesis(target, position) if target[position] == "(" else None
                collapsed += "()" if end is not None else target[position]
                position = end + 1 if end is not None else position + 1
            return (target, body[index + 1:]) if TARGET.fullmatch(collapsed) else None
    return None


def statement_effect(statement: str, names: Names) -> tuple[str, Effect]:
    """The expressions a statement evaluates before its own operation, in
    evaluation order, and what the whole statement can do to Err. A procedure
    runs after its arguments are read, and a property is set after its value
    is read. Assigning a property of Err, as in Err.Description = x, reads
    and clears nothing of it."""
    if DECLARATION.match(statement) or END_WITH.match(statement):
        return "", Effect()
    if WITH.match(statement):
        return statement[4:], effect(statement[4:], names)
    pair = assignment(statement)
    if pair is not None:
        target, value = pair
        if ERR_TARGET.fullmatch(target):
            return value, effect(value, names)
        index = target.find("(")
        reads = value + (" , " + target[index:] if index > 0 else "")
        return reads, effect(value, names) | target_effect(target, names)
    # A procedure of the project named like a VBA statement, such as Beep,
    # shadows it, unless the statement is qualified, as in VBA.Beep.
    qualifier = VBA_HEAD.match(statement)
    unqualified = statement[qualifier.end():] if qualifier else statement
    builtin = BUILTIN_STATEMENT.match(unqualified)
    if builtin and (qualifier or unqualified.split(None, 1)[0].casefold() not in names.callables):
        operands = STATEMENT_SYNTAX.sub(" ", unqualified[builtin.end():])
        return operands, effect(operands, names) | Effect(fails=True)
    head_match = HEAD.match(statement)
    arguments = statement[head_match.end():] if head_match else statement
    return arguments, effect(arguments, names) | Effect(True, True)


def evaluated_before(text: str, position: int) -> str:
    """The part of the expression ``text`` that VBA evaluates before the
    operand at ``position``: the text before it, without the calls that it is
    an argument of and the operators that it is an operand of, which run
    after it. An operator earlier in the text may also run after it, as & in
    a & b / Err.Number; it is kept, so that a read is never claimed stale on
    the strength of a guess."""
    before = list(text[:position])
    opened: list[int] = []
    for index, char in enumerate(text[:position]):
        if char == "(":
            opened.append(index)
        elif char == ")" and opened:
            opened.pop()
    for index in opened:
        call = ENCLOSING_CALL.search(text[:index])
        start = call.start() if call else index
        before[start:index + 1] = " " * (index + 1 - start)
    result = "".join(before).rstrip()
    while (operator := TRAILING_OPERATOR.search(result)) is not None:
        result = result[:operator.start()].rstrip()
    return result


def target_effect(target: str, names: Names) -> Effect:
    """What assigning to ``target`` can do to Err, once the value is read.

    A variable of the procedure or the module (and the procedure's own name)
    is a plain copy, an array element can fail, and a statement target such
    as Date or Mid(Text, 1, 1) can fail. Any other target is a property: a
    Property Let or Set of the project, which shadows a statement target of
    its name, a member (x.y = 1, .y = 1 in a With block), or a property of
    the host, such as Caption, which runs code. VBA.Mid(Text, 1, 1) is Mid."""
    qualifier = VBA_HEAD.match(target)
    if qualifier:
        rest = target[qualifier.end():]
        index = rest.find("(")
        if (rest[:index] if index > 0 else rest).rstrip("%&!#@$ ").casefold() in BUILTIN_TARGETS:
            return (effect(rest[index:], names) if index > 0 else Effect()) | Effect(fails=True)
    index = target.find("(")
    name = (target[:index] if index > 0 else target).rstrip("%&!#@$ ").casefold()
    indices = effect(target[index:], names) if index > 0 else Effect()
    if PLAIN_TARGET.fullmatch(target) or (index > 0 and PLAIN_TARGET.fullmatch(name)):
        if name in names.values:
            return indices | Effect(fails=index > 0)
        if name in BUILTIN_TARGETS and name not in names.callables:
            return indices | Effect(fails=True)
    return indices | Effect(True, True)


def handler_labels(procedure: Procedure) -> set[str]:
    """The labels that an On Error GoTo of the procedure names."""
    return {match.group(1).casefold() for _, _, statements in procedure.units
            for statement in statements for match in HANDLER.finditer(statement)
            if match.group(1) not in {"0"}}


def stale_message(procedure: str, reads: list[re.Match[str]], cause: str) -> str:
    fields = list(dict.fromkeys(f"Err.{match.group(1).capitalize()}" for match in reads))
    return (f"{', '.join(fields)} read in {procedure} after {cause}; "
            "save Err.Number, Err.Source and Err.Description into variables first")


def stale_reads(reads: str, segment: Segment, names: Names) -> list[re.Match[str]]:
    """The reads of Err in ``reads`` that are stale, each judged on what the
    statement evaluates before it as well as on ``segment``.

    After On Error Resume Next or outside a handler, code that runs before a
    read and can fail may be the error it checks: the read is stale only when
    Err already was and no such code runs. In a handler, a read is stale when
    Err already was, or when a call or member access runs before it in the
    same statement, as in Report Helper(), Err.Number. A call that a read is
    an argument of, as in CLng(Err.Number), runs after it."""
    stale = []
    for match in ERR_READ.finditer(reads):
        before = effect(evaluated_before(reads, match.start()), names)
        if segment.follows_last_failure:
            if segment.stale and not before.fails:
                stale.append(match)
        elif segment.stale or before.changes:
            stale.append(match)
    return stale


def check_statement(path: str, procedure: Procedure, line: int, statement: str,
                    segment: Segment, names: Names) -> list[Finding]:
    """Read one statement against ``segment``, then apply its effect on Err."""
    findings: list[Finding] = []
    if JOIN.match(statement):
        segment.restart()
        return findings
    if ON_ERROR.match(statement):
        segment.stale = (line, "which cleared Err with an On Error statement")
        segment.resume_next = bool(RESUME_NEXT.match(statement))
        return findings
    branch = BRANCH.match(statement)
    if branch:
        reads = IF_KEYWORDS.sub(" , ", statement) if IF_START.match(statement) else statement
        total = Effect()
    else:
        reads, total = statement_effect(statement, names)
    stale = stale_reads(reads, segment, names)
    if stale:
        cause = (f"line {segment.stale[0]}, {segment.stale[1]}" if segment.stale else
                 "a call or member access that the same statement evaluates first")
        findings.append(Finding(path, line, stale_message(procedure.name, stale, cause)))
    if branch:
        segment.restart()
    elif segment.follows_last_failure:
        if total.fails:
            segment.stale = None
    elif total.changes and segment.stale is None:
        segment.stale = (line, "which may have changed Err with a call or member access")
    return findings


def check_procedure(path: str, procedure: Procedure, names: Names) -> list[Finding]:
    handlers = handler_labels(procedure)
    segment = Segment()
    findings: list[Finding] = []
    # Whether each open With block is on Err, whose .Number reads Err.Number.
    with_err: list[bool] = []
    for line, label, statements in procedure.units:
        if label is not None:
            segment.restart(in_handler=label.casefold() in handlers)
        for statement in statements:
            if with_err and with_err[-1]:
                statement = WITH_MEMBER.sub("Err.", statement)
            findings += check_statement(path, procedure, line, statement, segment, names)
            if WITH.match(statement):
                with_err.append(bool(WITH_ERR.match(statement)))
            elif END_WITH.match(statement) and with_err:
                with_err.pop()
            if IF_START.match(statement):
                # The rest of the line belongs to a single-line If's branches.
                segment.restart()
                break
    return findings


def procedures(source: str) -> list[Procedure]:
    """Each procedure's logical lines: ``(line, label, statements)``, strings blanked."""
    found: list[Procedure] = []
    current: Procedure | None = None
    for line, _, kind, code in logical_units(source.split("\n"), blank_strings=True):
        if kind == "directive":
            continue
        label = label_at_start(code) if current is not None else None
        if label is not None:
            prefix = LABEL_PREFIX.match(code)
            code = code[prefix.end():] if prefix else code
        owner, statements = current, []
        for statement in split_statements(code):
            opener = is_procedure(statement)
            if opener:
                current = owner = Procedure(opener.group(2), line)
                found.append(current)
            elif PROC_CLOSE.match(statement):
                current = None
            elif current is not None and statement:
                statements.append(statement)
        if owner is not None and (statements or label is not None):
            owner.units.append((line, label, statements))
    return found


def procedure_names(module: Module, project: Names, line: int) -> Names:
    """The names the procedure at ``line`` resolves. Its values are its
    parameters and variables, its own name (its result), the module's
    variables, constants and Enum members, and those of ``project``: the
    names of the components and the public values of the standard modules.
    It can call the procedures, properties and Declares of its module and
    the public ones of the standard modules, in ``project``."""
    procedure = next((item for item in module.procedures if item.line == line), None)
    local = procedure.names | {procedure.name.casefold()} if procedure else set()
    return Names(frozenset(local | (module.names - module.callables) | project.values),
                 frozenset(module.callables | project.callables))


def analyze(modules: dict[str, str]) -> tuple[list[Finding], int]:
    """The findings of one environment's sources, and how many procedures it has."""
    scanned = {path: scan(path, source) for path, source in modules.items()}
    standard = [module for module in scanned.values() if module.standard]
    project = Names(
        frozenset({Path(path).stem.casefold() for path in modules} | {
            name for module in standard for name in module.public - module.public_callables}),
        frozenset(name for module in standard for name in module.public_callables))
    findings: list[Finding] = []
    count = 0
    for path, source in modules.items():
        for procedure in procedures(source):
            count += 1
            names = procedure_names(scanned[path], project, procedure.line)
            findings += check_procedure(path, procedure, names)
    return findings, count


def check(root: Path) -> tuple[list[Finding], int, int]:
    components = tracked_components(root)
    by_environment: dict[str, dict[str, str]] = {}
    for path in components:
        # Unreadable directives are vba-conditionals findings, not repeated here.
        sources, _ = reachable_sources(path, read_component(root, path))
        for environment, source in sources.items():
            by_environment.setdefault(environment, {})[path] = source
    environments: dict[Finding, list[str]] = {}
    count = 0
    for environment, modules in by_environment.items():
        found, procedures_seen = analyze(modules)
        count = max(count, procedures_seen)
        for finding in found:
            names = environments.setdefault(finding, [])
            if environment not in names:
                names.append(environment)
    findings = [Finding(item.path, item.line, f"{item.message} [{', '.join(names)}]")
                for item, names in environments.items()]
    return findings, len(components), count


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components, count = check(args.root.resolve())
    return report("vba-errors", findings, f"{components} VBA component(s), {count} procedure(s)")


if __name__ == "__main__":
    sys.exit(main())
