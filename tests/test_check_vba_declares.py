"""Valid and invalid fixtures for the native declaration check (tools/check_vba_declares.py).

Each fixture is a throw-away Git repository, built as in test_check_vba.py,
holding one synthetic VBA test module that calls Windows exports through
``Declare`` and a register of native declarations. The module splits its
declarations and one call sequence with ``#If VBA7``, so that a defect in one
branch must be reported in that branch's environments only.
"""

from __future__ import annotations

import contextlib
import io
import json
import re
import sys
import unittest
from collections.abc import Mapping
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))

import check_repository  # noqa: E402
import check_vba_conditionals  # noqa: E402
import check_vba_declares  # noqa: E402
import check_vba_identifiers  # noqa: E402
import check_vba_jumps  # noqa: E402
from test_check_vba import repository  # noqa: E402

MODULE = "tests/modules/TEST_Native.bas"
ALL = ("vba6-win32", "vba7-win32", "vba7-win64")
VBA7 = ("vba7-win32", "vba7-win64")
LEGACY = ("vba6-win32",)

# TlsGetValue returns zero both for a slot that holds zero and on failure, so
# the register marks its zero result as ambiguous; SetEnvironmentVariableW
# returns a BOOL, whose zero always means failure.
BASE = '''Attribute VB_Name = "TEST_Native"
Option Explicit
Option Private Module

#If VBA7 Then
Private Declare PtrSafe Function NativeTlsGetValue Lib "kernel32" Alias "TlsGetValue" _
    (ByVal dwTlsIndex As Long) As LongPtr
Private Declare PtrSafe Function SetEnvironmentVariableW Lib "kernel32" _
    (ByVal lpName As LongPtr, ByVal lpValue As LongPtr) As Long
Private Declare PtrSafe Sub NativeSetLastError Lib "kernel32" Alias "SetLastError" _
    (ByVal dwErrCode As Long)
#Else
Private Declare Function NativeTlsGetValue Lib "kernel32" Alias "TlsGetValue" _
    (ByVal dwTlsIndex As Long) As Long
Private Declare Function SetEnvironmentVariableW Lib "kernel32" _
    (ByVal lpName As Long, ByVal lpValue As Long) As Long
Private Declare Sub NativeSetLastError Lib "kernel32" Alias "SetLastError" _
    (ByVal dwErrCode As Long)
#End If


Private Function ReadSlot(ByVal SlotIndex As Long, ByRef SlotOut As Variant, _
    ByRef FailMsg As String) As Boolean
#If VBA7 Then
    Dim SlotValue As LongPtr
#Else
    Dim SlotValue As Long
#End If
    Dim DllError As Long
    On Error GoTo Failed
#If VBA7 Then
    NativeSetLastError 0
    SlotValue = NativeTlsGetValue(SlotIndex)
#Else
    Call NativeSetLastError(0)
    SlotValue = NativeTlsGetValue(SlotIndex)
#End If
    DllError = Err.LastDllError
    If SlotValue = 0 And DllError <> 0 Then
        FailMsg = "TlsGetValue failed: Err.LastDllError " & CStr(DllError) ' Err.LastDllError
        Exit Function
    End If
    SlotOut = SlotValue
    ReadSlot = True
    Exit Function

Failed:
    FailMsg = Err.Description
End Function


Private Function ClearVariable(ByVal VariableName As String, ByRef FailMsg As String) As Boolean
    Dim ApiOK As Long
    Dim DllError As Long
    On Error GoTo Failed
    ApiOK = SetEnvironmentVariableW(StrPtr(VariableName), 0)
    DllError = Err.LastDllError
    If ApiOK = 0 Then
        FailMsg = "SetEnvironmentVariableW failed with error " & CStr(DllError)
    End If
    ClearVariable = (ApiOK <> 0)
    Exit Function

Failed:
    FailMsg = Err.Description
End Function
'''

ENTRIES: list[dict[str, Any]] = [
    {"export": "TlsGetValue", "library": "kernel32", "zero_is_ambiguous": True},
    {"export": "SetEnvironmentVariableW", "library": "kernel32", "zero_is_ambiguous": False},
    {"export": "SetLastError", "library": "kernel32", "zero_is_ambiguous": False},
]

VBA7_CALL = "    NativeSetLastError 0\n    SlotValue = NativeTlsGetValue(SlotIndex)\n#Else"
LEGACY_CALL = "    Call NativeSetLastError(0)\n    SlotValue = NativeTlsGetValue(SlotIndex)\n#End If"
SLOT_CAPTURE = "    DllError = Err.LastDllError\n    If SlotValue"
CLEAR_CAPTURE = "    DllError = Err.LastDllError\n    If ApiOK = 0 Then\n"
LAST_ERROR = '''

Private Function LastError() As Long
    Dim DllError As Long
    DllError = Err.LastDllError
    LastError = DllError
End Function
'''


def register(entries: object = None, **envelope: object) -> str:
    data: dict[str, object] = {"schema_version": 1,
                               "declarations": ENTRIES if entries is None else entries}
    data.update(envelope)
    return json.dumps(data, indent=2) + "\n"


def changed(old: str, new: str, source: str = BASE) -> str:
    assert old in source, old
    return source.replace(old, new, 1)


def findings(source: str | None = BASE, register_text: str | None = None) -> list[str]:
    files: dict[str, str | bytes] = {}
    if source is not None:
        files[MODULE] = source
    if register_text is not None or source is not None:
        files[check_vba_declares.REGISTER] = register_text or register()
    with repository(files) as root:
        return [str(finding) for finding in check_vba_declares.check(root)[0]]


def environments(finding: str) -> set[str]:
    match = re.search(r"\[([^\]]*)\]$", finding)
    return set(match.group(1).split(", ")) if match else set()


QUALIFIED = BASE.replace("NativeSetLastError 0", "TEST_Native.NativeSetLastError 0") \
    .replace("Call NativeSetLastError(0)", "Call test_native.NativeSetLastError(0)") \
    .replace("= NativeTlsGetValue(", "= TEST_Native.NativeTlsGetValue(") \
    .replace("= SetEnvironmentVariableW(", "= TEST_Native . SetEnvironmentVariableW(")
CLEAR_CALL = "    ApiOK = SetEnvironmentVariableW(StrPtr(VariableName), 0)\n"
CONDITION = "inside a condition, a single-line If or a larger expression"

# A correct capture, written another way: each must pass in every environment.
CAPTURE_VALID: dict[str, str] = {
    "calls-qualified-by-the-module-name": QUALIFIED,
    "statement-call-with-call": changed(
        CLEAR_CALL, "    Call SetEnvironmentVariableW(StrPtr(VariableName), 0)\n"),
    "statement-call-without-call": changed(
        CLEAR_CALL, "    SetEnvironmentVariableW StrPtr(VariableName), 0\n"),
    "statement-call-with-a-parenthesized-argument": changed(
        CLEAR_CALL, "    TEST_Native.SetEnvironmentVariableW (StrPtr(VariableName)), 0\n"),
    "line-continuations": changed(
        CLEAR_CAPTURE, CLEAR_CAPTURE.replace("= Err", "= _\n        Err"), changed(
            CLEAR_CALL, "    ApiOK = SetEnvironmentVariableW( _\n        StrPtr(VariableName), 0)\n")),
    "call-and-capture-inside-a-with-block": changed(
        CLEAR_CALL + CLEAR_CAPTURE,
        "    With ThisWorkbook\n        ApiOK = SetEnvironmentVariableW(StrPtr(.Name), 0)\n"
        "        DllError = Err.LastDllError\n    End With\n    If ApiOK = 0 Then\n"),
    "nested-parentheses-and-a-comment": changed(
        VBA7_CALL, VBA7_CALL.replace("(SlotIndex)\n", "((SlotIndex) + CLng(0)) ' (x)\n")),
}

# (source, register, the texts the findings contain, each finding at least
# one of them, and the environments that report them).
CAPTURE_INVALID: dict[str, tuple[str, str | None, str | tuple[str, ...], tuple[str, ...]]] = {
    "statement-between-call-and-capture": (
        changed(CLEAR_CAPTURE, "    FailMsg = vbNullString\n" + CLEAR_CAPTURE),
        None, "not directly after a native call", ALL),
    "on-error-between-call-and-capture": (
        changed(CLEAR_CAPTURE, "    On Error Resume Next\n" + CLEAR_CAPTURE),
        None, "after the On Error statement", ALL),
    "label-between-call-and-capture": (
        changed(CLEAR_CAPTURE, "Retry:\n" + CLEAR_CAPTURE),
        None, "after the label of line", ALL),
    "capture-moved-into-the-failure-branch": (
        changed(CLEAR_CAPTURE, "    If ApiOK = 0 Then\n        DllError = Err.LastDllError\n"),
        None, "not directly after a native call", ALL),
    "capture-inside-an-expression": (
        changed(CLEAR_CAPTURE, CLEAR_CAPTURE.replace("Err.LastDllError",
                                                     "Err.LastDllError + 0")),
        None, "inside a larger statement", ALL),
    "capture-into-a-parameter": (
        changed(CLEAR_CAPTURE, CLEAR_CAPTURE.replace("DllError =", "FailMsg =")),
        None, "not a Dim or Static variable", ALL),
    "capture-before-any-native-call": (
        BASE + LAST_ERROR, None, "before any native call", ALL),
    "missing-clear-in-both-branches": (
        changed(LEGACY_CALL, LEGACY_CALL.replace("    Call NativeSetLastError(0)\n", ""),
                changed(VBA7_CALL, VBA7_CALL.replace("    NativeSetLastError 0\n", ""))),
        None, "without SetLastError 0 directly before it", ALL),
    "clear-not-adjacent-in-the-vba7-branch": (
        changed(VBA7_CALL, VBA7_CALL.replace("0\n", "0\n    DllError = 0\n", 1)),
        None, "without SetLastError 0 directly before it", VBA7),
    "non-zero-clear-in-the-legacy-branch": (
        changed(LEGACY_CALL, LEGACY_CALL.replace("(0)", "(5)")),
        None, "without SetLastError 0 directly before it", LEGACY),
    "expression-after-the-ambiguous-call-in-the-vba7-branch": (
        changed(VBA7_CALL, VBA7_CALL.replace("(SlotIndex)\n", "(SlotIndex) + (1)\n")),
        None, ("without assigning its result to a local variable", CONDITION), VBA7),
    "second-call-after-the-ambiguous-call-in-the-legacy-branch": (
        changed(LEGACY_CALL, LEGACY_CALL.replace("(SlotIndex)\n", "(SlotIndex) Or CLng(0)\n")),
        None, ("without assigning its result to a local variable", CONDITION), LEGACY),
    "qualified-ambiguous-call-without-clear-or-capture": (
        QUALIFIED.replace("    TEST_Native.NativeSetLastError 0\n", "")
        .replace("    Call test_native.NativeSetLastError(0)\n", "")
        .replace(SLOT_CAPTURE, "    If SlotValue"),
        None, ("without SetLastError 0 directly before it",
               "without <local> = Err.LastDllError directly after it"), ALL),
    "ambiguous-call-qualified-by-another-module": (
        changed(VBA7_CALL, VBA7_CALL.replace("= Native", "= TEST_Other.Native")),
        None, "after the statement of line 33, not directly after a native call", VBA7),
    "ambiguous-call-in-a-single-line-if": (
        changed(VBA7_CALL, VBA7_CALL.replace("    SlotValue =",
                                             "    If SlotIndex > 0 Then SlotValue =")),
        None, ("without assigning its result to a local variable", CONDITION), VBA7),
    "call-in-a-single-line-if-that-clears-the-error": (
        changed(CLEAR_CALL, "    If SetEnvironmentVariableW(StrPtr(VariableName), 0) = 0 "
                            "Then Err.Clear\n"),
        None, CONDITION, ALL),
    "call-and-capture-in-a-single-line-if": (
        changed(CLEAR_CALL + CLEAR_CAPTURE,
                "    If VariableName <> vbNullString Then ApiOK = SetEnvironmentVariableW("
                "StrPtr(VariableName), 0): DllError = Err.LastDllError\n    If ApiOK = 0 Then\n"),
        None, CONDITION, ALL),
    "call-in-an-if-condition-before-the-capture": (
        changed(CLEAR_CALL + CLEAR_CAPTURE,
                "    If SetEnvironmentVariableW(StrPtr(VariableName), 0) = 0 Then\n"
                "        DllError = Err.LastDllError\n"),
        None, CONDITION, ALL),
    "call-in-a-select-case-before-the-capture": (
        changed(CLEAR_CALL + CLEAR_CAPTURE,
                "    Select Case SetEnvironmentVariableW(StrPtr(VariableName), 0)\n"
                "    Case 0\n        DllError = Err.LastDllError\n    End Select\n"
                "    If ApiOK = 0 Then\n"),
        None, "not directly after a native call", ALL),
    "call-compared-in-an-assignment": (
        changed(CLEAR_CALL, "    ApiOK = SetEnvironmentVariableW(StrPtr(VariableName), 0) = 0\n"),
        None, CONDITION, ALL),
    "with-err-clears-before-the-capture": (
        changed(CLEAR_CAPTURE, "    With Err\n        .Clear\n        DllError = .LastDllError\n"
                               "    End With\n    If ApiOK = 0 Then\n"),
        None, "reads LastDllError through a With block", ALL),
    "with-err-capture": (
        changed(CLEAR_CAPTURE, "    With Err\n        DllError = .LastDllError\n"
                               "    End With\n    If ApiOK = 0 Then\n"),
        None, "reads LastDllError through a With block", ALL),
    "capture-through-the-vba-qualifier": (
        changed(CLEAR_CAPTURE, CLEAR_CAPTURE.replace("= Err.", "= VBA.Err.")),
        None, "a qualifier other than Err", ALL),
    "result-into-a-parameter-in-the-vba7-branch": (
        changed(VBA7_CALL, VBA7_CALL.replace("SlotValue =", "SlotOut =")),
        None, "without assigning its result to a local variable", VBA7),
    "ambiguous-call-without-capture": (
        changed(SLOT_CAPTURE, "    FailMsg = vbNullString\n" + SLOT_CAPTURE),
        None, "", ALL),
    "get-last-error-declared-in-the-vba7-branch": (
        changed("#Else\nPrivate Declare Function NativeTlsGetValue",
                'Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long\n'
                "#Else\nPrivate Declare Function NativeTlsGetValue"),
        None, "GetLastError", VBA7),
    "unregistered-declare": (
        BASE, register(ENTRIES[:1] + ENTRIES[2:]),
        "imports SetEnvironmentVariableW from kernel32, which docs/native-declarations.json "
        "does not list", ALL),
    "register-names-the-vba-name-instead-of-the-alias": (
        BASE, register([{**ENTRIES[0], "export": "NativeTlsGetValue"}, *ENTRIES[1:]]),
        "Declare NativeTlsGetValue imports TlsGetValue from kernel32", ALL),
    "register-names-another-library": (
        BASE, register([*ENTRIES[:2], {**ENTRIES[2], "library": "user32"}]),
        "Declare NativeSetLastError imports SetLastError from kernel32", ALL),
}

REGISTER_INVALID: dict[str, tuple[str, str]] = {
    "not-json": ("{", "is not valid JSON"),
    "wrong-schema-version": (register(schema_version=2),
                             "needs exactly schema_version 1 and a declarations list"),
    "unknown-top-level-key": (register(kind="native"),
                              "needs exactly schema_version 1 and a declarations list"),
    "declarations-not-a-list": (register({}),
                                "needs exactly schema_version 1 and a declarations list"),
    "entry-not-an-object": (register(["TlsGetValue"]), "declarations[0]: is not an object"),
    "missing-key": (register([{"export": "TlsGetValue", "library": "kernel32"}]),
                    "zero_is_ambiguous is missing"),
    "unknown-key": (register([{**ENTRIES[0], "ordinal": 3}]), "unknown key 'ordinal'"),
    "reserved-arguments": (register([{**ENTRIES[0], "arguments": ["Long"]}]),
                           "arguments is reserved for the export's signature"),
    "reserved-returns": (register([{**ENTRIES[0], "returns": "pointer"}]),
                         "returns is reserved for the export's signature"),
    "bad-export": (register([{**ENTRIES[0], "export": "Tls GetValue"}]),
                   "export is not an export name"),
    "bad-library": (register([{**ENTRIES[0], "library": " kernel32"}]),
                    "library is not a library name"),
    "integer-flag": (register([{**ENTRIES[0], "zero_is_ambiguous": 1}]),
                     "zero_is_ambiguous is not true or false"),
    "duplicate-ignoring-case-and-suffix": (
        register([*ENTRIES, {"export": "TLSGETVALUE", "library": "Kernel32.dll",
                             "zero_is_ambiguous": True}]),
        "declarations[3]: TLSGETVALUE from Kernel32.dll is listed twice"),
    "get-last-error-entry": (
        register([{"export": "GetLastError", "library": "kernel32", "zero_is_ambiguous": False}]),
        "GetLastError is never declared, so it has no entry"),
}


class DeclareTest(unittest.TestCase):
    def assertReported(self, found: list[str], text: str | tuple[str, ...],
                       expected: tuple[str, ...]) -> None:
        texts = (text,) if isinstance(text, str) else text
        self.assertTrue(found, "no finding")
        for finding in found:
            self.assertTrue(finding.startswith(f"{MODULE}:"), finding)
            self.assertTrue(any(item in finding for item in texts), finding)
        for item in texts:
            self.assertTrue(any(item in finding for finding in found), (item, found))
        self.assertEqual(set().union(*map(environments, found)), set(expected), found)

    def test_repository_register_is_valid_and_empty(self) -> None:
        data = json.loads((ROOT / check_vba_declares.REGISTER).read_text(encoding="utf-8"))
        self.assertEqual(data, {"schema_version": 1, "declarations": []})
        found, _, _, registered = check_vba_declares.check(ROOT)
        self.assertEqual((found, registered), ([], 0))

    def test_correct_capture_passes_in_every_branch(self) -> None:
        self.assertEqual(findings(), [])
        with repository({MODULE: BASE, check_vba_declares.REGISTER: register()}) as root:
            _, components, declares, registered = check_vba_declares.check(root)
        self.assertEqual((components, declares, registered), (1, 6, 3))

    def test_fixture_passes_the_other_vba_checks(self) -> None:
        files: Mapping[str, str | bytes] = {MODULE: BASE, "docs/PUBLIC_API.txt": "# None\n"}
        with repository(files) as root:
            for check in (check_vba_conditionals.check, check_vba_jumps.check,
                          check_vba_identifiers.check):
                with self.subTest(check=check.__module__):
                    self.assertEqual([str(item) for item in check(root)[0]], [])

    def test_export_matches_ignoring_case_and_the_dll_suffix(self) -> None:
        upper = [{**ENTRIES[0], "export": "TLSGETVALUE", "library": "KERNEL32.DLL"},
                 *ENTRIES[1:]]
        self.assertEqual(findings(register_text=register(upper)), [])
        self.assertEqual(findings(BASE.replace('Lib "kernel32"', 'Lib "kernel32.dll"')), [])

    def test_other_forms_of_a_correct_capture_pass(self) -> None:
        for name, source in CAPTURE_VALID.items():
            with self.subTest(name=name):
                self.assertEqual(findings(source), [])

    def test_strings_and_comments_are_not_reads(self) -> None:
        # BASE quotes Err.LastDllError in a string and in a comment of ReadSlot.
        self.assertIn('"TlsGetValue failed: Err.LastDllError "', BASE)
        self.assertIn("' Err.LastDllError", BASE)
        self.assertEqual(findings(), [])

    def test_invalid_capture_sequences_are_reported_in_their_environments(self) -> None:
        for name, (source, register_text, text, expected) in CAPTURE_INVALID.items():
            with self.subTest(name=name):
                self.assertReported(findings(source, register_text), text, expected)

    def test_ambiguous_call_without_capture_names_the_export(self) -> None:
        source, register_text, _, _ = CAPTURE_INVALID["ambiguous-call-without-capture"]
        found = findings(source, register_text)
        self.assertTrue(any("ReadSlot calls NativeTlsGetValue (TlsGetValue), whose zero "
                            "result is ambiguous, without <local> = Err.LastDllError "
                            "directly after it" in item for item in found), found)
        self.assertTrue(any("ReadSlot reads Err.LastDllError after the statement of line"
                            in item for item in found), found)

    def test_get_last_error_is_reported_with_its_line(self) -> None:
        source, _, _, _ = CAPTURE_INVALID["get-last-error-declared-in-the-vba7-branch"]
        line = source.split("\n").index(
            'Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long') + 1
        self.assertIn(f"{MODULE}:{line}: Declare GetLastError imports GetLastError; copy "
                      "Err.LastDllError directly after the native call instead "
                      "[vba7-win32, vba7-win64]", findings(source))

    def test_unregistered_declare_is_reported_in_each_branch(self) -> None:
        source, register_text, _, _ = CAPTURE_INVALID["unregistered-declare"]
        lines = source.split("\n")
        vba7 = lines.index("Private Declare PtrSafe Function SetEnvironmentVariableW "
                           'Lib "kernel32" _') + 1
        legacy = lines.index('Private Declare Function SetEnvironmentVariableW Lib "kernel32" _') \
            + 1
        found = findings(source, register_text)
        self.assertEqual(sorted(found), sorted(
            f"{MODULE}:{line}: Declare SetEnvironmentVariableW imports SetEnvironmentVariableW "
            f"from kernel32, which docs/native-declarations.json does not list [{names}]"
            for line, names in ((vba7, "vba7-win32, vba7-win64"), (legacy, "vba6-win32"))))

    def test_unreadable_declare_is_reported(self) -> None:
        source = changed('Private Declare Function SetEnvironmentVariableW Lib "kernel32" _',
                         "Private Declare Function SetEnvironmentVariableW _")
        self.assertTrue(any("cannot read the Lib clause of Declare SetEnvironmentVariableW "
                            "[vba6-win32]" in item for item in findings(source)))

    def test_register_must_exist(self) -> None:
        with repository({MODULE: BASE}) as root:
            found = [str(item) for item in check_vba_declares.check(root)[0]]
        self.assertIn("docs/native-declarations.json: is missing; restore it with an empty "
                      "declarations list", found)

    def test_invalid_registers_are_reported(self) -> None:
        for name, (text, expected) in REGISTER_INVALID.items():
            with self.subTest(name=name):
                found = findings(None, text)
                self.assertTrue(any(item.startswith(check_vba_declares.REGISTER) and
                                    expected in item for item in found), found)

    def test_main_exits_0_when_valid_and_1_when_invalid(self) -> None:
        invalid = CAPTURE_INVALID["capture-inside-an-expression"][0]
        for source, expected in ((BASE, 0), (invalid, 1)):
            files = {MODULE: source, check_vba_declares.REGISTER: register()}
            with self.subTest(expected=expected), repository(files) as root, \
                    contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(check_vba_declares.main(["--root", str(root)]), expected)
                verdict = output.getvalue().splitlines()[-1]
                self.assertIn(("PASS" if expected == 0 else "FAIL") + " (1 VBA component(s), "
                              "6 Declare statement(s), 3 registered export(s); static check, "
                              "nothing compiled or run)", verdict)

    def test_repository_integrity_runs_it_as_a_blocking_check(self) -> None:
        listed = {check.name: check for check in check_repository.checks(
            check_repository.Context(ROOT))}
        self.assertEqual(listed["vba-declares"].command,
                         [sys.executable, "tools/check_vba_declares.py"])
        self.assertFalse(listed["vba-declares"].advisory)


if __name__ == "__main__":
    unittest.main()
