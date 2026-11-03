Attribute VB_Name = "TEST_Relations"
'==============================================================================
' MODULE: TEST_Relations
'------------------------------------------------------------------------------
' PURPOSE
'   Illustrative relation tests: checks that need no external expected
'   value, run on a stand-in for the calculation under test, the total of a
'   list of amounts. Each case asserts that its relation holds, then that
'   the same comparison rejects a deliberately corrupted output, so that a
'   relation which would pass any output is seen. Nothing here is a
'   calculation of the workbook or a regulatory value.
'
' PUBLIC SURFACE
'   RunRelationTests runs the suite. Option Private Module keeps it out of
'   the workbook's external surface. The suite is not certified: it is not
'   listed in .github/excel-evidence-policy.json.
'
' DEPENDENCIES
'   TEST_Harness.
'
' STATE OWNERSHIP
'   None: the cases build their inputs in memory and change no Excel state.
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises
'   TEST_Harness.ERR_TEST_FAILED, or the unexpected error that escaped the
'   runner. A refused start raises TEST_Harness.ERR_TEST_DIRTY_START.
'
' USAGE
'   In a development workbook holding the modules of tests/modules/, type
'   TEST_Relations.RunRelationTests in the Immediate window. Replace
'   IllustrativeTotal with a call to the core procedure the relations test
'   (docs/methodology/TEST_HARNESS.md, Relation tests).
'
' UPDATED
'   2026-11-03
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the tests inside the project.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once.
        Private Const SUITE_NAME            As String = "relations"
        Private Const EXPECTED_CASES        As Long = 4
        Private Const EXPECTED_ASSERTIONS   As Long = 8
    'The tolerance of every comparison, in currency units and as a proportion
    '(docs/methodology/TEST_CASES.md, Tolerance).
        Private Const ABSOLUTE_TOLERANCE    As Double = 0.000001
        Private Const RELATIVE_TOLERANCE    As Double = 0.000000001
    'The shift that corrupts an output: one currency unit, far outside both
    'bounds for the totals of these cases, which stay below 2,000,000.
        Private Const CORRUPTION            As Double = 1#
    'The amount by which the monotonicity case raises one input.
        Private Const INCREMENT             As Double = 1000#


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunRelationTests()
'
'==============================================================================
'                               RunRelationTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun. A passing
'   run ends with RESULT=PASS; suite=relations; operation=PASS;
'   completeness=COMPLETE; cases=4; assertions=8; failures=0; cleanup=PASS;
'   recovery=NONE.
'
' ERROR POLICY
'   An error escaping a case handler is passed to EndRun, which reports it
'   and re-raises it unchanged. BeginRun's refusal propagates directly.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrNumber        As Long      'Runner error number; 0 for none
    Dim ErrSource        As String    'Runner error source
    Dim ErrDescription   As String    'Runner error description

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        TEST_Harness.BeginRun SUITE_NAME, "NORMAL", EXPECTED_CASES, EXPECTED_ASSERTIONS
        On Error GoTo RunFailed
        CaseZeroInput
        CaseAdditivity
        CaseSymmetry
        CaseMonotonicity

Finish:
        On Error GoTo 0
        TEST_Harness.EndRun ErrNumber, ErrSource, ErrDescription
        Exit Sub

RunFailed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        Resume Finish

End Sub


'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseZeroInput()
'
'==============================================================================
'                                CaseZeroInput
'------------------------------------------------------------------------------
' PURPOSE
'   Zero inputs give a zero total, matched within the absolute bound only;
'   a total shifted by CORRUPTION is rejected. 2 assertions.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Amounts(1 To 3)   As Double    'Three zero amounts, as VBA initializes them
    Dim Total             As Double    'Total of Amounts

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "zero-input"
        Total = IllustrativeTotal(Amounts)
        TEST_Harness.AssertNear "holds", 0#, Total, ABSOLUTE_TOLERANCE, RELATIVE_TOLERANCE
        TEST_Harness.AssertTrue "detects-corruption", _
            Not TEST_Harness.WithinTolerance(0#, Total + CORRUPTION, ABSOLUTE_TOLERANCE, _
            RELATIVE_TOLERANCE), "a corrupted total of zero amounts was accepted"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseAdditivity()
'
'==============================================================================
'                                CaseAdditivity
'------------------------------------------------------------------------------
' PURPOSE
'   The total of two lists together equals the sum of their totals; a
'   combined total shifted by CORRUPTION is rejected. 2 assertions.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FirstList(1 To 2)    As Double    'The first list
    Dim SecondList(1 To 3)   As Double    'The second list
    Dim BothLists(1 To 5)    As Double    'Both lists together
    Dim Index                As Long      'Position in a list
    Dim Whole                As Double    'Total of BothLists
    Dim Parts                As Double    'Sum of the totals of FirstList and SecondList

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Amounts that are not exact in binary, so the sums differ in their last
    'bits and the comparison needs its tolerance.
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "additivity"
        FirstList(1) = 600000.1
        FirstList(2) = 400000.2
        SecondList(1) = 250000.3
        SecondList(2) = 125000.4
        SecondList(3) = 0.5
        For Index = 1 To 2
            BothLists(Index) = FirstList(Index)
        Next Index
        For Index = 1 To 3
            BothLists(2 + Index) = SecondList(Index)
        Next Index
        Whole = IllustrativeTotal(BothLists)
        Parts = IllustrativeTotal(FirstList) + IllustrativeTotal(SecondList)
        TEST_Harness.AssertNear "holds", Parts, Whole, ABSOLUTE_TOLERANCE, RELATIVE_TOLERANCE
        TEST_Harness.AssertTrue "detects-corruption", _
            Not TEST_Harness.WithinTolerance(Parts, Whole + CORRUPTION, ABSOLUTE_TOLERANCE, _
            RELATIVE_TOLERANCE), "a corrupted combined total was accepted"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseSymmetry()
'
'==============================================================================
'                                 CaseSymmetry
'------------------------------------------------------------------------------
' PURPOSE
'   The order of the inputs does not change the total; a total of the
'   reversed list shifted by CORRUPTION is rejected. 2 assertions.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Forward(1 To 4)    As Double    'The list
    Dim Reversed(1 To 4)   As Double    'The same list in reverse order
    Dim Index              As Long      'Position in the list
    Dim Expected           As Double    'Total of Forward
    Dim Actual             As Double    'Total of Reversed

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "symmetry"
        Forward(1) = 0.1
        Forward(2) = 1000000.7
        Forward(3) = 0.2
        Forward(4) = 333333.3
        For Index = 1 To 4
            Reversed(Index) = Forward(5 - Index)
        Next Index
        Expected = IllustrativeTotal(Forward)
        Actual = IllustrativeTotal(Reversed)
        TEST_Harness.AssertNear "holds", Expected, Actual, ABSOLUTE_TOLERANCE, RELATIVE_TOLERANCE
        TEST_Harness.AssertTrue "detects-corruption", _
            Not TEST_Harness.WithinTolerance(Expected, Actual + CORRUPTION, ABSOLUTE_TOLERANCE, _
            RELATIVE_TOLERANCE), "a corrupted total of the reversed list was accepted"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseMonotonicity()
'
'==============================================================================
'                               CaseMonotonicity
'------------------------------------------------------------------------------
' PURPOSE
'   Raising one input by INCREMENT never lowers the total; a raised total
'   corrupted as if the increment had been subtracted is rejected.
'   2 assertions.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Original(1 To 3)   As Double    'The list
    Dim Raised(1 To 3)     As Double    'The same list with its second amount raised
    Dim Index              As Long      'Position in the list
    Dim OriginalTotal      As Double    'Total of Original
    Dim RaisedTotal        As Double    'Total of Raised

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "monotonicity"
        Original(1) = 500000#
        Original(2) = 250000.25
        Original(3) = 125000.125
        For Index = 1 To 3
            Raised(Index) = Original(Index)
        Next Index
        Raised(2) = Raised(2) + INCREMENT
        OriginalTotal = IllustrativeTotal(Original)
        RaisedTotal = IllustrativeTotal(Raised)
        TEST_Harness.AssertTrue "holds", RaisedTotal >= OriginalTotal, _
            "raised total " & CStr(RaisedTotal) & " is below " & CStr(OriginalTotal)
        TEST_Harness.AssertTrue "detects-corruption", _
            Not (RaisedTotal - 2# * INCREMENT >= OriginalTotal), _
            "a corrupted raised total was accepted"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


'
'------------------------------------------------------------------------------
'
'                         EXAMPLE PROCEDURE UNDER TEST
'
'------------------------------------------------------------------------------
'

Private Function IllustrativeTotal( _
    ByRef Amounts() As Double) _
    As Double
'
'==============================================================================
'                              IllustrativeTotal
'------------------------------------------------------------------------------
' PURPOSE
'   The sum of Amounts: a stand-in for the calculation the relations test.
'   Replace it with a call to that calculation's core procedure.
'
' INPUTS
'   Amounts: amounts in one currency, read only; VBA passes an array by
'   reference.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index   As Long      'Position in Amounts
    Dim Total   As Double    'Sum so far

'------------------------------------------------------------------------------
' ADD UP
'------------------------------------------------------------------------------
        For Index = LBound(Amounts) To UBound(Amounts)
            Total = Total + Amounts(Index)
        Next Index
        IllustrativeTotal = Total

End Function
