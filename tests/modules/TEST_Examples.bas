Attribute VB_Name = "TEST_Examples"
'==============================================================================
' MODULE: TEST_Examples
'------------------------------------------------------------------------------
' PURPOSE
'   Neutral example suite of TEST_Harness: one case per kind of assertion,
'   one case that records the results of the illustrative test cases, and
'   the reference implementation of the cleanup pattern of
'   docs/VBA_HOUSE_STYLE.md (Operation and cleanup outcomes) with its
'   failure-path cases. Nothing here is a calculation of the workbook or a
'   regulatory value.
'
' PUBLIC SURFACE
'   RunExampleTests is the documented test entry point.
'   RunExampleTestsWithInjectedFailure and RunExampleTestsWithSkippedCase
'   show what a failed and an incomplete run report. CertifyExamples
'   certifies all three and saves their logs, the certification log and the
'   results of the normal run as the evidence of an Excel run. Each takes
'   the folder of the test cases, tests/fixtures/ of the checkout.
'   Option Private Module keeps them out of the workbook's external surface.
'
' DEPENDENCIES
'   TEST_Harness; the VBA and Excel object models.
'
' STATE OWNERSHIP
'   Owns the mode and the fixtures folder of the running suite. Reads the
'   illustrative test cases of tests/fixtures/. RunExampleOperation changes
'   Application.Calculation and Application.ScreenUpdating and restores them
'   before it returns, on the success and the error path.
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises ERR_TEST_FAILED,
'   or the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START. CertifyExamples records those errors
'   instead of raising them, saves every log, and raises only when the
'   certification did not pass.
'
' USAGE
'   In a development workbook holding the modules of tests/modules/, type
'   TEST_Examples.RunExampleTests "<checkout>\tests\fixtures" in the
'   Immediate window. To record an Excel run, type
'   TEST_Examples.CertifyExamples "<folder>", "<checkout>\tests\fixtures"
'   instead (INSTALLATION.md, Recording an Excel run).
'
' UPDATED
'   2026-11-04
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
        Private Const SUITE_NAME              As String = "examples"
        Private Const EXPECTED_CASES          As Long = 14
        Private Const EXPECTED_ASSERTIONS     As Long = 34
    'Errors of the example procedures, test-only, as literals equal to
    'vbObjectError + n (see TEST_Harness, Module constants).
        Private Const ERR_EXAMPLE_INVALID     As Long = &H80040FF0   'vbObjectError + 4080: Input outside the domain
        Private Const ERR_EXAMPLE_OPERATION   As Long = &H80040FF1   'vbObjectError + 4081: Primary error of the operation
        Private Const ERR_EXAMPLE_RESTORE     As Long = &H80040FF2   'vbObjectError + 4082: Simulated restore failure
        Private Const ERR_CLEANUP_FAILED      As Long = &H80040FF3   'vbObjectError + 4083: Only the cleanup failed
        Private Const ERR_EXAMPLE_HOST        As Long = &H80040FF4   'vbObjectError + 4084: Simulated capture or change failure
        Private Const ERR_EXAMPLE_FIXTURE     As Long = &H80040FF5   'vbObjectError + 4085: Test case file not as expected
    'The quantity of the illustrative test cases (tests/expected/).
        Private Const ILLUSTRATIVE_QUANTITY   As String = "illustrative_total_notional_amount"

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mMode       As String    'NORMAL, INJECTED_FAILURE or SKIPPED_CASE
        Private mFixtures   As String    'Folder of the test cases, tests/fixtures/


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunExampleTests( _
    ByVal Fixtures As String)
'
'==============================================================================
'                               RunExampleTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run the example suite. A passing run ends with
'   RESULT=PASS; suite=examples; operation=PASS; completeness=COMPLETE;
'   cases=14; assertions=34; failures=0; cleanup=PASS; recovery=NONE.
'
' INPUTS
'   Fixtures: the full path of tests/fixtures/ in the checkout, without a
'   trailing separator.
'
' UPDATED
'   2026-11-04
'==============================================================================
'
        RunSuite "NORMAL", Fixtures

End Sub


Public Sub RunExampleTestsWithInjectedFailure( _
    ByVal Fixtures As String)
'
'==============================================================================
'                      RunExampleTestsWithInjectedFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Show the failure path: exact-double expects a wrong value. Expected
'   end: RESULT=FAIL; ...; operation=FAIL; completeness=COMPLETE; cases=14;
'   assertions=34; failures=1, then the error ERR_TEST_FAILED.
'
' INPUTS
'   Fixtures: as RunExampleTests.
'
' UPDATED
'   2026-11-04
'==============================================================================
'
        RunSuite "INJECTED_FAILURE", Fixtures

End Sub


Public Sub RunExampleTestsWithSkippedCase( _
    ByVal Fixtures As String)
'
'==============================================================================
'                        RunExampleTestsWithSkippedCase
'------------------------------------------------------------------------------
' PURPOSE
'   Show that an incomplete run is never a pass: the last case is skipped.
'   Expected end: RESULT=FAIL; ...; operation=PASS;
'   completeness=INCOMPLETE; cases=13; assertions=32; failures=1, the
'   failure being harness.completeness, then ERR_TEST_FAILED.
'
' INPUTS
'   Fixtures: as RunExampleTests.
'
' UPDATED
'   2026-11-04
'==============================================================================
'
        RunSuite "SKIPPED_CASE", Fixtures

End Sub


Public Sub CertifyExamples( _
    ByVal Folder As String, _
    ByVal Fixtures As String)
'
'==============================================================================
'                               CertifyExamples
'------------------------------------------------------------------------------
' PURPOSE
'   The certification entry point: refuse state left by an earlier run,
'   run the suite in its three modes, save each complete log in Folder
'   (examples.log, injected.log and skipped.log) and the results of the
'   normal run (results.json), then save certification.log.
'   tools/check_excel_run.py checks the folder against the commit.
'
' INPUTS
'   Folder: an existing folder outside the checkout or in the ignored
'   test-results/ directory, without a trailing backslash. Fixtures: the
'   full path of tests/fixtures/ in the same checkout, without a trailing
'   separator.
'
' OUTCOMES
'   The normal run must pass and be complete. The injected-failure and
'   skipped-case runs must fail and reach this caller as
'   TEST_Harness.ERR_TEST_FAILED with a description: they are the
'   self-test that a failed run is never silent.
'
' ERROR POLICY
'   The error that ends each run is recorded, not raised, so that every log
'   is written. EndCertification raises ERR_TEST_FAILED after writing
'   certification.log when the certification did not pass. A refused start
'   raises, before any run, TEST_Harness.ERR_TEST_DIRTY_START when a run is
'   active or recovery is required, or TEST_Harness.ERR_TEST_USAGE while an
'   earlier certification is still active. An error of SaveLog or
'   SaveResults propagates; run TEST_Harness.ResetTests before trying again.
'
' UPDATED
'   2026-11-04
'==============================================================================
'
        TEST_Harness.BeginCertification SUITE_NAME
        CertifyMode "NORMAL", True, Folder & "\examples.log", Fixtures
        TEST_Harness.SaveResults Folder & "\results.json"
        CertifyMode "INJECTED_FAILURE", False, Folder & "\injected.log", Fixtures
        CertifyMode "SKIPPED_CASE", False, Folder & "\skipped.log", Fixtures
        TEST_Harness.EndCertification Folder & "\certification.log"

End Sub


'
'------------------------------------------------------------------------------
'
'                                    RUNNER
'
'------------------------------------------------------------------------------
'

Private Sub RunSuite( _
    ByVal Mode As String, _
    ByVal Fixtures As String)
'
'==============================================================================
'                                   RunSuite
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun.
'
' INPUTS
'   Mode: NORMAL, INJECTED_FAILURE or SKIPPED_CASE. Fixtures: the folder of
'   the test cases, as RunExampleTests.
'
' ERROR POLICY
'   An error escaping a case handler is passed to EndRun, which reports it
'   and re-raises it unchanged. BeginRun's refusal propagates directly.
'
' UPDATED
'   2026-11-04
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
        mMode = Mode
        mFixtures = Fixtures
        TEST_Harness.BeginRun SUITE_NAME, Mode, EXPECTED_CASES, EXPECTED_ASSERTIONS
        On Error GoTo RunFailed
        CaseExactDouble
        CaseAbsoluteTolerance
        CaseRelativeTolerance
        CaseExactText
        CaseErrorValue
        CaseRaisedError
        CaseIllustrativeResults
        CaseCleanupRestoresState
        CaseCleanupKeepsPrimaryError
        CaseCleanupPartialCapture
        CaseCleanupChangeFailure
        CaseCleanupPrimaryOnly
        CaseCleanupEachRestorationFailure
        If mMode <> "SKIPPED_CASE" Then
            CaseCleanupOnlyFailure
        End If

Finish:
        On Error GoTo 0
        mMode = vbNullString
        mFixtures = vbNullString
        TEST_Harness.EndRun ErrNumber, ErrSource, ErrDescription
        Exit Sub

RunFailed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        Resume Finish

End Sub


Private Sub CertifyMode( _
    ByVal Mode As String, _
    ByVal MustPass As Boolean, _
    ByVal Path As String, _
    ByVal Fixtures As String)
'
'==============================================================================
'                                 CertifyMode
'------------------------------------------------------------------------------
' PURPOSE
'   Run the suite in Mode, save its log to Path, then certify how the run
'   reached this caller.
'
' INPUTS
'   MustPass: True for the normal run, False for a deliberate failure mode.
'   Fixtures: the folder of the test cases, as RunExampleTests.
'
' ERROR POLICY
'   The run's own error is captured, printed and certified after the log is
'   saved; an error of SaveLog propagates.
'
' UPDATED
'   2026-11-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RunErrorNumber        As Long      'Error that ended the run; 0 for none
    Dim RunErrorDescription   As String    'Its description

'------------------------------------------------------------------------------
' RUN AND SAVE
'------------------------------------------------------------------------------
        On Error Resume Next
        RunSuite Mode, Fixtures
        RunErrorNumber = Err.Number
        RunErrorDescription = Err.Description
        On Error GoTo 0
        TEST_Harness.SaveLog Path
        TEST_Harness.Echo "SAVED=" & Path & "; mode=" & Mode & "; error=" & CStr(RunErrorNumber) & _
            IIf(RunErrorNumber = 0, vbNullString, " " & RunErrorDescription)
        TEST_Harness.CertifyRun Mode, MustPass, RunErrorNumber, RunErrorDescription

End Sub


'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseExactDouble()
'
'==============================================================================
'                               CaseExactDouble
'------------------------------------------------------------------------------
' PURPOSE
'   Exact equality on a result that is exact in binary. 1 assertion.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Expected   As Double    'Expected sum; wrong on purpose when injected

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "exact-double"
        Expected = 0.75
        If mMode = "INJECTED_FAILURE" Then
            Expected = 0.76
        End If
        TEST_Harness.AssertEqualDouble "half-plus-quarter", Expected, 0.5 + 0.25
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseAbsoluteTolerance()
'
'==============================================================================
'                            CaseAbsoluteTolerance
'------------------------------------------------------------------------------
' PURPOSE
'   The absolute bound: 0.1 + 0.2 is not exactly 0.3 in binary, and near
'   zero only the absolute bound counts. 2 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "tolerance-absolute"
        TEST_Harness.AssertNear "tenth-plus-fifth", 0.3, 0.1 + 0.2, 0.000000000000001, 0#
        TEST_Harness.AssertNear "near-zero", 0#, 1E-20, 0.000000000001, 0.000000001
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseRelativeTolerance()
'
'==============================================================================
'                            CaseRelativeTolerance
'------------------------------------------------------------------------------
' PURPOSE
'   The relative bound governs large values, the bound is inclusive, and a
'   value outside both bounds is rejected. 3 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "tolerance-relative"
        TEST_Harness.AssertNear "large-value", 1000000000#, 1000000000.1, 0.000001, 0.000000001
        TEST_Harness.AssertTrue "bound-is-inclusive", _
            TEST_Harness.WithinTolerance(2#, 2.5, 0.5, 0#), "|2.5 - 2| = 0.5 must be within 0.5"
        TEST_Harness.AssertTrue "outside-both-bounds", _
            Not TEST_Harness.WithinTolerance(1#, 1.1, 0.01, 0.01), _
            "|1.1 - 1| = 0.1 must exceed max(0.01, 0.01 * 1)"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseExactText()
'
'==============================================================================
'                                CaseExactText
'------------------------------------------------------------------------------
' PURPOSE
'   Text is compared exactly, case included. 2 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "text-exact"
        TEST_Harness.AssertEqualText "concatenation", "CASE-1", "CASE" & "-" & CStr(1)
        TEST_Harness.AssertTrue "case-matters", _
            StrComp("a", "A", vbBinaryCompare) <> 0, "a and A must differ"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseErrorValue()
'
'==============================================================================
'                                CaseErrorValue
'------------------------------------------------------------------------------
' PURPOSE
'   An expected worksheet error value, returned rather than raised.
'   1 assertion.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "error-value"
        TEST_Harness.AssertErrorValue "zero-denominator", xlErrDiv0, ExampleRatio(1#, 0#)
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseRaisedError()
'
'==============================================================================
'                               CaseRaisedError
'------------------------------------------------------------------------------
' PURPOSE
'   An expected raised error, captured on the line after the call, and the
'   same call with a valid input raising nothing. 2 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result        As Double    'Value of the checked square root
    Dim RaisedError   As Long      'Err.Number right after the call

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "raised-error"
        On Error Resume Next
        Result = ExampleSquareRoot(-1#)
        RaisedError = Err.Number
        On Error GoTo CaseFailed
        TEST_Harness.AssertRaised "negative-input", ERR_EXAMPLE_INVALID, RaisedError
        On Error Resume Next
        Result = ExampleSquareRoot(4#)
        RaisedError = Err.Number
        On Error GoTo CaseFailed
        TEST_Harness.AssertRaised "valid-input", 0, RaisedError
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseIllustrativeResults()
'
'==============================================================================
'                           CaseIllustrativeResults
'------------------------------------------------------------------------------
' PURPOSE
'   The results of an Excel run: read the illustrative test cases from the
'   fixtures folder, add up the notional amounts of each, and record each
'   total with TEST_Harness.RecordResult, for the accuracy contract ACC-01
'   (tests/accuracy/contracts.json). tools/check_accuracy.py compares the
'   totals with their expected values; this case checks only that each file
'   held amounts. 3 assertions.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Total     As Double    'Sum of the notional amounts of one file
    Dim Records   As Long      'Amounts read from that file

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "illustrative-results"
        Total = SumJsonNumbers(FixturePath("illustrative-format-example.json"), _
            "notional_amount", Records)
        TEST_Harness.AssertTrue "format-example-amounts", Records > 0, _
            "illustrative-format-example.json holds no notional_amount"
        TEST_Harness.RecordResult "illustrative-format-example", vbNullString, _
            ILLUSTRATIVE_QUANTITY, vbNullString, vbNullString, Total
        Total = SumCsvColumn(FixturePath( _
            "illustrative-file-import-example/crlf-line-endings/positions.csv"), "notional", Records)
        TEST_Harness.AssertTrue "crlf-variant-amounts", Records > 0, _
            "the crlf-line-endings positions.csv holds no notional amount"
        TEST_Harness.RecordResult "illustrative-file-import-example", "crlf-line-endings", _
            ILLUSTRATIVE_QUANTITY, vbNullString, vbNullString, Total
        Total = SumCsvColumn(FixturePath( _
            "illustrative-file-import-example/lf-line-endings/positions.csv"), "notional", Records)
        TEST_Harness.AssertTrue "lf-variant-amounts", Records > 0, _
            "the lf-line-endings positions.csv holds no notional amount"
        TEST_Harness.RecordResult "illustrative-file-import-example", "lf-line-endings", _
            ILLUSTRATIVE_QUANTITY, vbNullString, vbNullString, Total
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseCleanupRestoresState()
'
'==============================================================================
'                           CaseCleanupRestoresState
'------------------------------------------------------------------------------
' PURPOSE
'   A successful operation restores what it changed and reports no cleanup
'   failure. 2 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details       As String     'Cleanup failures reported by the operation
    Dim SavedScreen   As Boolean    'ScreenUpdating before the operation

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "cleanup-restores-state"
        SavedScreen = Application.ScreenUpdating
        RunExampleOperation vbNullString, False, False, Details
        TEST_Harness.AssertEqualText "no-cleanup-failure", vbNullString, Details
        TEST_Harness.AssertTrue "screen-updating-restored", _
            Application.ScreenUpdating = SavedScreen, "ScreenUpdating differs after the operation"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseCleanupKeepsPrimaryError()
'
'==============================================================================
'                         CaseCleanupKeepsPrimaryError
'------------------------------------------------------------------------------
' PURPOSE
'   When the operation and its cleanup both fail, the caller receives the
'   primary error, and the cleanup failure is still reported. 3 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details       As String    'Cleanup failures reported by the operation
    Dim RaisedError   As Long      'Err.Number right after the call

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "cleanup-keeps-primary-error"
        On Error Resume Next
        RunExampleOperation "work", True, True, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
        TEST_Harness.AssertRaised "primary-error", ERR_EXAMPLE_OPERATION, RaisedError
        TEST_Harness.AssertTrue "calculation-failure-reported", _
            InStr(Details, "Calculation not restored") > 0, "details='" & Details & "'"
        TEST_Harness.AssertTrue "screen-failure-reported", _
            InStr(Details, "ScreenUpdating not restored") > 0, "details='" & Details & "'"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseCleanupPartialCapture()
'
'==============================================================================
'                          CaseCleanupPartialCapture
'------------------------------------------------------------------------------
' PURPOSE
'   When capture fails part-way, nothing is restored: no setting was
'   changed, and an uncaptured value is never restored. Both restore seams
'   are set, so any restoration attempt would be reported. 2 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details       As String    'Cleanup failures reported by the operation
    Dim RaisedError   As Long      'Err.Number right after the call

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "cleanup-partial-capture"
        On Error Resume Next
        RunExampleOperation "capture-screen", True, True, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
        TEST_Harness.AssertRaised "capture-error", ERR_EXAMPLE_HOST, RaisedError
        TEST_Harness.AssertEqualText "nothing-restored", vbNullString, Details
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseCleanupChangeFailure()
'
'==============================================================================
'                          CaseCleanupChangeFailure
'------------------------------------------------------------------------------
' PURPOSE
'   When a change fails after its restoration was registered, that setting
'   is still restored, as is the one changed before it, and the change's
'   error stays the primary error. 3 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details       As String     'Cleanup failures reported by the operation
    Dim RaisedError   As Long       'Err.Number right after the call
    Dim SavedScreen   As Boolean    'ScreenUpdating before the operation

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "cleanup-change-failure"
        SavedScreen = Application.ScreenUpdating
        On Error Resume Next
        RunExampleOperation "change-calculation", True, False, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
        TEST_Harness.AssertRaised "change-error", ERR_EXAMPLE_HOST, RaisedError
        TEST_Harness.AssertTrue "calculation-restore-attempted", _
            InStr(Details, "Calculation not restored") > 0, "details='" & Details & "'"
        TEST_Harness.AssertTrue "screen-updating-restored", _
            Application.ScreenUpdating = SavedScreen, "ScreenUpdating differs after the operation"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseCleanupPrimaryOnly()
'
'==============================================================================
'                            CaseCleanupPrimaryOnly
'------------------------------------------------------------------------------
' PURPOSE
'   When only the operation fails and every restoration succeeds, the
'   caller receives the primary error, no cleanup failure is reported and
'   both settings are restored: a failed operation stays failed even if its
'   cleanup succeeded. 4 assertions.
'
' UPDATED
'   2026-11-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details            As String           'Cleanup failures reported by the operation
    Dim RaisedError        As Long             'Err.Number right after the call
    Dim SavedCalculation   As XlCalculation    'Calculation before the operation
    Dim SavedScreen        As Boolean          'ScreenUpdating before the operation

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "cleanup-primary-only"
        SavedCalculation = Application.Calculation
        SavedScreen = Application.ScreenUpdating
        On Error Resume Next
        RunExampleOperation "work", False, False, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
        TEST_Harness.AssertRaised "primary-error", ERR_EXAMPLE_OPERATION, RaisedError
        TEST_Harness.AssertEqualText "no-cleanup-failure", vbNullString, Details
        TEST_Harness.AssertTrue "calculation-restored", _
            Application.Calculation = SavedCalculation, "Calculation differs after the operation"
        TEST_Harness.AssertTrue "screen-updating-restored", _
            Application.ScreenUpdating = SavedScreen, "ScreenUpdating differs after the operation"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseCleanupEachRestorationFailure()
'
'==============================================================================
'                      CaseCleanupEachRestorationFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Each restoration failure on its own is reported, names only its
'   setting and does not stop the other restoration. 4 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details       As String    'Cleanup failures reported by the operation
    Dim RaisedError   As Long      'Err.Number right after the call

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "cleanup-each-restoration-failure"
        On Error Resume Next
        RunExampleOperation vbNullString, True, False, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
        TEST_Harness.AssertRaised "calculation-cleanup-error", ERR_CLEANUP_FAILED, RaisedError
        TEST_Harness.AssertTrue "only-calculation-reported", _
            InStr(Details, "Calculation not restored") > 0 _
            And InStr(Details, "ScreenUpdating") = 0, "details='" & Details & "'"
        On Error Resume Next
        RunExampleOperation vbNullString, False, True, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
        TEST_Harness.AssertRaised "screen-cleanup-error", ERR_CLEANUP_FAILED, RaisedError
        TEST_Harness.AssertTrue "only-screen-reported", _
            InStr(Details, "ScreenUpdating not restored") > 0 _
            And InStr(Details, "Calculation") = 0, "details='" & Details & "'"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseCleanupOnlyFailure()

'==============================================================================
'                            CaseCleanupOnlyFailure
'------------------------------------------------------------------------------
' PURPOSE
'   When only the cleanup fails, the operation raises ERR_CLEANUP_FAILED,
'   never success. 2 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details       As String    'Cleanup failures reported by the operation
    Dim RaisedError   As Long      'Err.Number right after the call

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "cleanup-only-failure"
        On Error Resume Next
        RunExampleOperation vbNullString, True, True, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
        TEST_Harness.AssertRaised "cleanup-error", ERR_CLEANUP_FAILED, RaisedError
        TEST_Harness.AssertTrue "failure-reported", Len(Details) > 0, "no cleanup detail returned"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


'
'------------------------------------------------------------------------------
'
'                         EXAMPLE PROCEDURES UNDER TEST
'
'------------------------------------------------------------------------------
'

Private Function ExampleRatio( _
    ByVal Numerator As Double, _
    ByVal Denominator As Double) _
    As Variant
'
'==============================================================================
'                                 ExampleRatio
'------------------------------------------------------------------------------
' PURPOSE
'   Numerator / Denominator, or the worksheet error #DIV/0! for a zero
'   denominator, as a worksheet function returns errors (house style).
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        If Denominator = 0# Then
            ExampleRatio = CVErr(xlErrDiv0)
        Else
            ExampleRatio = Numerator / Denominator
        End If

End Function


Private Function ExampleSquareRoot( _
    ByVal Value As Double) _
    As Double
'
'==============================================================================
'                              ExampleSquareRoot
'------------------------------------------------------------------------------
' PURPOSE
'   Square root of Value >= 0; a negative Value raises ERR_EXAMPLE_INVALID,
'   as a core procedure raises (house style).
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        If Value < 0# Then
            Err.Raise ERR_EXAMPLE_INVALID, "TEST_Examples.ExampleSquareRoot", _
                "Value must be >= 0; got " & CStr(Value) & "."
        End If
        ExampleSquareRoot = Sqr(Value)

End Function


Private Function FixturePath( _
    ByVal Relative As String) _
    As String
'
'==============================================================================
'                                 FixturePath
'------------------------------------------------------------------------------
' PURPOSE
'   The full path of a file of tests/fixtures/, given relative to it with
'   "/" separators. The separator is the one the fixtures folder uses: "/"
'   when it has only forward slashes, as in LibreOffice on Linux, otherwise
'   "\".
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Separator   As String    'Path separator of the fixtures folder

'------------------------------------------------------------------------------
' COMPOSE
'------------------------------------------------------------------------------
        If Len(mFixtures) = 0 Then
            Err.Raise ERR_EXAMPLE_FIXTURE, "TEST_Examples.FixturePath", _
                "No fixtures folder: pass the full path of tests/fixtures."
        End If
        Separator = "\"
        If InStr(mFixtures, "/") > 0 And InStr(mFixtures, "\") = 0 Then
            Separator = "/"
        End If
        FixturePath = mFixtures & Separator & Replace(Relative, "/", Separator)

End Function


Private Function ReadFixture( _
    ByVal Path As String) _
    As String
'
'==============================================================================
'                                 ReadFixture
'------------------------------------------------------------------------------
' PURPOSE
'   The bytes of an ASCII test case file, as text with its line endings
'   kept.
'
' ERROR POLICY
'   A missing or unreadable file raises the file error; a byte above 127
'   raises ERR_EXAMPLE_FIXTURE. The file is closed first.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FileNumber         As Integer    'Handle of the open file
    Dim Text               As String     'Content of the file
    Dim Index              As Long       'Position in Text
    Dim ErrorNumber        As Long       'File error number
    Dim ErrorSource        As String     'File error source
    Dim ErrorDescription   As String     'File error description

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Access Read raises on a missing file instead of creating it.
        FileNumber = FreeFile
        Open Path For Binary Access Read As #FileNumber
        On Error GoTo ReadFailed
        Text = Space$(LOF(FileNumber))
        Get #FileNumber, , Text
        Close #FileNumber
        On Error GoTo 0

'------------------------------------------------------------------------------
' REQUIRE ASCII
'------------------------------------------------------------------------------
        For Index = 1 To Len(Text)
            If AscW(Mid$(Text, Index, 1)) > 127 Or AscW(Mid$(Text, Index, 1)) < 0 Then
                Err.Raise ERR_EXAMPLE_FIXTURE, "TEST_Examples.ReadFixture", _
                    Path & " is not ASCII at byte " & CStr(Index) & "."
            End If
        Next Index
        ReadFixture = Text
        Exit Function

'------------------------------------------------------------------------------
' HANDLE A READ ERROR
'------------------------------------------------------------------------------
ReadFailed:
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
        On Error GoTo 0
        Close #FileNumber
        Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Function


Private Function SumJsonNumbers( _
    ByVal Path As String, _
    ByVal Key As String, _
    ByRef Records As Long) _
    As Double
'
'==============================================================================
'                                SumJsonNumbers
'------------------------------------------------------------------------------
' PURPOSE
'   The sum of every number that follows "Key": in a JSON fixture, and how
'   many there are. Enough for the illustrative fixture, whose amounts are
'   plain numbers; not a JSON parser.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Text       As String    'Content of the fixture
    Dim Marker     As String    'The key and its colon
    Dim Position   As Long      'Position of the next marker
    Dim Total      As Double    'Sum so far

'------------------------------------------------------------------------------
' ADD UP
'------------------------------------------------------------------------------
    'Val reads the number after the marker with a period as the decimal
    'separator, whatever the regional settings, and skips the blanks.
        Text = ReadFixture(Path)
        Marker = """" & Key & """:"
        Records = 0
        Position = InStr(1, Text, Marker, vbBinaryCompare)
        Do While Position > 0
            Total = Total + Val(Mid$(Text, Position + Len(Marker)))
            Records = Records + 1
            Position = InStr(Position + Len(Marker), Text, Marker, vbBinaryCompare)
        Loop
        SumJsonNumbers = Total

End Function


Private Function SumCsvColumn( _
    ByVal Path As String, _
    ByVal Column As String, _
    ByRef Records As Long) _
    As Double
'
'==============================================================================
'                                 SumCsvColumn
'------------------------------------------------------------------------------
' PURPOSE
'   The sum of one column of a comma-separated file with a header row, and
'   how many rows there are. Lines end in CR LF or LF; empty lines are
'   ignored. Enough for the illustrative positions file, whose fields hold
'   no comma or quote; not a general CSV reader.
'
' ERROR POLICY
'   Raises ERR_EXAMPLE_FIXTURE when the header lacks Column or a row lacks
'   the field.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Lines     As Variant    'The file's lines
    Dim Header    As Variant    'The header's fields
    Dim Fields    As Variant    'One row's fields
    Dim Target    As Long       'Index of Column, -1 until found
    Dim Index     As Long       'Line or field index
    Dim Total     As Double     'Sum so far

'------------------------------------------------------------------------------
' FIND THE COLUMN
'------------------------------------------------------------------------------
        Lines = Split(Replace(ReadFixture(Path), vbCr, vbNullString), vbLf)
        Header = Split(Lines(0), ",")
        Target = -1
        For Index = 0 To UBound(Header)
            If Header(Index) = Column Then
                Target = Index
            End If
        Next Index
        If Target < 0 Then
            Err.Raise ERR_EXAMPLE_FIXTURE, "TEST_Examples.SumCsvColumn", _
                Path & " has no column " & Column & "."
        End If

'------------------------------------------------------------------------------
' ADD UP
'------------------------------------------------------------------------------
        Records = 0
        For Index = 1 To UBound(Lines)
            If Len(Lines(Index)) > 0 Then
                Fields = Split(Lines(Index), ",")
                If UBound(Fields) < Target Then
                    Err.Raise ERR_EXAMPLE_FIXTURE, "TEST_Examples.SumCsvColumn", _
                        Path & " line " & CStr(Index + 1) & " has no " & Column & "."
                End If
                Total = Total + Val(Fields(Target))
                Records = Records + 1
            End If
        Next Index
        SumCsvColumn = Total

End Function


Private Sub RunExampleOperation( _
    ByVal FailAt As String, _
    ByVal FailRestoreCalculation As Boolean, _
    ByVal FailRestoreScreen As Boolean, _
    ByRef CleanupDetails As String)
'
'==============================================================================
'                             RunExampleOperation
'------------------------------------------------------------------------------
' PURPOSE
'   Reference implementation of the cleanup pattern of
'   docs/VBA_HOUSE_STYLE.md (Operation and cleanup outcomes): capture each
'   setting, register its restoration before changing it, restore only what
'   was registered on both paths, preserve the primary error and report
'   every cleanup failure.
'
' INPUTS
'   FailAt: test seam; empty for none, or where to raise ERR_EXAMPLE_HOST
'   or ERR_EXAMPLE_OPERATION: "capture-screen" (after Calculation was
'   captured), "change-calculation" (after its restoration was registered)
'   or "work".
'   FailRestoreCalculation, FailRestoreScreen: test seams; the restore
'   helper restores its setting and then reports a simulated failure, so
'   that the failure paths can be tested without leaving Excel changed.
'
' RETURNS
'   CleanupDetails: the cleanup failures joined by "; ", empty on success;
'   set also when an error is raised.
'
' STATE OWNERSHIP
'   Changes Application.ScreenUpdating to False and Application.Calculation
'   to manual, and restores each one it registered to the value it captured.
'
' ERROR POLICY
'   Re-raises the primary error after cleanup; when only the cleanup failed,
'   raises ERR_CLEANUP_FAILED with the details.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SavedCalculation       As XlCalculation    'Calculation mode before the operation
    Dim SavedScreen            As Boolean          'ScreenUpdating before the operation
    Dim RestoreCalculationDue  As Boolean          'Calculation captured and about to change
    Dim RestoreScreenDue       As Boolean          'ScreenUpdating captured and about to change
    Dim ErrNumber              As Long             'Primary error number; 0 for none
    Dim ErrSource              As String           'Primary error source
    Dim ErrDescription         As String           'Primary error description

'------------------------------------------------------------------------------
' CAPTURE, REGISTER AND CHANGE
'------------------------------------------------------------------------------
    'A restoration is registered only for a captured setting, just before
    'its change, so a change that fails part-way is still restored and an
    'unchanged setting is never touched.
        CleanupDetails = vbNullString
        On Error GoTo Failed
        SavedCalculation = Application.Calculation
        If FailAt = "capture-screen" Then
            Err.Raise ERR_EXAMPLE_HOST, "TEST_Examples.RunExampleOperation", _
                "Simulated failure capturing ScreenUpdating."
        End If
        SavedScreen = Application.ScreenUpdating
        RestoreScreenDue = True
        Application.ScreenUpdating = False
        RestoreCalculationDue = True
        If FailAt = "change-calculation" Then
            Err.Raise ERR_EXAMPLE_HOST, "TEST_Examples.RunExampleOperation", _
                "Simulated failure changing Calculation."
        End If
        Application.Calculation = xlCalculationManual

'------------------------------------------------------------------------------
' WORK
'------------------------------------------------------------------------------
        If FailAt = "work" Then
            Err.Raise ERR_EXAMPLE_OPERATION, "TEST_Examples.RunExampleOperation", _
                "Simulated failure of the operation."
        End If

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
CleanUp:
        On Error GoTo 0
        If RestoreCalculationDue Then
            CleanupDetails = AppendDetail(CleanupDetails, _
                RestoreCalculation(SavedCalculation, FailRestoreCalculation))
        End If
        If RestoreScreenDue Then
            CleanupDetails = AppendDetail(CleanupDetails, _
                RestoreScreenUpdating(SavedScreen, FailRestoreScreen))
        End If
        If ErrNumber <> 0 Then
            Err.Raise ErrNumber, ErrSource, ErrDescription
        End If
        If Len(CleanupDetails) > 0 Then
            Err.Raise ERR_CLEANUP_FAILED, "TEST_Examples.RunExampleOperation", CleanupDetails
        End If
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE THE PRIMARY ERROR
'------------------------------------------------------------------------------
Failed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        Resume CleanUp

End Sub


Private Function RestoreCalculation( _
    ByVal SavedCalculation As XlCalculation, _
    ByVal SimulateFailure As Boolean) _
    As String
'
'==============================================================================
'                              RestoreCalculation
'------------------------------------------------------------------------------
' PURPOSE
'   Restore Application.Calculation; contain any error.
'
' RETURNS
'   Empty on success, otherwise what failed.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        On Error GoTo Failed
        Application.Calculation = SavedCalculation
        If SimulateFailure Then
            Err.Raise ERR_EXAMPLE_RESTORE, "TEST_Examples.RestoreCalculation", "simulated"
        End If
        Exit Function

Failed:
        RestoreCalculation = "Calculation not restored: " & Err.Description

End Function


Private Function RestoreScreenUpdating( _
    ByVal SavedScreen As Boolean, _
    ByVal SimulateFailure As Boolean) _
    As String
'
'==============================================================================
'                            RestoreScreenUpdating
'------------------------------------------------------------------------------
' PURPOSE
'   Restore Application.ScreenUpdating; contain any error.
'
' RETURNS
'   Empty on success, otherwise what failed.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        On Error GoTo Failed
        Application.ScreenUpdating = SavedScreen
        If SimulateFailure Then
            Err.Raise ERR_EXAMPLE_RESTORE, "TEST_Examples.RestoreScreenUpdating", "simulated"
        End If
        Exit Function

Failed:
        RestoreScreenUpdating = "ScreenUpdating not restored: " & Err.Description

End Function


Private Function AppendDetail( _
    ByVal First As String, _
    ByVal Second As String) _
    As String
'
'==============================================================================
'                                 AppendDetail
'------------------------------------------------------------------------------
' PURPOSE
'   Join two failure texts with "; ", skipping empty ones.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        If Len(First) = 0 Then
            AppendDetail = Second
        ElseIf Len(Second) = 0 Then
            AppendDetail = First
        Else
            AppendDetail = First & "; " & Second
        End If

End Function
