Attribute VB_Name = "TEST_Harness"
'==============================================================================
' MODULE: TEST_Harness
'------------------------------------------------------------------------------
' PURPOSE
'   Deterministic regression harness for the workbook's VBA test suites:
'   one run per suite, stable case identifiers, exact and tolerance
'   assertions, expected error values and raised errors, expected case and
'   assertion counts, a complete key=value log, and the certification of a
'   suite's runs as the evidence of an Excel run. It holds no calculation
'   of the workbook.
'
' PUBLIC SURFACE
'   Suite lifecycle: BeginRun, BeginCase, RecordCaseError, EndRun.
'   Assertions: AssertEqualDouble, AssertNear, AssertEqualText, AssertTrue,
'   AssertErrorValue, AssertRaised; WithinTolerance is the comparison rule.
'   Log: Say, Echo, SaveLog, LastResult. Results: RecordResult, SaveResults.
'   Recovery: ResetTests. Certification: BeginCertification, CertifyRun,
'   EndCertification. Option Private Module keeps all of them out of the
'   workbook's external surface.
'
' DEPENDENCIES
'   The VBA and Excel object models only: no reference, workbook fixture,
'   worksheet or external framework.
'
' STATE OWNERSHIP
'   Owns the run counters, the case identifiers, the failure records, the
'   retained log, the recorded results, the certification records and the
'   active-run, recovery-required and active-certification flags. Reads four
'   Application properties, the number of open workbooks and the caption of
'   the active window to prove a run left them unchanged; never changes
'   them. Reads the build stamp's commit from ThisWorkbook for SaveResults;
'   never writes a workbook or worksheet.
'
' ERROR POLICY
'   Assertion failures are recorded and the run continues. EndRun re-raises
'   an unexpected runner error unchanged after cleanup and reporting, or
'   raises ERR_TEST_FAILED for any other failed outcome. A run that starts
'   while another is active, or while recovery is required, is refused with
'   ERR_TEST_DIRTY_START. Misuse of the lifecycle raises ERR_TEST_USAGE.
'   EndCertification writes the certification log first, then raises
'   ERR_TEST_FAILED when the certification did not pass.
'
' LOG FORMAT
'   One record per line, KEY=value; fields separated by "; ". The keys and
'   their order are documented in docs/methodology/TEST_HARNESS.md (Log
'   format, Certification log). SaveResults writes the JSON results file of
'   docs/methodology/TEST_CASES.md (Accuracy contracts).
'
' COMPATIBILITY
'   Excel VBA on Windows. The log records the host, its version, the Office
'   bitness and the VBA generation.
'
' USAGE
'   Called by the suites in tests/modules/, for example
'   TEST_Examples.RunExampleTests and TEST_Examples.CertifyExamples; SaveLog
'   and ResetTests are also typed in the Immediate window. It has no suite
'   of its own.
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
    'Require explicit declarations; keep the harness inside the project.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Test-only error numbers, kept apart from the production range in use.
    'Written as literals equal to vbObjectError + n: LibreOffice, the optional
    'smoke host, cannot evaluate vbObjectError in a constant expression.
        Public Const ERR_TEST_DIRTY_START   As Long = &H80040FE0   'vbObjectError + 4064: Run refused: active or recovery required
        Public Const ERR_TEST_FAILED        As Long = &H80040FE1   'vbObjectError + 4065: Run finished without PASS
        Public Const ERR_TEST_USAGE         As Long = &H80040FE2   'vbObjectError + 4066: Lifecycle used out of order
    'The defined name that carries the build stamp's commit
    '(tools/devbuild/workbook.json, stamp.commit).
        Private Const STAMP_COMMIT_NAME     As String = "build_commit"

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'One run at a time. Counters and the log stay readable after EndRun,
    'until the next BeginRun or ResetTests.
        Private mRunActive            As Boolean        'True from BeginRun to EndRun
        Private mRecoveryRequired     As Boolean        'True until ResetTests after an unsafe end
        Private mRecoveryReason       As String         'Why recovery is required
        Private mSuite                As String         'Name of the running or last suite
        Private mExpectedCases        As Long           'Cases a complete run starts
        Private mExpectedAssertions   As Long           'Assertions a complete run evaluates
        Private mCaseCount            As Long           'Cases started in this run
        Private mAssertionCount       As Long           'Assertions evaluated in this run
        Private mFailureCount         As Long           'Failures of any kind in this run
        Private mErrorCount           As Long           'Unexpected errors in this run
        Private mCurrentCase          As String         'Identifier of the running case
        Private mCaseIds              As Collection     'Case identifiers seen in this run
        Private mPrimaryFailure       As String         'First failure recorded in this run
        Private mLastResult           As String         'RESULT record of the last finished run
        Private mLog                  As Collection     'Log since the last BeginRun or ResetTests
        Private mResults              As Collection     'Results recorded since the last BeginRun or ResetTests
        Private mSnapshotValid        As Boolean        'True once all four properties were read
        Private mInitialCalculation   As XlCalculation  'Calculation at BeginRun
        Private mInitialAlerts        As Boolean        'DisplayAlerts at BeginRun
        Private mInitialEvents        As Boolean        'EnableEvents at BeginRun
        Private mInitialScreen        As Boolean        'ScreenUpdating at BeginRun
        Private mInitialWorkbooks     As Long           'Open workbooks at BeginRun
        Private mInitialWindow        As String         'Active window caption at BeginRun
        Private mCertActive           As Boolean        'True from BeginCertification to EndCertification
        Private mCertSuite            As String         'Suite being certified
        Private mCertLog              As Collection     'Certification records
        Private mCertModes            As Long           'Runs certified
        Private mCertFailures         As Long           'Runs that did not end as required
        Private mCertMandatory        As Long           'Runs required to pass
        Private mCertComplete         As Long           'Required runs that were complete


'
'------------------------------------------------------------------------------
'
'                               SUITE LIFECYCLE
'
'------------------------------------------------------------------------------
'

Public Sub BeginRun( _
    ByVal SuiteName As String, _
    ByVal Mode As String, _
    ByVal ExpectedCases As Long, _
    ByVal ExpectedAssertions As Long)
'
'==============================================================================
'                                   BeginRun
'------------------------------------------------------------------------------
' PURPOSE
'   Start one suite run: refuse an unsafe start, reset the run state, record
'   the host state to compare at the end and write the log header.
'
' INPUTS
'   SuiteName: stable suite identifier, written in the RUN and RESULT records.
'   Mode: NORMAL, or the name of a deliberate failure mode of the suite.
'   ExpectedCases, ExpectedAssertions: counts of a complete run, >= 1.
'
' ERROR POLICY
'   Raises ERR_TEST_DIRTY_START, after logging why, when a run is active or
'   recovery is required; raises ERR_TEST_USAGE for counts below 1. A host
'   property that cannot be read leaves the snapshot invalid, which fails
'   cleanup at EndRun; it does not stop the run.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' REFUSE AN UNSAFE START
'------------------------------------------------------------------------------
    'An interrupted run or an unverified cleanup must be acknowledged with
    'ResetTests before a new run may overwrite its state. A project reset
    '(End or Reset in the VBE) clears this state, so such a stop is not seen.
        If mRunActive Then
            mRecoveryRequired = True
            mRecoveryReason = "run of " & mSuite & " still active (interrupted?)"
        End If
        If mRecoveryRequired Then
            mLastResult = "RESULT=FAIL; suite=" & SuiteName & "; operation=NOT_RUN" & _
                "; completeness=INCOMPLETE; cases=0; assertions=0; failures=0" & _
                "; cleanup=NOT_RUN; recovery=REQUIRED"
            Say mLastResult
            Say "RECOVERY=REQUIRED; reason=" & mRecoveryReason & _
                "; action=check Excel state, then run TEST_Harness.ResetTests"
            Err.Raise ERR_TEST_DIRTY_START, "TEST_Harness.BeginRun", _
                "Recovery required: " & mRecoveryReason & ". Run TEST_Harness.ResetTests."
        End If
        If ExpectedCases < 1 Or ExpectedAssertions < 1 Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness.BeginRun", _
                "Expected counts must be at least 1; got cases=" & CStr(ExpectedCases) & _
                ", assertions=" & CStr(ExpectedAssertions) & "."
        End If

'------------------------------------------------------------------------------
' RESET AND CLAIM THE RUN
'------------------------------------------------------------------------------
        ResetRunState
        Set mLog = New Collection
        Set mResults = New Collection
        mSuite = SuiteName
        mExpectedCases = ExpectedCases
        mExpectedAssertions = ExpectedAssertions
        mRunActive = True

'------------------------------------------------------------------------------
' SNAPSHOT HOST STATE
'------------------------------------------------------------------------------
    'Mark the snapshot valid only after every read, so that cleanup never
    'compares a value that was not read.
        On Error GoTo SnapshotFailed
        mInitialCalculation = Application.Calculation
        mInitialAlerts = Application.DisplayAlerts
        mInitialEvents = Application.EnableEvents
        mInitialScreen = Application.ScreenUpdating
        mInitialWorkbooks = Application.Workbooks.Count
        mInitialWindow = ActiveWindowCaption()
        mSnapshotValid = True

WriteHeader:
        On Error GoTo 0
        Say "RUN=" & SuiteName & "; mode=" & Mode & _
            "; expected_cases=" & CStr(ExpectedCases) & _
            "; expected_assertions=" & CStr(ExpectedAssertions) & _
            "; started=" & Format$(Now, "yyyy-mm-dd hh:mm:ss")
        Say "ENVIRONMENT=" & EnvironmentSummary()
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE A FAILED SNAPSHOT
'------------------------------------------------------------------------------
SnapshotFailed:
        mSnapshotValid = False
        Resume WriteHeader

End Sub


Public Sub BeginCase( _
    ByVal CaseId As String)
'
'==============================================================================
'                                  BeginCase
'------------------------------------------------------------------------------
' PURPOSE
'   Count and log the start of one case.
'
' INPUTS
'   CaseId: lower-case kebab-case identifier, unique within the run and
'   stable across runs, so that logs can be compared case by case.
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. A malformed or repeated identifier
'   is recorded as a failure; the case still counts and runs.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' REGISTER THE CASE
'------------------------------------------------------------------------------
        RequireActiveRun "BeginCase"
        mCaseCount = mCaseCount + 1
        mCurrentCase = CaseId
        Say "CASE=" & CaseId
        If Not IsCaseId(CaseId) Then
            RecordFailure "harness.case-id", "malformed case identifier '" & CaseId & "'"
        ElseIf IsKnownCase(CaseId) Then
            RecordFailure "harness.case-id", "repeated case identifier '" & CaseId & "'"
        Else
            mCaseIds.Add CaseId, CaseId
        End If

End Sub


Public Sub RecordCaseError()
'
'==============================================================================
'                               RecordCaseError
'------------------------------------------------------------------------------
' PURPOSE
'   Record the error that ended the running case, so that the remaining
'   cases still run. Call it first thing in the case's error handler.
'
' INPUTS
'   The Err object, read before anything else can reset it.
'
' SIDE EFFECTS
'   Counts an unexpected error and a failure; evaluates no assertion.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrorNumber        As Long      'Number of the case error
    Dim ErrorSource        As String    'Source of the case error
    Dim ErrorDescription   As String    'Description of the case error

'------------------------------------------------------------------------------
' RECORD
'------------------------------------------------------------------------------
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
        RequireActiveRun "RecordCaseError"
        mErrorCount = mErrorCount + 1
        Say "ERROR=" & mCurrentCase & "; number=" & CStr(ErrorNumber) & _
            "; source=" & ErrorSource & "; description=" & ErrorDescription
        RecordFailure mCurrentCase & ".unexpected-error", _
            "number=" & CStr(ErrorNumber) & "; source=" & ErrorSource

End Sub


Public Sub EndRun( _
    ByVal RunErrorNumber As Long, _
    ByVal RunErrorSource As String, _
    ByVal RunErrorDescription As String)
'
'==============================================================================
'                                    EndRun
'------------------------------------------------------------------------------
' PURPOSE
'   Finish the run: check completeness, release the run, verify the host
'   state, write the summary and raise the outcome.
'
' INPUTS
'   RunErrorNumber, RunErrorSource, RunErrorDescription: the error that
'   escaped the suite's runner, or 0 and empty strings when none did.
'
' OUTCOMES
'   operation: what ran: PASS, FAIL (an assertion or a case identifier
'   failed) or ERROR (an unexpected error occurred). completeness: COMPLETE
'   only when both expected counts were met exactly. cleanup: PASS when the host state is unchanged.
'   recovery: REQUIRED when the host state changed or could not be checked,
'   or when a start was refused while this run was active.
'   RESULT=PASS needs operation PASS, COMPLETE, cleanup PASS and no recovery.
'
' ERROR POLICY
'   Re-raises the runner error unchanged, after reporting, so that cleanup
'   never replaces the primary failure; otherwise raises ERR_TEST_FAILED
'   unless RESULT=PASS. Raises ERR_TEST_USAGE outside a run.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CleanupDetail   As String     'Reason for the cleanup outcome
    Dim CleanupPassed   As Boolean    'True when the host state is unchanged
    Dim Completed       As Boolean    'True when both expected counts were met
    Dim Operation       As String     'PASS, FAIL or ERROR
    Dim Verdict         As String     'PASS or FAIL

'------------------------------------------------------------------------------
' COLLECT THE OUTCOMES
'------------------------------------------------------------------------------
        RequireActiveRun "EndRun"
        mCurrentCase = "run"
        If RunErrorNumber <> 0 Then
            mErrorCount = mErrorCount + 1
            Say "ERROR=run; number=" & CStr(RunErrorNumber) & "; source=" & RunErrorSource & _
                "; description=" & RunErrorDescription
            RecordFailure "runner.unexpected-error", _
                "number=" & CStr(RunErrorNumber) & "; source=" & RunErrorSource
        End If

    'The operation outcome covers what ran; completeness and cleanup are
    'judged separately below, so that each is reported on its own.
        If mErrorCount > 0 Then
            Operation = "ERROR"
        ElseIf mFailureCount > 0 Then
            Operation = "FAIL"
        Else
            Operation = "PASS"
        End If

    'Exact counts: a skipped case or assertion, or an extra one, means the
    'run did not execute the suite as written.
        Completed = (mCaseCount = mExpectedCases) And _
                    (mAssertionCount = mExpectedAssertions)
        If Not Completed Then
            RecordFailure "harness.completeness", _
                "expected cases=" & CStr(mExpectedCases) & _
                ", assertions=" & CStr(mExpectedAssertions) & _
                "; observed cases=" & CStr(mCaseCount) & _
                ", assertions=" & CStr(mAssertionCount)
        End If

        CleanupPassed = VerifyHostState(CleanupDetail)
        mRunActive = False
        If Not CleanupPassed Then
            mRecoveryRequired = True
            mRecoveryReason = CleanupDetail
        End If

        If Operation = "PASS" And Completed And CleanupPassed And Not mRecoveryRequired Then
            Verdict = "PASS"
        Else
            Verdict = "FAIL"
        End If

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
        Say "CASES=" & CStr(mCaseCount) & "; expected=" & CStr(mExpectedCases)
        Say "ASSERTIONS=" & CStr(mAssertionCount) & "; expected=" & CStr(mExpectedAssertions)
        Say "FAILURES=" & CStr(mFailureCount) & "; errors=" & CStr(mErrorCount)
        If Len(mPrimaryFailure) > 0 Then
            Say "PRIMARY_FAILURE=" & mPrimaryFailure
        End If
        Say "CLEANUP=" & IIf(CleanupPassed, "PASS", "FAIL") & "; detail=" & CleanupDetail
        If mRecoveryRequired Then
            Say "RECOVERY=REQUIRED; reason=" & mRecoveryReason & _
                "; action=check Excel state, then run TEST_Harness.ResetTests"
        Else
            Say "RECOVERY=NONE"
        End If
        mLastResult = "RESULT=" & Verdict & "; suite=" & mSuite & _
            "; operation=" & Operation & _
            "; completeness=" & IIf(Completed, "COMPLETE", "INCOMPLETE") & _
            "; cases=" & CStr(mCaseCount) & _
            "; assertions=" & CStr(mAssertionCount) & _
            "; failures=" & CStr(mFailureCount) & _
            "; cleanup=" & IIf(CleanupPassed, "PASS", "FAIL") & _
            "; recovery=" & IIf(mRecoveryRequired, "REQUIRED", "NONE")
        Say mLastResult

'------------------------------------------------------------------------------
' RAISE THE OUTCOME
'------------------------------------------------------------------------------
        If RunErrorNumber <> 0 Then
            Err.Raise RunErrorNumber, RunErrorSource, RunErrorDescription
        End If
        If Verdict <> "PASS" Then
            Err.Raise ERR_TEST_FAILED, "TEST_Harness.EndRun", _
                "Suite " & mSuite & " did not pass; see the log."
        End If

End Sub


Public Sub ResetTests()
'
'==============================================================================
'                                  ResetTests
'------------------------------------------------------------------------------
' PURPOSE
'   Acknowledge an interrupted run or an unverified cleanup, so that the
'   next run may start.
'
' USAGE
'   Run only after the interrupted run has stopped and the Excel state has
'   been checked by hand; it restores nothing itself.
'
' SIDE EFFECTS
'   Clears the run state, both flags and an unfinished certification; logs
'   the reset.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        ResetRunState
        mRecoveryRequired = False
        mRecoveryReason = vbNullString
        mCertActive = False
        Set mCertLog = Nothing
        Set mLog = New Collection
        Set mResults = New Collection
        Say "RESET=DONE; run_active=no; recovery=NONE"

End Sub


'
'------------------------------------------------------------------------------
'
'                                  ASSERTIONS
'
'------------------------------------------------------------------------------
'

Public Sub AssertEqualDouble( _
    ByVal AssertionName As String, _
    ByVal Expected As Double, _
    ByVal Actual As Double)
'
'==============================================================================
'                              AssertEqualDouble
'------------------------------------------------------------------------------
' PURPOSE
'   Require exact Double equality, for results that are exact in binary.
'
' INPUTS
'   AssertionName: stable name within the case. Expected, Actual: compared.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        LogAssertion AssertionName, (Actual = Expected), _
            "expected=" & NumberText(Expected) & "; actual=" & NumberText(Actual)

End Sub


Public Sub AssertNear( _
    ByVal AssertionName As String, _
    ByVal Expected As Double, _
    ByVal Actual As Double, _
    ByVal AbsoluteTolerance As Double, _
    ByVal RelativeTolerance As Double)
'
'==============================================================================
'                                  AssertNear
'------------------------------------------------------------------------------
' PURPOSE
'   Require |Actual - Expected| <= max(AbsoluteTolerance, RelativeTolerance
'   * |Expected|), the rule of docs/methodology/TEST_CASES.md (Tolerance).
'
' INPUTS
'   AbsoluteTolerance, RelativeTolerance: >= 0, both always stated. Near
'   zero the absolute bound governs; both 0 demands exact equality.
'
' ERROR POLICY
'   A negative tolerance fails the assertion; nothing is raised.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        If AbsoluteTolerance < 0# Or RelativeTolerance < 0# Then
            LogAssertion AssertionName, False, "tolerances must not be negative"
        Else
            LogAssertion AssertionName, _
                WithinTolerance(Expected, Actual, AbsoluteTolerance, RelativeTolerance), _
                "expected=" & NumberText(Expected) & "; actual=" & NumberText(Actual) & _
                "; absolute=" & NumberText(AbsoluteTolerance) & _
                "; relative=" & NumberText(RelativeTolerance)
        End If

End Sub


Public Sub AssertEqualText( _
    ByVal AssertionName As String, _
    ByVal Expected As String, _
    ByVal Actual As String)
'
'==============================================================================
'                               AssertEqualText
'------------------------------------------------------------------------------
' PURPOSE
'   Require the same text, compared binary: case and spaces count.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        LogAssertion AssertionName, (StrComp(Actual, Expected, vbBinaryCompare) = 0), _
            "expected='" & Expected & "'; actual='" & Actual & "'"

End Sub


Public Sub AssertTrue( _
    ByVal AssertionName As String, _
    ByVal Condition As Boolean, _
    ByVal Detail As String)
'
'==============================================================================
'                                  AssertTrue
'------------------------------------------------------------------------------
' PURPOSE
'   Require a condition the suite evaluated; Detail is logged on failure.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        LogAssertion AssertionName, Condition, Detail

End Sub


Public Sub AssertErrorValue( _
    ByVal AssertionName As String, _
    ByVal ExpectedCode As Long, _
    ByVal Actual As Variant)
'
'==============================================================================
'                               AssertErrorValue
'------------------------------------------------------------------------------
' PURPOSE
'   Require a specific worksheet error value, such as CVErr(xlErrValue).
'
' INPUTS
'   ExpectedCode: an xlErr constant. Actual: the value returned.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
    'Check the type first, then compare the error values as text
    '("Error 2007"): comparing a value with an error value by "=" raises
    'error 13, Type mismatch.
        If Not IsError(Actual) Then
            LogAssertion AssertionName, False, _
                "expected error " & CStr(ExpectedCode) & "; actual type=" & TypeName(Actual)
        Else
            LogAssertion AssertionName, (CStr(Actual) = CStr(CVErr(ExpectedCode))), _
                "expected error " & CStr(ExpectedCode) & "; actual=" & CStr(Actual)
        End If

End Sub


Public Sub AssertRaised( _
    ByVal AssertionName As String, _
    ByVal ExpectedNumber As Long, _
    ByVal ActualNumber As Long)
'
'==============================================================================
'                                 AssertRaised
'------------------------------------------------------------------------------
' PURPOSE
'   Require that a call raised a specific error number. The suite captures
'   Err.Number on the line after the call; 0 means nothing was raised.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        LogAssertion AssertionName, (ActualNumber = ExpectedNumber), _
            "expected error=" & CStr(ExpectedNumber) & "; actual error=" & CStr(ActualNumber)

End Sub


Public Function WithinTolerance( _
    ByVal Expected As Double, _
    ByVal Actual As Double, _
    ByVal AbsoluteTolerance As Double, _
    ByVal RelativeTolerance As Double) _
    As Boolean
'
'==============================================================================
'                               WithinTolerance
'------------------------------------------------------------------------------
' PURPOSE
'   The comparison rule of AssertNear, exposed so that suites can test it.
'
' RETURNS
'   True when |Actual - Expected| <= max(AbsoluteTolerance,
'   RelativeTolerance * |Expected|); the bound is inclusive.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Bound   As Double    'Larger of the two bounds

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        Bound = RelativeTolerance * Abs(Expected)
        If AbsoluteTolerance > Bound Then
            Bound = AbsoluteTolerance
        End If
        WithinTolerance = (Abs(Actual - Expected) <= Bound)

End Function


'
'------------------------------------------------------------------------------
'
'                                CERTIFICATION
'
'------------------------------------------------------------------------------
'

Public Sub BeginCertification( _
    ByVal SuiteName As String)
'
'==============================================================================
'                              BeginCertification
'------------------------------------------------------------------------------
' PURPOSE
'   Start the certification of one suite: refuse state left by an earlier
'   run, then keep a certification log apart from the run logs.
'
' INPUTS
'   SuiteName: the suite identifier, as in its RUN records.
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE while another certification is active, and
'   ERR_TEST_DIRTY_START when a run is active or recovery is required; it
'   changes no run state before either check.
'
' UPDATED
'   2026-11-02
'==============================================================================
'

'------------------------------------------------------------------------------
' REFUSE LEFTOVER STATE
'------------------------------------------------------------------------------
        If mCertActive Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness.BeginCertification", _
                "The certification of " & mCertSuite & " is still active. " & _
                "Run TEST_Harness.ResetTests."
        End If
        If mRunActive Then
            mRecoveryRequired = True
            mRecoveryReason = "run of " & mSuite & " still active (interrupted?)"
        End If
        If mRecoveryRequired Then
            Err.Raise ERR_TEST_DIRTY_START, "TEST_Harness.BeginCertification", _
                "Recovery required: " & mRecoveryReason & ". Run TEST_Harness.ResetTests."
        End If

'------------------------------------------------------------------------------
' CLAIM THE CERTIFICATION
'------------------------------------------------------------------------------
        mCertActive = True
        mCertSuite = SuiteName
        mCertModes = 0
        mCertFailures = 0
        mCertMandatory = 0
        mCertComplete = 0
        mLastResult = vbNullString
        Set mCertLog = New Collection
        CertSay "CERTIFY=" & SuiteName & "; started=" & Format$(Now, "yyyy-mm-dd hh:mm:ss")

End Sub


Public Sub CertifyRun( _
    ByVal Mode As String, _
    ByVal MustPass As Boolean, _
    ByVal RaisedNumber As Long, _
    ByVal RaisedDescription As String)
'
'==============================================================================
'                                  CertifyRun
'------------------------------------------------------------------------------
' PURPOSE
'   Record how one run of the suite reached its caller, and whether that is
'   how the mode must end.
'
' INPUTS
'   Mode: the run's mode. MustPass: True for a run that must pass, False
'   for a deliberate failure mode. RaisedNumber, RaisedDescription: the
'   error that reached the caller of the suite's entry point, or 0 and an
'   empty string when none did.
'
' OUTCOMES
'   A run that must pass ends with no error and a RESULT=PASS that is
'   COMPLETE. A deliberate failure ends with RESULT=FAIL and reaches its
'   caller as ERR_TEST_FAILED with a description: the self-test that a
'   failure is never silent. A run with no RESULT of its own fails.
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a certification. A run that is still
'   active did not end: it requires recovery and is recorded as failed, so
'   that the certification still writes its log.
'
' UPDATED
'   2026-11-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Completeness   As String     'COMPLETE, INCOMPLETE or NONE
    Dim Passed         As Boolean    'True when the run ended as required
    Dim Verdict        As String     'PASS, FAIL or NONE

'------------------------------------------------------------------------------
' READ THE RUN'S RESULT
'------------------------------------------------------------------------------
        RequireCertification "CertifyRun"
        If mRunActive Then
            mRunActive = False
            mRecoveryRequired = True
            mRecoveryReason = "run of " & mSuite & " did not end"
            mLastResult = vbNullString
        End If
        If Len(mLastResult) = 0 Then
            Verdict = "NONE"
            Completeness = "NONE"
        Else
            Verdict = IIf(Left$(mLastResult, 12) = "RESULT=PASS;", "PASS", "FAIL")
            Completeness = IIf(InStr(mLastResult, "; completeness=COMPLETE;") > 0, _
                               "COMPLETE", "INCOMPLETE")
        End If
        mLastResult = vbNullString

'------------------------------------------------------------------------------
' JUDGE AND RECORD
'------------------------------------------------------------------------------
        If MustPass Then
            Passed = (RaisedNumber = 0) And (Verdict = "PASS") And (Completeness = "COMPLETE")
            mCertMandatory = mCertMandatory + 1
            If Completeness = "COMPLETE" Then
                mCertComplete = mCertComplete + 1
            End If
        Else
            Passed = (RaisedNumber = ERR_TEST_FAILED) And (Len(RaisedDescription) > 0) _
                     And (Verdict = "FAIL")
        End If
        mCertModes = mCertModes + 1
        If Not Passed Then
            mCertFailures = mCertFailures + 1
        End If
        CertSay "MODE=" & Mode & "; expected=" & IIf(MustPass, "PASS", "FAIL") & _
            "; error=" & CStr(RaisedNumber) & "; result=" & Verdict & _
            "; completeness=" & Completeness & "; status=" & IIf(Passed, "PASS", "FAIL")

End Sub


Public Sub EndCertification( _
    ByVal Path As String)
'
'==============================================================================
'                               EndCertification
'------------------------------------------------------------------------------
' PURPOSE
'   Finish the certification: write its verdict, save the certification log
'   and only then raise the outcome.
'
' INPUTS
'   Path: full path of the certification log, created or overwritten.
'
' OUTCOMES
'   completeness: COMPLETE when every run that must pass was complete.
'   CERTIFICATION=PASS needs at least one such run, COMPLETE, and every run
'   ending as its mode requires.
'
' ERROR POLICY
'   A file error is re-raised after the file is closed. Otherwise raises
'   ERR_TEST_FAILED, after the log is written, unless the certification
'   passed. Raises ERR_TEST_USAGE outside a certification.
'
' UPDATED
'   2026-11-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Completed   As Boolean    'True when every required run was complete
    Dim Verdict     As String     'PASS or FAIL

'------------------------------------------------------------------------------
' JUDGE
'------------------------------------------------------------------------------
        RequireCertification "EndCertification"
        Completed = (mCertMandatory > 0) And (mCertComplete = mCertMandatory)
        If Completed And mCertFailures = 0 Then
            Verdict = "PASS"
        Else
            Verdict = "FAIL"
        End If
        CertSay "CERTIFICATION=" & Verdict & "; suite=" & mCertSuite & _
            "; completeness=" & IIf(Completed, "COMPLETE", "INCOMPLETE") & _
            "; modes=" & CStr(mCertModes) & "; failures=" & CStr(mCertFailures)
        mCertActive = False

'------------------------------------------------------------------------------
' WRITE THE EVIDENCE, THEN RAISE
'------------------------------------------------------------------------------
        WriteRecords Path, mCertLog
        If Verdict <> "PASS" Then
            Err.Raise ERR_TEST_FAILED, "TEST_Harness.EndCertification", _
                "Certification of " & mCertSuite & " did not pass; see " & Path & "."
        End If

End Sub


'
'------------------------------------------------------------------------------
'
'                                     LOG
'
'------------------------------------------------------------------------------
'

Public Sub Say( _
    ByVal Text As String)
'
'==============================================================================
'                                     Say
'------------------------------------------------------------------------------
' PURPOSE
'   Write one log record to the Immediate window and keep it for SaveLog.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        If mLog Is Nothing Then
            Set mLog = New Collection
        End If
        mLog.Add Text
        Echo Text

End Sub


Public Sub Echo( _
    ByVal Text As String)
'
'==============================================================================
'                                     Echo
'------------------------------------------------------------------------------
' PURPOSE
'   Print one line to the Immediate window, without keeping it in the log.
'
' ERROR POLICY
'   A host without an Immediate window, such as LibreOffice run headless,
'   raises on Debug.Print; that failure is ignored, because the kept log,
'   not the window, is the record of a run.
'
' UPDATED
'   2026-11-02
'==============================================================================
'
        On Error Resume Next
        Debug.Print Text
        On Error GoTo 0

End Sub


Public Sub SaveLog( _
    ByVal Path As String)
'
'==============================================================================
'                                   SaveLog
'------------------------------------------------------------------------------
' PURPOSE
'   Write every record kept since the last BeginRun or ResetTests to a text
'   file. The Immediate window keeps only its last 200 lines; the file is
'   the complete log.
'
' INPUTS
'   Path: full path of the file, created or overwritten; outside the
'   checkout or in the ignored test-results/ directory.
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE when no record is kept; a file error is re-raised
'   after the file is closed.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

        If mLog Is Nothing Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness.SaveLog", "No log record is kept."
        End If
        WriteRecords Path, mLog

End Sub


Public Function LastResult() _
    As String
'
'==============================================================================
'                                  LastResult
'------------------------------------------------------------------------------
' PURPOSE
'   The RESULT record of the last finished or refused run; empty before
'   the first, and after CertifyRun has read it.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        LastResult = mLastResult

End Function


'
'------------------------------------------------------------------------------
'
'                                   RESULTS
'
'------------------------------------------------------------------------------
'

Public Sub RecordResult( _
    ByVal CaseId As String, _
    ByVal VariantName As String, _
    ByVal Quantity As String, _
    ByVal Dataset As String, _
    ByVal Record As String, _
    ByVal Value As Double)
'
'==============================================================================
'                                 RecordResult
'------------------------------------------------------------------------------
' PURPOSE
'   Keep one computed value of the running suite for SaveResults: the
'   result of one expected output of a test case, which
'   tools/check_accuracy.py --results compares with its expected value.
'
' INPUTS
'   CaseId: the test case's id (tests/fixtures/). VariantName: the variant of
'   a file-based case, or an empty string for a JSON fixture. Quantity: the
'   output's quantity. Dataset, Record: the dataset and record id of a
'   record-scoped output, or two empty strings for a result of the whole
'   calculation. Value: the result, written with the 15 significant digits
'   of Str$.
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. A repeated result is kept; the
'   accuracy check rejects it.
'
' UPDATED
'   2026-11-03
'==============================================================================
'
        RequireActiveRun "RecordResult"
        mResults.Add "{""case"": " & JsonText(CaseId) & _
            ", ""variant"": " & JsonTextOrNull(VariantName) & _
            ", ""quantity"": " & JsonText(Quantity) & _
            ", ""scope"": " & ScopeJson(Dataset, Record) & _
            ", ""value"": " & NumberText(Value) & "}"

End Sub


Public Sub SaveResults( _
    ByVal Path As String)
'
'==============================================================================
'                                 SaveResults
'------------------------------------------------------------------------------
' PURPOSE
'   Write the results recorded since the last BeginRun or ResetTests as the
'   results file of docs/methodology/TEST_CASES.md (Accuracy contracts), with
'   the commit of the workbook's build stamp. The file is ASCII, a subset of
'   UTF-8, with LF line endings.
'
' INPUTS
'   Path: full path of the file, created or overwritten; outside the
'   checkout or in the ignored test-results/ directory.
'
' OUTCOMES
'   A workbook without a readable build stamp writes an empty commit, and a
'   build with local changes writes its marked commit; the accuracy check
'   rejects both, because the results then name no exact commit.
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE before the first run; a file error is re-raised
'   after the file is closed.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Text    As String     'The whole file
    Dim Entry   As Variant    'One recorded result
    Dim Count   As Long       'Results written so far

'------------------------------------------------------------------------------
' COMPOSE
'------------------------------------------------------------------------------
        If mResults Is Nothing Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness.SaveResults", "No run has recorded results."
        End If
        Text = "{" & vbLf & "  ""schema_version"": 1," & vbLf & _
            "  ""kind"": ""test-results""," & vbLf & _
            "  ""commit"": " & JsonText(StampedCommit()) & "," & vbLf & _
            "  ""results"": ["
        For Each Entry In mResults
            Count = Count + 1
            Text = Text & IIf(Count = 1, vbNullString, ",") & vbLf & "    " & CStr(Entry)
        Next Entry
        Text = Text & IIf(Count = 0, vbNullString, vbLf & "  ") & "]" & vbLf & "}" & vbLf

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        WriteText Path, Text

End Sub


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub LogAssertion( _
    ByVal AssertionName As String, _
    ByVal Passed As Boolean, _
    ByVal Detail As String)
'
'==============================================================================
'                                 LogAssertion
'------------------------------------------------------------------------------
' PURPOSE
'   Count one assertion, log its outcome and record it when it failed.
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. An assertion outside a case is
'   counted and recorded as a failure.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        RequireActiveRun "an assertion"
        mAssertionCount = mAssertionCount + 1
        If Len(mCurrentCase) = 0 Then
            RecordFailure "harness.no-case", "assertion '" & AssertionName & "' outside a case"
        End If
        If Passed Then
            Say "ASSERT=" & mCurrentCase & "/" & AssertionName & "; status=PASS"
        Else
            Say "ASSERT=" & mCurrentCase & "/" & AssertionName & "; status=FAIL; " & Detail
            RecordFailure mCurrentCase & "/" & AssertionName, Detail
        End If

End Sub


Private Sub RecordFailure( _
    ByVal FailureId As String, _
    ByVal Detail As String)
'
'==============================================================================
'                                RecordFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Count and log one failure; the first one of a run is its primary failure.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        mFailureCount = mFailureCount + 1
        If Len(mPrimaryFailure) = 0 Then
            mPrimaryFailure = FailureId & "; " & Detail
        End If
        Say "FAILURE=" & FailureId & "; " & Detail

End Sub


Private Sub RequireActiveRun( _
    ByVal Caller As String)
'
'==============================================================================
'                               RequireActiveRun
'------------------------------------------------------------------------------
' PURPOSE
'   Refuse a lifecycle call or an assertion made outside BeginRun..EndRun.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        If Not mRunActive Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness", _
                "No test run is active for " & Caller & "; call BeginRun first."
        End If

End Sub


Private Sub RequireCertification( _
    ByVal Caller As String)
'
'==============================================================================
'                             RequireCertification
'------------------------------------------------------------------------------
' PURPOSE
'   Refuse a certification call made outside
'   BeginCertification..EndCertification.
'
' UPDATED
'   2026-11-02
'==============================================================================
'
        If Not mCertActive Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness", _
                "No certification is active for " & Caller & _
                "; call BeginCertification first."
        End If

End Sub


Private Sub CertSay( _
    ByVal Text As String)
'
'==============================================================================
'                                   CertSay
'------------------------------------------------------------------------------
' PURPOSE
'   Write one certification record to the Immediate window and keep it for
'   EndCertification.
'
' UPDATED
'   2026-11-02
'==============================================================================
'
        mCertLog.Add Text
        Echo Text

End Sub


Private Sub WriteRecords( _
    ByVal Path As String, _
    ByVal Records As Collection)
'
'==============================================================================
'                                 WriteRecords
'------------------------------------------------------------------------------
' PURPOSE
'   Write records to a text file, one per line.
'
' INPUTS
'   Path: full path of the file, created or overwritten. Records: the
'   records, in order.
'
' ERROR POLICY
'   A file error is re-raised after the file is closed.
'
' UPDATED
'   2026-11-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FileNumber         As Integer    'Handle of the open file
    Dim Entry              As Variant    'One record
    Dim ErrorNumber        As Long       'File error number
    Dim ErrorSource        As String     'File error source
    Dim ErrorDescription   As String     'File error description

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        FileNumber = FreeFile
        Open Path For Output As #FileNumber
        On Error GoTo WriteFailed
        For Each Entry In Records
            Print #FileNumber, CStr(Entry)
        Next Entry
        Close #FileNumber
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE A WRITE ERROR
'------------------------------------------------------------------------------
WriteFailed:
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
        On Error GoTo 0
        Close #FileNumber
        Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Sub


Private Sub WriteText( _
    ByVal Path As String, _
    ByVal Text As String)
'
'==============================================================================
'                                  WriteText
'------------------------------------------------------------------------------
' PURPOSE
'   Write text to a file exactly as given, without a line ending added.
'
' INPUTS
'   Path: full path of the file, created or overwritten. Text: ASCII text.
'
' ERROR POLICY
'   A file error is re-raised after the file is closed.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FileNumber         As Integer    'Handle of the open file
    Dim ErrorNumber        As Long       'File error number
    Dim ErrorSource        As String     'File error source
    Dim ErrorDescription   As String     'File error description

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'The trailing semicolon keeps Print from adding CR LF.
        FileNumber = FreeFile
        Open Path For Output As #FileNumber
        On Error GoTo WriteFailed
        Print #FileNumber, Text;
        Close #FileNumber
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE A WRITE ERROR
'------------------------------------------------------------------------------
WriteFailed:
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
        On Error GoTo 0
        Close #FileNumber
        Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Sub


Private Function StampedCommit() _
    As String
'
'==============================================================================
'                                StampedCommit
'------------------------------------------------------------------------------
' PURPOSE
'   The commit of the workbook's build stamp: the value of the defined name
'   STAMP_COMMIT_NAME in ThisWorkbook, the workbook that holds this module.
'
' RETURNS
'   The stamp as written by tools/devbuild/build_workbook.py, or an empty
'   string when the workbook has no such name or it cannot be read.
'
' UPDATED
'   2026-11-03
'==============================================================================
'
        On Error GoTo ReadFailed
        StampedCommit = CStr(ThisWorkbook.Names(STAMP_COMMIT_NAME).RefersToRange.Value)
        Exit Function

ReadFailed:
        StampedCommit = vbNullString

End Function


Private Function JsonText( _
    ByVal Value As String) _
    As String
'
'==============================================================================
'                                   JsonText
'------------------------------------------------------------------------------
' PURPOSE
'   A JSON string literal of Value in ASCII: the quote, the backslash, the
'   control characters and every character above code 126 are escaped.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index       As Long      'Position in Value
    Dim Character   As String    'Character at Index
    Dim Code        As Long      'Its UTF-16 code unit, 0 to 65535
    Dim Text        As String    'Escaped text so far

'------------------------------------------------------------------------------
' ESCAPE
'------------------------------------------------------------------------------
        For Index = 1 To Len(Value)
            Character = Mid$(Value, Index, 1)
            Code = AscW(Character) And &HFFFF&
            If Character = """" Or Character = "\" Then
                Text = Text & "\" & Character
            ElseIf Code < 32 Or Code > 126 Then
                Text = Text & "\u" & Right$("000" & Hex$(Code), 4)
            Else
                Text = Text & Character
            End If
        Next Index
        JsonText = """" & Text & """"

End Function


Private Function JsonTextOrNull( _
    ByVal Value As String) _
    As String
'
'==============================================================================
'                                JsonTextOrNull
'------------------------------------------------------------------------------
' PURPOSE
'   JSON null for an empty string, otherwise the string literal of Value.
'
' UPDATED
'   2026-11-03
'==============================================================================
'
        If Len(Value) = 0 Then
            JsonTextOrNull = "null"
        Else
            JsonTextOrNull = JsonText(Value)
        End If

End Function


Private Function ScopeJson( _
    ByVal Dataset As String, _
    ByVal Record As String) _
    As String
'
'==============================================================================
'                                  ScopeJson
'------------------------------------------------------------------------------
' PURPOSE
'   The scope of a result as the expected-results format writes it: null for
'   the whole calculation, otherwise {"dataset": ..., "record": ...}, the only
'   shape tools/check_accuracy.py matches with a record-scoped output.
'
' UPDATED
'   2026-11-03
'==============================================================================
'
        If Len(Dataset) = 0 And Len(Record) = 0 Then
            ScopeJson = "null"
        Else
            ScopeJson = "{""dataset"": " & JsonText(Dataset) & ", ""record"": " & _
                JsonText(Record) & "}"
        End If

End Function


Private Function ActiveWindowCaption() _
    As String
'
'==============================================================================
'                             ActiveWindowCaption
'------------------------------------------------------------------------------
' PURPOSE
'   The caption of Excel's active window, or an empty string when there is
'   none, so that a run which leaves another window active is seen.
'
' UPDATED
'   2026-11-02
'==============================================================================
'
        If Application.ActiveWindow Is Nothing Then
            ActiveWindowCaption = vbNullString
        Else
            ActiveWindowCaption = Application.ActiveWindow.Caption
        End If

End Function


Private Function VerifyHostState( _
    ByRef Detail As String) _
    As Boolean
'
'==============================================================================
'                               VerifyHostState
'------------------------------------------------------------------------------
' PURPOSE
'   Compare the four Application properties, the number of open workbooks
'   and the caption of the active window with their BeginRun values.
'
' RETURNS
'   True when all are unchanged; Detail names the changed ones, or why they
'   could not be compared. The harness never restores them: it never
'   changed them, and the procedure that did owns their restoration.
'
' UPDATED
'   2026-11-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Changed   As String    'Names of the changed properties

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        If Not mSnapshotValid Then
            Detail = "Excel state not compared: it could not be read at BeginRun"
            VerifyHostState = False
            Exit Function
        End If
        On Error GoTo ReadFailed
        If Application.Calculation <> mInitialCalculation Then Changed = Changed & " Calculation"
        If Application.DisplayAlerts <> mInitialAlerts Then Changed = Changed & " DisplayAlerts"
        If Application.EnableEvents <> mInitialEvents Then Changed = Changed & " EnableEvents"
        If Application.ScreenUpdating <> mInitialScreen Then Changed = Changed & " ScreenUpdating"
        If Application.Workbooks.Count <> mInitialWorkbooks Then Changed = Changed & " Workbooks"
        If ActiveWindowCaption() <> mInitialWindow Then Changed = Changed & " ActiveWindow"
        On Error GoTo 0
        If Len(Changed) = 0 Then
            Detail = "Excel state unchanged (Calculation, DisplayAlerts, EnableEvents, " & _
                "ScreenUpdating, Workbooks, ActiveWindow)"
            VerifyHostState = True
        Else
            Detail = "Excel state changed during the run:" & Changed
            VerifyHostState = False
        End If
        Exit Function

'------------------------------------------------------------------------------
' HANDLE A FAILED READ
'------------------------------------------------------------------------------
ReadFailed:
        Detail = "Excel state not compared: read failed with error " & CStr(Err.Number) & _
            " (" & Err.Description & ")"
        VerifyHostState = False

End Function


Private Function IsCaseId( _
    ByVal CaseId As String) _
    As Boolean
'
'==============================================================================
'                                   IsCaseId
'------------------------------------------------------------------------------
' PURPOSE
'   Whether CaseId is lower-case kebab-case: a-z and 0-9 in groups joined by
'   single hyphens, as the test-case files use (TEST_CASES.md, Files).
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Character   As String    'Character under test
    Dim i           As Long      'Position in CaseId

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        If Len(CaseId) = 0 Or Left$(CaseId, 1) = "-" Or Right$(CaseId, 1) = "-" Then
            IsCaseId = False
            Exit Function
        End If
        If InStr(CaseId, "--") > 0 Then
            IsCaseId = False
            Exit Function
        End If
        For i = 1 To Len(CaseId)
            Character = Mid$(CaseId, i, 1)
            If Not (Character Like "[a-z0-9-]") Then
                IsCaseId = False
                Exit Function
            End If
        Next i
        IsCaseId = True

End Function


Private Function IsKnownCase( _
    ByVal CaseId As String) _
    As Boolean
'
'==============================================================================
'                                 IsKnownCase
'------------------------------------------------------------------------------
' PURPOSE
'   Whether CaseId already started in this run.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Found   As Variant    'Item stored under CaseId, when present

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
    'A Collection raises error 5 for an unknown key; that error is the answer.
        On Error Resume Next
        Found = mCaseIds.Item(CaseId)
        IsKnownCase = (Err.Number = 0)
        On Error GoTo 0

End Function


Private Function NumberText( _
    ByVal Value As Double) _
    As String
'
'==============================================================================
'                                  NumberText
'------------------------------------------------------------------------------
' PURPOSE
'   A Double as text that does not depend on the regional settings: Str$
'   always uses a period as the decimal separator. Str$ drops the zero
'   before the point (" .75"), which is put back ("0.75", "-0.75").
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Text   As String    'Str$ result without its sign space

'------------------------------------------------------------------------------
' FORMAT
'------------------------------------------------------------------------------
        Text = Trim$(Str$(Value))
        If Left$(Text, 1) = "." Then
            Text = "0" & Text
        ElseIf Left$(Text, 2) = "-." Then
            Text = "-0" & Mid$(Text, 2)
        End If
        NumberText = Text

End Function


Private Function HostProperty( _
    ByVal PropertyName As String) _
    As String
'
'==============================================================================
'                                 HostProperty
'------------------------------------------------------------------------------
' PURPOSE
'   One Application property as text, or an empty string when the host
'   cannot read it.
'
' UPDATED
'   2026-11-02
'==============================================================================
'
        On Error GoTo ReadFailed
        Select Case PropertyName
            Case "Name"
                HostProperty = Application.Name
            Case "Version"
                HostProperty = Application.Version
            Case "Build"
                HostProperty = CStr(Application.Build)
            Case "OperatingSystem"
                HostProperty = Application.OperatingSystem
        End Select
        Exit Function

ReadFailed:
        HostProperty = vbNullString

End Function


Private Function EnvironmentSummary() _
    As String
'
'==============================================================================
'                              EnvironmentSummary
'------------------------------------------------------------------------------
' PURPOSE
'   Describe the Excel host and the compiled VBA environment.
'
' RETURNS
'   host, version, build, os, office (bitness) and runtime (VBA
'   generation); the last two come from conditional compilation, not from
'   Windows. A host property that cannot be read is left empty, so that the
'   run still starts and tools/check_excel_run.py rejects the evidence.
'
' UPDATED
'   2026-11-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Bitness   As String    'Office bitness
    Dim Runtime   As String    'VBA generation

'------------------------------------------------------------------------------
' DESCRIBE
'------------------------------------------------------------------------------
#If Win64 Then
        Bitness = "64-bit"
#Else
        Bitness = "32-bit"
#End If
#If VBA7 Then
        Runtime = "VBA7"
#Else
        Runtime = "VBA6"
#End If
        EnvironmentSummary = "host=" & HostProperty("Name") & "; version=" & HostProperty("Version") & _
            "; build=" & HostProperty("Build") & "; os=" & HostProperty("OperatingSystem") & _
            "; office=" & Bitness & "; runtime=" & Runtime

End Function


Private Sub ResetRunState()
'
'==============================================================================
'                                ResetRunState
'------------------------------------------------------------------------------
' PURPOSE
'   Return the counters, case identifiers, failures and flags of a run to
'   their idle values; the recovery flag and the log are left to callers.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        mRunActive = False
        mSuite = vbNullString
        mExpectedCases = 0
        mExpectedAssertions = 0
        mCaseCount = 0
        mAssertionCount = 0
        mFailureCount = 0
        mErrorCount = 0
        mCurrentCase = vbNullString
        Set mCaseIds = New Collection
        mPrimaryFailure = vbNullString
        mSnapshotValid = False

End Sub
