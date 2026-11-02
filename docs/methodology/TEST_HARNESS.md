# 🧪 VBA test harness

The harness runs the workbook's VBA test suites in Excel and writes a complete,
line-based log of each run. A certification runs a suite in every mode,
writes the evidence of the run, and raises only afterwards.
`tools/check_excel_run.py` then checks that evidence against one commit,
without Excel. [Recording an Excel run](../../INSTALLATION.md#recording-an-excel-run)
gives the procedure; this guide defines the modules, the outcomes and the
records.

## Modules

| Module | Holds |
| --- | --- |
| `tests/modules/TEST_Harness.bas` | The run lifecycle (`BeginRun`, `BeginCase`, `RecordCaseError`, `EndRun`), the assertions (`AssertEqualDouble`, `AssertNear`, `AssertEqualText`, `AssertTrue`, `AssertErrorValue`, `AssertRaised`, with `WithinTolerance` as the comparison rule), the log (`Say`, `SaveLog`, `LastResult`), recovery (`ResetTests`) and certification (`BeginCertification`, `CertifyRun`, `EndCertification`) |
| `tests/modules/TEST_Examples.bas` | The example suite: one case per kind of assertion, and the reference implementation of the [cleanup pattern](../VBA_HOUSE_STYLE.md#operation-and-cleanup-outcomes) with its failure paths. It holds no calculation of the workbook |

Both use `Option Private Module`, so they never become part of the
workbook's external surface. A production workbook contains neither.

A project suite follows the example suite: one `Case` procedure per case,
each starting with `BeginCase` and ending in an error handler that calls
`RecordCaseError`; a runner that calls every case once, between `BeginRun`
and `EndRun`; and the expected case and assertion counts as constants. The
expected values of a numerical case come from its
[test case](TEST_CASES.md), with the tolerance that the case states.

## Run outcomes

`EndRun` reports four outcomes separately:

| Outcome | Values | Meaning |
| --- | --- | --- |
| `operation` | `PASS`, `FAIL`, `ERROR` | What ran: every assertion passed, an assertion or a case identifier failed, or an unexpected error occurred |
| `completeness` | `COMPLETE`, `INCOMPLETE` | `COMPLETE` only when both expected counts were met exactly. A skipped case is `INCOMPLETE`, whatever the cases that ran did ([completeness policy](TEST_CASES.md#completeness-policy)) |
| `cleanup` | `PASS`, `FAIL` | Whether Excel's state is as it was at `BeginRun`: `Calculation`, `DisplayAlerts`, `EnableEvents`, `ScreenUpdating`, the number of open workbooks and the active window |
| `recovery` | `NONE`, `REQUIRED` | `REQUIRED` when cleanup failed or could not be checked, or when a run was interrupted |

`RESULT=PASS` needs `operation=PASS`, `COMPLETE`, cleanup `PASS` and no
recovery. Any other run raises `ERR_TEST_FAILED`, or the unexpected error that
escaped the runner, only after its summary is logged.

A run that starts while another is active, or while recovery is required, is
refused with `ERR_TEST_DIRTY_START`. Check Excel's state by hand, then type
`TEST_Harness.ResetTests`; it restores nothing itself.

## Modes and certification

The example suite has three modes:

| Mode | Log | Must end |
| --- | --- | --- |
| `NORMAL` | `examples.log` | `RESULT=PASS`, `COMPLETE` |
| `INJECTED_FAILURE` | `injected.log` | `RESULT=FAIL`, `operation=FAIL`, `COMPLETE`, one failure: the designated assertion |
| `SKIPPED_CASE` | `skipped.log` | `RESULT=FAIL`, `operation=PASS`, `INCOMPLETE`, one failure: `harness.completeness` |

`TEST_Examples.CertifyExamples "<folder>"` is the certification entry point:

1. `BeginCertification` refuses to start while a run or a certification is
   active, or while recovery is required.
2. Each mode runs. Its error is recorded, not raised, and its log is saved,
   so that every log is written whatever the run did.
3. `CertifyRun` records how the run reached its caller. The normal run must
   pass and be complete. Each deliberate failure must reach its caller as
   `ERR_TEST_FAILED` with a description. This is the self-test that a failed
   run is never silent.
4. `EndCertification` writes `certification.log`, then raises
   `ERR_TEST_FAILED` if the certification did not pass.

## Log format

One record per line, `KEY=value`, with further fields separated by `; `.
`RUN` and `ENVIRONMENT` come first; `CASE`, `ASSERT`, `ERROR` and `FAILURE`
follow as the run reaches them; the summary records come last, in the order
of the table:

| Record | Fields |
| --- | --- |
| `RUN=<suite>` | `mode` (`NORMAL` or a deliberate failure mode), `expected_cases`, `expected_assertions`, `started` |
| `ENVIRONMENT=` | `host`, `version`, `build` (`Application.Build`), `os`, `office` (`64-bit` or `32-bit`), `runtime` (`VBA7` or `VBA6`) |
| `CASE=<case-id>` | One per case, in run order |
| `ASSERT=<case-id>/<name>` | `status` (`PASS` or `FAIL`); on failure, the expected and actual values |
| `FAILURE=<id>` | One per failure: an assertion, `<case-id>.unexpected-error`, `runner.unexpected-error`, `harness.case-id`, `harness.no-case` or `harness.completeness`, with its detail |
| `ERROR=<case-id>` or `ERROR=run` | `number`, `source`, `description` of an unexpected error |
| `CASES=<n>`, `ASSERTIONS=<n>` | `expected` |
| `FAILURES=<n>` | `errors` |
| `PRIMARY_FAILURE=<id>` | The first failure of the run, when there is one |
| `CLEANUP=PASS` or `FAIL` | `detail`: the state compared, or what changed |
| `RECOVERY=NONE` or `REQUIRED` | `reason`, `action`, when required |
| `RESULT=PASS` or `FAIL` | `suite`, `operation`, `completeness`, `cases`, `assertions`, `failures`, `cleanup`, `recovery` |

`RESULT` is always the last record of a run. A refused start writes a
`RESULT=FAIL` with `operation=NOT_RUN` and a `RECOVERY=REQUIRED` record, and
`ResetTests` writes `RESET=DONE`. Case identifiers and assertion names are
fixed in the suite's source, so two logs of the same commit can be compared
line by line, apart from `started` and `ENVIRONMENT`.

## Certification log

`certification.log` holds one record per line, in this order:

| Record | Fields |
| --- | --- |
| `CERTIFY=<suite>` | `started` |
| `MODE=<mode>` | One per mode, in run order: `expected` (`PASS` or `FAIL`), `error` (the number that reached the caller, `0` for none), `result` (`PASS`, `FAIL`, or `NONE` when the run wrote no `RESULT`), `completeness` (`COMPLETE`, `INCOMPLETE` or `NONE`), `status` (`PASS` when the run ended as its mode requires) |
| `CERTIFICATION=PASS` or `FAIL` | `suite`, `completeness` (`COMPLETE` when every run that must pass was complete), `modes`, `failures` |

## Evidence folder

| File | Written by |
| --- | --- |
| `examples.log`, `injected.log`, `skipped.log` | `CertifyExamples`, one per mode |
| `certification.log` | `EndCertification` |
| `exported/` | You: every VBA component exported back from the workbook after the run, each form with its `.frx` |
| `host.txt` | You: `windows=`, `references=` and `compile=` |
| `export-manifest.json` | `tools/devbuild/export_manifest.py`: the SHA-256 of the workbook and of every exported file, which binds the export to the workbook |

`tools/check_excel_run.py` holds the layout, the modes, the supported
references and the required Office bitness as constants at its top. Change
them there, together with the suite and this guide.

## What it proves

The example suite exercises the harness itself: assertions, tolerances, error
values, raised errors, completeness, cleanup and certification. A passing run
validates no calculation of the workbook. Without Excel,
`tests/test_vba_harness.py` checks what the source can show: the example
suite's counts equal its cases and assertions, every case runs once with a
unique identifier, every record in this guide is written, and
`EndCertification` writes the log before it raises. `check_excel_run.py`
proves only that an evidence folder is consistent with a commit. It cannot see
Excel: the compile result, the references and the Windows build are the
runner's own record.
