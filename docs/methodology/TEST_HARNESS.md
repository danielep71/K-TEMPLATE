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
| `tests/modules/TEST_Harness.bas` | The run lifecycle (`BeginRun`, `BeginCase`, `RecordCaseError`, `EndRun`), the assertions (`AssertEqualDouble`, `AssertNear`, `AssertEqualText`, `AssertTrue`, `AssertErrorValue`, `AssertRaised`, with `WithinTolerance` as the comparison rule), the log (`Say`, `Echo`, `SaveLog`, `LastResult`), the [results](#results) (`RecordResult`, `SaveResults`), recovery (`ResetTests`) and certification (`BeginCertification`, `CertifyRun`, `EndCertification`) |
| `tests/modules/TEST_Examples.bas` | The example suite: one case per kind of assertion, one case that reads the illustrative [test cases](TEST_CASES.md) and records their totals, and the reference implementation of the [cleanup pattern](../VBA_HOUSE_STYLE.md#operation-and-cleanup-outcomes) with its failure paths. It holds no calculation of the workbook |

Both use `Option Private Module`, so they never become part of the
workbook's external surface. A production workbook contains neither.

A project suite follows the example suite: one `Case` procedure per case,
each starting with `BeginCase` and ending in an error handler that calls
`RecordCaseError`; a runner that calls every case once, between `BeginRun`
and `EndRun`; and the expected case and assertion counts as constants. The
expected values of a numerical case come from its
[test case](TEST_CASES.md), with the tolerance that the case states. A case
that computes an output of an accuracy contract also records it with
`RecordResult`.

The example suite's entry points take the full path of `tests/fixtures/` in
the checkout, from which its results case reads the illustrative test cases:
`RunExampleTests "<fixtures>"`, and so on.

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

`TEST_Examples.CertifyExamples "<folder>", "<fixtures>"` is the certification
entry point:

1. `BeginCertification` refuses to start while a run or a certification is
   active, or while recovery is required.
2. Each mode runs. Its error is recorded, not raised, and its log is saved,
   so that every log is written whatever the run did. After the normal run,
   `SaveResults` saves its results as `results.json`.
3. `CertifyRun` records how the run reached its caller. The normal run must
   pass and be complete. Each deliberate failure must reach its caller as
   `ERR_TEST_FAILED` with a description. This is the self-test that a failed
   run is never silent.
4. `EndCertification` writes `certification.log`, then raises
   `ERR_TEST_FAILED` if the certification did not pass.

A project suite is certified in its normal run only. Its certification entry
point takes the same two arguments and follows the same steps with one mode:

```vb
Public Sub CertifyProject(ByVal Folder As String, ByVal Fixtures As String)
        TEST_Harness.BeginCertification SUITE_NAME
        ' Run the suite in NORMAL mode, capturing its error as CertifyMode does,
        TEST_Harness.SaveLog Folder & "\project.log"
        TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, RunErrorDescription
        TEST_Harness.EndCertification Folder & "\project-certification.log"
End Sub
```

List the suite under `suites` in the [evidence policy](#evidence-policy), with
the names of its two logs. Each suite declares its counts for the Office
bitness the policy supports: a case or assertion inside `#If Win64` counts
only where that branch is compiled, and the check reads the suite the same
way.

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

## Results

`RecordResult` keeps one result of the running suite: its case, its variant
(an empty string for a JSON fixture), its quantity, the dataset and record of
its scope (two empty strings for a result of the whole calculation) and its
value. It raises `ERR_TEST_USAGE` outside a run.
The results belong to the run, like its log: `BeginRun` and `ResetTests`
clear them.

`SaveResults "<path>"` writes them as the results file of the
[accuracy contracts](TEST_CASES.md#accuracy-contracts): ASCII JSON with LF
line endings, an empty variant or scope written as `null`, a record scope as
`{"dataset": …, "record": …}`, and each value
with the 15 significant digits of `Str$`, whatever the regional settings.
Its `commit` is the value of the defined name `build_commit` in the
workbook that holds the harness, the build stamp of
`tools/devbuild/build_workbook.py`. A workbook without that name writes an
empty commit, and a build with local changes writes its marked stamp; the
accuracy check rejects both. This is the only part of a workbook the harness
reads, and it writes none.

## Evidence folder

| File | Written by |
| --- | --- |
| `examples.log`, `injected.log`, `skipped.log` | `CertifyExamples`, one per mode |
| `results.json` | `SaveResults`, after the normal run |
| `certification.log` | `EndCertification` |
| A log and a certification log per project suite, as `suites` names them | Its certification entry point |
| `exported/` | By hand only: every VBA component exported back from the workbook after the run, each form with its `.frx`. Without it, `check_excel_run.py --workbook` proves identity by rebuilding the commit |
| `host.txt` | You, or `tools/devbuild/excel_run.py`: the [host record](#host-record), also for a run that stopped early |
| `export-manifest.json` | By hand only: `tools/devbuild/export_manifest.py`, the SHA-256 of the workbook and of every exported file, which binds the export to the workbook |

<a id="host-record"></a>

### Host record

`host.txt` holds one `key=value` line per field, in this order:

| Field | Value |
| --- | --- |
| `operator` | Who ran it |
| `started`, `finished` | When the run started and finished, in ISO 8601 with the UTC offset, such as `2026-11-03T14:05:09+01:00` |
| `windows` | The Windows version, as `winver` shows it |
| `excel` | `Application.Version` and `Application.Build`, such as `16.0 build 17928` |
| `bitness` | `64-bit` or `32-bit`: the bitness of Excel |
| `automation_security` | `Application.AutomationSecurity` as found: `Low`, `ByUI` or `ForceDisable` (1, 2 or 3). It is read, never set |
| `vba_project_access` | *Trust access to the VBA project object model* as found: `TRUSTED` or `NOT_TRUSTED` |
| `references` | The ticked references, separated by `; ` |
| `compile` | `PASS` when **Debug → Compile VBAProject** completed, `ON_DEMAND` for a run without Trust access, or what the compile reported |
| `outcome` | How the run ended: `PASS`, `FAIL`, `IMPORT_FAILED`, `COMPILE_FAILED`, `EXECUTION_TIMEOUT`, `INCOMPLETE`, `NOT_ENABLED` or `UNAVAILABLE` |

`PASS` means every certification entry point returned, and `FAIL` that one
raised. The other outcomes say where the run stopped: the components could
not be imported, the project did not compile, Excel did not finish in time,
the run stopped before its end, the macros were not enabled, or Excel could
not be started. A run that stopped early writes `UNKNOWN` for a value it could
not read; the record of a `PASS` has none.

`check_excel_run.py` rejects a record that lacks a field or holds a value
outside this table, a time without its UTC offset, a finish before the start,
or, for a `PASS`, an `excel` or `bitness` other than the logs' `ENVIRONMENT`.
It accepts any other outcome only as evidence of that outcome: the check then
fails, whatever the rest of the folder holds.

## Evidence policy

[`.github/excel-evidence-policy.json`](../../.github/excel-evidence-policy.json)
holds what an Excel run must leave as evidence:

- `harness` and `suite`: the harness module, the example suite's name, source
  module and certification entry point, and the case and assertion counts it
  declares as `EXPECTED_CASES` and `EXPECTED_ASSERTIONS`;
- `suites`: the project suites, each with its name, source module,
  certification entry point, counts, `log` and `certification_log`, run in
  `NORMAL` mode after the example suite; empty in the template;
- `modes`: the three modes in run order, the log each one saves and whether
  it must pass (only `NORMAL`);
- `evidence`: the names of the certification log, the export folder, the host
  record, the export manifest and the results file in the evidence folder;
- `host`: the exact references and the required Office bitness.

`tools/check_excel_run.py` reads the policy from the candidate commit, so a
run is judged by the policy of the commit it names, and rejects a policy that
is missing or malformed. It reads the harness and every suite at the exact
paths the policy names, as the policy's Office bitness compiles them, and
fails when the commit has no file there, or when a suite's counts or entry
point differ from the policy. Every log of every suite must come from one
host. A policy of version 1, which older commits hold, has no `suites`. The same policy is read
by `tools/devbuild/export_manifest.py`, for the names of the export folder and
the manifest, and by `tools/render_docs.py`, for the expected results in the
installation guide. `tests/test_check_excel_run.py` keeps the policy equal to
the suite, to the host policy of the installation guide and to the evidence
files the guide names. To change a mode, a file name, a reference or the
bitness, change the policy together with the suite, the installation guide
and this guide.

## What it proves

The example suite exercises the harness itself: assertions, tolerances, error
values, raised errors, completeness, cleanup, certification and the results
file. Its results case reads the illustrative test cases, whose contract is
illustrative, so a passing run validates no calculation of the workbook.
Without Excel, `tests/test_vba_harness.py` checks what the source can show:
- the example suite's counts equal its cases and assertions;
- every case runs once with a unique identifier;
- every record in this guide is written;
- `EndCertification` writes the log before it raises;
- the results case records exactly the outputs of the contracts' cases, from
  files that exist;
- the harness reads the stamp name the build writes.

`check_excel_run.py` proves that an evidence folder is consistent with a
commit, and measures the commit's accuracy contracts on its results file. It
cannot see Excel: the host record, from the operator to the outcome, is the
runner's own record.
