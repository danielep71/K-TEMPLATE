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
(an empty string for a JSON fixture), its quantity, its scope (an empty
string for none) and its value. It raises `ERR_TEST_USAGE` outside a run.
The results belong to the run, like its log: `BeginRun` and `ResetTests`
clear them.

`SaveResults "<path>"` writes them as the results file of the
[accuracy contracts](TEST_CASES.md#accuracy-contracts): ASCII JSON with LF
line endings, an empty variant or scope written as `null`, and each value
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
| `exported/` | You: every VBA component exported back from the workbook after the run, each form with its `.frx` |
| `host.txt` | You: `windows=`, `references=` and `compile=` |
| `export-manifest.json` | `tools/devbuild/export_manifest.py`: the SHA-256 of the workbook and of every exported file, which binds the export to the workbook |

## Evidence policy

[`.github/excel-evidence-policy.json`](../../.github/excel-evidence-policy.json)
holds what an Excel run must leave as evidence:

- `harness` and `suite`: the harness module, the suite's name, source module
  and certification entry point, and the case and assertion counts it
  declares as `EXPECTED_CASES` and `EXPECTED_ASSERTIONS`;
- `modes`: the three modes in run order, the log each one saves and whether
  it must pass (only `NORMAL`);
- `evidence`: the names of the certification log, the export folder, the host
  record, the export manifest and the results file in the evidence folder;
- `host`: the exact references and the required Office bitness.

`tools/check_excel_run.py` reads the policy from the candidate commit, so a
run is judged by the policy of the commit it names, and rejects a policy that
is missing or malformed. It reads the harness and the suite at the exact paths
the policy names, and fails when the commit has no file there, or when the
suite's counts or entry point differ from the policy. The same policy is read
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
cannot see Excel: the compile result, the references and the Windows build
are the runner's own record.
