# ✅ Validation plan

This plan says how each [calculation contract](CALCULATION_CONTRACTS.md) is
shown to hold, and what evidence a result needs before anyone relies on it.

## Responsibilities

Implementing a calculation, checking it independently and accepting it are
separate steps, recorded separately even when one person does all three.
Code review and passing repository checks are necessary, but they are not
evidence that a figure is right: the static checks never run Excel or a
calculation.

## Independent reference values

An expected result is computed outside the workbook: by hand, in a separate
spreadsheet that shares no formulas with it, or by a script, for example in
Python or R. Record how each value was derived and who derived it. A
reference script is a test tool, never a runtime dependency of the workbook.

Each comparison has an absolute and a relative tolerance, with the reason for
them: the source's own rounding, or the precision of `Double` arithmetic.

The [numerical test cases](TEST_CASES.md) record each expected value with its
reference class (`published`, `independent` or `illustrative`), its source,
its derivation and its tolerance. Only `published` and `independent` values
validate a rule.

## Validation matrix

| Calculation | Independent checks | Evidence | Status |
| --- | --- | --- | --- |
| *Illustrative:* `CALC-01` | Hand-computed totals for a synthetic portfolio; empty input; invalid input rejected | Evidence record per release candidate | Not started |

The [traceability table](#traceability) gives the status of each rule of a
calculation.

Cover, for each calculation, with the case categories of the
[test-case format](TEST_CASES.md#categories):

- typical cases and the boundaries of every rule;
- invalid and missing input, which the [data contract](DATA_CONTRACT.md)
  rejects;
- failure: earlier outputs are withdrawn;
- repeatability: the same input gives the same output.

## Traceability

Each rule is traced from the paragraph of its source to the step of the
calculation contract that states it, to the VBA procedure that implements it
and to the test cases that test it, with its validation status:

| Source | Paragraph | Calculation step | VBA procedure | Test cases | Status |
| --- | --- | --- | --- | --- | --- |
| *Illustrative:* `SRC-01` | Article 1(1) | `CALC-01.1` | — | `illustrative-format-example`, `illustrative-file-import-example` | Not implemented |

- **Source and paragraph:** a `SRC-` ID of the [source register](SOURCES.md)
  and the exact paragraph, as in the step's *Source* column.
- **Calculation step:** `CALC-NN.<step>`, a step of the *Rules* table of
  that [calculation contract](CALCULATION_CONTRACTS.md). A rule implemented
  in several steps has one row per step.
- **VBA procedure:** `Module.Procedure` of a component under `src/`, or `—`
  while none implements the rule. Name every procedure that applies the rule.
- **Test cases:** the IDs of the [test cases](TEST_CASES.md) that test the
  step, or `—`. Each case tests the row's calculation.
- **Status:** one of the four below.

| Status | Meaning | The row names |
| --- | --- | --- |
| Validated | A `published` or `independent` case passes for the rule on an Excel run of the release candidate ([completeness policy](TEST_CASES.md#completeness-policy)) | A procedure, and a case with a `published` or `independent` output whose `rule` is the row's step |
| Exercised | Cases run the rule, but none is a `published` or `independent` case that passed: they are `illustrative`, or the reference cases have not passed yet | A procedure and at least one case |
| Not validated | Implemented, but no case tests it | A procedure and no case |
| Not implemented | No procedure implements it yet; its cases may come first | No procedure |

`python tools/check_test_cases.py`, which **Repository integrity** runs,
fails a row whose source is not in the register, whose step belongs to no
declared calculation, whose procedure is not in a component under `src/`,
whose case does not exist or tests another calculation, or whose status
contradicts what the row names. It cannot see an Excel run: whether a
`Validated` rule's case passed on the release candidate is the
[evidence record](#evidence-record)'s to show.

## Evidence record

Each run of a check in Excel records:

- the commit of the source and the workbook it was built into;
- the case, its expected value, that value's provenance and the tolerance;
- the observed value and the outcome;
- the Excel build, its bitness and the Windows version;
- who ran it, who reviewed it, and when.

A check that was not run is recorded as `NOT RUN` with the reason. It is
never a pass.

An Excel run of the harness keeps the host, who ran it, when and how it
ended in its [host record](TEST_HARNESS.md#host-record), also when it stopped
early. `tools/check_excel_run.py` rejects a record that lacks one of them, and
accepts an outcome other than `PASS` only as evidence of that outcome.

### What an accuracy contract proves

A frozen [accuracy contract](TEST_CASES.md#accuracy-contracts) whose results
pass on an Excel run of the candidate shows that, on that host, the worst
error of the calculation over its calibration and holdout cases stays within
the threshold. The holdout cases played no part in setting the threshold,
and their values are `published` or `independent`. It proves nothing outside
those cases, regime and host. A draft contract, and any contract on
`illustrative` values, validates nothing. Record each contract's worst errors
with the evidence.

<a id="host-behaviour"></a>

## Host behaviour

The code relies on how Excel and the VBE behave: which settings an operation
changes, how a name is re-cased, what an import needs, how a window or an
event behaves. Such behaviour can differ by Excel build, channel, bitness and
policy. Record each behaviour that the code or a guide relies on here, with
whether it was measured:

| ID | Behaviour | Status | Probe | Host | Date | Result | Issue |
| --- | --- | --- | --- | --- | --- | --- | --- |
| `HB-01` | The VBE re-cases every use of a name to its latest spelling, so an exported module can differ from its source only in the letter case of code outside strings and comments | `assumed` | The next Excel run by hand: the `Note:` lines of `tools/check_excel_run.py` | — | — | Not measured: no Excel run has recorded a re-cased export yet. `tools/check_excel_run.py` tolerates it: it compares that code without letter case and reports each re-cased identifier as a note | #245 |
| `HB-02` | Excel started through COM can follow the user's macro setting (`AutomationSecurity` is `msoAutomationSecurityByUI`, set by policy). A workbook it opens then has its macros disabled, and `Application.Run` fails with *Excel 4.0 (XLM) macros have been disabled*, because the VBA procedure is not found | `measured` | By hand: in PowerShell, open the built workbook through COM and `Run` `'<workbook>'!TEST_Harness.LastResult` | Microsoft 365 Excel 16.0 build 20430, 64-bit; Windows 11 Enterprise 25H2, build 26200.9711 | 2026-11-03 | `AutomationSecurity` was `ByUI`, and `Run` failed with that message. `tools/devbuild/excel_run.py` changes no macro setting: it waits for **Enable Content** | #204 |
| `HB-03` | `Application.Run` reaches a public procedure of an `Option Private Module` module when the name is qualified with the workbook (`'Book.xlsm'!Module.Procedure`) | `measured` | As `HB-02`, in a session whose macros were enabled | As `HB-02` | 2026-11-03 | `'Workbook-development.xlsm'!TEST_Harness.LastResult` ran | #204 |
| `HB-04` | In a workbook opened through COM with its macros disabled, **Enable Content** enables them without reopening, and `Application.Run` then reaches its procedures | `measured` | `python tools/devbuild/excel_run.py` on a host with `HB-02`: click **Enable Content** in the yellow bar | As `HB-02` | 2026-11-03 | Excel showed the yellow bar; after the click, the certification ran to `excel-run: PASS` at 50c0658 without reopening the workbook | #204 |

- **Status:** `measured` when a probe ran on the named host and its output is
  kept; `assumed` otherwise.
- **Probe:** a `TEST_` procedure under `tests/modules/` that observes the
  behaviour and prints what it saw, run with the VBA test harness. A
  measurement found by hand names the steps instead.
- **Host:** the Excel product, version and build, its channel and bitness,
  and the Windows version, as in the [evidence record](#evidence-record).
- **Result:** what was observed, and the model it supports. A measurement
  holds for the named host only; it is never generalised to others.

A guide that states host behaviour the code relies on cites its `HB-` entry.
An `assumed` entry is listed under *Limits* in the
[release certification](../release/RELEASE_CERTIFICATION.md) until it is
measured.

## Open decisions

| Decision | Options | Issue |
| --- | --- | --- |
| *Illustrative:* which tool produces the reference values | By hand, separate spreadsheet, script | Link |
