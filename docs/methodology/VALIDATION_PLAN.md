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

Cover, for each calculation, with the case categories of the
[test-case format](TEST_CASES.md#categories):

- typical cases and the boundaries of every rule;
- invalid and missing input, which the [data contract](DATA_CONTRACT.md)
  rejects;
- failure: earlier outputs are withdrawn;
- repeatability: the same input gives the same output.

## Evidence record

Each run of a check in Excel records:

- the commit of the source and the workbook it was built into;
- the case, its expected value, that value's provenance and the tolerance;
- the observed value and the outcome;
- the Excel build, its bitness and the Windows version;
- who ran it, who reviewed it, and when.

A check that was not run is recorded as `NOT RUN` with the reason. It is
never a pass.

<a id="host-behaviour"></a>

## Host behaviour

The code relies on how Excel and the VBE behave: which settings an operation
changes, how a name is re-cased, what an import needs, how a window or an
event behaves. Such behaviour can differ by Excel build, channel, bitness and
policy. Record each behaviour that the code or a guide relies on here, with
whether it was measured:

| ID | Behaviour | Status | Probe | Host | Date | Result | Issue |
| --- | --- | --- | --- | --- | --- | --- | --- |
| *Illustrative:* `HB-01` | The VBE re-cases every use of a name to its latest spelling | `assumed` | — | — | — | — | Link |

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
