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

## Validation matrix

| Calculation | Independent checks | Evidence | Status |
| --- | --- | --- | --- |
| *Illustrative:* `CALC-01` | Hand-computed totals for a synthetic portfolio; empty input; invalid input rejected | Evidence record per release candidate | Not started |

Cover, for each calculation:

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

## Open decisions

| Decision | Options | Issue |
| --- | --- | --- |
| *Illustrative:* which tool produces the reference values | By hand, separate spreadsheet, script | Link |
