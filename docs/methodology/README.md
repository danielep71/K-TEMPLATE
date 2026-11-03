# 📐 Methodology

These guides state, before and while the workbook is built, what each
calculation relies on, accepts, produces and how it is validated. Code review
and green checks show that the repository is consistent; only these guides,
and the evidence they require, show that the figures are right.

| Guide | States |
| --- | --- |
| [📚 Sources](SOURCES.md) | Every regulatory and technical source a rule relies on, the version used, and whether it has been verified. |
| [🧾 Data contract](DATA_CONTRACT.md) | The input datasets, their layout, units and validation, and how an import is traced. |
| [🧮 Calculation contracts](CALCULATION_CONTRACTS.md) | For each calculation: inputs, rules with their sources, outputs, limits and failure behaviour. |
| [✅ Validation plan](VALIDATION_PLAN.md) | Who checks what, the independent checks for each calculation, how each rule traces to its procedure, cases and status, what an evidence record contains, and which Excel host behaviour has been measured. |
| [🧪 Numerical test cases](TEST_CASES.md) | The format of the synthetic inputs and expected results, with their provenance and tolerances, that the checks enforce. |
| [📊 Workbook contract](WORKBOOK_CONTRACT.md) | The workbook's sheets, code names, defined names and areas, what each operation writes, and how a failure marks earlier results. |
| [🧪 VBA test harness](TEST_HARNESS.md) | The harness that runs the VBA suites in Excel, its outcomes, the log and certification formats, and the evidence an Excel run leaves. |

Read the first five in that order: a rule cites a source, consumes the data
contract's fields, and is validated as the plan describes, with test cases in
the format the fifth guide defines. The workbook contract then places those
inputs and results in the workbook, and the harness guide says how the VBA
suites run in Excel and what evidence a run leaves.

## Support matrix

The workbook's scope, feature by feature:

| Feature | Status | Calculation or input | Notes |
| --- | --- | --- | --- |
| *Illustrative:* Example total per counterparty | Supported | `CALC-01` | Implementation and validation status per rule: [traceability](VALIDATION_PLAN.md#traceability) |

- **Supported:** the workbook calculates it, as the named
  [calculation contract](CALCULATION_CONTRACTS.md) states. Whether each rule
  is implemented and validated is in the
  [traceability table](VALIDATION_PLAN.md#traceability).
- **Not supported:** out of scope. The contract's *Limits and non-goals* say
  what happens when the input holds it: it is reported as unsupported scope,
  never calculated silently
  ([input and output contracts](../REPOSITORY_STRUCTURE.md#input-and-output-contracts)).
- **Entered by the user:** the workbook does not calculate it: the user
  enters the value, which the [data contract](DATA_CONTRACT.md) or the
  [workbook contract](WORKBOOK_CONTRACT.md) defines and validates, and the
  workbook uses it as entered. Name that input in the third column.

## Filling in the guides

Each table holds one row marked *Illustrative*. Replace it with the project's
entries, keep the columns, and record each decision in the guide's
*Decisions* table with the issue where it was taken. A guide with only its
illustrative row states that nothing has been decided yet.

Byte-preserved copies of sources go under `docs/methodology/sources/`, whose
Git and editor rules are already set; [Sources](SOURCES.md) describes the
naming and the register entry each copy needs.
