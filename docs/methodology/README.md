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
| [✅ Validation plan](VALIDATION_PLAN.md) | Who checks what, the independent checks for each calculation, and what an evidence record contains. |

Read them in that order: a rule cites a source, consumes the data contract's
fields, and is validated as the plan describes.

## Filling in the guides

Each table holds one row marked *Illustrative*. Replace it with the project's
entries, keep the columns, and record each decision in the guide's
*Decisions* table with the issue where it was taken. A guide with only its
illustrative row states that nothing has been decided yet.

Byte-preserved copies of sources go under `docs/methodology/sources/`, whose
Git and editor rules are already set; [Sources](SOURCES.md) describes the
naming and the register entry each copy needs.
