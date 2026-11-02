# 🧮 Calculation contracts

A contract states what one calculation must produce, from which inputs, under
which rules. The VBA implements the contract, the
[validation plan](VALIDATION_PLAN.md) tests against it, and a change to a
result it defines is a `behavior-change`.

## Common rules

State once, for every calculation:

- **Units and signs:** the currency, unit and sign convention of every amount,
  and whether rates are fractions or percentages.
- **Dates:** the reporting date, the day-count conventions, and how
  maturities and residual terms are measured.
- **Precision and rounding:** VBA `Double` arithmetic throughout, rounding
  only where a source requires it, with the method and the number of digits.
- **Currency conversion:** which rates, from which date and source.
- **Failure:** a calculation that cannot complete reports why, and withdraws
  its earlier outputs instead of leaving them as if current.

## Contract template

Copy this section for each calculation.

### *Illustrative:* `CALC-01` Example calculation

**Purpose:** what the figure is, and where it is reported.

**Inputs:**

| Input | Dataset and field | Unit |
| --- | --- | --- |
| Example amount | `DS-01.notional` | Currency of `DS-01.currency` |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | Example: sum the amounts per counterparty | `SRC-01`, Article 1(1) |

**Outputs:**

| Output | Unit | Scope | Rounding |
| --- | --- | --- | --- |
| Example total | Reporting currency | Per counterparty | None |

**Limits and non-goals:** cases the calculation does not cover, and what
happens if they occur in the input.

**Failure behaviour:** which conditions stop it, and what it reports.

**Validation:** the checks in the [validation plan](VALIDATION_PLAN.md) that
cover it.

## Decisions

| Date | Decision | Reason | Issue |
| --- | --- | --- | --- |
| *Illustrative* | How a rule the source leaves open is applied | Why | Link |
