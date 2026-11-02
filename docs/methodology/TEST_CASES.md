# 🧪 Numerical test cases

How the numerical test cases are written: the files, the inputs and their
units, the expected results and their provenance, tolerances, case
categories, invalid inputs, and when a rule counts as validated. A case
tests one [calculation contract](CALCULATION_CONTRACTS.md), with inputs laid
out as the [data contract](DATA_CONTRACT.md) defines, against values whose
provenance the [source register](SOURCES.md) records. The
[validation plan](VALIDATION_PLAN.md) says which cases each calculation needs.

`python tools/check_test_cases.py` checks the format, and **Repository
integrity** runs it ([checking](#checking)). It never runs a calculation.

## Files

A case is one input file and one expected-result file:

| File | Holds |
| --- | --- |
| `tests/fixtures/<case-id>.json` | Synthetic inputs, grouped by dataset |
| `tests/expected/<case-id>.json` | Expected results, each with its provenance |
| `tests/derivations/<name>.py` | Optional: a script that computes expected values from the documented method, separately from the code under test |

- Case IDs are lowercase kebab-case, equal to the file names, and never
  reused.
- A case whose expected values are all `illustrative` starts with
  `illustrative-`, and only such a case may.
- Every fixture has an expected file, and every expected file a fixture.
- Files are UTF-8 JSON, stored with LF. All data is **synthetic**: never real
  positions, counterparties, clients, amounts or market data.

`illustrative-format-example` shows the format and holds no regulatory value.
It uses the illustrative entries of the methodology guides; remove it when you
replace them.

<a id="fixture-format"></a>

## Fixture format

```json
{
  "schema_version": 1,
  "kind": "test-fixture",
  "id": "illustrative-format-example",
  "synthetic": true,
  "description": "What the case exercises, in one sentence",
  "category": "nominal",
  "calculation": "CALC-01",
  "reporting_date": "2026-01-02",
  "reporting_currency_code": "EUR",
  "datasets": {
    "DS-01": [
      {"id": "P1", "inputs": {"currency_code": "EUR", "notional_amount": 600000}},
      {"id": "P2", "inputs": {"currency_code": "EUR", "notional_amount": 400000}}
    ]
  }
}
```

| Rule | Detail |
| --- | --- |
| Envelope | `schema_version` (1), `kind` (`test-fixture`), `id`, `synthetic` (`true`), `description`, `category`, `calculation`, `reporting_date`, `reporting_currency_code` and `datasets` are required, and no other key is allowed. An `invalid` case relaxes this for the calculation inputs ([invalid inputs](#invalid-inputs)) |
| Calculation | A `CALC-` ID declared by a heading of the [calculation contracts](CALCULATION_CONTRACTS.md) |
| Datasets | A non-empty object whose keys are `DS-` IDs declared in the *Datasets* table of the [data contract](DATA_CONTRACT.md); each holds a non-empty list of records |
| Records | Each record has an `id`, unique in the fixture, and an `inputs` object |
| Inputs | Keys are lower snake case and end in a unit suffix (below). The field tables of the data contract decide which keys exist |
| Currencies | `reporting_currency_code` is the currency of results. A record whose `inputs` hold an `_amount` key also holds `currency_code`, the currency of all its amounts, so a currency mismatch is always visible |
| Illustrative entries | A `CALC-` or `DS-` ID marked *Illustrative* in its guide may be used only by an `illustrative-` case |

**Unit suffixes.** A key without one of these suffixes is rejected:

| Suffix | Value |
| --- | --- |
| `_amount` | Number in the record's `currency_code`, unscaled |
| `_rate`, `_weight`, `_factor` | Decimal number: `0.05`, never `5` meaning 5 % |
| `_years` | Number of years |
| `_count` | Whole number, at least 0 |
| `_date` | Text, ISO `YYYY-MM-DD` |
| `_flag` | `true` or `false` |
| `_code` | Text code; one ending `currency_code` is three upper-case letters (ISO 4217) |
| `_id` | Text identifier |

`null` means not applicable. Files are strict JSON, with no duplicate key,
and numbers are finite in every category: `NaN`, `Infinity` and a number that
overflows to infinity, such as `1e400`, are rejected. Test a non-finite input
in a VBA test module instead.

<a id="expected-format"></a>

## Expected-result format

```json
{
  "schema_version": 1,
  "kind": "test-expected",
  "fixture": "illustrative-format-example",
  "outputs": [
    {
      "quantity": "illustrative_total_notional_amount",
      "scope": null,
      "rule": null,
      "value": 1000000,
      "unit": "EUR",
      "tolerance": {"absolute": 0.000001, "relative": 0.000000001},
      "reference": {
        "class": "illustrative",
        "source": null,
        "locator": null,
        "derivation": "Why the value exists and how it was obtained",
        "script": null,
        "derived_by": "Name",
        "reviewed_by": null,
        "date": "2026-11-02"
      }
    }
  ]
}
```

| Rule | Detail |
| --- | --- |
| Envelope | `schema_version` (1), `kind` (`test-expected`), `fixture` (the case ID) and a non-empty `outputs` list are required |
| Scope | `null` for a result of the whole calculation, or `{"dataset": "DS-NN", "record": "<id>"}` for a result of one record of the fixture |
| Quantity | Lower snake case, named as the calculation contract's outputs. Only an `illustrative-` case may use a name starting with `illustrative_` |
| Number | `value`, `unit` and `tolerance`. `unit` is a currency code for an amount, `1` for a rate, weight or factor, or `years` |
| Text | `text`, compared exactly, with no `unit` or `tolerance`. A rejection is a text output whose quantity ends in `_status` |
| Rule | A step of the case's calculation, `CALC-NN.<step>`, such as `CALC-01.2` for step 2 of its *Rules* table. It may be `null` only for an `illustrative` value |
| Reference | One per output; see [reference classes](#reference-classes) |

<a id="reference-classes"></a>

## Reference classes

Every expected value carries exactly one class. Only the first two are
numerical-validation references.

| Class | Meaning | Required fields |
| --- | --- | --- |
| `published` | Printed in a registered source, for example a worked example | `source` (a `SRC-` register ID) and an exact `locator` (article, paragraph, table, printed page) |
| `independent` | Computed by an external implementation by another author, registered in the source register with an `IMPL-` ID | `source` (`IMPL-…`), `locator`, `derivation`, `derived_by` (the implementation and its version) |
| `illustrative` | Shows the format or a behaviour; **not** a validated figure | `derivation`, saying why it exists and how the value was obtained |

- A `published` value never cites an `IMPL-` source, and an `independent`
  value always does. A source marked *Illustrative* in the register supports
  only an `illustrative` value.
- A value derived within the project, even from a registered text, stays
  `illustrative`, including one computed by a script under
  `tests/derivations/`.
- `script`, when not `null`, names a file under `tests/derivations/` that
  computed the value; it must exist.
- `reviewed_by` may be `null`, but record who reviewed a value when someone
  other than its author did.
- **Examples are never benchmarks** unless their source is registered and the
  value is printed there. Test output reports `illustrative` results
  separately and never as validation.

<a id="derivation-scripts"></a>

## Derivation scripts

A script under `tests/derivations/` computes expected values from the
documented method, written from the contracts and never from the code under
test. An output names it in `reference.script`, and
`python tools/check_derivations.py`, which **Repository integrity** runs,
compares the two:

- it runs each named script once per case, as
  `python -I tests/derivations/<name>.py tests/fixtures/<case-id>.json`;
- the script prints one JSON object, `{"outputs": [...]}`, whose items have the
  `quantity` and `scope` of an expected output and its `value` or `text`;
- every output that names the script must appear in what it prints, and match
  within the output's [tolerance](#tolerance), or exactly for a text.

A failing script, output that is not JSON, a missing quantity or a mismatch
fails the check. A match shows only that the script and the expected file
agree: the value stays `illustrative`, because the project derived it. The
scripts use the Python standard library only and are never a runtime
dependency of the workbook. `tests/derivations/illustrative_total.py` shows
the format for `illustrative-format-example`.

<a id="tolerance"></a>

## Tolerance

An actual result `a` matches the expected value `e` when

```text
|a − e| ≤ max(absolute, relative × |e|)
```

Both bounds are stated on every numeric output, as finite numbers of at least
0; there is no hidden default. Both 0 means an exact comparison. Record the
project's default bounds here once they are decided, with the issue that
decided them. A common starting point:

| Reference | `absolute` | `relative` |
| --- | --- | --- |
| `independent` or `illustrative`, full precision | `1e-6` currency units, `1e-9` for a rate, weight or factor | `1e-9` |
| `published`, rounded | Half a unit of the last printed digit: printed in thousands → `500` | `0` |

- **Near zero** the absolute bound governs: an expected `0` is matched within
  `absolute` only, so an `absolute` of 0 demands an exact zero.
- **Domain rules apply before tolerance:** a quantity that must not be
  negative fails if it is negative, however small.
- Results are compared unrounded, unless the contract rounds them.

<a id="categories"></a>

## Case categories

Each fixture declares one `category`:

| Category | Covers |
| --- | --- |
| `nominal` | Typical inputs within the domain |
| `boundary` | Values on a domain edge: a zero amount, a single record, a value at a cap, a result that must be exactly zero |
| `invalid` | Inputs that must be rejected: missing or unknown fields, a value outside its range, an unknown code or currency |

<a id="invalid-inputs"></a>

### Invalid inputs

- An `invalid` case is expected to be **rejected**, never calculated: its
  expected file holds at least one `_status` text output, and no numeric
  output for what it rejects.
- An `invalid` case separates the harness metadata, which must always be
  valid, from the calculation inputs, which may break any input rule, because
  that is what the case tests:
  - **Always valid:** `schema_version`, `kind`, `id`, `synthetic`,
    `description`, `category` and `reporting_date`, and the expected file.
  - **May be invalid, missing or unknown:** `calculation`,
    `reporting_currency_code` and `datasets`, with their records, keys, unit
    suffixes and values.
  - The `description` names the input rule the case breaks.
- The status texts are fixed by the calculation contract. A rejection never
  substitutes a default value.
- Inputs that JSON cannot express, such as non-finite numbers or wrong VBA
  types, are tested in a VBA test module instead of a fixture.

<a id="completeness-policy"></a>

## Completeness policy

- A rule is **validated** only when at least one `published` or
  `independent` case passes for it. `illustrative` cases never count.
- Each implemented rule also needs at least one `boundary` case, and each
  input domain at least one `invalid` case, as a fixture or a VBA test.
- A run that does not execute every expected case and assertion is
  **incomplete, not a pass**.
- A release states which rules are validated, under which version of their
  sources, and lists every rule that is implemented but not validated
  ([release certification](../release/RELEASE_CERTIFICATION.md)).

<a id="checking"></a>

## Checking

`tools/check_test_cases.py` checks the files under `tests/fixtures/` and
`tests/expected/`. With no case present it passes and says so.

It checks:

- strict JSON, with no duplicate key and only finite numbers, in every
  category;
- the harness metadata of every fixture;
- outside an `invalid` case, the calculation inputs:
  - the calculation and the datasets, against the methodology guides;
  - record identifiers;
  - keys, unit suffixes and value types;
  - currencies;
- each expected file's name and envelope;
- its outputs:
  - their scope against the fixture;
  - their form, unit and tolerance;
  - their rule, against the case's calculation;
- the reference class, its source against the register, and its script;
- the `illustrative` naming rules, and a `_status` text output in every
  `invalid` case.

It does not check:

- whether an input key or a quantity name belongs to the contracts;
- that an `invalid` case has no numeric output for what it rejects;
- any value: nothing is computed or compared. `tools/check_derivations.py`
  compares the values that a [derivation script](#derivation-scripts) computes,
  and running the cases on the workbook is the job of a VBA test harness in
  Excel.
