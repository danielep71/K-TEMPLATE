# 🧾 Data contract

This contract defines every input the workbook accepts: where it comes from,
its layout and units, what makes it invalid, and how an import is traced.
The [calculation contracts](CALCULATION_CONTRACTS.md) refer to its datasets
and fields by name.

## Datasets

| ID | Description | Provider | Frequency | Granularity | File or sheet |
| --- | --- | --- | --- | --- | --- |
| *Illustrative:* `DS-01` | Example positions | Upstream system or manual input | Per reporting date | One row per position | `positions.csv` |

## File format

State, for each file the workbook imports:

- the encoding, delimiter, decimal separator and quoting;
- the date format, ISO 8601 (`YYYY-MM-DD`) unless the provider imposes
  another;
- whether a header row is required, and whether column order matters;
- how a missing value is written, kept distinct from zero.

## Fields

One table per dataset:

| Field | Type | Unit or currency | Required | Allowed values or range | Notes |
| --- | --- | --- | --- | --- | --- |
| *Illustrative:* `notional` | Decimal | Currency in field `currency` | Yes | ≥ 0 | Sign is given by `direction` |

State the unit of every amount, rate and period explicitly: a percentage, a
decimal fraction and basis points are different fields, not different
formats.

## Validation and invalid input

| Rule | Rejects | Effect |
| --- | --- | --- |
| *Illustrative:* `DV-01` | A required field that is empty | The import stops and lists every row that fails |

- Each rule has an identifier that the workbook reports with its finding.
- Invalid input is reported, never silently corrected, rounded or defaulted.
- State whether a rule stops the whole import or excludes the row, and where
  excluded rows are listed.
- After a failed import, the workbook shows no results from earlier data as if
  they were current.

## Point-in-time data

Use only data that was known on the reporting date. State, for each dataset,
which date decides that: the reporting date, a market data cut-off, or the
date a record was booked.

## Volume

Excel worksheets hold at most 1,048,576 rows and 16,384 columns. State the
expected and the largest supported volume, and how the workbook imports data
beyond a sheet's limits, if it needs to.

## Lineage

Each import records, where the workbook can read it later:

- the file name and its SHA-256;
- the number of rows read, accepted and rejected;
- the reporting date and the import time;
- the commit of the workbook's source, when the build records it.

## Synthetic data

The repository holds synthetic data only: no real positions, counterparties,
clients or amounts. Test fixtures say they are synthetic and how they were
made.

## Decisions

| Date | Decision | Reason | Issue |
| --- | --- | --- | --- |
| *Illustrative* | How a missing value is written in a file | Why | Link |
