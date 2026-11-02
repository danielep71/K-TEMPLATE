# 📊 Workbook contract

This contract defines what the workbook looks like:

- its sheets and their code names;
- the defined names the code addresses;
- which areas are inputs and which the code writes;
- how each operation reads and writes them;
- how a failed run marks earlier results as no longer current;
- what the build creates.

The [data contract](DATA_CONTRACT.md) defines the inputs and the
[calculation contracts](CALCULATION_CONTRACTS.md) the results; this contract
says where they sit in the workbook. The
[repository structure](../REPOSITORY_STRUCTURE.md) says which components may
touch a sheet, and the [house style](../VBA_HOUSE_STYLE.md) how they are
named.

## Principles

These defaults hold unless a decision below replaces one:

- **Built from source, never committed.** The build creates the workbook from
  the exported source of one commit, as the
  [installation guide](../../INSTALLATION.md) describes, and saves it outside
  the checkout or under `test-results/`. Only a macro-free
  `src/workbook/Template.xlsx` may be tracked.
- **Datasets where the data contract says.** The *File or sheet* column of the
  data contract's *Datasets* table says whether a dataset is read from a file
  into memory or typed on a sheet. A dataset read from a file is never copied
  onto a sheet.
- **No calculation formulas on sheets.** VBA writes every calculated value,
  through the public facade, as a value. Input areas may use data validation;
  conditional formats and charts may read written values.
- **Adapters own the sheets.** Only the Excel adapters (`M_` modules) and the
  document modules in `src/workbook/` read or write a worksheet. They address
  a sheet by its code name and a range by its defined name, never by tab name
  or a hard-coded address. Core, classes and the facade never touch a
  workbook object.
- **No partial results.** An operation validates first, calculates in memory,
  then clears and writes all of its result areas. On failure it writes none
  of them and marks them `FAILED` ([status and staleness](#status-and-staleness)).

## Sheets

Tab order is the order of the table. A sheet is visible unless a decision
below says why it is hidden.

| # | Tab name | Code name | Role | Code writes | Content |
| ---: | --- | --- | --- | --- | --- |
| *Illustrative:* 1 | `Readme` | `shReadme` | Reference | Build only | What the workbook is, its limits, and the build stamp |
| *Illustrative:* 2 | `Inputs` | `shInputs` | Input | No | Reporting date and run settings ([areas](#areas)) |

- **Code name:** `sh` plus the role, as the house style requires. The
  workbook's own document module is `ThisWorkbook`.
- **Role:**
  - `Reference`: what the workbook is, its limits and its build;
  - `Input`: typed by users;
  - `Parameter`: values a calibration writes and later runs read;
  - `Output`: results;
  - `Diagnostics`: findings, logs and test results;
  - `Test`: present only in a development build.
- **Code writes:**
  - `No`: users edit the sheet, and code never writes it;
  - `Yes`: code replaces its content, and users never edit it;
  - `Append only`: code adds rows, and never changes or deletes them;
  - `Build only`: only the build writes it.

## Layout

State the layout every sheet shares. A common default:

- The first rows carry the workbook's title band and the sheet title. Every
  sheet that code writes has a status line, named `<Sheet>_Status`.
- A table's first header cell carries the name `<Sheet>_<Table>`, and its data
  starts on the row below. Tables that grow downward sit side by side, so
  that none can overwrite another.
- Input cells look different from cells that code writes.
- A table that would exceed a worksheet's 1,048,576 rows fails the operation
  before anything is written; it is never truncated.

## Defined names

| Name | Sheet | Refers to | Written by |
| --- | --- | --- | --- |
| *Illustrative:* `build_commit` | `shReadme` | The commit the workbook was built from | The build |
| *Illustrative:* `build_time` | `shReadme` | The commit's committer time | The build |
| *Illustrative:* `build_configuration` | `shReadme` | `production` or `development` | The build |
| *Illustrative:* `cfg_reporting_date` | `shInputs` | The value cell of the reporting date | Users |

- Every name is workbook-scoped and visible, and refers to one cell of this
  workbook by an absolute reference. External and relative references are
  not allowed.
- **Settings:** `cfg_<code>`, on the value cell of each setting.
- **Table anchors:** `<Sheet>_<Table>`, on a table's first header cell. No
  table is named `Status`, so an anchor never collides with a status line.
- **State:** `Run_Status` and the status line of each sheet that code writes.
- **Build stamp:** `build_commit`, `build_time` and `build_configuration`, on
  the reference sheet.
- The build creates every name. Before an operation, the adapter checks that
  each name it uses exists and refers to the expected sheet; a missing or
  moved name stops the operation before any calculation.

## Areas

| Area | Sheet | Kind | Written by | Holds |
| --- | --- | --- | --- | --- |
| *Illustrative:* `Inputs_Settings` | `shInputs` | Input | Users | One row per setting: code, value, unit and the contract that uses it |

- **Input** areas hold what users type: settings, parameters chosen by hand,
  and datasets the data contract places on a sheet. Each cell's type, unit
  and valid range come from the data contract or the calculation contract
  that reads it. A value that fails validation stops the operation; it is
  never corrected.
- **Parameter** areas hold values one operation writes and later operations
  read, with the operation and run that wrote them.
- **Result** areas hold the outputs of the calculation contracts, in the units
  and scopes they define. Every row names the scope and the run it belongs
  to.
- **Diagnostics** areas hold findings, the operation log and, in a development
  build, test results.

## Operations

| Entry point | Reads | Writes | Requires |
| --- | --- | --- | --- |
| *Illustrative:* `M_Run.RunCalculation` | `Inputs_Settings` and the datasets of `CALC-01` | Its result areas, findings and log | A valid reporting date |

- Each entry point is an Excel adapter procedure that calls the facade; it
  holds no regulatory formula.
- It captures and restores the Excel settings it changes, and reports both
  the operation and the cleanup outcome
  ([house style](../VBA_HOUSE_STYLE.md#operation-and-cleanup-outcomes)).
- A cancelled selection, such as a closed file dialog, changes nothing.

## Status and staleness

`Run_Status` and each `<Sheet>_Status` hold `CURRENT`, `STALE` or `FAILED`,
with the run ID and the time of the operation that set it.

| Area kind | After success | After failure |
| --- | --- | --- |
| Result | Replaced and `CURRENT` | Cleared and `FAILED` |
| Imported data and its lineage | Replaced, by an import only | Unchanged |
| Findings | Rewritten | Rewritten, so the findings that explain a failure are shown |
| Append only (parameters, log) | One row or set appended | Nothing appended, except one log row |

- A successful import marks every result that depends on it `STALE`. A
  rejected import keeps the previous data and still marks those results
  `STALE`, so no result from earlier data is shown as current, as the data
  contract requires.
- A result records a fingerprint of the inputs it used. Each operation
  compares it with the current inputs and marks the result `STALE` on any
  difference. Without worksheet event handlers, an edit made after a run is
  flagged at the next operation, not when it is typed; state which applies.

## Build

The build writes, from one commit:

- the sheets of [Sheets](#sheets), in order, with their tab and code names;
- every defined name, the formats, the data validation of the input areas and
  the print areas;
- the build stamp on the reference sheet: the commit, its committer time
  (not the clock, so that a rebuild is byte-identical) and the configuration.
  A working tree with changes is stamped as not evidence-grade;
- the VBA components of that commit, as the
  [installation guide](../../INSTALLATION.md) lists them.

It writes no data and no results. A production build has no `Test` sheet, test
module or example; a development build adds them.

`tools/devbuild/build_workbook.py` builds the sheets, code names, defined
names, stamp and VBA project. It reads them from
`tools/devbuild/workbook.json`, which must match the tables above. Formats,
data validation and print areas are not built yet: add them to the builder
when a decision here defines them. The
[installation guide](../../INSTALLATION.md#building-a-workbook-from-source)
says how to run it.

## Decisions

| Date | Decision | Reason | Issue |
| --- | --- | --- | --- |
| *Illustrative* | Whether a sheet is hidden | Why | Link |
