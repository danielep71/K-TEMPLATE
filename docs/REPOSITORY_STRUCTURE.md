# 🗂️ Repository structure

This guide says where each kind of file lives, which components may call
which, and how the supported VBA API is recorded. The
[house style](VBA_HOUSE_STYLE.md) covers how the code is written; the
[methodology guides](methodology/README.md) cover what it must calculate.

Paths marked *reserved* are conventions: create each directory with its first
real file, never as an empty placeholder. The Git, editor and ignore rules for
them are already set in `.gitattributes`, `.editorconfig` and `.gitignore`.

## Locations

| Location | Holds | In the template |
| --- | --- | --- |
| `.github/` | Workflows, issue forms, labels and Dependabot configuration | Present |
| `docs/` | Governance, setup, structure, style and the API inventory | Present |
| `docs/methodology/` | Sources, data contract, calculation contracts and validation plan | Present, as guides to fill in |
| `docs/methodology/sources/` | Byte-preserved copies of regulatory sources | Reserved |
| `src/core/` | Calculations and pure validation, host-independent standard modules | Reserved |
| `src/classes/` | Host-independent classes, if the design needs them | Reserved |
| `src/modules/` | The public facade: the supported entry points | Reserved |
| `src/workbook/` | Excel adapters, `ThisWorkbook` and sheet document modules, and the optional macro-free `Template.xlsx` | Reserved |
| `src/forms/` | UserForms, each `.frm` with its `.frx` | Reserved |
| `tests/` | Python tests of the repository tooling (`test_*.py`) | Present |
| `tests/modules/` | VBA test modules | Reserved |
| `tests/fixtures/`, `tests/expected/` | Synthetic inputs and expected results in the [test-case format](methodology/TEST_CASES.md) | Present, with one illustrative case of each form |
| `tests/derivations/` | Scripts that compute expected values from the documented method ([derivation scripts](methodology/TEST_CASES.md#derivation-scripts)) | Present, with one illustrative script for each case |
| `tools/` | Repository checks, setup and build tooling, in Python | Present |
| `examples/modules/` | Synthetic VBA examples that use only the supported API | Reserved |
| `test-results/` | Disposable output of local and CI runs | Ignored by Git |

- **Source first.** The exported source is authoritative. A change made in a
  workbook is exported and reviewed like any other change.
- **No built workbooks.** Workbooks are ignored, except
  `src/workbook/Template.xlsx`, a macro-free template, and the copies under
  `docs/methodology/sources/`. A built workbook is saved outside the checkout
  or under `test-results/`.
- **No unreviewed binaries.** **Repository integrity** runs
  `tools/check_binaries.py`, which fails on any tracked file that
  `.gitattributes` treats as binary, or that contains a NUL byte, outside those
  paths and the `.frx` beside its `.frm`. To keep another binary, such as an
  image for the documentation, review it and add its path, SHA-256 and reason
  to `.github/binary-allowlist.json`; a later change to the file fails until
  the entry is updated.
- **Forms.** A form's `.frx` stays beside its `.frm`.

## Dependencies

The Excel adapter calls the public facade. The facade validates its contract
and delegates to the core, which may use host-independent classes. Nothing
calls back the other way, so there are no cycles.

- Core and classes have no Excel object model, worksheet, range, user
  interface, file or network access, and no mutable run state.
- The facade orchestrates and validates; it does not repeat formulas. It
  returns structured results and errors.
- Workbook and form code maps cells and events to facade inputs and displays
  results; it contains no regulatory formula.
- Tests may call the core directly. Examples use only the supported facade.
- The workbook never calls Python. Reference values are derived from the
  methodology, not by translating the VBA
  ([validation plan](methodology/VALIDATION_PLAN.md#independent-reference-values)).
- Repository tooling is not a second implementation of the calculation.

## Public API inventory

[PUBLIC_API.txt](PUBLIC_API.txt) lists the supported declarations; in the
template it lists none. A `Public` declaration is supported only if it is
listed there. `Option Private Module` limits visibility between projects; it
is not a security boundary. Test and example modules are never part of a
production workbook.

Each supported declaration is recorded as two tab-separated lines:

```text
<component>	<kind>	<name>
# SIG	<component>	<kind>	<name>	<normalized declaration>
```

- **Kinds:** `Function`, `Sub`, `Property Get`, `Property Let`,
  `Property Set`, `Const`, `Enum` or `Type`.
- **Normalized declaration:** the declaration with its white space reduced to
  single spaces.
- Blank lines and other lines starting with `#` are ignored. Public members of
  a class follow the same format once the class is supported.

A change to the supported API needs an issue and an update, in the same pull
request, to:

- the inventory;
- the [calculation contract](methodology/CALCULATION_CONTRACTS.md);
- the tests.

There are no supported entry points until these are reviewed together with
the implementation.

## Input and output contracts

Before the first facade entry point exists, its inputs and outputs are defined
in the [data contract](methodology/DATA_CONTRACT.md) and the
[calculation contracts](methodology/CALCULATION_CONTRACTS.md). For every field
they give:

- the type, unit, currency and date basis;
- the valid range and sign convention;
- how a missing value is treated.

A contract also names:

- the approach and the version of the rules it implements;
- where the inputs come from, and the reporting and calculation dates;
- the validation failures and the result status;
- the output units and the intermediate values kept for audit.

It distinguishes invalid input, unsupported scope and numerical failure. It
never substitutes a default silently, turns missing values into zero, or
presents a partial result as complete.
