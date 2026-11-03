# 🧩 VBA house style

These conventions apply to every VBA component of the workbook. The
[repository structure](REPOSITORY_STRUCTURE.md) says where each kind of
component lives and which may call which; the
[installation guide](../INSTALLATION.md) says how source moves between the
repository and Excel. The template holds only the
[VBA test harness](methodology/TEST_HARNESS.md) and its example suite in
`tests/modules/`, so the examples below name no existing component.

## Naming and visibility

Replace `<REGIME>` with the project's short name, such as `SACCR` or `LCR`,
once, in this table. The
[static checks](#static-checks) read the facade prefix from this table; while
it is `<REGIME>_`, they accept any project prefix that is not a reserved one
(`CORE_`, `TEST_`, `EX_`, `M_`).

| Element | Convention | Example |
| --- | --- | --- |
| Core standard module | `CORE_` plus subject | `CORE_Validation` |
| Public facade module | `<REGIME>_` plus subject | `<REGIME>_Functions` |
| Host-independent class | `C` plus noun | `CCalculationResult` |
| Excel adapter module | `M_` plus subject | `M_Run` |
| Document module | `ThisWorkbook`, or `sh` plus role | `shInputs` |
| UserForm | `frm` plus role | `frmRunOptions` |
| Test module | `TEST_` plus subject | `TEST_Harness` |
| Example module | `EX_` plus subject | `EX_Inputs` |
| Procedure | PascalCase verb phrase | `ValidateInput` |
| Parameter or local variable | Descriptive PascalCase | `CalculationCurrency` |
| Constant, error constant | `UPPER_SNAKE_CASE`, `ERR_` prefix | `ERR_INVALID_INPUT` |
| Enum and member | `<REGIME>` plus noun; short member prefix | `<REGIME>RunStatus`, `rsFailed` |

- A component's name equals its exported `Attribute VB_Name` and its file
  name stem.
- VBA names are case-insensitive, and the VBE re-cases every use of a name to
  its latest spelling. Review case changes in a diff before accepting them.
- Single-letter names are for loop counters only.
- There are no mutable public variables.

Every component starts with `Option Explicit`. Core and test standard modules
also use `Option Private Module`. Procedures are `Private` unless another
component calls them. `Public` does not mean supported: only the declarations
listed in [PUBLIC_API.txt](PUBLIC_API.txt) are the supported API, and a change
to one updates the inventory, its contract and its tests together.

- Declare array bounds explicitly; never use `Option Base`.
- Prefer specific types to `Variant`. Use `Variant` only at a boundary that
  must carry missing, error or mixed values, and validate it there.
- Pass `ByVal` by default; document every `ByRef` argument a procedure
  changes.
- Core code has no Excel object model access and no mutable run state.

A native `Declare` needs its own compatibility review: `PtrSafe` alone does not
make a declaration correct. Use `LongPtr` for pointers and handles, and the
widths the API defines for every other argument; never replace `Long` with
`LongPtr` wholesale.

## Contracts, units and errors

Each procedure's header states, where they apply: `PURPOSE`, `INPUTS`,
`RETURNS`, `ERROR POLICY`, `DEPENDENCIES`, `STATE OWNERSHIP` and `REFERENCE`,
with an `UPDATED` date. `INPUTS` gives each argument's meaning, type, unit,
currency or date basis, valid range and ownership. `REFERENCE` cites the
[calculation contract](methodology/CALCULATION_CONTRACTS.md) and the source it
implements. Comments explain invariants and reasons, not obvious assignments.

- Use `Double` for amounts and proportions, `Long` for counts, `Boolean` for
  flags, `String` for identifiers and currency codes, and `Date` for dates on
  the host side. Where a date crosses a boundary as text, state its format.
- Names or contracts distinguish decimal rates from basis points, and amounts
  in currency units from scaled display amounts.
- Never invent a rounding rule, day count, regulatory parameter or valid
  range: record it in the calculation contract first.
- Never turn missing input into zero silently, and use no undocumented
  default.

Errors raised by the project use named `ERR_` constants, allocated once in one
documented range: `vbObjectError + 2048` through `vbObjectError + 4063`. The
range `vbObjectError + 4064` through `vbObjectError + 4095` is reserved for the
test harness and its example suite (`TEST_Harness`, `TEST_Examples`).
`Err.Source` names the component and procedure. A message describes the
condition without repeating confidential input. Error codes distinguish
invalid input, unsupported scope, numerical failure, and host or cleanup
failure.

Core procedures raise documented errors. Facade entry points translate them as
their contract says: a macro returns or reports a structured failure; a
worksheet function may return a documented `CVErr` value in a `Variant`.
Worksheet functions change no Excel state and show no dialog. A facade never
turns an unexpected failure into an amount that looks valid.

`On Error Resume Next` guards one operation only. Check and clear `Err`
immediately after it, then restore the intended handler. Never use it around
a whole calculation or cleanup sequence.

## Operation and cleanup outcomes

An operation that changes host state reports both outcomes to its caller:

| Field | Meaning |
| --- | --- |
| `OperationSucceeded` | Whether the intended work completed |
| `PrimaryError` | The saved error number, source and description, or none |
| `CleanupSucceeded` | Whether every required restoration succeeded |
| `CleanupErrors` | Every restoration that failed, with its error |
| `RecoveryRequired` | True when host state could not be restored |

A successful operation whose cleanup failed is not a success, and a failed
operation stays failed even if its cleanup succeeded. Callers check both.
These are required fields of the result, not an existing type.

- Capture each setting before changing it, and record which settings were
  captured and which were changed. Register the restoration before attempting
  the change, so that a setter that fails half way is still restored.
- Never restore a value that was not captured. Restore only what the
  operation tried to change, to the captured value; a nested operation keeps
  its caller's state.
- Normal and error paths both enter cleanup. Save `Err.Number`, `Err.Source`
  and `Err.Description` into variables before any other call.
- Restore each setting through its own helper and handler, and attempt every
  restoration even when one fails. Cleanup never replaces the saved primary
  error.
- Return both outcomes, or publish the cleanup failures through a documented
  output before raising the saved primary error again. If only cleanup fails,
  report a cleanup failure, not success.
- Never reset `Calculation`, `EnableEvents`, `ScreenUpdating`, `DisplayAlerts`
  or `StatusBar` to an assumed default. Capture only what the operation
  changes. Sheets and other objects an operation creates need the same
  explicit cleanup.

Failure-path tests cover partial capture, a failed change, a primary failure,
each failed restoration, and a primary and cleanup failure together.

## Source representation

- Exported `.bas`, `.cls` and `.frm` files are ASCII only, without a byte
  order mark. Represent other characters at run time, for example with
  `ChrW`, or keep them in a separately specified resource, so the source does
  not depend on a Windows code page.
- A form's binary `.frx` stays beside its `.frm`.
- Git stores VBA source with LF line endings and checks it out with CRLF
  (`.gitattributes`). Keep the `VERSION`, `BEGIN`, `END` and `Attribute` lines
  that the VBE writes.
- Indent with four spaces; no tabs and no trailing white space. Keep
  procedure declarations, `End` statements and labels at their outer level.
  Separate procedures with two blank lines.
- Use `MODULE SETTINGS`, `DECLARE`, `INITIALIZE` and `FAIL` sections where they
  make a routine easier to follow, not as empty ceremony.
- A generated component names its generator and is regenerated, never edited
  by hand.

`python tools/format_vba.py --write` applies the mechanical part of these
rules to the `.bas`, `.cls` and `.frm` files under `src/`, `tests/modules/` and
`examples/modules/`:

- CRLF line endings and one final line ending;
- no trailing white space, and four spaces for each indentation tab;
- procedure declarations, their `End` statements and labels at the outer
  level;
- two blank lines after each procedure's `End` statement.

It never changes a string literal, and it keeps the `VERSION`, `BEGIN`, `END`,
`Object` and `Attribute` lines that the VBE writes. **Repository integrity**
runs it with `--check`. Re-import a module it rewrote into the VBE before
committing, so that the workbook and the source stay the same. It neither
compiles nor checks the other rules of this guide.

## Static checks

**Repository integrity** runs six static checks on every tracked `.bas`,
`.cls` and `.frm` file. Each also runs on its own, for example
`python tools/check_vba_source.py`; with no VBA source, each passes and says
it examined no component.

| Check | Enforces |
| --- | --- |
| `tools/check_vba_source.py` | Each component in its home with its role's prefix and a `VB_Name` equal to its file name, names unique ignoring case and at most 31 characters; ASCII only; `Option Explicit`, and `Option Private Module` in core and test modules; no `Option Base`, tab or trailing white space; closed string literals; each form beside its `.frx`; LF storage with CRLF checkout; no workbook committed except `src/workbook/Template.xlsx` and the regulatory source copies under `docs/methodology/sources/` |
| `tools/check_vba_conditionals.py` | Well-formed `#If`, `#ElseIf`, `#Else` and `#End If` blocks, evaluated in each compilation environment (VBA6 Win32, VBA7 Win32 and VBA7 Win64); every `Declare` reachable under VBA7 is `PtrSafe` |
| `tools/check_vba_jumps.py` | Every `GoTo`, `GoSub`, `Resume` and `On Error GoTo` target is a label of the same procedure, in each environment; no label is defined twice in a procedure |
| `tools/check_vba_public_api.py` | Explicit visibility on every declaration; the public declarations of `src/modules/` match [PUBLIC_API.txt](PUBLIC_API.txt) one to one, with normalized signatures; no public variable or public `Declare` in a standard module; no public name in two standard modules, which Excel reports as an ambiguous name |
| `tools/check_vba_dependencies.py` | The [dependency direction](REPOSITORY_STRUCTURE.md#dependencies): core and classes have no Excel object model, file, registry or process access, dynamic call or reference to a component above them; core has no module-level or `Static` state; examples use only the facade |
| `tools/check_vba_identifiers.py` | In each environment: every name a procedure assigns (`=`, `Let`, `Set`, `LSet`, `RSet`, `For`, `For Each`) is declared, by the procedure, the module, a form control or a standard module's public name, and `ReDim` declares as in VBA; every statement-level call (`Name args`, `Call Name(args)`) names a procedure or `Declare` of the module or a public one of a standard module, or a documented built-in; no procedure, parameter, variable, constant, Enum, Enum member, Type, `Declare` or `Event` is named like a VBA keyword such as `Empty`, `Date`, `Error` or `Name`; a statement it cannot read is a finding |

What they cannot prove:

- they neither compile nor run VBA, and resolve no call: a name-based rule
  can miss a call made through a variable or an object;
- they check the three compilation environments statically, not on a 32-bit
  or 64-bit host;
- a private procedure name may repeat across modules, as VBA allows; only a
  public name in two standard modules is a collision;
- the identifier check reads neither member access (`x.y = 1`, `.y` in a
  `With` block) nor qualified calls (`Module.Name`, `obj.Method`), nor calls
  inside an expression (`x = Helper(1)`), and it does not see the controls of
  a worksheet, which a document module may name. Its built-in callees, such as
  `MsgBox` and `DoEvents`, are listed in `BUILTIN_CALLS`, a form's or document
  module's own methods, such as `Hide` or `Protect`, in `HOST_CALLS`, and the
  host properties assignable without a declaration in `HOST_TARGETS`:
  `Range`, `Cells` and `Selection` anywhere, a UserForm's `Caption` or
  `Width` in a form, a workbook's `Saved` or a worksheet's `Visible` in a
  document module. Extend the list when code calls or assigns a built-in it
  does not know. A procedure of another VBA project that this one references
  is not seen either; call it qualified by its module;
- the formatter and these checks enforce the mechanical rules only; review
  covers the rest of this guide.

Changed VBA is imported, compiled and run on the supported host before its
behaviour is accepted. Repository checks are not compilation. Record the full
commit, the Excel build and bitness, the references and the results, as the
[validation plan](methodology/VALIDATION_PLAN.md#evidence-record) describes.
