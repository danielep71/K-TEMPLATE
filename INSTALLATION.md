# 📥 Installation: VBA import and export

This guide moves VBA source between the repository and Excel, in both
directions, without losing line endings, metadata or form resources. It also
covers building a workbook from source, running the VBA test harness,
recording an Excel run, upgrading to a new revision, troubleshooting and
removal. The
[repository structure](docs/REPOSITORY_STRUCTURE.md) says where each component
lives, and the [house style](docs/VBA_HOUSE_STYLE.md) says how it is written.
The repository's own checks are in the
[contributing guide](CONTRIBUTING.md#-local-checks).

## 🖥️ Host policy

Adapt this section to the project and record the decision in an issue.

- **Target host:** Microsoft 365 Excel desktop on Windows, 64-bit. Windows
  32-bit, older perpetual versions, Mac and Excel for the web are not
  supported until a recorded decision and evidence add them.
- **References:** only the defaults of the target installation: Visual Basic
  For Applications, the Microsoft Excel 16.0 Object Library, OLE Automation
  and the Microsoft Office 16.0 Object Library. Any other reference needs a
  reviewed decision, and a change to the host references of
  `.github/excel-evidence-policy.json`. Before compiling, check **Tools → References**
  for entries marked *MISSING*.
- **Python** is used for repository tooling and for deriving reference values,
  never by the workbook.
- **Macro security:** follow your organization's policy. Nothing here needs
  a lower security level, or *Trust access to the VBA project object model*.

Record each host on which a commit has been compiled and run:

| Tested environment | Commit | Evidence |
| --- | --- | --- |
| Microsoft 365 Excel 16.0 build 20430, 64-bit; Windows 11 Enterprise 25H2, build 26200.9711; `tools/devbuild/excel_run.py`, `compile=ON_DEMAND` | `50c065859c7f566ef37946cca9c6206447951145` | #204: `excel-run: PASS`, `Binding: VERIFIED`, ACC-01 worst error 0.0 |

## 🔖 Start from an exact revision

```sh
git clone https://github.com/OWNER/REPOSITORY.git
cd REPOSITORY
git switch BRANCH
git rev-parse HEAD
git status --short
```

Record the full SHA, and start only when `git status --short` prints nothing.

Import from a Git checkout. `.gitattributes` checks VBA source out with CRLF
line endings, which the VBE needs. A raw download or a source archive holds
the stored LF files and skips that conversion. Never mix files from two
revisions.

## ⬇️ Importing VBA into Excel

1. **Workbook.** Create a blank workbook, or copy `src/workbook/Template.xlsx`
   if the project has one. Save it as `.xlsm` outside the checkout or under
   `test-results/`, which Git ignores.
2. **Project.** Open the VBE (**Alt+F11**) and select that workbook's
   project explicitly, never whichever one is active. Check the references.
3. **Modules and classes.** Import the `.bas` and `.cls` files of
   `src/core/`, `src/classes/` and `src/modules/`, then the standard `.bas`
   adapters of `src/workbook/`.
   - Import in file name order, so that every import is the same.
   - Skip a directory that does not exist.
   - Start from a fresh workbook, so that no obsolete or duplicate component
     remains.
4. **Forms.** Import each `.frm` of `src/forms/` with its `.frx` beside it;
   never import a `.frx` on its own.
5. **Document modules.** Do not import the `.cls` file of `ThisWorkbook` or a
   sheet as an ordinary class.
   - Find the existing document component with the same code name. Create a
     sheet only where the build instructions call for one.
   - Paste the file's code into that component, without its `VERSION`,
     `BEGIN` and `END` header and without its `Attribute` lines.
   - Hidden attributes need a reviewed round trip before such a component is
     accepted; never drop them silently.
6. **Tests.** A development workbook also imports `tests/modules/` and,
   when the project has it, `examples/modules/`, the reserved location of
   synthetic VBA examples. A production workbook never contains tests or
   examples.
7. **Compile and record.** Compile only once every component is present:
   **Debug → Compile VBAProject**. Then save the workbook and run the
   documented tests. Record the full SHA, the Excel build and bitness, the
   Windows version, the references, and what you ran with its output, as the
   [validation plan](docs/methodology/VALIDATION_PLAN.md#evidence-record)
   describes.

Manual import needs no *Trust access to the VBA project object model*.

## ⬆️ Exporting changes

1. **Export.** Export each changed component into a temporary folder first,
   keeping the VBE header and `Attribute` lines. Check that each file name
   stem equals its `Attribute VB_Name`.
2. **Document modules.** Export the component as `.cls` and compare it with
   the matching file in `src/workbook/`. Review differences in the code name
   and in hidden attributes, then update the file. Never replace the
   component by importing the `.cls`.
3. **Forms.** Review a changed `.frm` and its `.frx` together and copy them
   as a pair. Never edit a `.frx` by hand.
4. **Compare.** Compare the exports with the exact checkout. Reject non-ASCII
   source, unexpected changes in letter case and unexpected metadata changes.
   Then copy the intended changes into their source paths; Git normalizes the
   line endings when you add them.
5. **Check.** Run the [local checks](CONTRIBUTING.md#-local-checks) before
   committing. They check the repository, not VBA compilation or the results
   of a calculation.

<a id="building-a-workbook-from-source"></a>

## 🏗️ Building a workbook from source

`tools/devbuild/build_workbook.py` builds the workbook of steps 1 to 6 above,
apart from forms (step 4), in one command, from the checkout:

```sh
python tools/devbuild/build_workbook.py                            # development
python tools/devbuild/build_workbook.py --configuration production
```

It writes `test-results/build/Workbook-<configuration>.xlsm`, which Git
ignores, with:

- the sheets, code names and defined names of the
  [workbook contract](docs/methodology/WORKBOOK_CONTRACT.md), read from
  `tools/devbuild/workbook.json`;
- the layout of `src/workbook/Template.xlsx`, when the project has one: the
  build starts from a copy of it and keeps its formats and protection
  ([build](docs/methodology/WORKBOOK_CONTRACT.md#build)). Without it, the
  sheets are blank;
- the build stamp on the reference sheet: the commit, its committer time and
  the configuration. A tracked change, or an untracked file in a source
  directory, adds *NOT evidence-grade* to the commit;
- a VBA project written from source. The production configuration takes the
  standard and class modules of `src/`. The development configuration adds
  `tests/modules/` and `examples/modules/`. Each sheet's document module
  comes from `src/workbook/<code name>.cls`, when that file exists.

Excel compiles the project from source when the workbook opens
([HB-05](docs/methodology/VALIDATION_PLAN.md#host-behaviour)), so the
build needs no *Trust access to the VBA project object model*. Two builds of
the same commit and working tree are byte-identical. The build stops on a form,
which needs a designer the builder does not write; import forms by hand
(step 4). It also stops on a module too dense to fit the VBA compression
format; import that module by hand.

1. Unblock the file if it was copied from another machine, open it, enable
   macros, and check **Tools → References**.
2. Compile with **Debug → Compile VBAProject**.
3. Changes made in the VBE are not read back: export them as
   [exporting changes](#%EF%B8%8F-exporting-changes) describes, then build
   again.

If Excel rejects or strips the embedded project, import by hand as in
[importing VBA into Excel](#%EF%B8%8F-importing-vba-into-excel), and record
the failure in an issue.

**Early warning in LibreOffice (optional).** Where LibreOffice and its
Python UNO bridge are installed, `python3 tools/devbuild/lo_smoke.py` builds
the development workbook and runs the harness certification in headless
LibreOffice. It reports only failures that LibreOffice is not known to
cause. LibreOffice is not a supported host, and this is never Excel
evidence.

<a id="running-the-harness"></a>

## 🧪 Running the harness

In a **development** workbook built as above, with the modules of
`tests/modules/`, open the Immediate window (**Ctrl+G**) and type, with the
full path of `tests\fixtures` in your checkout and no trailing backslash:

```text
TEST_Examples.RunExampleTests "C:\path\to\checkout\tests\fixtures"
```

It runs the example suite through `TEST_Harness` and writes one record per
line. One case reads the illustrative test cases from that folder and
records their totals as results. The two entry points below take the same
argument. A passing run ends with:

<!-- BEGIN generated: example-suite-results -->
```text
RESULT=PASS; suite=examples; operation=PASS; completeness=COMPLETE; cases=14; assertions=34; failures=0; cleanup=PASS; recovery=NONE
```

Anything else, or the error *Suite examples did not pass*, is a failure.
Two more entry points show the failure paths and are expected to fail:

| Entry point | Expected end |
| --- | --- |
| `TEST_Examples.RunExampleTestsWithInjectedFailure` | `RESULT=FAIL`, `operation=FAIL`, `completeness=COMPLETE`, `failures=1`, then *Suite examples did not pass* |
| `TEST_Examples.RunExampleTestsWithSkippedCase` | `RESULT=FAIL`, `operation=PASS`, `completeness=INCOMPLETE`, `cases=13`, `assertions=32`, `failures=1`, then *Suite examples did not pass* |
<!-- END generated: example-suite-results -->

- **Complete log.** The Immediate window keeps only its last 200 lines. To
  keep the whole run, type `TEST_Harness.SaveLog "<folder>\examples.log"`
  after it, with a folder outside the checkout or the ignored
  `test-results/` directory.
- **Recovery.** A run that stopped half way, or that left Excel's state
  changed, leaves `recovery=REQUIRED`, and the next run is refused. Check
  **File → Options → Formulas** (calculation), the open workbooks and the
  other settings the log names, then type `TEST_Harness.ResetTests`.
- **End or Reset.** **End** in an error dialog, or **Run → Reset** in the
  VBE, clears the harness's state with the rest of the project's, so the
  stop is not detected: the next run starts normally and takes Excel's state
  as it finds it as its baseline. After such a stop, check Excel by hand
  before the next run. In the Immediate window,
  `? Application.Calculation, Application.ScreenUpdating, Application.DisplayAlerts, Application.EnableEvents`
  prints the four settings the harness compares (`-4105` is automatic
  calculation), and the open workbooks and the active window must be those
  you had before the run. Put back anything the stopped run changed.
- **What it proves.** The suite exercises the harness itself and validates no
  calculation of the workbook. The [harness guide](docs/methodology/TEST_HARNESS.md)
  defines the outcomes and the log format.

<a id="recording-an-excel-run"></a>

## 🧾 Recording an Excel run

This is the evidence that a commit's VBA was imported, compiled and run on the
supported host, for a release and for any issue or pull request that cites
an Excel run. It needs no self-hosted runner, no change to the
macro-security level and no *Trust access to the VBA project object model*.

### One command

On Windows with Excel, Python 3.11 or later and Git, in a clone with no local
changes (`git status --short` prints nothing), at the candidate commit:

```sh
python tools/devbuild/excel_run.py
```

It builds the development workbook, and stops if the build is not
evidence-grade. Then it:

1. empties `test-results\excel-run\`;
2. opens the workbook read-only in Excel. If Excel shows the yellow bar
   *SECURITY WARNING Macros have been disabled*, click **Enable Content**
   (or **File → Info → Enable Content**). This happens where a policy makes
   programs follow your macro setting
   ([HB-02](docs/methodology/VALIDATION_PLAN.md#host-behaviour)). The command
   changes no macro setting, and waits up to `--enable-timeout` seconds (300
   by default) for that click, which enables the macros without reopening
   the workbook ([HB-04](docs/methodology/VALIDATION_PLAN.md#host-behaviour));
3. runs the certification of step 5 below, with that folder and
   `tests\fixtures`;
4. closes the workbook without saving, and quits Excel;
5. writes `host.txt`, the
   [host record](docs/methodology/TEST_HARNESS.md#host-record): you as the
   operator (your Windows user name, or `--operator "<name>"`), the start and
   finish times, Windows from the registry, the Excel build and bitness, the
   `AutomationSecurity` and the Trust access it found, which it reads and
   never changes, the four references the build gives the project,
   `compile=ON_DEMAND` and the outcome;
6. runs the check of step 9 with `--workbook` and `--require-binding`.

It prints the report and saves it as `test-results\excel-run-report.txt`.
It must end with `Binding: VERIFIED` and `excel-run: PASS`; record it as
step 10 says.

There is no export: the check proves identity by rebuilding the commit in a
temporary worktree. The workbook Excel ran must hold the same parts as that
build, byte for byte once decompressed, the VBA project included. Excel
compiles the project from source on open
([HB-05](docs/methodology/VALIDATION_PLAN.md#host-behaviour)), so the code
that ran is the commit's.

Without Trust access no automated run can compile the whole project. VBA
then compiles on demand only the modules the run calls, and `compile=ON_DEMAND`
says so; a module the run never calls is not compiled. `check_excel_run.py`
accepts `ON_DEMAND` only with identity by rebuild, `--workbook` and no
`exported\` folder, as this command checks it. A run by hand compiles with
**Debug → Compile VBAProject**, which needs no Trust access, and records
`compile=PASS`.

Excel opens visibly while the run lasts. Apart from **Enable Content**, leave
it alone: a dialog it shows stops the run until you close it, and the command
gives up after `--timeout` seconds (600 by default). A run that stops early
still writes `host.txt`, with the outcome `NOT_ENABLED` when the macros were
not enabled in time, `EXECUTION_TIMEOUT` after `--timeout`, `UNAVAILABLE`
when Excel could not be started or `INCOMPLETE` when the run stopped before
its end. That record is evidence of how the run stopped, never of a pass;
run the command again.

Use the steps below instead when the project has forms, when the build
fails, or when you need a full compile. A workbook you import by hand instead
of building it has no build stamp, so the check fails while the commit holds
accuracy contracts, as step 3 says.

### By hand

1. **Fix the candidate.** In a clone with no local changes
   (`git status --short` prints nothing), type `git switch --detach <commit>`,
   then `git rev-parse HEAD` for the full SHA.
2. **Make the evidence folder**, for example `test-results\excel-run\` in the
   checkout (ignored by Git), with an empty `exported\` folder inside.
3. **Build a development workbook** from that checkout with
   `python tools/devbuild/build_workbook.py`, as in
   [building a workbook from source](#building-a-workbook-from-source). Use
   [importing VBA into Excel](#%EF%B8%8F-importing-vba-into-excel) only if the
   project has forms or the build fails. Then only the source identity is
   checked, not the binding of steps 7 and 9. Such a workbook has no
   `build_commit` stamp, so `SaveResults` writes an empty commit, which the
   [accuracy check](docs/methodology/TEST_HARNESS.md#results) rejects: while
   the commit holds `tests/accuracy/contracts.json`, as the template does, the
   run of a workbook imported by hand ends `excel-run: FAIL` with *commit is
   not a full lower-case commit SHA*, whatever else it shows.
4. **Compile** with **Debug → Compile VBAProject**.
5. **Certify.** In the Immediate window type, with the full path of the
   evidence folder and of `tests\fixtures` in the same checkout, both with no
   trailing backslash:

   ```text
   TEST_Examples.CertifyExamples "C:\path\to\test-results\excel-run", "C:\path\to\checkout\tests\fixtures"
   ```

   It refuses to start if an earlier run left state behind. It runs the
   suite normally, with the injected failure and with the skipped case, and
   saves each complete log as `examples.log`, `injected.log` and
   `skipped.log`. After the normal run it saves `results.json`: the
   results that run recorded, with the commit of the workbook's
   `build_commit` stamp. Then it writes `certification.log`. Each run prints a
   `SAVED=` line: the normal run with `error=0`, the two failure runs with
   the error *Suite examples did not pass*. Those errors are recorded, not
   raised, so that every log is written. The certification ends with
   `CERTIFICATION=PASS`; otherwise it raises *Certification of examples did
   not pass* after writing its log. Then type the certification entry point
   of each project suite that `.github/excel-evidence-policy.json` lists under
   `suites`, in order, with the same two arguments; each saves its log, its
   certification log and, when its entry names one, its results file.
6. **Export the modules back.** In the VBE, right-click each component →
   **Export File**, into `exported\`, under its own name. Export every VBA
   component of the commit, the test modules included; a form exports as its
   `.frm` with its `.frx`. Do not export a document module that the build
   generated and that has no file in `src/workbook/`. Do not edit the
   exported files.
7. **Bind the export to the workbook.** Close the workbook, then record the
   SHA-256 of the workbook and of every exported file:

   ```sh
   python tools/devbuild/export_manifest.py test-results/build/Workbook-development.xlsm test-results/excel-run
   ```

   It writes `export-manifest.json` in the folder. Do not open or save the
   workbook again before step 9.
8. **Write `host.txt`** in the folder:

   ```text
   operator=<your name>
   started=2026-11-03T14:05:09+01:00
   finished=2026-11-03T14:21:40+01:00
   windows=Windows 11 Pro, version 24H2, OS build 26100.4061
   excel=16.0 build 17928
   bitness=64-bit
   automation_security=Low
   vba_project_access=NOT_TRUSTED
   references=Visual Basic For Applications; Microsoft Excel 16.0 Object Library; OLE Automation; Microsoft Office 16.0 Object Library
   compile=PASS
   outcome=PASS
   ```

   The [host record](docs/methodology/TEST_HARNESS.md#host-record) defines
   each field:
   - `started` and `finished` are when you began step 4 and ended step 6,
     with your UTC offset.
   - `windows` is what `winver` shows.
   - `excel` is what `? Application.Version & " build " & Application.Build`
     prints in the Immediate window, and `bitness` what **File → Account →
     About Excel** shows.
   - `automation_security` is what `? Application.AutomationSecurity` prints:
     1 is `Low`, 2 `ByUI`, 3 `ForceDisable`. Only read it; never change it.
   - `vba_project_access` is `TRUSTED` when **File → Options → Trust Center →
     Trust Center Settings → Macro Settings** ticks *Trust access to the VBA
     project object model*, `NOT_TRUSTED` otherwise. Change nothing there.
   - `references` lists the boxes ticked in **Tools → References**, separated
     by `; `.
   - Write `compile=PASS` only if step 4 completed with no error; otherwise
     write what it reported. Never write `ON_DEMAND`: the check accepts it
     only with identity by rebuild, never with `exported\`.
   - `outcome` is `PASS` when step 5 ended with `CERTIFICATION=PASS` for
     every suite, and `FAIL` when a certification raised. When you stop
     early, write the record anyway, with `IMPORT_FAILED` (step 3),
     `COMPILE_FAILED` (step 4), `NOT_ENABLED`, `INCOMPLETE` or `UNAVAILABLE`
     (Excel would not start), and `UNKNOWN` for a value you could not read.
9. **Check** the folder against the commit, with Python 3.11 or later and
   Git:

   ```sh
   python tools/check_excel_run.py test-results/excel-run --commit <commit> \
     --workbook test-results/build/Workbook-development.xlsm --require-binding
   ```

   It must end with `Binding: VERIFIED` and `excel-run: PASS`. It prints
   the commit, the host and the Git blob id of every imported component. When
   the commit holds accuracy contracts, an `Accuracy:` line gives each
   contract's worst error over its calibration and its holdout cases. For
   a workbook imported by hand, omit `--workbook` and `--require-binding`:
   the binding is then `UNVERIFIED`, which a release does not accept, and
   while the commit holds accuracy contracts the run fails, because its
   results name no commit (step 3).
10. **Record.** Copy the report into the issue, the pull request or the
    [release certification](docs/release/RELEASE_CERTIFICATION.md), and add a
    row to the tested environments of the
    [host policy](#%EF%B8%8F-host-policy). The folder stays outside Git: its
    exported modules would be second copies of the VBA components.

`check_excel_run.py` reads only the folder and the commit:

| It fails when | Because |
| --- | --- |
| An exported module differs from the commit's file line for line, a form's `.frx` differs byte for byte, or one is missing or extra. Code outside strings and comments is compared without letter case, because the VBE re-cases a name ([HB-01](docs/methodology/VALIDATION_PLAN.md#host-behaviour)): each re-cased identifier is reported on a `Note:` line instead. In a document module of `src/workbook/`, the attributes Excel writes itself (`VB_Base`, `VB_GlobalNameSpace`, `VB_Creatable`, `VB_PredeclaredId`, `VB_Exposed`, `VB_TemplateDerived`, `VB_Customizable`) are not compared; `VB_Name` is | The workbook was not built from that commit |
| A log is missing or empty, does not end with `RESULT`, holds more than one run, or a line is not a `KEY=value` record | The log is incomplete |
| A record is missing, extra or out of the order in which the commit's suite writes it, or an unexpected error (`ERROR`) is logged | The run did not execute the suite as written |
| The normal run is `INCOMPLETE` | A mandatory case or assertion did not run |
| The normal run is not `RESULT=PASS`, or its cases and assertions differ from the commit's suite in name or order | The run did not pass as written |
| `CLEANUP` is not `PASS` or `RECOVERY` is not `NONE` in any of the three logs | Cleanup failed |
| The two failure runs do not end as [documented](#running-the-harness), or fail anything but their own failure: the injected run its one designated assertion, with the expected value the suite injects and the same values in each record, the skipped-case run only `harness.completeness` | The failure paths did not behave |
| `certification.log` is missing, a mode did not end as it must, a failure run did not reach its caller as `ERR_TEST_FAILED`, or it does not end with `CERTIFICATION=PASS` and `COMPLETE` | The certification did not pass |
| `ENVIRONMENT` lacks a field, differs between the logs, or Office is not 64-bit | The host is unknown or not the supported target |
| `host.txt` lacks a field or holds a value the [host record](docs/methodology/TEST_HARNESS.md#host-record) does not allow, a time lacks its UTC offset, the run finished before it started, the references are not exactly the four of the [host policy](#%EF%B8%8F-host-policy), `compile` is not `PASS` (or `ON_DEMAND` with identity by rebuild: `--workbook` and no `exported\` folder, as the one command records it), `operator` is `UNKNOWN`, or `excel` or `bitness` differs from the logs' `ENVIRONMENT` | The host record is incomplete |
| `outcome` in `host.txt` is not `PASS` | The run stopped early or failed: the record is evidence of that outcome, never of a pass |
| Without `exported\`, a part of the workbook given with `--workbook` differs from a rebuild of the commit, or a part is missing or extra | It was saved after the build, or not built from that commit (`Binding: FAIL`) |
| `export-manifest.json` does not match the exported files or the workbook, or the workbook's `build_commit` is not the commit of a clean build | The export is not bound to a workbook built from the commit (`Binding: FAIL`) |
| With `--require-binding`, no manifest or no `--workbook` | The binding is `UNVERIFIED` |
| The commit holds `tests/accuracy/contracts.json`, and `results.json` or the results file a project suite names is missing or names another commit, a build with local changes or, for a workbook imported by hand, no commit; or two files hold the same result, the files together lack a result, a result lies outside its output's tolerance, or a contract's worst error exceeds its threshold | The run did not measure the contracts, or they do not hold ([accuracy contracts](docs/methodology/TEST_CASES.md#accuracy-contracts)) |

It cannot see Excel: the host record, from the operator to the outcome, is
the runner's own record. A PASS says the evidence is consistent
with the commit. It validates a calculation only as far as its accuracy
contracts and their reference values do; the template's illustrative
contract validates none.

## 🔄 Upgrading to a new revision

A workbook is never patched in place. To move to a new revision, build a new
workbook from it and carry your inputs across:

1. **Read the changes.** Read the changelog between the two revisions, and the
   `behavior-change` entries first. They name changed results, input layouts
   and output layouts, so you know which differences to expect.
2. **Keep the old workbook.** Keep it unchanged and record its commit, so you
   can compare results.
3. **Build the new one.** Import every component from the new revision into a
   fresh workbook, as [Importing VBA into Excel](#%EF%B8%8F-importing-vba-into-excel)
   describes. Never import some components from one revision and some from
   another, and never paste new source beside old components.
4. **Carry the inputs across.**
   - Re-import the input files the [data contract](docs/methodology/DATA_CONTRACT.md)
     defines.
   - Copy the values of each input area that the
     [workbook contract](docs/methodology/WORKBOOK_CONTRACT.md) lists, area by
     area, by its defined name. A sheet, area or name that the new revision
     renamed or removed is listed in the changelog. Never copy results:
     recalculate them.
5. **Compile and compare.** Compile, run the documented tests, and compare the
   results with the old workbook. Explain every difference with a changelog
   entry, or treat it as a defect. Record the run as for a new installation.

A copy with local changes is a fork. Export its changes first, compare them
with both revisions, and reapply them deliberately; never assume they survive
a rebuild.

## 🩺 Troubleshooting

| Symptom | Check |
| --- | --- |
| A compile error or a missing procedure | Every component comes from one revision and was imported; the references match the [host policy](#%EF%B8%8F-host-policy) |
| *Ambiguous name detected* | An obsolete or duplicate component remains. Start again from a fresh workbook |
| A reference marked *MISSING* | The workbook was built on another host. Remove it, or record a decision for it |
| A form without its controls | Its `.frm` was imported without the matching `.frx` beside it |
| Source with LF line endings, or one long line | It came from a raw download or an archive instead of a [Git checkout](#-start-from-an-exact-revision) |
| A document module duplicated as a class | Its `.cls` was imported instead of pasted ([step 5](#%EF%B8%8F-importing-vba-into-excel)) |
| Excel left with screen updating, events or calculation changed after an error | Read the operation's cleanup outcome ([house style](docs/VBA_HOUSE_STYLE.md#operation-and-cleanup-outcomes)); never reset settings to assumed defaults |
| Results marked `STALE` or `FAILED` | Read the findings and the log; the [workbook contract](docs/methodology/WORKBOOK_CONTRACT.md#status-and-staleness) says what each status means |
| A result differs from its expected value | Check the commit, the inputs and the tolerance of the [test case](docs/methodology/TEST_CASES.md) before changing code |

If Excel's state is uncertain, save your data separately, close Excel, and
reproduce the problem in a fresh session with synthetic data before changing
code.

## 🗑️ Removing the workbook's components

1. Run any documented cleanup, so that the workbook leaves Excel's settings as
   it found them.
2. Remove every imported module, class and form, and any test or example
   module.
3. Delete the code pasted into `ThisWorkbook` and each sheet's document module
   ([step 5](#%EF%B8%8F-importing-vba-into-excel)), keeping the components
   themselves. Otherwise their event handlers stay active and call
   procedures that no longer exist.
4. Compile what remains, then reopen the workbook and check that it opens with
   no error.

Removing components does not remove:

- data on worksheets, defined names or sheets the build created;
- references added by hand;
- trusted locations, macro settings or anything else outside the workbook.

Remove only what this workbook created, and record anything you keep.

## 🧹 Git hygiene

These commands check the line-ending, binary and ignore rules on probe paths.
The paths are probes; the files need not exist.

```sh
git check-attr text eol diff merge -- src/core/CORE_Example.bas src/forms/frmExample.frx
git check-ignore -v --no-index test-results/run.log sample.xlsm
```

Expected results:

- `.bas` is `text` with `eol: crlf` and `diff: vba`;
- `.frx` is binary, with `text` unset;
- both probe files are reported as ignored.

To check that the stored files follow the line-ending rules, start from a
clean checkout and run:

```sh
git add --renormalize .
git diff --cached --name-only
```

The second command prints nothing when every file follows the rules. A file it
lists would be normalized: inspect it, then either commit the normalization as
a change of its own or undo the staging with `git restore --staged .`. Before
any commit, stage only the intended changes and run `git diff --cached --check`.

Ignore rules do not remove files that Git already tracks, and they are not a
secret scanner.

To show procedure names in the hunk headers of VBA diffs, configure each
clone once with the `diff.vba.xfuncname` setting given in `.gitattributes`.
