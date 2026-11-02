# 📥 Installation: VBA import and export

This guide moves VBA source between the repository and Excel, in both
directions, without losing line endings, metadata or form resources. The
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
  For Applications, the Microsoft Excel Object Library, OLE Automation and the
  Microsoft Office Object Library. Any other reference needs a reviewed
  decision. Before compiling, check **Tools → References** for entries marked
  *MISSING*.
- **Python** is used for repository tooling and for deriving reference values,
  never by the workbook.
- **Macro security:** follow your organization's policy. Nothing here needs
  a lower security level, or *Trust access to the VBA project object model*.

Record each host on which a commit has been compiled and run:

| Tested environment | Commit | Evidence |
| --- | --- | --- |
| None yet | — | — |

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
6. **Tests.** A development workbook may also import `tests/modules/`. A
   production workbook never contains tests or examples.
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
