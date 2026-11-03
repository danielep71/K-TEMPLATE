# 📜 Changelog

Notable changes to this repository, from a user's point of view. The format
follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions
follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html). It is
neither a commit log nor release evidence. **No version has been released.**

## 🧭 Maintenance policy

- Add a material change under `[Unreleased]` in the pull request that makes
  it. A change nobody would notice, such as a typo or a refactor with no
  effect, needs no entry.
- Describe the observable result: what exists, what a user or contributor
  can now do, what changed in a contract, and what to do about it.
- Link the issue or pull request that holds the detail.
- Record only evidence that was produced. A static check is not an Excel run:
  say what was and was not run in Excel.
- Until a version is released, edit an existing entry rather than adding one
  that describes a change to unreleased work: an entry describes what the
  release contains.
- Entries move into a dated section only when a release is prepared, as
  [RELEASING.md](RELEASING.md) describes. A published section is not
  rewritten; a demonstrable error is corrected in review, with a line
  `Corrected after release on YYYY-MM-DD: <reason>` in that section.

<a id="date-and-version-rules"></a>

### Date and version rules

`python tools/check_changelog.py` checks the headings, dates, order, link
references and `VERSION`, and that this introduction, the *Current status* of
[RELEASING.md](RELEASING.md#current-status) and the README describe the same
release state; **Repository integrity** also checks that every link reference
points to this repository. Once a release's tag exists, its section must equal
that section at the tag, unless it holds the dated correction line above; a
correction is reported as a note. What a version number means is checked
in review.

| Rule | Standard |
| --- | --- |
| Version | `MAJOR.MINOR.PATCH`, without a leading `v` in headings |
| Release heading | `## [X.Y.Z] - YYYY-MM-DD`, with a real calendar date |
| Order | `[Unreleased]` first, then releases from newest to oldest, by version and date |
| Link references | One for `[Unreleased]` and one for every release; a release's points to `.../releases/tag/vX.Y.Z` |
| `0.y.z` | Initial development: no compatibility is promised between versions |
| Patch, minor, major | From `1.0.0`: a compatible correction, a compatible capability, an incompatible change to a calculation result, an input or output layout, or the supported VBA API |

| Category | Use for |
| --- | --- |
| **Added** | New capabilities, files, checks or tests |
| **Changed** | Changes to existing behaviour, contracts, tooling or documentation |
| **Fixed** | Corrected defects |
| **Security** | Safely disclosed security corrections |
| **Validation** | Evidence actually produced |
| **Known limitations** | Deliberate, unresolved boundaries |

Use only the categories an entry needs.

## [Unreleased]

Changes made before this changelog existed are described by the closed
milestones and their pull requests.

### Added

- `tools/check_vba_calls.py` checks VBA calls that resolve to a procedure of
  the project, in statements and in expressions, qualified or not: too many
  or too few arguments, counting `Optional`, `ParamArray` and named
  arguments, a named argument the procedure does not declare, and a
  qualified call to a `Private` procedure of another module. It reports
  nothing for a `With`-block member, an object member, an array element, a
  `Property` procedure or a built-in. **Repository integrity** runs it as
  advisory: a finding is a warning, not a failure, until the house style's
  rule makes it blocking (#249).
- **Repository integrity** fails a committed Office package, outside the
  regulatory source copies, that carries macros (a `vbaProject` part, a
  reference to one or a macro-enabled content type) or personal metadata (an
  author, last editor, company or manager in `docProps/`) (#241).
- Dependabot version updates wait 14 days after a release (`cooldown`), and a
  workflow control test requires it. The pull request guide describes how
  security updates, the manual actionlint update and the files GitHub reads
  only from the default branch behave with a release branch (#239).
- `tools/devbuild/excel_run.py` records an Excel run with one command on
  Windows. It builds the workbook, runs the certification in Excel through
  COM, closes without saving and writes `host.txt`, then checks the
  evidence. It needs no Trust access: with `--workbook` and no export,
  `tools/check_excel_run.py` proves identity by rebuilding the commit and
  comparing the workbook's parts byte for byte, and accepts `compile=ON_DEMAND`,
  because only the modules the run calls are compiled. It changes no macro
  setting: where Excel opens the workbook with macros disabled, it waits for
  **Enable Content** (#204).
- Accuracy contracts (`tests/accuracy/contracts.json`, checked by
  `tools/check_accuracy.py` in **Repository integrity**). A contract states
  the worst error a calculation guarantees for one regime, over separate
  calibration and holdout cases. Freezing it needs enough `published` or
  `independent` holdout cases, and a frozen threshold grows only with a new
  decision. With `--results`, the check measures the results of an Excel run.
  The template's illustrative contract stays a draft (#162). An Excel run now
  writes those results: a suite records them with `TEST_Harness.RecordResult`,
  and the certification saves them as `results.json`, with the commit of the
  workbook's build stamp. The example suite records the illustrative cases'
  totals from `tests/fixtures/`, so its entry points take that folder.
  `tools/check_excel_run.py` requires the file when the commit holds
  contracts, and fails the run when they do not hold (#184).
- A **Release** workflow: on a pushed `vX.Y.Z` tag it verifies the tag against
  the changelog, `VERSION`, the default branch, the tag message and
  **Repository integrity**. It builds a reproducible source bundle with its
  checksums, attests their build provenance, and drafts the GitHub Release.
  A private repository attests only when the repository variable
  `RELEASE_ATTESTATIONS` is `enabled`, with GitHub Enterprise Cloud; otherwise
  the draft says that the bundle is not attested. A manual run rehearses a
  ref without attesting or releasing (#161).
- `tools/render_docs.py`, which renders documentation facts the code already
  knows into marked blocks: the list of **Repository integrity** checks in
  `CONTRIBUTING.md` and the example suite's expected results in
  `INSTALLATION.md`. **Repository integrity** fails while a block is stale
  (#163).
- An OpenSSF Scorecard workflow, which scores the default branch's
  supply-chain posture weekly and on each push to it. It publishes to code
  scanning and the Scorecard service on public repositories only (#164).
- `tools/devbuild/build_workbook.py`, which builds a macro-enabled workbook
  from the source of one commit, with the sheets, code names and defined
  names of the workbook contract, a build stamp, and a VBA project written
  from source, so it needs no *Trust access to the VBA project object model*.
  A rebuild is byte-identical, and a working tree with changes is stamped as
  not evidence-grade. `tools/devbuild/export_manifest.py` and
  `check_excel_run.py --workbook` bind an export to the workbook it came from
  (`VERIFIED`, `FAIL` or `UNVERIFIED`), and a release requires `VERIFIED`. An
  optional LibreOffice smoke test runs the harness certification as an early
  warning, never as evidence (#100).
- A VBA test harness and an example suite in `tests/modules/`, with a
  certification entry point. It refuses state left by an earlier run,
  reports completeness apart from the verdict, and checks that Excel's state,
  the open workbooks and the active window are unchanged. It writes every log
  before it raises, and checks that each deliberate failure reaches its
  caller. `tools/check_excel_run.py` checks the evidence of an Excel run
  against one commit, without Excel, and rejects an incomplete run. The
  [harness guide](docs/methodology/TEST_HARNESS.md) and
  [recording an Excel run](INSTALLATION.md#recording-an-excel-run) describe
  them, and the check is a release gate (#79).
- Six static checks of exported VBA source, which **Repository integrity**
  runs:
  - placement, names, encoding, options, layout and storage;
  - `#If` structure and `PtrSafe` in each compilation environment;
  - procedure-local jump targets;
  - visibility and the public API inventory, including a public name in two
    modules;
  - the dependency direction;
  - identifiers, in each compilation environment: a name a procedure assigns
    that nothing declares, a statement that calls a procedure no module
    provides, or a declaration named like a VBA keyword, such as
    `Dim Empty()`. A statement it cannot read is a finding (#173).

  The [house style](docs/VBA_HOUSE_STYLE.md#static-checks) lists what they
  cannot prove (#70).
- `tools/check_repository.py`, one command that runs every **Repository
  integrity** check locally and in CI. It writes a log per check and a
  `report.json` with the commit, the dirty state and the SHA-256 of every
  tracked file. A summary of each check appears on the run page, and CI keeps
  the reports as a 30-day artifact. White space is now checked in every
  committed file, not only in the pushed changes (#73).
- `tools/format_vba.py`, which brings exported VBA source into the
  house-style layout with `--write`, and which **Repository integrity** runs
  with `--check`. It changes only line endings, trailing white space,
  indentation tabs, the level of procedure and label lines and the blank
  lines between procedures, never a string literal (#130).
- A release process: [RELEASING.md](RELEASING.md), this changelog, a release
  certification record and a milestone closeout record under `docs/release/`,
  and `tools/check_changelog.py`, which **Repository integrity** runs (#69).
- A numerical test-case format, `docs/methodology/TEST_CASES.md`, and
  `tools/check_test_cases.py`, which **Repository integrity** runs (#75).
  A case is a JSON fixture, or a file-based case whose variants hold input
  files named and encoded as in the data contract, stored byte for byte
  (#117). An illustrative case of each form is under `tests/fixtures/` and
  `tests/expected/`.
- `tools/check_binaries.py`, which **Repository integrity** runs: a tracked
  binary file outside the approved paths fails unless
  `.github/binary-allowlist.json` lists it with its current SHA-256 (#113).
- `tools/check_derivations.py`, which **Repository integrity** runs: it runs
  the derivation scripts named by expected values and compares their results
  within tolerance, once per variant for a file-based case, with an
  illustrative script for each example case (#115, #117).

- A host behaviour register in `docs/methodology/VALIDATION_PLAN.md`: each
  Excel or VBE behaviour that the code relies on is `assumed` or `measured`,
  with the probe, the host and the result. Assumed entries are listed as a
  limit in the release certification (#133).
- `docs/methodology/WORKBOOK_CONTRACT.md`, a workbook contract to fill in: the
  sheets and code names, defined names, areas, operations, status after a
  failure, and what the build creates (#118).

### Changed

- mypy also reports unreachable code (`warn_unreachable`), such as a branch
  that type narrowing has already excluded; strict mode alone does not (#191).
- Label synchronization no longer deletes a label that an issue or pull request,
  open or closed, still carries: with `prune: true`, it checks every label it
  would delete and stops before any change. A name with a comma is refused,
  because its use cannot be queried (#188).
- `tools/check_changelog.py` freezes released sections: once the tag
  `vX.Y.Z` exists, the section must equal its content at the tag. A section
  with a reviewed line `Corrected after release on YYYY-MM-DD: <reason>` may
  differ, and the difference is reported as a note. Deleting or renaming a
  tagged section fails. A release without a tag yet, or a copy without Git
  history, is a note (#172).
- The workflow control tests check every `uses:` reference a workflow can
  reach, not only those of job steps. That covers job-level calls of reusable
  workflows, local actions followed into their metadata, and Docker images,
  which need an `@sha256:` digest. Each remote pin also needs a version comment
  (#174).
- What an Excel run must leave as evidence moves from constants in
  `tools/check_excel_run.py` to `.github/excel-evidence-policy.json`: the
  harness, the suite with its counts and certification entry point, the modes
  and their logs, the evidence files, the references and the Office bitness.
  The check reads the policy of the candidate commit, rejects a missing or
  malformed one, and fails when the suite disagrees with it (#160).
- `tools/check_changelog.py` also fails when the prose describes another
  release state than the changelog: the "No version has been released"
  sentence, the *Current status* of `RELEASING.md`, placeholder lines such as
  `TBD`, and a "Release Candidate" mention or unchecked task in the README
  (#131). A status row must name the newest release as a whole version
  number, so `10.2.0` does not pass for `0.2.0` (#142).
- The pull request template has an *Excel run* block for changes that need
  Excel: the commit, host, compile result, suite, evidence, and a completion
  state that is `PASS`, `FAIL`, `INCOMPLETE` or `NOT RUN`.
  [PULL_REQUESTS.md](docs/repo_governance/PULL_REQUESTS.md#excel-run) defines
  the states (#134).
- `INSTALLATION.md` explains how to upgrade to a new revision by rebuilding
  the workbook and carrying its inputs across, how to diagnose common import
  and run failures, and what removing the workbook's components leaves
  behind (#132).
- `docs/licenses/provenance.json` records the source repository, license and
  revision of each adapted file, or why the revision is not recorded; a test
  keeps it in step with the third-party notice (#116).
- **Repository integrity** proves that actionlint rejects an invalid workflow,
  and the workflow control tests reject duplicate YAML keys in the workflows,
  Dependabot's configuration and the issue forms (#114).

### Fixed

- `tools/setup_repository.py` no longer reports paused Dependabot security
  updates as on: they count as drift, the setup record shows `on (paused)`,
  and an apply that leaves them paused fails its read-back (#190).
- A backtick, pipe or line break in a label name, colour or API error no
  longer breaks the code span or table row of the label synchronization and
  drift summaries (#189).

### Validation

- First Excel run of the template, with `tools/devbuild/excel_run.py` at
  50c0658 on Microsoft 365 Excel 16.0 build 20430, 64-bit, Windows 11
  Enterprise 25H2: `excel-run: PASS`, binding `VERIFIED` by rebuild, ACC-01
  worst error 0.0, `compile=ON_DEMAND`. The host behaviours `HB-02` to `HB-04`
  of the validation plan are measured on that host (#204, #211).

[Unreleased]: https://github.com/danielep71/MyTemplate/commits/main
