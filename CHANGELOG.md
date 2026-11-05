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
  [RELEASING.md](RELEASING.md) describes. After that, until the tag exists,
  a fix for a release gate that failed on the candidate records its change in
  the dated section being released
  ([step 2](RELEASING.md#2-integrate-and-freeze-the-candidate)). A published section is not
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

- `tools/check_doc_commands.py` checks the commands that the tracked
  Markdown cites, in inline code and in shell or unmarked fenced blocks:
  each `python tools/...` script and each `python -m` module of the
  repository must be tracked, and each subcommand and option must exist in
  the script's `argparse` parser, which it reads from the source without
  importing or running it. It joins continued lines, skips placeholders
  such as `<folder>` and `X.Y.Z`, and reports a parser it cannot read
  instead of passing it. **Repository integrity** runs it, so a renamed
  option or a removed tool fails until the guides follow; only the
  commands of a tool marked **Remove once run**, once it is deleted, are
  skipped (#273).
- **Repository integrity** has a `secret-files` check: it fails on a tracked
  file named as a secret or key file (`.env` and `.env.*` except
  `.env.example`; `.pem`, `.key`, `.p12`, `.pfx`, `.pvk`, `.jks`,
  `.keystore`; `id_rsa*` and `id_ed25519*` except their `.pub` keys;
  `.kdbx`), also one force-added past `.gitignore`, and says to remove the
  file, rotate the secret and follow *Handling a finding* in
  `DISTRIBUTION.md`. `.gitignore` now ignores every one of these names
  (#276).
- `tools/release.py verify-published vX.Y.Z --repository OWNER/REPOSITORY
  [--milestone N] [--asset FILE=SHA256]` checks a published release through
  `gh`, read-only, and prints one line per fact: the tag is annotated and
  points at the commit its `Candidate` line names, **Repository integrity**
  succeeded on it, the release is published and not a draft, its assets are
  exactly the source bundle, `SHA256SUMS` and each `--asset`, with the
  SHA-256 values that `SHA256SUMS` and `--asset` record, `gh attestation
  verify` passes on the bundle, and the milestone is closed with no open
  issue. It fails on any of these. `RELEASING.md` step 7, the milestone
  closeout and the release certification now use it; the download, the
  unblock and the Excel checks of a published workbook stay manual (#274).

- `tools/initialize_repository.py` adapts a repository created from the
  template in one diff: it deletes the template-only section of `AGENTS.md`,
  points the contact links and the traffic badges at the new repository, or
  removes the badges, sets or removes Dependabot's `target-branch`, replaces
  the changelog's entries and link references with an empty `[Unreleased]`
  section, and replaces `<REGIME>` in the house style. Without `--apply` it
  only lists the changes; with it, it refuses a working tree that differs from
  the commit. It changes nothing on GitHub, and a second run changes nothing.
  A test initializes a copy of the template and runs **Repository integrity**
  on it with `--template false`. README step 2 and the setup guide name the
  command, with the manual edits as the fallback; the initializer and its test
  can be deleted once run (#272).
- `tools/check_vba_calls.py` checks VBA calls that resolve to a procedure of
  the project, in statements and in expressions, qualified or not: too many
  or too few arguments, counting `Optional`, `ParamArray` and named
  arguments, a named argument the procedure does not declare, and a
  qualified call to a `Private` procedure of another module. It reports
  nothing for a `With`-block member, an object member, an array element, a
  `Property` procedure or a built-in. **Repository integrity** runs it as
  advisory: a finding is a warning, not a failure, until the house style's
  rule makes it blocking (#249).
- The host record of an Excel run, `host.txt`, names the operator, the start
  and finish times with their UTC offset, the Excel build and bitness, the
  `AutomationSecurity` and the Trust access to the VBA project as found, and
  an outcome (`PASS`, `FAIL`, `IMPORT_FAILED`, `COMPILE_FAILED`,
  `EXECUTION_TIMEOUT`, `INCOMPLETE`, `NOT_ENABLED` or `UNAVAILABLE`), beside
  Windows, the references and the compile mode. `excel_run.py` writes it also
  when the run stops early, and reads the macro settings without changing
  them. `check_excel_run.py` rejects a record that lacks a field or leaves
  the operator unread (`operator=UNKNOWN`), and accepts an outcome other than
  `PASS` only as evidence of that outcome; a record written by hand in the
  earlier three-line form now fails (#244, #284).
- `tools/check_independent.py`, which **Repository integrity** runs: it runs
  each independent implementation registered for an `IMPL-` source in
  `tests/independent/adapters.json`, a file a project creates when it adds
  one, on the test cases, through an adapter that
  translates each case into the implementation's input, and compares the
  results with the expected values within their tolerance. It lists the cases
  an implementation cannot represent, with the reason, and skips an adapter
  whose implementation is not installed, with a note. The template registers
  none, so the check passes and says that nothing was cross-checked (#248).
- The VBA harness has patterns for three more kinds of test.
  `tests/modules/TEST_Relations.bas` holds illustrative relation tests
  (additivity, symmetry, monotonicity, zero input) on a stand-in total, each
  with a check that a deliberately corrupted output fails it; it is not
  certified, and `tests/test_vba_harness.py` keeps its counts and checks
  consistent. The harness guide gives the skeleton of an end-to-end case
  runner, which restores the input sheets it writes, and the conventions of a
  volume benchmark, which `check_excel_run.py` refuses as an evidence suite.
  Static checks only: the new module has not been run in Excel (#246).
- The methodology guides trace each rule. The methodology overview has a
  support matrix that marks each feature as supported, not supported or
  entered by the user. The validation plan has a traceability table that
  maps each `SRC-` paragraph to its `CALC-` step, its VBA procedure, its test
  cases and a status: Validated, Exercised, Not validated or Not implemented.
  Each table has one illustrative row. `check_test_cases.py` fails a row whose
  source, step, procedure under `src/` or case does not exist, whose case
  tests another calculation, or whose status contradicts what the row names
  (#247).
- `RELEASING.md` step 8 checks the published workbook, with a *Published
  workbook* section in the release certification: the attached workbook's
  build stamp names the candidate and `production`, a rebuild of the tag
  gives the same SHA-256 or the same parts, and a downloaded, unblocked copy
  shows no link to another workbook and no `#REF!` name, and survives run,
  save and reopen. An optional user manual for workbook recipients, by
  convention one folder per released version under `docs/user-manual/`, is
  built to a PDF with a tool the release records (Pandoc as the example; the
  template installs none) and listed among the certified artifacts (#250).
- With `src/workbook/Template.xlsx`, `tools/devbuild/build_workbook.py` builds
  the workbook from a copy of the template, keeping its formats, print areas
  and protection: it sets the code names, writes the stamp and adds the VBA
  project, and two builds stay byte-identical. Without a template it writes
  blank sheets as before. A protection matrix, by convention
  `docs/methodology/workbook-protection.json` and added with the template,
  states each sheet's protection options and its input (unlocked) and output
  (locked) ranges;
  `tools/check_workbook_protection.py` fails **Repository integrity** when the
  template differs from it, and passes as inactive while neither exists (#243).
- A machine-readable data contract, `docs/methodology/data-contract.json`,
  holds each dataset's file or sheet, encoding and delimited-text format, and
  each field's type, unit or currency field, required flag and allowed values
  or range. The *Datasets* table and the field tables of `DATA_CONTRACT.md`
  are rendered from it, and **Repository integrity** fails while they differ.
  `tools/check_data_contract.py`, which it runs, checks the contract and, with
  `--dataset`, validates input files against it, reporting each finding with a
  stable code such as `DC-TYPE` or `DC-MISSING-COLUMN`.
  `tools/check_test_cases.py` now runs that validator on every variant file of
  a file-based case outside an `invalid` case. The illustrative dataset `DS-01`
  lists the three columns its example file already holds (#238).
- The evidence of an Excel run can certify project suites beside the example
  suite: the evidence policy, now version 2, lists them under `suites`, each
  with its certification entry point, counts and two logs. `excel_run.py` runs
  every entry point, `check_excel_run.py` checks each suite's normal run and
  certification from the same host, and every suite is read as the policy's
  Office bitness compiles it, so `#If Win64` cases count only there (#237).
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
  and its certification saves them with the commit of the workbook's build
  stamp, the example suite as `results.json` and a project suite in the
  `results` file its entry in `.github/excel-evidence-policy.json` names. The
  example suite records the illustrative cases' totals from `tests/fixtures/`,
  so its entry points take that folder. `tools/check_excel_run.py` requires
  every results file when the commit holds contracts, measures them together,
  rejects a result that two files hold, and fails the run when the contracts
  do not hold. A workbook imported by hand has no build stamp, so its run
  fails while contracts exist (#184, #284).
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
- Seven static checks of exported VBA source, which **Repository integrity**
  runs, the call check as advisory:
  - placement, names, encoding, options, layout and storage;
  - `#If` structure and `PtrSafe` in each compilation environment;
  - procedure-local jump targets;
  - visibility and the public API inventory, including a public name in two
    modules;
  - the dependency direction;
  - identifiers, in each compilation environment: a name a procedure assigns
    that nothing declares, a statement that calls a procedure no module
    provides, or a declaration named like a VBA keyword, such as
    `Dim Empty()`. A statement it cannot read is a finding (#173);
  - the calls to the project's own procedures, as the
    `tools/check_vba_calls.py` entry above describes (#249).

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

- The **Traffic alert** job reads the `traffic-history` commit that its
  collection pushed, or merged onto when nothing changed, which the collect
  job outputs as `history-sha`, not the branch's current tip; it stops when
  that output is missing, and a workflow control test keeps the binding
  (#304).
- External links must use HTTPS: `tools/check_links.py` fails a plain
  `http://` link in a Markdown link, a reference definition, an autolink or an
  HTML attribute, outside code, in **Repository integrity**, and the weekly
  **Link check** requests links over HTTPS only and fails a redirect to
  another protocol. A 403, 429 or 5xx answer stays a warning, but one listed
  in the optional `.github/link-exceptions.json`, with a reason and an
  `expires` date, fails once that date has passed; **Repository integrity**
  fails while the file is malformed or lists a URL that no Markdown file
  links (#275).
- **Repository integrity** cancels the run of a pull request commit that a
  newer commit replaces; a push run is never cancelled. CI runs
  `tools/check_repository.py` with `--ci`, which fails when, after the
  checks, the working tree differs from the commit. mypy checks `tests/` as
  well as `tools/`; four unannotated test files need no annotations, but their
  code is checked. The bug form has an optional field for evidence, a
  workaround and recovery, and the enhancement form one for alternatives
  considered (#251).
- `tools/check_excel_run.py` compares the modules exported by hand with the
  commit without the letter case of code outside strings and comments, as the
  VBE re-cases a name (`HB-01`, still `assumed`), and reports each re-cased
  identifier on a `Note:` line. In a document module it no longer compares the
  attributes Excel writes itself, such as `VB_Base`; `VB_Name`, string
  literals, comments and any other difference still fail (#245).
- `SECURITY.md` describes the runtime risk surface: what the test harness,
  the example suite and `tools/devbuild/excel_run.py` read, write and delete,
  the Excel settings they need (enabled macros, no Trust access, no change to
  a macro setting), which access the dependency rules forbid in core and
  classes, and what a generated repository states for its own entry points.
  It adds a data-and-secrets section (what a workbook can hide, and that a
  leaked secret is revoked or rotated before it is reported) and a safe
  harbor for good-faith research. The setup guide tells a generated
  repository to adapt the runtime section (#242).
- The pull request template asks for the calculation impact of a change: the
  source of each changed rule and the class of the expected values that test
  it, or "No calculation change". Its checklist asks for the `[Unreleased]`
  changelog entry. `CONTRIBUTING.md` states the rules of evidence (expected
  values never come from the code under test; a skipped test is not a pass),
  what a commit message says, and what a contract change is and what it
  updates (#240).
- mypy also reports unreachable code (`warn_unreachable`), such as a branch
  that type narrowing has already excluded; strict mode alone does not (#191).
- Label synchronization no longer deletes a label that an issue or pull request,
  open or closed, still carries: with `prune: true`, it checks every label it
  would delete and stops before any change. A name with a comma, which the
  API's label filter cannot query, is checked against the labels of every
  issue and pull request, up to 10,000; a malformed response or the limit
  stops it (#188, #305).
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

- The release closeout can finish: the release certification, the
  distribution review and the milestone closeout are issues of the next
  milestone, and milestone N is closed in step 7, before publication, so
  `tools/release.py verify-published --milestone N` can pass and **Issue
  triage** still finds an open milestone for those issues; the checklist has
  that gate, and every `verify-published` command names `--repository`. The
  manual cleanup lists every edit the initializer makes, restoring the
  changelog's "No version has been released" sentence, deleting `VERSION`
  if it exists and resetting *Current status* included, and says which of
  them **Repository integrity** checks; the initializer's `--default-branch`
  help says what the link shows. The setup guide's handover table covers
  every tracked file, names both traffic tests and describes each test of
  `tests/test_repository_consistency.py`; the structure guide covers the
  root files, the reserved `VERSION` and all of `.github/`; the changelog
  policy, the contributing guide and the pull request template
  say where a release-gate fix records its change, the label guide matches
  the workflows' write-permission list, and `tests/independent/adapters.json`
  is described as a file to create (#302, #318).
- The methodology guides agree with each other, the illustrative data and
  the repository's configuration. The validation plan gives "independent" the
  test-case format's meaning: values the project computes by hand, in a
  spreadsheet or with its own script are `illustrative` cross-checks, and a
  rule is validated only by `published` or `independent` values; the
  validation matrix uses the traceability statuses. The data contract's file
  format states the fixed rules its validator applies. `.gitattributes`,
  `.editorconfig` and the guides name the source register in `SOURCES.md`.
  The illustrative settings area is typed only in its value column `C5`, with
  its size in the contract, and the protection example and the case-runner
  skeleton follow it. The illustrative expected values cite rule
  `CALC-01.1`, the *Outputs* template names each output's quantity, default
  tolerances go in the calculation contracts' common rules, a fixture
  record's `id` carries the dataset's identifier field, and the protection
  check is described as it compares options (#301).
- A project suite can be certified with a runner of its own: its `suites`
  entry may name it as `runner` (`RunSuite` by default), and the harness
  guide states the shape `tools/check_excel_run.py` reads, so the relation
  tests can be certified with `"runner": "RunRelationTests"`; a `RunCase`
  runner is not certifiable as it stands. The Excel evidence documents agree
  with the code: `SECURITY.md` lists the harness's read of `build_commit`,
  `UNKNOWN` is allowed in any record that did not pass but never in
  `operator`, `started`, `finished` or `outcome`, a one-command run records
  the build's references, and the installation guide's rejection table
  covers the project suites. The host behaviour register has `HB-06` (End
  or Reset clears the harness's state) and `HB-07` (the Immediate window
  keeps 200 lines), both `assumed`, and the guides cite them and `HB-01`;
  `HOST_TARGETS` and the house style's examples are described consistently
  (#300).
- `tools/release.py notes` and `bundle` report a failure on a FAIL line
  instead of a traceback, and `verify` reports a missing tag as a finding.
  `tools/check_links.py --external` skips the repository's own links locally
  as in CI, reading the repository from the `origin` remote when
  `GITHUB_REPOSITORY` is unset. `tools/devbuild/lo_smoke.py` exits 2 for an
  unknown option or a missing workbook before it starts LibreOffice, and 1
  without a traceback when LibreOffice cannot run the workbook. The weekly
  **OpenSSF Scorecard** run is no longer skipped: a scheduled event carries
  no repository in its payload, so the job exempts it from the branch test
  and reads the visibility from the API (#299).
- The VBA harness guides, module headers and pull request template agree
  with the code. The example suite has a `cleanup-primary-only` case, in
  which only the operation fails and every restoration succeeds: the primary
  error reaches the caller, no cleanup failure is reported and both settings
  are restored. The suite now has 14 cases and 34 assertions, and the
  evidence policy, the installation guide and the LibreOffice smoke test
  follow; the case has not run in Excel. The installation and harness
  guides say that **End** or **Reset** in the VBE clears the run state, so
  that stop is not detected, and what to check by hand afterwards, and the
  harness guide gives the two records of a refused start, with
  `operation=NOT_RUN` and `cleanup=NOT_RUN`. The host behaviour register has
  `HB-05`, still `assumed`, for Excel compiling a project written from
  source when the workbook opens, cited where the guides and tools rely on
  it; `tools/devbuild/excel_run.py` cites `HB-03` and `HB-04`, and the
  `HB-02` host records that its channel was not noted. The import steps
  include `examples/modules/`, the build is said to leave forms out, the
  house style says what the call and identifier checks each resolve, and the
  pull request template and guide admit `compile=ON_DEMAND` and ask for the
  host record's `outcome` of a run that stopped early (#285).
- The methodology guides agree with their illustrative example and the
  tools. `CALC-01` sums the notional amounts of the whole calculation, as its
  data, cases and derivations do, and the support matrix's *Supported* means
  in scope, with each rule's status in the traceability table. The workbook
  contract declares `Inputs_Settings` as a one-cell table anchor, which
  `tools/devbuild/workbook.json` builds on `Inputs!B4`, and the case-runner
  skeleton derives the area it captures and restores from that anchor.
  `Run_Status` and the status lines are declared with the first sheet that
  code writes. The contract also lists `sheet` among the protection options
  that default to `false`, as `tools/check_workbook_protection.py` does, keeps
  the reporting-date label in the example's `output`, names code names `sh`
  plus the tab name in PascalCase, as the house style now does, and states
  one failure behaviour: no new result, the result areas cleared and marked
  `FAILED`. The validation plan names only the `nominal`, `boundary` and
  `invalid` case categories and leaves failure and repeatability to VBA
  tests, and the rules for filling in the guides name the guides with a
  *Decisions* table and the two tables that differ (#286).
- Without PyYAML, **Repository integrity** skips `codeql-default-branch`, and
  `--strict` fails it, instead of failing it with an import error; the unit
  tests that read YAML are skipped too. The release certification accepts
  the evidence of a `tools/devbuild/excel_run.py` run for the workbook built
  from the candidate (`Binding: VERIFIED` by rebuild and the workbook's
  SHA-256) and for the references (`host.txt`). `SECURITY.md` lists every
  read of `excel_run.py`, the setup guide lists `tests/test_excel_run.py`
  with the devbuild tools, and the repository structure names
  `docs/release/`, `docs/licenses/` and the harness guide (#287).
- **Repository integrity** passes in a repository that retired the setup
  tooling: the test that compares the guide's *Remove once run* rows with
  `tools/check_doc_commands.py` is skipped once
  `docs/setup/REPOSITORY_SETUP.md` is deleted, and the retirement steps say
  so and name `tools/check_doc_commands.py` with the link check (#298).
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
