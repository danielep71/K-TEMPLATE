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
  rewritten; a demonstrable error is corrected with a dated note.

<a id="date-and-version-rules"></a>

### Date and version rules

`python tools/check_changelog.py` checks the headings, dates, order, link
references and `VERSION`, and that this introduction, the *Current status* of
[RELEASING.md](RELEASING.md#current-status) and the README describe the same
release state; **Repository integrity** also checks that every link reference
points to this repository. What a version number means is checked
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

- `tools/check_changelog.py` also fails when the prose describes another
  release state than the changelog: the "No version has been released"
  sentence, the *Current status* of `RELEASING.md`, placeholder lines such as
  `TBD`, and a "Release Candidate" mention or unchecked task in the README
  (#131).
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

[Unreleased]: https://github.com/danielep71/MyTemplate/commits/main
