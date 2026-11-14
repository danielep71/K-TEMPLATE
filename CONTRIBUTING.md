# 🛠️ Contributing

Thank you for helping. This guide is the entry point; the detailed rules live
in the guides it links, and this file does not repeat them.

By taking part you agree to the [code of conduct](CODE_OF_CONDUCT.md). Report a
suspected vulnerability privately, as [SECURITY.md](SECURITY.md) describes, never
in an issue or pull request.

## 🐞 Issues

Open issues with the issue forms. Each issue is triaged automatically: it needs
an assignee, exactly one priority label, an open milestone, a target end date
earlier than the milestone's due date, and a dependency section. The
[issue guide](docs/repo_governance/ISSUE_LABELS.md) explains the labels, the
priorities (`P1` only for a confirmed calculation bug) and how an issue is
closed, with a *Resolution* section.

## 🌿 Branches and pull requests

- While a release branch such as `release/1.2.0` is active, start task branches
  from it and open the pull request against it; otherwise use the default
  branch. The [release branches](docs/repo_governance/PULL_REQUESTS.md#-release-branches)
  section gives the bases, issue keywords and merge methods.
- Fill in the pull request template. Its *Calculation impact* section names
  the source and the class of the expected values behind a changed
  calculation, or says there is none. Its *Validated* section lists the checks
  you ran, and for a defect fix its [negative control](#-negative-controls);
  write NOT RUN or NOT APPLICABLE with a reason instead of claiming
  evidence that was not produced.
- Add the `CHANGELOG.md` entry under `[Unreleased]` in the same pull request,
  as its [maintenance policy](CHANGELOG.md#-maintenance-policy) describes; a
  fix for a failed release gate before the tag goes in the dated section
  being released instead.
  `tools/check_changelog.py` checks the format, not that the change is
  recorded.
- Every pull request needs a passing **Repository integrity** check and a
  review of its final head, as the
  [pull request guide](docs/repo_governance/PULL_REQUESTS.md) describes.

## ✍️ Commit messages

- The subject says what changed, after a type prefix such as `feat:`, `fix:`,
  `docs:`, `ci:` or `chore:`.
- The body says why, and what was validated: the checks you ran, and what was
  not run, such as an Excel run.
- The last line references the issue with `Refs #N`.
- A closing keyword, such as `Resolves #N`, `Fixes #N` or `Closes #N`, belongs
  only in a pull request into the default branch, or a commit that lands there,
  while no release branch is active, and only for an issue that it completes.
  GitHub closes an issue when a commit whose message names it with a keyword
  reaches the default branch. A commit merged into a release branch reaches it
  at the integration, whatever the state of the issue by then, and it cannot be
  changed after the merge. The
  [release branches](docs/repo_governance/PULL_REQUESTS.md#-release-branches)
  section explains how the **Closing keywords** workflow checks this.

A squashed pull request becomes one commit, and its message follows the same
rules: its title and body become the message, so they use `Refs #N` too.

## 📏 Rules of evidence

- Expected values come from the documented method and its sources, never from
  running the code under test, and a value derived within the project stays
  `illustrative` ([reference classes](docs/methodology/TEST_CASES.md#reference-classes),
  [derivation scripts](docs/methodology/TEST_CASES.md#derivation-scripts)).
- A skipped test is not a pass. A run that does not execute every expected
  case and assertion is incomplete, even if everything that ran passed
  ([completeness policy](docs/methodology/TEST_CASES.md#completeness-policy)).
- A static check is not an Excel run. Record only evidence that was produced,
  on the commit it was produced for.
- A test that passes before and after a fix proves nothing about the defect.
  A fix shows its test failing on the defective code first
  ([negative controls](#-negative-controls)).

## 🔬 Negative controls

A pull request that fixes a defect shows each new or changed test that proves
the fix failing on the defective code, then passing on the fix. Make the
defective code in one of two ways:

- **The base.** Check the base out in a detached worktree, for example with
  `git worktree add --detach ../base-check origin/main`, copy only the new or
  changed test files into it, and run them there. Remove it afterwards with
  `git worktree remove ../base-check`.
- **A mutant.** When the base cannot run the test, for example because the
  test needs code that the fix adds, put the defect back into a copy of the
  fixed file:

  ```sh
  python tools/make_mutant.py src/core/CORE_Example.bas --find "a - b" --replace "a + b"
  ```

  The tool writes the copy under `test-results/mutants/`, which Git ignores,
  at the source's relative path, and never over the source or through a
  symbolic link. It refuses unless the text occurs exactly once, overlapping
  occurrences included, and it keeps every other byte, CRLF line endings and
  `Attribute VB_Name` included. It prints the changed lines
  and the SHA-256 of the source and of the mutant. The text is literal; for a
  text that spans lines, pass `--find-file` and `--replace-file` with files
  that hold it. It exits with 0 when it writes the mutant, 1 when it refuses
  and 2 on a usage error; its docstring lists every refusal.

  To run a mutated VBA component, import the mutant into a development
  workbook in place of its source component, as
  [importing VBA into Excel](INSTALLATION.md#%EF%B8%8F-importing-vba-into-excel)
  describes, and run the suite. Build or import the workbook from source again
  for the run on the fix. For another file, copy the mutant over its source in
  a detached worktree of your branch and run the test there.

Record in *Validated* which way you used, the command, and the failing result
on the defective code; for a mutant, also its SHA-256. When no test can fail on
the defect, write NOT APPLICABLE with the reason.

A negative control proves one fix. `tools/verification_depth.json` keeps the
standing record: for each check of **Repository integrity** that can fail the
run, what it runs, its negative proofs, the violations they show rejected and
its state. A proof is a unittest id, such as
`tests.test_check_links.CheckTests.test_missing_files_anchors_and_paths_are_errors`,
of a test that shows the check rejecting a bad input, or `probe` for a check
whose probe the run makes the tool reject first. The state is `green`, or
`green_in_ci` when a proof is skipped where a tool it needs is missing and runs
in CI. The `verification-depth` check, `python tools/check_verification_depth.py`,
takes the checks from `tools/check_repository.py` as CI runs them, less the
advisory ones, and fails when one has no entry; when an entry names no such
check, is listed twice, has a missing or empty field, a checker that differs
from what the check runs, another state or a claim too short to be specific;
or when a proof names a test that `python -m unittest discover -s tests` does
not collect. It does not run the proofs: the `tests` check does. A pull request
that adds a check, or renames or deletes a proof, updates the inventory in the
same change; to write a missing proof, make its defective input with
`tools/make_mutant.py`.

## 📐 Contract changes

A contract change alters what a documented contract promises its users:

- a dataset, field, layout, unit or validation rule of the
  [data contract](docs/methodology/DATA_CONTRACT.md);
- the inputs, rules, sources or results of a
  [calculation contract](docs/methodology/CALCULATION_CONTRACTS.md);
- a sheet, code name, defined name or written area of the
  [workbook contract](docs/methodology/WORKBOOK_CONTRACT.md);
- a declaration in the [supported VBA API](docs/PUBLIC_API.txt);
- a frozen [accuracy contract](docs/methodology/TEST_CASES.md#accuracy-contracts).

The pull request that makes it updates the contract, its
tests and the documents that describe it together, and its changelog entry
says what changed and what a user must do. A change to computed results or to
an input or output layout carries the `behavior-change` label; from `1.0.0`,
an incompatible one needs a major version
([version rules](CHANGELOG.md#date-and-version-rules)). A change that alters
none of these is not a contract change.

## 🧪 Local checks

Run these before opening a pull request. The tests need Node.js for the label
scripts. The tests that read YAML and the `codeql-default-branch` check need the
pinned configuration parser; without it they are skipped locally.
The quality checks need Ruff and mypy, whose hashes cover CPython 3.13 on
Linux x86_64, and actionlint:

```sh
python -m pip install --require-hashes -r .github/requirements.txt
python -m pip install --only-binary=:all: --require-hashes -r tools/requirements-quality-ci.txt
python tools/check_repository.py
```

For actionlint, download the release that `ACTIONLINT_VERSION` names in
`.github/workflows/repository-integrity.yml` and put it on your `PATH`. On
Linux x86_64, first check the archive against `ACTIONLINT_LINUX_AMD64_SHA256`
there, as CI does.

The last command runs every check of **Repository integrity**, the tests
included, and writes a log for each, `report.json` and `summary.md` under
`test-results/integrity/`, which Git ignores. It runs these checks, in this
order. `OWNER/REPOSITORY` is the repository that `--repository` names, as CI
passes it, or else the GitHub repository of the `origin` remote; without
either, `changelog` runs without `--repository`.

<!-- BEGIN generated: integrity-checks -->
| Check | Runs |
| --- | --- |
| `tests` | `python -m unittest discover -s tests -v` |
| `ruff` | `python -m ruff check --no-cache tools tests` |
| `mypy` | `python -m mypy --no-incremental` |
| `actionlint` | `actionlint -no-color`, once the probe `actionlint -no-color .github/actionlint/invalid-workflow.yml` has failed with `unexpected key "bogus-key"` |
| `labels-sync-self-test` | `node .github/scripts/labels-sync.mjs --self-test` |
| `labels-catalogue` | `node .github/scripts/labels-sync.mjs --mode validate` |
| `labels-drift-self-test` | `node .github/scripts/labels-drift.mjs --self-test` |
| `issue-configuration` | `python tools/check_issues.py --check-config` |
| `template-only-content` | Built in: A generated repository has deleted the template-only section of AGENTS.md. |
| `issue-contact-links` | Built in: The issue forms' contact links point to this repository. |
| `readme-traffic-badges` | Built in: The README's traffic badges read this repository's traffic. |
| `binaries` | `python tools/check_binaries.py` |
| `secret-files` | Built in: No tracked file is named as a secret or private key file. |
| `vba-format` | `python tools/format_vba.py --check` |
| `vba-source` | `python tools/check_vba_source.py` |
| `vba-conditionals` | `python tools/check_vba_conditionals.py` |
| `vba-jumps` | `python tools/check_vba_jumps.py` |
| `vba-public-api` | `python tools/check_vba_public_api.py` |
| `vba-dependencies` | `python tools/check_vba_dependencies.py` |
| `vba-identifiers` | `python tools/check_vba_identifiers.py` |
| `vba-errors` | `python tools/check_vba_errors.py` |
| `vba-declares` | `python tools/check_vba_declares.py` |
| `vba-calls` | `python tools/check_vba_calls.py` (advisory: findings warn, never fail the run) |
| `data-contract` | `python tools/check_data_contract.py` |
| `test-cases` | `python tools/check_test_cases.py` |
| `derivations` | `python tools/check_derivations.py` |
| `independent-implementations` | `python tools/check_independent.py` |
| `accuracy` | `python tools/check_accuracy.py` |
| `workbook-protection` | `python tools/check_workbook_protection.py` |
| `changelog` | `python tools/check_changelog.py --repository OWNER/REPOSITORY` |
| `generated-docs` | `python tools/render_docs.py --check` |
| `restated-counts` | `python tools/check_restated_counts.py` |
| `markdown-links` | `python tools/check_links.py` |
| `documented-commands` | `python tools/check_doc_commands.py` |
| `verification-depth` | `python tools/check_verification_depth.py` |
| `dependabot-target-branches` | Built in: Every Dependabot target-branch exists. |
| `codeql-default-branch` | Built in: codeql.yml lists the default branch, which branch filters cannot name. |
| `required-documents` | Built in: The documents the repository requires exist and are not empty. |
| `whitespace-committed` | `git diff --check 4b825dc642cb6eb9a060e54bf8d69288fbee4904 HEAD` |
| `whitespace-local` | `git diff --check HEAD` |
<!-- END generated: integrity-checks -->

`python tools/render_docs.py --write` regenerates this table and the other
blocks between `BEGIN generated` and `END generated` markers from the code;
**Repository integrity** fails while one differs from a fresh render. Edit
the renderer in `tools/render_docs.py`, never the block. Outside a block, link
a count's source instead of copying the count: `restated-counts` fails on a
count of cases, assertions, modes, checks, datasets or finding codes written
by hand, as the docstring of `tools/check_restated_counts.py` lists. A check whose tool or GitHub
context is missing locally, such as Ruff or the default branch, is reported as
skipped, as is `dependabot-target-branches` when a `target-branch` has no
remote-tracking branch and the `origin` remote cannot be asked whether it
exists; `--strict` fails it instead, as CI does. CI also passes `--ci`, which
adds a last check, `clean-working-tree`: it fails when a file is changed,
deleted, or added and not ignored, so that the results describe the commit
itself. To run a single check, use
its own command, for example `python -m unittest discover -s tests` or
`python tools/format_vba.py --check`.

`python tools/format_vba.py --write` fixes what the last command reports;
re-import a rewritten module into the VBE before committing
([house style](docs/VBA_HOUSE_STYLE.md#source-representation)).

To run Ruff and strict mypy on their own, see the
[setup guide](docs/setup/REPOSITORY_SETUP.md#-maintain-the-baseline).

## 📄 Data and licensing

- Use synthetic data only: no real portfolio, client, counterparty or
  credential data in code, tests, issues, screenshots or workbooks.
- Contributions are licensed under the repository's [LICENSE](LICENSE). Files
  listed in the [third-party notice](docs/licenses/THIRD_PARTY_NOTICES.txt) keep
  the license the notice states for them.
