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
- Fill in the pull request template. Its *Validated* section lists the checks
  you ran; write NOT RUN or NOT APPLICABLE with a reason instead of claiming
  evidence that was not produced.
- Every pull request needs a passing **Repository integrity** check and a
  review of its final head, as the
  [pull request guide](docs/repo_governance/PULL_REQUESTS.md) describes.

## 🧪 Local checks

Run these before opening a pull request. The tests need Node.js for the label
scripts and the pinned configuration parser for the workflow control tests:

```sh
python -m pip install --require-hashes -r .github/requirements.txt
python tools/check_repository.py
```

The second command runs every check of **Repository integrity**, the tests
included, and writes a log for each, `report.json` and `summary.md` under
`test-results/integrity/`, which Git ignores. It runs these checks, in this order:

<!-- BEGIN generated: integrity-checks -->
| Check | Runs |
| --- | --- |
| `tests` | `python -m unittest discover -s tests -v` |
| `ruff` | `python -m ruff check --no-cache tools tests` |
| `mypy` | `python -m mypy --no-incremental` |
| `actionlint` | `actionlint -no-color` |
| `labels-sync-self-test` | `node .github/scripts/labels-sync.mjs --self-test` |
| `labels-catalogue` | `node .github/scripts/labels-sync.mjs --mode validate` |
| `labels-drift-self-test` | `node .github/scripts/labels-drift.mjs --self-test` |
| `issue-configuration` | `python tools/check_issues.py --check-config` |
| `template-only-content` | Built in: A generated repository has deleted the template-only section of AGENTS.md. |
| `issue-contact-links` | Built in: The issue forms' contact links point to this repository. |
| `readme-traffic-badges` | Built in: The README's traffic badges read this repository's traffic. |
| `binaries` | `python tools/check_binaries.py` |
| `vba-format` | `python tools/format_vba.py --check` |
| `vba-source` | `python tools/check_vba_source.py` |
| `vba-conditionals` | `python tools/check_vba_conditionals.py` |
| `vba-jumps` | `python tools/check_vba_jumps.py` |
| `vba-public-api` | `python tools/check_vba_public_api.py` |
| `vba-dependencies` | `python tools/check_vba_dependencies.py` |
| `vba-identifiers` | `python tools/check_vba_identifiers.py` |
| `test-cases` | `python tools/check_test_cases.py` |
| `derivations` | `python tools/check_derivations.py` |
| `changelog` | `python tools/check_changelog.py` |
| `generated-docs` | `python tools/render_docs.py --check` |
| `markdown-links` | `python tools/check_links.py` |
| `dependabot-target-branches` | Built in: Every Dependabot target-branch exists. |
| `codeql-default-branch` | Built in: codeql.yml lists the default branch, which branch filters cannot name. |
| `required-documents` | Built in: The documents the repository requires exist and are not empty. |
| `whitespace-committed` | `git diff --check 4b825dc642cb6eb9a060e54bf8d69288fbee4904 HEAD` |
| `whitespace-local` | `git diff --check HEAD` |
<!-- END generated: integrity-checks -->

`python tools/render_docs.py --write` regenerates this table and the other
blocks between `BEGIN generated` and `END generated` markers from the code;
**Repository integrity** fails while one differs from a fresh render. Edit
the renderer in `tools/render_docs.py`, never the block. A check whose tool or GitHub
context is missing locally, such as Ruff or the default branch, is reported as
skipped; `--strict` fails it instead, as CI does. To run a single check, use
its own command, for example `python -m unittest discover -s tests` or
`python tools/format_vba.py --check`.

`python tools/format_vba.py --write` fixes what the last command reports;
re-import a rewritten module into the VBE before committing
([house style](docs/VBA_HOUSE_STYLE.md#source-representation)).

Ruff, strict mypy and actionlint also run in **Repository integrity**; the
[setup guide](docs/setup/REPOSITORY_SETUP.md#-maintain-the-baseline) explains how
to run them locally with the same pinned versions.

## 📄 Data and licensing

- Use synthetic data only: no real portfolio, client, counterparty or
  credential data in code, tests, issues, screenshots or workbooks.
- Contributions are licensed under the repository's [LICENSE](LICENSE). Files
  listed in the [third-party notice](docs/licenses/THIRD_PARTY_NOTICES.txt) keep
  the license the notice states for them.
