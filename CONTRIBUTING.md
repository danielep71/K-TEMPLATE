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
python -m unittest discover -s tests
python tools/check_links.py
python tools/format_vba.py --check
```

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
