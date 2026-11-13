<div align="center">

<img src="docs/assets/social-preview.png" alt="K-TEMPLATE: the starting point for regulatory risk calculation workbooks in Excel VBA, for SA-CCR, SFT, FRTB, LCR and NSFR" width="100%">

# 🧰 K-TEMPLATE

### The starting point for regulatory risk calculation workbooks in Excel VBA

**SA-CCR · SFT · FRTB · LCR · NSFR · Exported VBA source · Reproducible builds · Excel-run evidence · Governed releases**

<br>

[![Excel VBA](https://img.shields.io/badge/Excel_VBA-exported_source-217346?style=for-the-badge)](#-vba-development)
[![Windows](https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge)](INSTALLATION.md)
[![Python](https://img.shields.io/badge/Tooling-Python_3.11+-3776AB?style=for-the-badge&logo=python&logoColor=white)](#-quality-gates)
[![GitHub Actions](https://img.shields.io/badge/CI-GitHub_Actions-2088FF?style=for-the-badge&logo=githubactions&logoColor=white)](#-quality-gates)
[![License](https://img.shields.io/badge/License-MPL_2.0-2ea44f?style=for-the-badge)](#-license)
<br>[![Clones, including CI](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/danielep71/K-TEMPLATE/traffic-history/data/badges/clones.json&style=flat-square)](docs/repo_governance/TRAFFIC.md#badges)
[![Unique visitors, last 14 days](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/danielep71/K-TEMPLATE/traffic-history/data/badges/visitors.json&style=flat-square)](docs/repo_governance/TRAFFIC.md#badges)
[![Daily unique visitors since tracking began](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/danielep71/K-TEMPLATE/traffic-history/data/badges/visits.json&style=flat-square)](docs/repo_governance/TRAFFIC.md#badges)

<br>

**Text-first and reviewable: exported VBA, JSON contracts and Markdown evidence, every change through a pull request**

[What it is](#-what-this-template-is)
&nbsp;·&nbsp;
[What you get](#-what-you-get)
&nbsp;·&nbsp;
[Quick start](#-start-a-new-repository)
&nbsp;·&nbsp;
[Lifecycle](#-from-source-to-release)
&nbsp;·&nbsp;
[Quality gates](#-quality-gates)
&nbsp;·&nbsp;
[Documentation](#-documentation)
&nbsp;·&nbsp;
[Status](#-status)
&nbsp;·&nbsp;
[Template wiki](https://github.com/danielep71/K-TEMPLATE/wiki)

---

</div>

## ✨ What this template is

I use this template to start repositories for regulatory risk calculation
workbooks, such as SA-CCR, SFT, FRTB, LCR and NSFR. A workbook built from it
is not stored as a macro-enabled file: it is built from exported VBA source and
an optional macro-free template, so every line of calculation code is
reviewable in a pull request, and every Excel run leaves evidence that a check
binds to one commit.

The template holds no calculation of its own. It provides the repository around
one: the source layout and house style, the methodology guides to fill in, the
test harness and its evidence, the release process, and the GitHub governance
that keeps them honest.

## 🎯 What you get

| Area | What the template provides |
| --- | --- |
| Source control for Excel | Editing and Git policies for workbooks, exported VBA source and byte-preserved regulatory source copies |
| VBA quality | Static checks of exported VBA (layout, names, `#If` structure, jump targets, visibility, dependency direction, identifiers, `Err` copies, native declarations, calls) and a layout formatter, run on every change |
| Reproducible builds | A macro-enabled workbook built from the source of one commit, without *Trust access to the VBA project object model*, with the export bound to the workbook it came from |
| Testing and evidence | A VBA test harness with a certification entry point, numerical test cases and accuracy contracts, and an offline check that the evidence of an Excel run matches one commit |
| Methodology | Guides for the sources, the data contract, the calculation and workbook contracts and the validation plan, with a traceability table from each rule to its tests |
| Python quality gates | Ruff, strict mypy and actionlint, plus tests that hold the workflows to pinned actions, explicit permissions and the required check |
| Links | Relative links, anchors and HTTPS on every change; external links weekly |
| Issues and pull requests | Issue labels, issue forms, triage automation, a pull request template with a final-head review, and a workflow that keeps milestones on issues only |
| GitHub settings | A documented settings baseline (general settings, security features, rulesets, Actions settings) that a script previews, applies and verifies, writing a dated evidence record |
| Security | A security policy, a review to record before distribution, CodeQL scanning of the Python, JavaScript and workflow code, an OpenSSF Scorecard measure, and Dependabot updates for the pinned Actions and the hash-locked Python dependencies |
| Releases | A release process with a changelog, a certification record, a milestone closeout and a reproducible, attested source bundle |
| Traffic | Optional daily collection of clones and visitors, with the badges above |

## 🚀 Start a new repository

GitHub copies files from a template, not its settings. In order:

1. Create the repository with **Use this template**. GitHub copies the
   template's default branch as it is at that moment, not a release tag:
   record the template commit the new repository came from, or the release
   tag when the default branch equals it, so that later template changes
   can be compared with it.
2. Adapt the files to the new repository, in a clean clone of it. The first
   command lists the changes and writes nothing; `--apply` makes them. Neither
   changes anything on GitHub. `PREFIX` is the project's short name, such as
   `LCR`:

   ```sh
   python tools/initialize_repository.py --repository OWNER/NEW-REPO --regime PREFIX
   python tools/initialize_repository.py --repository OWNER/NEW-REPO --regime PREFIX --apply
   ```

   Add `--release-branch BRANCH` to point Dependabot at an existing release
   branch, and `--remove-traffic-badges` to remove the badges. Review the diff
   and commit it.

   <details>
   <summary><b>What the initializer changes, and what Repository integrity checks</b></summary>

   <br>

   It deletes the template-only section of `AGENTS.md`, with its markers,
   points the two contact links in `.github/ISSUE_TEMPLATE/config.yml` and
   the traffic badges at the top of this README to the new repository,
   removes `target-branch` from `.github/dependabot.yml`, replaces the
   entries and the link references of `CHANGELOG.md` with an empty
   `[Unreleased]` section of the new repository and restores the
   introduction's "No version has been released" sentence if it is missing,
   deletes `VERSION` if the template has one (it exists once a template
   release is prepared) and sets the release rows of *Current status* in
   `RELEASING.md` to say that nothing is released yet, since both describe
   the template's releases, and replaces `<REGIME>` in the naming table of
   `docs/VBA_HOUSE_STYLE.md`. You can then delete
   `tools/initialize_repository.py` and `tests/test_initialize_repository.py`:
   nothing else uses them. Without the initializer, make the same edits by
   hand.

   **Repository integrity** fails while the `AGENTS.md` section remains, a
   contact link or traffic badge names another repository, `target-branch`
   names a branch that does not exist, or a link reference of `CHANGELOG.md`
   points to another repository. Once the changelog holds no dated release,
   it also fails while `VERSION` exists, the introduction lacks the "No
   version has been released" sentence, or *Current status* does not say that
   no version is released and that `VERSION` does not exist. It does not
   check the changelog's entries or `<REGIME>`: replace them by hand, or let
   the initializer do it.

   </details>

3. Enable Actions and let **Repository integrity** pass on the default branch.
4. Apply and verify the settings baseline, as the
   [repository setup guide](docs/setup/REPOSITORY_SETUP.md) describes. It
   needs Python 3.11+ and GitHub CLI authentication; applying needs
   administrator access, the `PLAN SHA-256` that the preview prints and a new
   journal path. Set the Actions settings it reports, which it never
   writes, in **Settings → Actions → General**:

   ```sh
   python tools/setup_repository.py --repo OWNER/NEW-REPO
   python tools/setup_repository.py --repo OWNER/NEW-REPO --apply \
     --approve-plan PLAN_SHA256 --journal docs/repo_governance/settings-journal-YYYY-MM-DD.json
   python tools/setup_repository.py --repo OWNER/NEW-REPO --check \
     --report docs/repo_governance/SETTINGS_EVIDENCE.md \
     --evidence docs/repo_governance/settings-evidence-{date}.json
   ```

5. Run **Sync issue labels** and create milestones with due dates
   ([issue labels](docs/repo_governance/ISSUE_LABELS.md)).
6. Adapt, replace or remove the remaining template files, as the setup
   guide's [handover table](docs/setup/REPOSITORY_SETUP.md#recommended-handover)
   says. The initializer leaves them as they are: among them the title,
   introduction, banner and wiki links of this README, the social preview
   in `docs/assets/`, the repository's **About** fields, and the
   illustrative methodology, cases and contracts.
7. Optionally, enable traffic collection
   ([traffic history](docs/repo_governance/TRAFFIC.md)).
8. If Codex reviews your pull requests, create its environment for the new
   repository ([pull request review](docs/repo_governance/PULL_REQUESTS.md)).

The [settings reference](docs/repo_governance/SETTINGS.md) and its
[JSON configuration](docs/setup/repository-settings.json) describe the baseline.

## 🔄 From source to release

Each step leaves something a later step, or a reviewer, can check:

```mermaid
flowchart LR
    A["Exported VBA source<br/>and contracts"] -->|pull request| B["Repository integrity<br/>static checks and tests"]
    B --> C["Build the workbook<br/>from one commit"]
    C --> D["Excel run<br/>harness certification"]
    D --> E["Evidence checked<br/>against the commit"]
    E --> F["Release<br/>certified and attested"]
```

| Step | Tool or record | What it proves |
| --- | --- | --- |
| Change | Pull request, with the [final-head review](docs/repo_governance/PULL_REQUESTS.md) | The reviewed commit is the one merged |
| Check | `python tools/check_repository.py` | Every **Repository integrity** check passes locally as in CI |
| Build | `tools/devbuild/build_workbook.py` | The workbook comes from the source of one commit, byte for byte reproducible |
| Run | `tools/devbuild/excel_run.py` | The harness certification ran in Excel, with a host record |
| Evidence | `python tools/check_excel_run.py` | The run's logs and results match the commit and the evidence policy |
| Release | [RELEASING.md](RELEASING.md), `python tools/release.py` | The tag, the changelog, the certification and the published assets agree |

## ✅ Quality gates

| Workflow | When it runs | What it checks |
| --- | --- | --- |
| **Repository integrity** | Every push and pull request; the required check | Python, VBA, workflow, link, changelog, test-case and governance checks, with the unit tests; also on the minimum supported Python |
| **CodeQL** | Pushes, pull requests and weekly | Code scanning of the Python, JavaScript and workflow code |
| **OpenSSF Scorecard** | Pushes to the default branch and weekly | The repository's supply-chain posture |
| **Link check** | Weekly | External links |
| **Sync issue labels** · **Detect issue-label drift** | Catalogue changes · daily | The live labels against `.github/labels.json`, never deleting one in use |
| **Detect settings drift** | Daily | The live rulesets and the other settings a read-only token can read against `docs/setup/repository-settings.json`; the rest is listed as not checked |
| **Issue triage** | Issue and milestone events | Assignee, one priority, an open milestone, a target end date and a dependency section on every open issue |
| **Keep milestones issue-only** | Pull request events and daily | Removes any milestone set on a pull request, and lists each removal |
| **Release** | A pushed `vX.Y.Z` tag | Verifies the tag, builds a reproducible source bundle, attests it where the repository allows it and drafts the release |
| **Traffic history** | Daily, optional | Collects clones and visitors for the badges above |

## 🧩 VBA development

- The [repository structure](docs/REPOSITORY_STRUCTURE.md) says where each VBA
  component lives and which may call which.
- The [house style](docs/VBA_HOUSE_STYLE.md) covers naming, contracts, errors
  and cleanup.
- The [public API inventory](docs/PUBLIC_API.txt) lists the supported entry
  points, and every other public declaration under `src/`.
- The [installation guide](INSTALLATION.md) moves source between the
  repository and Excel, and covers upgrades, troubleshooting and removal.

## 📦 Repository structure

```text
.github/   workflows, issue forms, labels, Dependabot and the Excel evidence policy
docs/      methodology, release records, governance, setup, house style, licenses
tests/     VBA harness and example suite, test cases, accuracy contracts, Python tests
tools/     static checks, workbook build and Excel run, release, setup and initializer
src/       the workbook's VBA source and macro-free template (reserved: added by a project)
```

The full layout and its rules are in
[`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md).

## 📚 Documentation

| Document | Covers |
| --- | --- |
| [`docs/methodology/`](docs/methodology/README.md) | Sources, data contract, calculation and workbook contracts, test cases, test harness, validation plan |
| [`INSTALLATION.md`](INSTALLATION.md) | Moving source between the repository and Excel, recording an Excel run, upgrades, troubleshooting and removal |
| [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) | Source layout, reserved paths, dependency direction |
| [`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md) | VBA naming, contracts, errors, cleanup and static checks |
| [`docs/repo_governance/`](docs/repo_governance/README.md) | Settings, issue labels, pull requests, traffic, distribution |
| [`docs/setup/REPOSITORY_SETUP.md`](docs/setup/REPOSITORY_SETUP.md) | Setting up a new repository and what to keep, adapt or remove |
| [`RELEASING.md`](RELEASING.md) | How a version is prepared, certified and published, and the current status |
| [`docs/release/`](docs/release/CLOSEOUT.md) | The milestone [closeout](docs/release/CLOSEOUT.md) and the [release certification](docs/release/RELEASE_CERTIFICATION.md) records, and the [maintenance audit](docs/release/MAINTENANCE_AUDIT.md) done before each release |
| [`CONTRIBUTING.md`](CONTRIBUTING.md) | Workflow, checks, rules of evidence and pull requests |
| [`SECURITY.md`](SECURITY.md) | Private vulnerability reporting, runtime risk surface and data handling |
| [`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md) | Participant behaviour and conduct reporting |
| [`CHANGELOG.md`](CHANGELOG.md) | Change history |
| [Template wiki](https://github.com/danielep71/K-TEMPLATE/wiki) | A guided tour of the template, page by page; the documents above govern where it differs |

## 📌 Status

The release state is recorded in one place: the
[*Current status*](RELEASING.md#current-status) of `RELEASING.md`, which
**Repository integrity** checks against [`CHANGELOG.md`](CHANGELOG.md).
Development continues on the active release branch through task branches and
pull requests; only a release is versioned and tagged.

## 🔐 Data and security

No real portfolio data: never commit or attach real trades, positions,
counterparty data, or any client or employer material. VBA runs with the
user's Office permissions: review the source and follow your organisation's
macro policy. Report suspected vulnerabilities privately as described in
[`SECURITY.md`](SECURITY.md).

## 🧭 Contributing

To take part, read the [contributing guide](CONTRIBUTING.md) and the
[code of conduct](CODE_OF_CONDUCT.md).

## 👤 Author

**Daniele Penza** — [@danielep71](https://github.com/danielep71)

## 📄 License

Repositories created from this template default to the [Mozilla Public License 2.0](LICENSE).
If you choose a different license for a new repository, replace `LICENSE` and keep
[`docs/licenses/`](docs/licenses/THIRD_PARTY_NOTICES.txt): the files listed there keep
the license the notice states for them, MPL-2.0 or MIT.
