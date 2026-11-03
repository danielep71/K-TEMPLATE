<div align="center">

# 🚀 Repository Setup

**From template files to a configured GitHub repository**

[![Scope](https://img.shields.io/badge/Scope-4_settings_%C2%B7_5_security_features_%C2%B7_5_rulesets-00338D?style=flat-square)](#%EF%B8%8F-what-this-applies)
[![Mode](https://img.shields.io/badge/Default-Preview_only-005EB8?style=flat-square)](#%EF%B8%8F-launch-sequence)
[![Python](https://img.shields.io/badge/Python-3.11%2B-3776AB?style=flat-square&logo=python&logoColor=white)](#-before-you-start)
[![Verification](https://img.shields.io/badge/Verification-Read_back_after_apply-00875A?style=flat-square)](#-verification-and-recovery)

[Prepare](#-before-you-start) · [Launch](#%EF%B8%8F-launch-sequence) · [Labels](#%EF%B8%8F-initialize-issue-labels) · [Coverage](#%EF%B8%8F-what-this-applies) · [Clean up](#-after-launch-keep-adapt-or-remove) · [Maintain](#-maintain-the-baseline)

</div>

---

> [!IMPORTANT]
> GitHub's **Use this template** copies files, not repository settings. Run this
> setup for each new destination after its first integrity workflow succeeds.
> The script does not run automatically and never creates or stores credentials.

## 🧭 Before you start

| Prerequisite | Ready when… |
| --- | --- |
| **Runtime** | Python **3.11+** and a recent [GitHub CLI](https://cli.github.com/) are installed. |
| **Destination** | The new repository exists and contains the template files. |
| **Template cleanup** | The template-only section of `AGENTS.md`, between its `template-only` markers, is deleted; the two `contact_links` URLs in `.github/ISSUE_TEMPLATE/config.yml` and the README traffic badges name the new repository, or the badges are removed; `target-branch` in `.github/dependabot.yml` names a branch of the new repository, or is removed; and the entries and link references of `CHANGELOG.md` are the new repository's. `Repository integrity` fails in a generated repository until all five are done. |
| **Identity** | `gh auth status --hostname github.com` shows an account with administrator access to the destination. Sign in with `gh auth login --hostname github.com`, or set `GH_TOKEN` to a fine-grained token. |
| **Token permissions** | A fine-grained token has destination repository **Administration: write**, **Contents: read**, and **Checks: read**. A workflow `GITHUB_TOKEN` cannot administer these rulesets. Never commit tokens. |
| **Integrity check** | Actions is enabled and `Repository integrity` has succeeded on the **current default-branch HEAD**. |
| **Workflow coverage** | The required job runs on every relevant PR without path filters. The supplied workflow runs on all branches, including a renamed default branch. |
| **Merge methods** | Repository-level merge commits and squash merging are enabled. |
| **Availability** | The destination's plan and visibility support repository rulesets. |

## ▶️ Launch sequence

Run from the copied repository directory. Replace `OWNER/NEW-REPO` with the
**new destination**, not the source template.

### ① Preview the changes

```sh
python tools/setup_repository.py --repo OWNER/NEW-REPO
```

Read-only by default. Review the printed destination, visibility, default branch,
and proposed API changes before applying them.

### ② Apply the baseline

```sh
python tools/setup_repository.py --repo OWNER/NEW-REPO --apply
```

The script creates missing rulesets, enables the security features GitHub offers
for the destination's visibility, patches differing general settings, and reads
back the result. Protection is installed before automatic branch deletion is enabled.

### ③ Verify the result

```sh
python tools/setup_repository.py --repo OWNER/NEW-REPO --check \
  --report docs/repo_governance/SETTINGS_EVIDENCE.md
```

A verified match returns **0**. `--report` writes the repository's **setup
record**: a dated Markdown file with each managed setting and security feature as
configured and as GitHub reports it, each managed ruleset with its ID and status,
and any unmanaged or inherited rulesets. `docs/repo_governance/SETTINGS_EVIDENCE.md`
is the conventional place to commit it in a generated repository; add anything the
script cannot read, such as decisions and exceptions, below it. `--report` works with
preview, `--check` and `--apply`; after apply it records the read-back state.
Running apply again against matching settings makes no writes.

Dependabot security updates count as on only when GitHub reports them enabled and
not paused. GitHub pauses them when Dependabot pull requests go unanswered. Paused
updates are drift: the plan prints a note, the record shows `on (paused)`, and an
apply that does not resume them fails its read-back.

<details>
<summary><strong>⚙️ Advanced: reconcile an existing named ruleset</strong></summary>

A differing managed ruleset blocks apply before any writes. After reviewing the
preview, explicitly allow replacement of its entire definition:

```sh
python tools/setup_repository.py --repo OWNER/NEW-REPO --apply --update-existing
```

Replacement removes additional rules or bypasses inside that named ruleset.
Rulesets with other names and inherited organization rules remain unchanged.
Duplicate managed names are rejected.

When you adopt an existing repository, its rulesets may cover the same branches
and tags under other names. The preview names every such unmanaged ruleset with
its target, ref patterns and enforcement. It stays active after apply, with its
own rules and bypasses, so delete the ones that a managed ruleset replaces in
**Settings → Rules → Rulesets**. The script never deletes a ruleset.

The destination is always explicit; it is never inferred from the current directory.
The default JSON path is relative to the script, so it also works from another
working directory. Use `--config PATH` for a reviewed baseline with the same schema.

</details>

## 🏷️ Initialize issue labels

After creating a repository, run **Actions → Sync issue labels → Run workflow** on the default branch. This installs the 20 labels defined in this repository, including exact descriptions and colours. See [labels and automation](../repo_governance/ISSUE_LABELS.md) for triage rules and instructions for replacing my default assignee account in a generated repository. Keep the label configuration, scripts, workflows, forms, and triage tests after setup; they are ongoing automation, not bootstrap leftovers.

## 🛡️ What this applies

The [JSON configuration](repository-settings.json) is the executable counterpart
of the [settings reference](../repo_governance/SETTINGS.md).

| Layer | Managed configuration |
| --- | --- |
| **General · 4 settings** | Delete merged head branches; suggest PR branch updates; disable Projects; disable the legacy downloads flag. |
| **Security · 5 features** | Dependabot alerts and security updates on every repository; private vulnerability reporting, secret scanning and push protection on a public repository (skipped with a note otherwise). |
| **Protection · 5 rulesets** | Default branch, release branches, independent release-branch deletion protection, release tags, and traffic history. |
| **Enforcement** | Active rulesets with no bypass actors or ref exclusions. `Protect main` targets `~DEFAULT_BRANCH`, not a fixed branch name. |
| **Verification** | All pages are read; full rule definitions are compared without response metadata; managed settings are read back after apply. |

**Outside this baseline:** visibility, template mode, Actions policy, Pages,
secrets, environments, labels, collaborators and integrations. Configure these
separately where required. Code scanning is configured by the CodeQL workflow,
not by the script; see *Recommended handover*.
The script creates no branches, tags, releases or PRs. Destination defaults and
inherited policies may impose additional restrictions.

## 🧹 After launch: keep, adapt or remove

> [!TIP]
> **Nothing must be deleted after setup.** Keeping the script and configuration
> lets you run `--check` later and reapply an approved baseline. Removing local
> setup files does not undo settings already saved on GitHub.

### Recommended handover

| File or setting | Action in the new repository | Reason / dependency |
| --- | --- | --- |
| `tools/check_binaries.py`, `tests/test_check_binaries.py`, `.github/binary-allowlist.json` | **Keep** | Fails **Repository integrity** on a tracked binary outside the approved paths. Add a reviewed binary, such as a documentation image, to the allowlist with its SHA-256 and reason. |
| `.gitattributes`, `.gitignore`, `.editorconfig` | **Keep** | Line-ending, binary-file, ignore and editor rules for workbooks, exported VBA source and regulatory source copies. Adjust them only with a layout change that needs it. |
| `.github/labels.json`, `.github/scripts/`, `.github/ISSUE_TEMPLATE/`, the label, triage and milestone workflows, `tools/check_issues.py` | **Keep** | Ongoing label and issue automation; see *Initialize issue labels* above. |
| `AGENTS.md` | **Adapt** | Delete the template-only section first; keep the remaining instructions that still apply to the project. |
| `CONTRIBUTING.md`, `CODE_OF_CONDUCT.md` | **Keep and adapt** | The contribution entry point and the Contributor Covenant 2.1 by reference, both shown in GitHub's community profile. Adapt the checks and reporting contact to the project. |
| `SECURITY.md` | **Keep and adapt** | Names the private reporting channel, response targets and safe harbor. Adjust supported versions and scope to the project. Rewrite *Runtime risk surface* for the project's macros as they are added: for each entry point of the production workbook, what it reads, writes, saves or deletes, and whether the facade, workbook or form code starts a process, creates an external object, calls a native library or reaches the network. |
| `docs/repo_governance/DISTRIBUTION.md` | **Keep and adapt** | The review to record before making the repository public, sharing a copy or publishing a release, with the full-history scan. Add the project's own search patterns. |
| `README.md` | **Adapt** | Replace the MyTemplate title and introduction with the project's purpose, installation and usage. Keep or replace setup links deliberately. |
| `LICENSE` | **Keep or replace** | MPL-2.0 by default. The integrity workflow requires a non-empty `LICENSE`. |
| `docs/VBA_HOUSE_STYLE.md`, `docs/REPOSITORY_STRUCTURE.md`, `docs/PUBLIC_API.txt`, `INSTALLATION.md` | **Keep and adapt** | VBA naming, contracts, errors and cleanup; component locations and dependency direction; the supported API inventory; import, export, upgrade, troubleshooting and removal. Replace `<REGIME>` with the project's short name, and adapt the host policy. Its tested environments are the template's: replace them with the runs of the new repository. |
| `RELEASING.md`, `CHANGELOG.md`, `docs/release/`, `tools/check_changelog.py`, `tests/test_check_changelog.py` | **Keep and adapt** | The release process, the changelog, and the certification and closeout records. Replace the changelog's entries and its link references with the new repository's: **Repository integrity** fails while a link reference points to another repository. Do not create `VERSION` before the first release. |
| `docs/methodology/` | **Keep and adapt** | Sources, data contract, calculation contracts, validation plan, the numerical test-case format, the workbook contract and the VBA test harness guide. Replace each illustrative row with the project's entries as decisions are taken. The *Datasets* table and the field tables of `DATA_CONTRACT.md` are rendered from `docs/methodology/data-contract.json`: replace its illustrative dataset `DS-01` there with the project's datasets, each with its file or sheet, encoding, delimiter and decimal separator and its fields, then run `python tools/render_docs.py --write`, and remove the illustrative file-based case in the same change, as the next row says. **Repository integrity** fails while a table differs from the file, and `python tools/check_data_contract.py --dataset DS-NN FILE` validates a real input file against it. The host behaviour entry `HB-01` concerns `tools/check_excel_run.py`, and `HB-02` to `HB-04` concern `tools/devbuild/excel_run.py` and hold for the host they name; keep them, delete the template's issue and commit references from their *Issue* and *Result* columns, which this repository's history cannot resolve, and record your own host when you measure it. |
| `tests/fixtures/`, `tests/expected/`, `tests/derivations/`, `tests/accuracy/`, `tools/check_test_cases.py`, `tools/check_data_contract.py`, `tools/check_derivations.py`, `tools/check_accuracy.py` and their tests | **Keep and adapt** | The numerical test cases, their format check, the data contract's reference validator, which the format check runs on every variant file of a file-based case, the comparison with the derivation scripts and the accuracy contracts, which **Repository integrity** runs. Remove the `illustrative-format-example` and `illustrative-file-import-example` cases and their scripts, `tests/derivations/illustrative_total.py` and `illustrative_file_total.py`, when you replace the illustrative entries of the methodology guides they use, and replace the illustrative contract `ACC-01` with the project's. In the same change, point `CaseIllustrativeResults` in `tests/modules/TEST_Examples.bas`, which reads those fixtures, at the replacement cases, or replace it and update the suite's counts in `.github/excel-evidence-policy.json` and the guides; `tests/test_vba_harness.py` fails while it reads a fixture that is gone. |
| `tools/check_independent.py`, `tests/test_check_independent.py` | **Keep; register an adapter when wanted** | Runs each independent implementation registered in `tests/independent/adapters.json` on the test cases and compares its results with the expected values; with no adapter, as in the template, it passes and says that nothing was cross-checked. To register one: give the implementation an `IMPL-` entry in `docs/methodology/SOURCES.md`; write an adapter script under `tests/independent/` that translates a fixture into the implementation's input, runs it and prints its outputs, or why it cannot represent the case or is not installed; then add its `source` and `command` to `tests/independent/adapters.json` ([independent implementations](../methodology/TEST_CASES.md#independent-implementations)). An adapter whose implementation is not installed is skipped with a note: install the implementation in the **Repository integrity** workflow for the comparison to run in CI. |
| `docs/licenses/` | **Keep** | Notice and license texts for the files listed in `THIRD_PARTY_NOTICES.txt`, which keep the license the notice states for them (MPL-2.0 or MIT) whatever license you choose. If you remove a listed file, remove it from the notice and from `provenance.json` too. That register records the source repository, license and revision of each adapted file, and a test keeps it in step with the notice; record the revision of any file you adapt, or why it is unknown. |
| `docs/repo_governance/SETTINGS.md` | **Keep and maintain** | Records the adopted baseline. The integrity workflow explicitly requires this file. |
| `docs/setup/repository-settings.json` | **Keep** | Enables future drift checks and reapplication. The setup tests read this file. |
| `tools/setup_repository.py` | **Keep** | Provides preview, apply and verification for future maintenance. |
| `tests/test_setup_repository.py` | **Keep while retaining setup tooling** | Imports the setup script and reads its JSON configuration. |
| `docs/setup/REPOSITORY_SETUP.md` | **Keep or retire after handover** | Optional operational guide; remove links to it if you retire it. |
| `.github/PULL_REQUEST_TEMPLATE.md` | **Keep and adapt** | Records linked issues, validation, final-head review and what was not validated. Add the project's own checks to *Validated*. |
| `.github/workflows/traffic-history.yml`, `tools/traffic_history.py`, `tests/test_traffic_history.py`, `docs/repo_governance/TRAFFIC.md` | **Keep; enable if wanted** | Optional traffic collection. It does nothing until you follow *Enabling* in `TRAFFIC.md`. Point the README's traffic badges at the new repository, or remove them; **Repository integrity** fails while they read another repository. If you never want it, remove all four files and the badges together, and in the same change: remove the files from `docs/licenses/THIRD_PARTY_NOTICES.txt` and `docs/licenses/provenance.json`; remove the two `traffic-history.yml` entries of `WRITE_PERMISSIONS` and the traffic test from `tests/test_workflow_controls.py`; and remove every link to `TRAFFIC.md`, which `python tools/check_links.py` lists. |
| `pyproject.toml`, `tools/requirements-quality-ci.txt` | **Keep** | Ruff and strict mypy settings and their hash-locked versions; `tests/test_quality_gates.py` proves the complexity ceiling, strict mode, the unreachable-code warning and actionlint are active, and `tests/strict_yaml.py` makes the workflow control tests reject duplicate YAML keys. |
| `.github/dependabot.yml`, `.github/requirements.txt` | **Keep** | Proposes updates to the SHA-pinned Actions, the hash-locked configuration parser that the workflow control tests use, and the hash-locked quality tools. Its `target-branch` names the active release branch: set it to this repository's release branch or remove it to target the default branch. **Repository integrity** fails while it names a branch that does not exist. |
| `.github/workflows/release.yml`, `tools/release.py`, `tests/test_release.py` | **Keep** | On a pushed `vX.Y.Z` tag, verifies it against the changelog, `VERSION`, the default branch, its message and **Repository integrity**. It then builds a reproducible source bundle, attests its provenance (on a private repository only with GitHub Enterprise Cloud and the repository variable `RELEASE_ATTESTATIONS` set to `enabled`) and drafts the release. A manual run is a rehearsal. Publication stays manual ([RELEASING.md](../../RELEASING.md)). |
| `.github/workflows/scorecard.yml` | **Keep; public repositories** | Scores the default branch's supply-chain posture with OpenSSF Scorecard weekly and on each push to it, and publishes the results to the Scorecard service and to code scanning. It is skipped on a private repository and is not a required check. If you never want the results published, delete it together with its `WRITE_PERMISSIONS` entry and its test in `tests/test_workflow_controls.py`. |
| `.github/workflows/codeql.yml` | **Keep** | Scans the Python, JavaScript and workflow code with CodeQL's security-extended queries on pushes and pull requests to `main` and `release/**`, and weekly; results appear under **Security → Code scanning**. Pull requests are analyzed without write permission: their SARIF is kept as a 30-day artifact and they show no code scanning annotations. Leave GitHub's default setup off, because it rejects results from this workflow. Free on a public repository; on a private one it needs GitHub Code Security and the repository variable `CODE_SCANNING` set to `enabled`, otherwise it skips the analysis. If the default branch is not `main`, rename it in the workflow's `push` and `pull_request` branch lists; **Repository integrity** fails while they omit it. |
| `tools/check_repository.py`, `tests/test_check_repository.py` | **Keep and evolve** | Runs every **Repository integrity** check, locally and in CI, and writes per-check logs, `report.json` (commit, dirty state, file hashes) and the job summary under `test-results/integrity/`; CI keeps them as a 30-day artifact. Add project checks to its list. |
| `tools/vba_lexer.py`, `tools/check_vba_*.py`, `tests/test_check_vba.py` | **Keep** | The seven static checks of exported VBA source, which **Repository integrity** runs, the call check as advisory. They read the facade prefix from the house style's naming table, so replace `<REGIME>` there. They pass while there is no VBA source. |
| `tests/modules/TEST_Harness.bas`, `tests/modules/TEST_Examples.bas`, `.github/excel-evidence-policy.json`, `tools/check_excel_run.py`, `tests/test_check_excel_run.py`, `tests/test_vba_harness.py` | **Keep** | The VBA test harness, its example suite, the evidence policy and the offline check of Excel-run evidence against it ([harness guide](../methodology/TEST_HARNESS.md#evidence-policy)). Add the project's suites beside the example suite, each with a certification entry point and an entry under `suites` in the evidence policy ([harness guide](../methodology/TEST_HARNESS.md#modes-and-certification)); keep the example suite, whose certification is the release evidence that the harness works on the host. |
| `tests/modules/TEST_Relations.bas` | **Keep and adapt** | The illustrative [relation tests](../methodology/TEST_HARNESS.md#relation-tests), each with its corruption check, on a stand-in total. Replace `IllustrativeTotal` with a call to the core procedure of your calculation and the relations with those its calculation contract implies, keeping a corruption check in every case; `tests/test_vba_harness.py` keeps the counts and the checks consistent, and skips them once the module is gone. Delete the module, and its row in the harness guide's *Modules* table, if no relation applies. |
| `tools/devbuild/`, `tests/test_devbuild.py` | **Keep and evolve** | Builds the workbook from source ([building a workbook from source](../../INSTALLATION.md#building-a-workbook-from-source)), binds an export to it, and records an Excel run with one command (`excel_run.py`). Replace the illustrative sheets and names of `tools/devbuild/workbook.json` together with the workbook contract; the tests keep them equal. |
| `tools/check_workbook_protection.py`, `tests/test_check_workbook_protection.py`, `tests/workbook_template.py` | **Keep** | Compares the protection of `src/workbook/Template.xlsx` with its protection matrix ([workbook contract](../methodology/WORKBOOK_CONTRACT.md#protection)), in **Repository integrity**. It passes while neither exists. |
| `tools/format_vba.py`, `tests/test_format_vba.py` | **Keep** | Brings exported VBA source into the house-style layout with `--write`; **Repository integrity** runs `--check`. It passes while there is no VBA source. |
| `tools/render_docs.py`, `tests/test_render_docs.py` | **Keep and evolve** | Renders the documentation blocks between `BEGIN generated` and `END generated` markers, such as the list of **Repository integrity** checks and the example suite's expected results; **Repository integrity** fails while one is stale. Add a renderer for each fact the project's code already knows, and run `--write` after changing it. |
| `tools/check_links.py`, `tests/test_check_links.py`, `.github/workflows/link-check.yml` | **Keep** | Check every Markdown link: relative links and anchors in **Repository integrity**, external links weekly. |
| `docs/repo_governance/PULL_REQUESTS.md` | **Keep and adapt** | Final-head review procedure that the pull request template refers to. Adapt it if the project uses another automated reviewer. |
| `.github/workflows/repository-integrity.yml` | **Keep and evolve** | Supplies the required `Repository integrity` check. Add project-specific checks as the project grows. |
| GitHub **Template repository** setting | **Review separately** | Choose whether the new project should itself be reusable as a template; the setup script does not change this flag. |

### Optional: retire the setup tooling

Use this route only in a **generated project** where you no longer want the setup
utilities. Keep them in the source template so future repositories receive them.

1. Run `--check` successfully and preserve the result and adopted configuration in
   the setup record.
2. Create a normal task branch and PR for cleanup. Protection remains active after
   setup; do not rely on direct pushes to the default branch.
3. In that same PR, add meaningful project checks to the workflow. Keep the job
   name **`Repository integrity`**, its coverage of all relevant PRs, the steps
   that set up Node.js and install the configuration parser, the quality tools
   and actionlint, the **Run the repository checks** step and the step that
   keeps its reports. Add project checks to `tools/check_repository.py`, and
   keep every check it already runs, as the generated table of
   [CONTRIBUTING.md](../../CONTRIBUTING.md#-local-checks) lists them, the
   template-cleanup checks included. Keep `.github/workflows/codeql.yml` and
   `.github/workflows/scorecard.yml`: they are not part of the setup tooling.
   Every test in `tests/` other than `tests/test_setup_repository.py` stays
   after the setup tests are removed; without those setup and install steps,
   the issue triage, workflow control and quality-gate tests fail.
4. Remove only the retired setup files, together:

   ```text
   tools/setup_repository.py
   tests/test_setup_repository.py
   docs/setup/repository-settings.json
   docs/setup/REPOSITORY_SETUP.md
   ```

5. Update `README.md`, `docs/repo_governance/README.md`,
   `docs/repo_governance/SETTINGS.md` and `CONTRIBUTING.md`: remove links and
   commands that reference deleted files, and explain how settings will be
   maintained going forward. Preserve the actual settings reference. In
   `CONTRIBUTING.md`, replace the link to this guide with the Ruff and mypy
   commands of *Maintain the baseline* below. In `docs/release/CLOSEOUT.md`,
   replace criterion 5, which runs the setup script, with the way you now
   verify the settings. `python tools/check_links.py` lists any link left
   behind.
6. Verify that the cleanup PR reports a passing **`Repository integrity`** check
   with the intended project checks, then merge it through the protected PR process.

> [!WARNING]
> Do not delete the integrity workflow or rename its required job during cleanup.
> GitHub would continue waiting for that check and block merges. Deleting only the
> script or JSON would also break the existing setup tests. Remove individual files,
> not entire `tools/`, `tests/`, `docs/` or `.github/` directories.

Cleanup is documented as a reviewed PR because replacing template tests with useful
project validation depends on the new project. There is no automatic deletion step.

## 🔎 Verification and recovery

| Exit code | Meaning | Next action |
| --- | --- | --- |
| **0** | Successful preview, or verified match in apply/check mode | Read the mode-specific message: preview success does not mean settings were applied. |
| **1** | Drift found by `--check` | Preview the differences and decide whether to reapply. |
| **2** | Invalid configuration, missing prerequisite, unavailable feature, API failure or read-back mismatch | Resolve the reported problem, then preview and retry. |

API operations are **not transactional**. If a later request fails, earlier changes
remain. There is no silent skip, no success claim for partial application and no
automatic rollback of protection. Rerun preview to see what remains.

A passing check on the current HEAD is a prerequisite, not proof of future workflow
coverage or destructive enforcement testing. Verification covers the documented
settings, not every GitHub setting or inherited policy.

## 🔄 Maintain the baseline

Update the JSON and Markdown together when changing policy, review the preview,
and run the offline tests before applying it. The tests also need Node.js, which
runs the label scripts. Install the pinned configuration parser once; without it
the workflow control tests are skipped locally:

```sh
python -m pip install --require-hashes -r .github/requirements.txt
python -m unittest discover -s tests -v
```

After that one-time installation, the tests require neither credentials nor network access.
**Toolchain baseline.** The workflows set exact versions: Python 3.13.18, Python
3.11.18 for the minimum-version job (the `python_version` in `pyproject.toml`), and
Node.js 24.22.2 without the package-manager cache. Dependabot does not update
these values: change them together, in every workflow, in one pull request. The
workflow control tests reject floating versions and mixed versions.

**Action pins.** Every `uses:` reference reachable from a workflow is
immutable:

- a remote action or reusable workflow, called from a step or from a job, is
  pinned to a full commit SHA;
- a Docker image (`docker://`) is pinned to an `@sha256:` digest;
- each pin carries its version as a comment on the same line
  (`# v7.0.1`), which Dependabot and reviewers read;
- a local reusable workflow must exist in `.github/workflows/`, and is checked
  as a workflow in its own right;
- a local action (`./path`) is followed into its `action.yml` or `action.yaml`,
  whose steps and image follow the same rules.

`tests/test_workflow_controls.py` fails on any other form, and on a reference
it cannot read with its comment, such as a flow mapping. A job that calls a
reusable workflow is exempt from the `timeout-minutes` rule, which GitHub does
not allow there.

**Repository integrity** also runs Ruff on `tools/` and `tests/`, strict mypy on
`tools/` (both configured in `pyproject.toml`, with `warn_unreachable` on top of
strict mode) and actionlint on the workflows. To run Ruff and mypy locally on
CPython 3.13 for Linux x86_64, the platform the hashes in
`tools/requirements-quality-ci.txt` cover:

```sh
python -m pip install --only-binary=:all: --require-hashes -r tools/requirements-quality-ci.txt
ruff check tools tests
mypy
```

The tests already check every relative link and anchor. To also request the
external links, as the weekly **Link check** workflow does:

```sh
python tools/check_links.py --external
```

Once setup tooling has been retired from a generated project, use that project's
replacement validation and settings-maintenance process instead.

---

**References:** [Repository updates](https://docs.github.com/en/rest/repos/repos#update-a-repository)
· [Rulesets](https://docs.github.com/en/rest/repos/rules)
· [GitHub CLI API](https://cli.github.com/manual/gh_api)

[↑ Back to top](#-repository-setup)
