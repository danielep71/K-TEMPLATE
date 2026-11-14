<div align="center">

# 🚀 Repository Setup

**From template files to a configured GitHub repository**

[![Scope](https://img.shields.io/badge/Scope-12_settings_%C2%B7_5_security_features_%C2%B7_5_rulesets-00338D?style=flat-square)](#%EF%B8%8F-what-this-applies)
[![Mode](https://img.shields.io/badge/Default-Preview_only-005EB8?style=flat-square)](#%EF%B8%8F-launch-sequence)
[![Python](https://img.shields.io/badge/Python-3.11%2B-3776AB?style=flat-square&logo=python&logoColor=white)](#-before-you-start)
[![Verification](https://img.shields.io/badge/Verification-Read_back_after_apply-00875A?style=flat-square)](#-verification-and-recovery)

[Prepare](#-before-you-start) · [Launch](#%EF%B8%8F-launch-sequence) · [Labels](#%EF%B8%8F-initialize-issue-labels) · [Coverage](#%EF%B8%8F-what-this-applies) · [Clean up](#-after-launch-keep-adapt-or-remove) · [Maintain](#-maintain-the-baseline)

</div>

---

> [!IMPORTANT]
> GitHub's **Use this template** copies files, not repository settings. Run this
> setup for each new destination after its first integrity workflow succeeds.
> No workflow applies the baseline: **Detect settings drift** only checks it,
> with a read-only token. The script never creates or stores credentials.

## 🧭 Before you start

| Prerequisite | Ready when… |
| --- | --- |
| **Runtime** | Python **3.11+** and a recent [GitHub CLI](https://cli.github.com/) are installed. |
| **Destination** | The new repository exists and contains the template files. **Use this template** copies the template's default branch as it is then, not a release tag: record the template commit it came from. |
| **Template cleanup** | The template-only section of `AGENTS.md`, between its `template-only` markers, is deleted; the two `contact_links` URLs in `.github/ISSUE_TEMPLATE/config.yml` and the README traffic badges name the new repository, or the badges are removed; `target-branch` in `.github/dependabot.yml` names a branch of the new repository, or is removed; and every link reference of `CHANGELOG.md` points to the new repository. `Repository integrity` fails in a generated repository until all five are done. Replace the changelog's entries with an empty `[Unreleased]` section, and restore the introduction's "No version has been released" sentence if it is missing; delete `VERSION` if the template has one, since it exists once a template release is prepared; reset the *Current status* of `RELEASING.md` to no release; replace `<REGIME>` in `docs/VBA_HOUSE_STYLE.md`; and remove the template's evidence, which cites its own commits and issues: the rows of the *Tested environment* table of `INSTALLATION.md`, and the *Issue* column of the *Host behaviour* table of `docs/methodology/VALIDATION_PLAN.md`, which becomes a dash. Once the changelog holds no dated release, `Repository integrity` fails while `VERSION` exists, the sentence is missing or *Current status* does not say that no version is released and that `VERSION` does not exist; it does not check the changelog's entries or `<REGIME>`. In a clean clone, `python tools/initialize_repository.py --repository OWNER/NEW-REPO --regime PREFIX` lists all these edits and makes them with `--apply`; see *Recommended handover* for its options. Without it, make them by hand. |
| **Identity** | `gh auth status --hostname github.com` shows an account that can read the destination; `--apply` needs one with administrator access to it. Sign in with `gh auth login --hostname github.com`, or set `GH_TOKEN` to a fine-grained token. |
| **Token permissions** | Preview and `--check` only read. A fine-grained token needs destination repository **Metadata: read**, and **Administration: read** for Dependabot alerts, Dependabot security updates and the five Actions settings. GitHub reports secret scanning and push protection only to administrators of the repository, and a ruleset's bypass list only to a token that can write the ruleset; without it, each managed ruleset that exists is unknown, and a missing one is reported as missing. A value the token cannot read is shown as unknown, with the endpoint and HTTP status, and counts as drift. `--apply` needs administrator access and a token with **Administration: write**, **Contents: read** and **Checks: read**. A workflow `GITHUB_TOKEN` cannot administer these rulesets; `--check --token-scope workflow` compares what it can read ([scheduled drift check](../repo_governance/SETTINGS.md#drift)). Never commit tokens. |
| **Integrity check** | Actions is enabled and `Repository integrity` has succeeded on the **current default-branch HEAD**. |
| **Workflow coverage** | The required job runs on every relevant PR without path filters. The supplied workflow runs on all branches, including a renamed default branch. |
| **Availability** | The destination's plan and visibility support repository rulesets. |

## ▶️ Launch sequence

Run from the copied repository directory. Replace `OWNER/NEW-REPO` with the
**new destination**, not the source template. Each command is on one line, so
it can be pasted into PowerShell or `sh` alike.

Run it on a task branch created from the default branch, such as
`docs/settings-evidence`:

```sh
git switch -c docs/settings-evidence
```

Step ② writes its journal, and step ③ its report and evidence, into
`docs/repo_governance/`, as files to commit. Step ② also creates `Protect main`,
after which the default branch accepts changes only through a pull request.
Commit the three files on the task branch and merge them through a pull request.

### ① Preview the changes

```sh
python tools/setup_repository.py --repo OWNER/NEW-REPO
```

Read-only by default, and it needs no administrator access. Review the printed
destination, visibility and default branch, then one `DIFF` line for each managed
value that differs from the configuration, and the proposed API changes that
follow them, before applying them:

```text
DIFF setting has_projects: observed true, configured false
DIFF security vulnerability_alerts: observed false, configured true [GET vulnerability-alerts (HTTP 404: off)]
DIFF ruleset "Protect main" (12345) rules.pull_request.allowed_merge_methods: observed ["merge", "squash"], configured ["merge"]
```

A security feature or Actions setting that GitHub does not report to your token is
`null`, with the endpoint and status in brackets; it counts as drift.

A configured ruleset that no ruleset of its name exists for, but that other
active rulesets already enforce at least as strictly, under any names, is
**covered**: the preview prints a `NOTE` line that names them instead of a `DIFF`,
it is not drift, and nothing is planned or written for it. A ruleset counts
towards coverage only when it has the same target, a reported and empty bypass
list, no excluded refs and every configured ref in its include (`~ALL`, or
`~DEFAULT_BRANCH` for the default branch); each configured rule must be present in
one of them with parameters at least as strict: more required approvals, fewer
allowed merge methods, a superset of the required checks. Different rules may be
covered by different rulesets. Coverage is judged on the rulesets as the plan
leaves them: a named ruleset that the plan replaces counts with its configured
definition, not with what its live definition enforces beyond it, and missing
rulesets are created before differing ones are replaced. A named ruleset with
stronger parameters covers itself and is `covered`, not `differs`; a weaker one
is `differs` and is replaced whatever other rulesets enforce. Otherwise the named
ruleset is `matches`, `missing` or `unknown`, as the `DIFF` lines show.

The plan ends with its digest, `PLAN SHA-256: <hex>`: the SHA-256 of the
repository name, the configuration's SHA-256, every observed managed value and
the exact planned writes. The same unchanged repository and configuration give the
same digest; `--check` prints it too, and `--report` and `--evidence` record it.
Apply needs it.

### ② Apply the baseline

```sh
python tools/setup_repository.py --repo OWNER/NEW-REPO --apply --approve-plan PLAN_SHA256 --journal docs/repo_governance/settings-journal-YYYY-MM-DD.json
```

Replace `PLAN_SHA256` with the digest that the preview printed. Apply reads the
repository again, rebuilds the plan and computes its digest before any write; when
it differs from the approved one, because the repository or the configuration
changed since the reviewed plan, apply refuses and asks you to preview again. An
approved plan with no steps is a valid apply: the journal is written, nothing is
written to GitHub, and the read-back runs as usual.

`--journal PATH` names a new file; an existing one is refused before any write.
The journal is JSON with the plan digest, the repository, the configuration's
SHA-256, the start time, each planned write with its method, endpoint and status
(`pending`, `done` or `failed`, with the error text), and the read-back result.
A ruleset write also names its ruleset, and once done records the ID that GitHub
returned for it. The journal is written atomically before the first write and
again after each write's result, so it exists before anything reaches GitHub. A
failing write stops the sequence and is recorded; nothing is retried or rolled
back. Keep the journal with the setup record, for example beside
`SETTINGS_EVIDENCE.md` as the path above does.

The script first sets the three merge methods where they differ, because GitHub
may reject a ruleset that allows a disabled one. It then creates missing rulesets,
replaces differing ones when you allow it, sets the security features GitHub offers for the destination's
visibility, patches the other differing general settings, and reads back the
result. Protection is installed before automatic branch deletion is enabled, so a
merged pull request cannot delete a release branch before its deletion rule exists.

Apply refuses before any write when the plan's digest differs from the approved
one, the journal path exists, you lack administrator access, the destination
is archived or disabled, its visibility differs from a `visibility` the
configuration sets, or a managed general setting, security feature or ruleset is unknown.
It never changes visibility: follow the
[distribution review](../repo_governance/DISTRIBUTION.md) first.

The script never writes the five Actions settings. When one still differs after
the rest is applied, the read-back fails and names each: set it in the
repository's **Settings → Actions → General**, then rerun `--check`.

### ③ Verify the result

```sh
python tools/setup_repository.py --repo OWNER/NEW-REPO --check --report docs/repo_governance/SETTINGS_EVIDENCE.md --evidence 'docs/repo_governance/settings-evidence-{date}.json'
```

A verified match returns **0**. `--report` writes the repository's **setup
record**: a dated Markdown file with the plan digest, each managed general setting, security feature
and Actions setting as configured and as GitHub reports it, each managed ruleset with its ID and status,
the rulesets that cover it included, and any unmanaged or inherited rulesets. `docs/repo_governance/SETTINGS_EVIDENCE.md`
is the conventional place to commit it in a generated repository; add anything the
script cannot read, such as decisions and exceptions, below it. `--report` works with
preview, `--check` and `--apply`; after apply it records the read-back state.
Running apply again against matching settings makes no writes.

`--evidence PATH` writes the same check as JSON, for a machine to read, whether or
not you also pass `--report`; a `{date}` in the path becomes the UTC date of the
check (`YYYY-MM-DD`). Keep such a path in single quotes, as above: PowerShell
does not pass the braces of an unquoted `{date}` through. The record names the
mode, repository, default branch and result, the configuration's path and
SHA-256, and each managed value as configured
and observed, with the source of each security feature and Actions setting, each
managed ruleset with its differences and the rulesets that cover it, the unmanaged
and inherited rulesets, the method and endpoint of each planned change, with the
ruleset's name for a ruleset change but without its payload, and the plan digest
(`plan_sha256`). It works in the same modes as `--report`. A public-only feature
or Actions setting on another visibility is recorded as
`not applied (public repositories only)` and does not count as drift.

`--summary PATH` appends a short Markdown summary of the comparison to a file,
such as `$GITHUB_STEP_SUMMARY`, in every mode.

The **Detect settings drift** workflow runs the check daily with its own
read-only token:

```sh
python tools/setup_repository.py --repo OWNER/NEW-REPO --check --token-scope workflow
```

`--token-scope workflow` compares only what a workflow `GITHUB_TOKEN` with
`contents: read` can read: each ruleset without its bypass list, the visibility,
the general settings GitHub returns to it and, on a public repository, private
vulnerability reporting. It never calls the endpoints that need **Administration:
read**. Each value it cannot read is printed as `NOT CHECKED` and is neither drift
nor a match: a value GitHub leaves out of its answer, or refuses with HTTP 401,
403 or 404. Any other failed request, such as a server error, is an API failure
that returns **2**. Drift in
what it reads returns **1**; otherwise it returns **0** with the result `PARTIAL`
while something is not checked. It plans nothing, prints no
digest and takes neither `--report` nor `--evidence`: the setup record comes only
from a check with an administrator's token. The
[settings reference](../repo_governance/SETTINGS.md#drift) lists what each section
needs.

Dependabot security updates count as on only when GitHub reports them enabled and
not paused. GitHub pauses them when Dependabot pull requests go unanswered. Paused
updates are drift: the plan prints a note, the record shows `on (paused)`, and an
apply that does not resume them fails its read-back.

<details>
<summary><strong>⚙️ Advanced: reconcile an existing named ruleset</strong></summary>

A named ruleset that differs and is not covered blocks apply before any writes.
After reviewing the preview, explicitly allow replacement of its entire definition:

```sh
python tools/setup_repository.py --repo OWNER/NEW-REPO --apply --update-existing --approve-plan PLAN_SHA256 --journal docs/repo_governance/settings-journal-YYYY-MM-DD.json
```

Replacement removes additional rules or bypasses inside that named ruleset; a
named ruleset that is `covered` is never replaced. Rulesets with other names and
inherited organization rules remain unchanged. Duplicate managed names are rejected.

When you adopt an existing repository, its rulesets may protect the same branches
and tags under other names. One that enforces a configured ruleset with no ruleset
of its name at least as strictly covers it, and the preview says so. The preview also names every unmanaged
ruleset with its target, ref patterns and enforcement. It stays active after
apply, with its own rules and bypasses, so delete the ones that a managed ruleset
replaces in **Settings → Rules → Rulesets**. The script never edits or deletes an
unmanaged ruleset.

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
| **General · 12 settings** | The three merge methods, auto-merge, branch update suggestions, deletion of merged head branches, web commit sign-off, and the Issues, Projects, Wiki, Discussions and legacy downloads features. Four differ from GitHub's defaults: delete merged head branches, suggest PR branch updates, disable Projects and the legacy downloads flag. |
| **Security · 5 features** | Dependabot alerts and security updates on every repository; private vulnerability reporting, secret scanning and push protection on a public repository (skipped with a note otherwise). |
| **Actions · 5 settings, checked only** | SHA pinning required, read-only default workflow permissions, no pull request approval by workflows, approval for every external contributor's fork pull request workflows (public repository only) and 90-day retention. Reported and verified, never written. |
| **Protection · 5 rulesets** | Default branch, release branches, independent release-branch deletion protection, release tags, and traffic history. The configuration may hold any number of uniquely named rulesets. Active rulesets without bypass actors or excluded refs that enforce a configured one at least as strictly, under any names, cover it: nothing is written for it. |
| **Visibility · optional** | `null` in the template: any visibility. Set to `public` or `private`, a different visibility is drift that apply refuses and never changes. |
| **Enforcement** | Active rulesets with no bypass actors or ref exclusions. `Protect main` targets `~DEFAULT_BRANCH`, not a fixed branch name. |
| **Verification** | All pages are read; full rule definitions are compared without response metadata; apply writes only the plan whose digest was approved, journaling each write; managed settings are read back after apply. |

**Outside this baseline:** template mode, the other Actions policies, Pages,
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
| `docs/assets/social-preview.svg`, `docs/assets/social-preview.png` | **Replace or remove** | The template's social preview and README banner, which name the template. Edit the SVG with the project's name and subtitle, export a 1280×640 PNG or larger at 2:1, under 1 MB, and update its SHA-256 in `.github/binary-allowlist.json`. Or delete both, the allowlist entry and the image at the top of `README.md`. |
| Repository **About** and **Social preview** | **Set** | GitHub copies neither from a template. Set the description, website and topics in the repository page's **About** settings, and upload the PNG in **Settings → General → Social preview**; GitHub offers no API for the preview. |
| `.gitattributes`, `.gitignore`, `.editorconfig` | **Keep** | Line-ending, binary-file, ignore and editor rules for workbooks, exported VBA source and regulatory source copies. Adjust them only with a layout change that needs it. |
| `.github/labels.json`, `.github/scripts/`, `.github/ISSUE_TEMPLATE/`, the label, triage and milestone workflows, `tools/check_issues.py`, `tests/test_check_issues.py`, `tests/test_labels_sync.py`, `docs/repo_governance/ISSUE_LABELS.md` | **Keep** | Ongoing label and issue automation and its guide; see *Initialize issue labels* above. |
| `tests/test_repository_consistency.py` | **Keep** | Cross-file checks in **Repository integrity**: the label guide against the catalogue, with every label's name, colour and description, each label once in *Using the labels*, and every label count the Markdown files state; Dependabot's labels against the catalogue and its assignee against the bug form's; office-type ignore rules for regulatory source copies; `.editorconfig` against `.gitattributes`, so that editors trim the trailing white space `git diff --check` reports in Markdown and leave alone the files Git keeps byte for byte; closed code fences in Markdown, and `--repository` in every `release.py verify-published` command outside the changelog; one GitHub API version; the third-party notice against the files and its provenance register; and where the template may name other repositories, skipped once the template-only section is gone. Adapt a check only together with the files it compares. |
| `tools/initialize_repository.py`, `tests/test_initialize_repository.py` | **Remove once run** | Makes the *Template cleanup* edits of *Before you start*: it also replaces the changelog's entries and `<REGIME>`, restores the changelog's "No version has been released" sentence if it is missing, deletes `VERSION` if it exists, resets the *Current status* of `RELEASING.md` and removes the template's evidence rows, as one diff to review and commit. Without `--apply` it lists them and writes nothing. With `--apply` it refuses a working tree that differs from the commit, and a second run changes nothing. It never changes GitHub or commits. By default it repoints the traffic badges, removes Dependabot's `target-branch` and links the changelog's `[Unreleased]` to `main`; `--remove-traffic-badges`, `--release-branch BRANCH` (an existing branch) and `--default-branch BRANCH` change that. Its test initializes a copy of the template and runs **Repository integrity** on it with `--template false`; it is skipped once the template-only section is gone. Delete both files together after committing the result: nothing else uses them. |
| `AGENTS.md` | **Adapt** | Delete the template-only section first; keep the remaining instructions that still apply to the project. |
| `CONTRIBUTING.md`, `CODE_OF_CONDUCT.md` | **Keep and adapt** | The contribution entry point and the Contributor Covenant 2.1 by reference, both shown in GitHub's community profile. Adapt the checks and reporting contact to the project. |
| `SECURITY.md` | **Keep and adapt** | Names the private reporting channel, response targets and safe harbor. Adjust supported versions and scope to the project. Rewrite *Runtime risk surface* for the project's macros as they are added: for each entry point of the production workbook, what it reads, writes, saves or deletes, and whether the facade, workbook or form code starts a process, creates an external object, calls a native library or reaches the network. |
| `docs/repo_governance/DISTRIBUTION.md` | **Keep and adapt** | The review to record before making the repository public, sharing a copy or publishing a release, with the full-history scan. Add the project's own search patterns. |
| `README.md` | **Adapt** | Replace the template's title and introduction with the project's purpose, installation and usage. Keep or replace setup links deliberately. |
| `LICENSE` | **Keep or replace** | MPL-2.0 by default. The integrity workflow requires a non-empty `LICENSE`. |
| `docs/VBA_HOUSE_STYLE.md`, `docs/REPOSITORY_STRUCTURE.md`, `docs/PUBLIC_API.txt`, `INSTALLATION.md` | **Keep and adapt** | VBA naming, contracts, errors and cleanup; component locations and dependency direction; the inventory of the supported and project-public VBA declarations; import, export, upgrade, troubleshooting and removal. Replace `<REGIME>` with the project's short name, and adapt the host policy. The initializer removes the template's tested environments: record the runs of the new repository. |
| `RELEASING.md`, `CHANGELOG.md`, `docs/release/`, `tools/check_changelog.py`, `tests/test_check_changelog.py` | **Keep and adapt** | The release process, the changelog, and the certification and closeout records. Replace the changelog's entries and its link references with the new repository's: **Repository integrity** fails while a link reference points to another repository. Keep or restore the introduction's "No version has been released" sentence, delete `VERSION` if the template has one (it exists once a template release is prepared), and reset *Current status* in `RELEASING.md` to no release; **Repository integrity** fails while these disagree with a changelog that holds no dated release. Create `VERSION` when you prepare the first release. |
| `docs/methodology/` | **Keep and adapt** | Sources, data contract, calculation contracts, validation plan, the numerical test-case format, the workbook contract and the VBA test harness guide. Replace each illustrative row with the project's entries as decisions are taken. The *Datasets* table and the field tables of `DATA_CONTRACT.md` are rendered from `docs/methodology/data-contract.json`: replace its illustrative dataset `DS-01` there with the project's datasets, each with its file or sheet, encoding, delimiter and decimal separator and its fields, then run `python tools/render_docs.py --write`. In the same change, remove both illustrative cases and their row of the *Traceability* table, and replace `ACC-01`, as the next row says. **Repository integrity** fails while a table differs from the file, and `python tools/check_data_contract.py --dataset DS-NN FILE` validates a real input file against it. The host behaviour entry `HB-01` concerns `tools/check_excel_run.py`, `HB-02` to `HB-04` concern `tools/devbuild/excel_run.py` and hold for the host they name, `HB-05` concerns `tools/devbuild/build_workbook.py` and the identity by rebuild of `tools/check_excel_run.py`, and `HB-06` and `HB-07` concern the harness and installation guides; keep them, delete the template's issue and commit references from their *Issue* and *Result* columns, which this repository's history cannot resolve, and record your own host when you measure it. |
| `tests/fixtures/`, `tests/expected/`, `tests/derivations/`, `tests/accuracy/`, `tools/check_test_cases.py`, `tools/check_data_contract.py`, `tools/check_derivations.py`, `tools/check_accuracy.py` and their tests | **Keep and adapt** | The numerical test cases, their format check, the data contract's reference validator, which the format check runs on every variant file of a file-based case, the comparison with the derivation scripts and the accuracy contracts, which **Repository integrity** runs. Remove the `illustrative-format-example` and `illustrative-file-import-example` cases, their scripts, `tests/derivations/illustrative_total.py` and `illustrative_file_total.py`, and the row of the *Traceability* table in `docs/methodology/VALIDATION_PLAN.md` that names them, when you replace the illustrative entries of the methodology guides they use, and replace the illustrative contract `ACC-01` with the project's. In the same change, point `CaseIllustrativeResults` in `tests/modules/TEST_Examples.bas`, which reads those fixtures, at the replacement cases, or replace it and update the suite's counts in `.github/excel-evidence-policy.json` and the guides; `tests/test_vba_harness.py` fails while it reads a fixture that is gone, or while the results it records differ from the outputs of the contracts' cases. |
| `tools/check_independent.py`, `tests/test_check_independent.py` | **Keep; register an adapter when wanted** | Runs each independent implementation registered in `tests/independent/adapters.json` on the test cases and compares its results with the expected values; with no adapter, as in the template, it passes and says that nothing was cross-checked. To register one: give the implementation an `IMPL-` entry in `docs/methodology/SOURCES.md`; write an adapter script under `tests/independent/` that translates a fixture into the implementation's input, runs it and prints its outputs, or why it cannot represent the case or is not installed; then create `tests/independent/adapters.json`, which the template does not ship, or add to it the adapter's `source` and `command` ([independent implementations](../methodology/TEST_CASES.md#independent-implementations)). An adapter whose implementation is not installed is skipped with a note: install the implementation in the **Repository integrity** workflow for the comparison to run in CI. |
| `docs/licenses/` | **Keep** | Notice and license texts for the files listed in `THIRD_PARTY_NOTICES.txt`, which keep the license the notice states for them (MPL-2.0 or MIT) whatever license you choose. If you remove a listed file, remove it from the notice and from `provenance.json` too. That register records the source repository, license and revision of each adapted file, and a test keeps it in step with the notice; record the revision of any file you adapt, or why it is unknown. |
| `docs/repo_governance/SETTINGS.md` | **Keep and maintain** | Records the adopted baseline. The integrity workflow explicitly requires this file. |
| `docs/setup/repository-settings.json` | **Keep** | Enables future drift checks and reapplication. The setup tests read this file. |
| `tools/setup_repository.py` | **Keep** | Provides preview, apply and verification for future maintenance. |
| `.github/workflows/settings-drift.yml` | **Keep** | Runs `setup_repository.py --check --token-scope workflow` daily with the read-only workflow token, and fails on drift in the rulesets and the other values that token can read; it lists the rest as not checked ([scheduled drift check](../repo_governance/SETTINGS.md#drift)). It fails until you apply the baseline. Remove it together with the setup tooling, as *Optional: retire the setup tooling* says. |
| `tests/test_setup_repository.py` | **Keep while retaining setup tooling** | Imports the setup script and reads its JSON configuration. |
| `docs/setup/REPOSITORY_SETUP.md` | **Keep or retire after handover** | Optional operational guide; remove links to it if you retire it. |
| `.github/PULL_REQUEST_TEMPLATE.md` | **Keep and adapt** | Records linked issues, validation, final-head review and what was not validated. Add the project's own checks to *Validated*. |
| `.github/workflows/traffic-history.yml`, `tools/traffic_history.py`, `tests/test_traffic_history.py`, `docs/repo_governance/TRAFFIC.md` | **Keep; enable if wanted** | Optional traffic collection. It does nothing until you follow *Enabling* in `TRAFFIC.md`. Point the README's traffic badges at the new repository, or remove them; **Repository integrity** fails while they read another repository. If you never want it, remove all four files and the badges together, and in the same change: remove the files from `docs/licenses/THIRD_PARTY_NOTICES.txt` and `docs/licenses/provenance.json`; remove the two `traffic-history.yml` entries of `WRITE_PERMISSIONS` and the two traffic tests, `test_traffic_app_credentials_stay_in_their_steps` and `test_traffic_alert_reads_the_commit_its_collection_wrote`, from `tests/test_workflow_controls.py`; and remove every link to `TRAFFIC.md`, which `python tools/check_links.py` lists. |
| `pyproject.toml`, `tools/requirements-quality-ci.txt`, `tests/test_quality_gates.py`, `tests/strict_yaml.py` | **Keep** | Ruff and strict mypy settings and their hash-locked versions; `tests/test_quality_gates.py` proves the complexity ceiling, strict mode, the unreachable-code warning and actionlint are active, and `tests/strict_yaml.py` makes the workflow control tests reject duplicate YAML keys. |
| `.github/dependabot.yml`, `.github/requirements.txt` | **Keep** | Proposes updates to the SHA-pinned Actions, the hash-locked configuration parser for the tests that read YAML and the `codeql-default-branch` check, and the hash-locked quality tools. Its `target-branch` names the active release branch: set it to this repository's release branch or remove it to target the default branch. **Repository integrity** fails while it names a branch that does not exist. |
| `.github/workflows/release.yml`, `tools/release.py`, `tests/test_release.py` | **Keep** | On a pushed `vX.Y.Z` tag, verifies it against the changelog, `VERSION`, the default branch, its message and **Repository integrity**; when the tagged commit holds a VBA component, the message must declare its Excel evidence, which `release.py evidence` checks and `release.py verify` checks again on the machine of the run before the push, or an invariant 4 deviation. It then builds a reproducible source bundle, attests its provenance (on a private repository only with GitHub Enterprise Cloud and the repository variable `RELEASE_ATTESTATIONS` set to `enabled`) and drafts the release. A manual run is a rehearsal. Publication stays manual; afterwards, `release.py verify-published` checks the published release ([RELEASING.md](../../RELEASING.md)). |
| `.github/workflows/scorecard.yml` | **Keep; public repositories** | Scores the default branch's supply-chain posture with OpenSSF Scorecard weekly and on each push to it, and publishes the results to the Scorecard service and to code scanning. It is skipped on a private repository and is not a required check. If you never want the results published, delete it together with its `WRITE_PERMISSIONS` entry and its test in `tests/test_workflow_controls.py`. |
| `.github/workflows/codeql.yml` | **Keep** | Scans the Python, JavaScript and workflow code with CodeQL's security-extended queries on pushes and pull requests to `main` and `release/**`, and weekly; results appear under **Security → Code scanning**. Pull requests are analyzed without write permission: their SARIF is kept as a 30-day artifact and they show no code scanning annotations. Leave GitHub's default setup off, because it rejects results from this workflow. Free on a public repository; on a private one it needs GitHub Code Security and the repository variable `CODE_SCANNING` set to `enabled`, otherwise it skips the analysis. If the default branch is not `main`, rename it in the workflow's `push` and `pull_request` branch lists; **Repository integrity** fails while they omit it. |
| `tools/check_repository.py`, `tests/test_check_repository.py`, `.github/actionlint/invalid-workflow.yml` | **Keep and evolve** | Runs every **Repository integrity** check, locally and in CI, and writes per-check logs, `report.json` (commit, dirty state, file hashes) and the job summary under `test-results/integrity/`; CI keeps them as a 30-day artifact. With `--ci`, which CI passes, it fails when the working tree differs from the commit. The `actionlint` check first lints `.github/actionlint/invalid-workflow.yml`, a workflow with a known error that never runs, and fails unless actionlint rejects it. Add project checks to its list. |
| `tools/vba_lexer.py`, `tools/check_vba_*.py`, `tests/test_check_vba.py` | **Keep** | The nine static checks of exported VBA source, which **Repository integrity** runs, the call check as advisory. They read the facade prefix from the house style's naming table, so replace `<REGIME>` there. They pass while there is no VBA source. |
| `docs/native-declarations.json`, `tools/check_vba_declares.py`, `tests/test_check_vba_declares.py` | **Keep and adapt** | The [register of native declarations](../REPOSITORY_STRUCTURE.md#native-declarations) and the static check that reads it. Add an entry for each export that a `Declare` imports, in the pull request that adds it, with its ABI in the optional `arguments` and `returns` keys; **Repository integrity** fails on a `Declare` the register does not list, or that differs from its export's ABI in a compilation environment. The register is empty in the template, and the check passes while there is no VBA source. |
| `tests/modules/TEST_Harness.bas`, `tests/modules/TEST_Examples.bas`, `.github/excel-evidence-policy.json`, `tools/check_excel_run.py`, `tests/test_check_excel_run.py`, `tests/test_vba_harness.py` | **Keep** | The VBA test harness, its example suite, the evidence policy and the offline check of Excel-run evidence against it ([harness guide](../methodology/TEST_HARNESS.md#evidence-policy)). Add the project's suites beside the example suite, each with a certification entry point and an entry under `suites` in the evidence policy ([harness guide](../methodology/TEST_HARNESS.md#modes-and-certification)); keep the example suite, whose certification is the release evidence that the harness works on the host. |
| `tools/vba_code_identity.py`, `tests/test_vba_code_identity.py` | **Keep** | Compares two versions of a VBA component by their code, without comments, for `tools/check_excel_run.py --carried-to`, which carries a run's evidence to a pull request commit that changes only comments or layout ([carried evidence](../../INSTALLATION.md#carried-evidence)). It refuses non-ASCII input and any component that uses `[` or `]` outside a string. Carried evidence never certifies a release. |
| `tests/modules/TEST_Relations.bas` | **Keep and adapt** | The illustrative [relation tests](../methodology/TEST_HARNESS.md#relation-tests), each with its corruption check, on a stand-in total. Replace `IllustrativeTotal` with a call to the core procedure of your calculation and the relations with those its calculation contract implies, keeping a corruption check in every case; `tests/test_vba_harness.py` keeps the counts and the checks consistent, and skips them once the module is gone. Delete the module, and its row in the harness guide's *Modules* table, if no relation applies. |
| `tools/devbuild/`, `tests/test_devbuild.py`, `tests/test_excel_run.py` | **Keep and evolve** | Builds the workbook from source ([building a workbook from source](../../INSTALLATION.md#building-a-workbook-from-source)), binds an export to it, and records an Excel run with one command (`excel_run.py`). Replace the illustrative sheets and names of `tools/devbuild/workbook.json` together with the workbook contract; the tests keep them equal. |
| `tools/check_workbook_protection.py`, `tests/test_check_workbook_protection.py`, `tests/workbook_template.py` | **Keep** | Compares the protection of `src/workbook/Template.xlsx` with its protection matrix ([workbook contract](../methodology/WORKBOOK_CONTRACT.md#protection)), in **Repository integrity**. It passes while neither exists. |
| `tools/format_vba.py`, `tests/test_format_vba.py` | **Keep** | Brings exported VBA source into the house-style layout with `--write`; **Repository integrity** runs `--check` on the tracked modules. It passes while there is no VBA source. If you import modules that are not in layout, list them in a `tools/vba_layout_baseline.json` [baseline](../VBA_HOUSE_STYLE.md#source-representation) instead of reformatting them all at once; the check then fails on any other module out of layout and on a listed module back in layout, so the list only shrinks. |
| `tools/make_mutant.py`, `tests/test_make_mutant.py` | **Keep** | Writes a mutant, a copy of a source file with one text replaced, under the ignored `test-results/mutants/` at the source's relative path, for the [negative control](../../CONTRIBUTING.md#-negative-controls) of a defect fix. It refuses unless the text occurs exactly once, keeps every other byte, CRLF line endings and `Attribute VB_Name` included, and prints the changed lines and the SHA-256 of the source and of the mutant. It never writes over a source file or into a folder that Git does not ignore. |
| `tools/render_docs.py`, `tests/test_render_docs.py` | **Keep and evolve** | Renders the documentation blocks between `BEGIN generated` and `END generated` markers, such as the list of **Repository integrity** checks and the example suite's expected results; **Repository integrity** fails while one is stale. Add a renderer for each fact the project's code already knows, and run `--write` after changing it. |
| `tools/check_links.py`, `tests/test_check_links.py`, `.github/workflows/link-check.yml` | **Keep** | Check every Markdown link: relative links, anchors and HTTPS in **Repository integrity**, external links weekly, with optional expiring exceptions in `.github/link-exceptions.json` ([maintain the baseline](#-maintain-the-baseline)). |
| `tools/check_doc_commands.py`, `tests/test_check_doc_commands.py` | **Keep** | Checks, in **Repository integrity**, that each `python tools/...` command and repository `python -m` command in the tracked Markdown runs a tracked script, and that its subcommands and options exist in the script's `argparse` parser, read from its source without running it. A parser it cannot read is a failure, not a pass. Commands of a **Remove once run** tool are skipped once its file is deleted. Keep each tool's parser literal so it stays readable. |
| `docs/repo_governance/README.md` | **Keep and adapt** | The index of the governance guides and the repository controls. Update its links when you add, retire or move a guide. |
| `docs/repo_governance/PULL_REQUESTS.md` | **Keep and adapt** | Final-head review procedure that the pull request template refers to. Adapt it if the project uses another automated reviewer. |
| `.github/workflows/repository-integrity.yml` | **Keep and evolve** | Supplies the required `Repository integrity` check. Add project-specific checks as the project grows. |
| `.github/workflows/closing-keywords.yml`, `tools/check_closing_keywords.py`, `tests/test_check_closing_keywords.py` | **Keep** | On each pull request, finds the closing keywords it would carry to the default branch ([release branches](../repo_governance/PULL_REQUESTS.md#-release-branches)). Into a release branch, any keyword in a commit, the title or the body fails; into the default branch from a release branch, a keyword that names an open issue fails unless the body lists the issue on a `Reopen after merge: #N` line. Any other pull request is reported as not checked. It reads issue states with the workflow token and `issues: read` and `pull-requests: read`, outside **Repository integrity**, and is not a required check. |
| `tests/test_workflow_controls.py` | **Keep and evolve** | Governance checks of the workflows and Dependabot, read with the pinned configuration parser: pinned Actions and credential-free checkouts, timeouts, explicit permissions and the jobs listed in `WRITE_PERMISSIONS`, the unconditional required check, concurrency, hash-locked installs, shared toolchain versions, Dependabot coverage, the guards of the label, closing keyword, CodeQL, Scorecard, settings drift, traffic and milestone workflows, and the milestone workflow's script, which clears milestones from pull requests only. Update it in the same change as the workflow it checks. The tests of the traffic alert script skip once `.github/workflows/traffic-history.yml` is removed. |
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
   the issue triage, workflow control and quality-gate tests fail. The test
   that compares the *Remove once run* rows of this guide with
   `tools/check_doc_commands.py` is skipped once the guide is deleted.
4. Remove only the retired setup files, together:

   ```text
   tools/setup_repository.py
   tests/test_setup_repository.py
   docs/setup/repository-settings.json
   docs/setup/REPOSITORY_SETUP.md
   .github/workflows/settings-drift.yml
   ```

   In the same change, remove
   `test_settings_drift_checks_on_a_schedule_with_the_read_only_token` from
   `tests/test_workflow_controls.py`.

5. Update `README.md`, `docs/repo_governance/README.md`,
   `docs/repo_governance/SETTINGS.md` and `CONTRIBUTING.md`: remove links and
   commands that reference deleted files, and explain how settings will be
   maintained going forward. Preserve the actual settings reference. In
   `CONTRIBUTING.md`, replace the link to this guide with the Ruff and mypy
   commands of *Maintain the baseline* below. In `docs/release/CLOSEOUT.md`,
   replace criterion 5, which runs the setup script, with the way you now
   verify the settings. `python tools/check_links.py` lists any link left
   behind, and `python tools/check_doc_commands.py` any command of a deleted
   tool.
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
| **0** | Successful preview, verified match in apply/check mode, or no drift in what `--token-scope workflow` read | Read the mode-specific message: preview success does not mean settings were applied, and `PARTIAL` means that the values listed as `NOT CHECKED` were not compared. |
| **1** | Drift found by `--check`, an unknown value included; with `--token-scope workflow`, drift in a value it read | Preview the differences and decide whether to reapply, or grant the access the unknown value needs. |
| **2** | Invalid configuration, a plan whose digest is not the approved one, an existing journal path, refused or failed apply prerequisite, unavailable feature, API failure or read-back mismatch, an Actions setting to set by hand included | Resolve the reported problem, then preview and retry. |

API operations are **not transactional**. If a later request fails, earlier changes
remain. There is no silent skip, no success claim for partial application and no
automatic rollback of protection. The journal records which writes were done and
which one failed; rerun preview to see what remains.

A passing check on the current HEAD is a prerequisite, not proof of future workflow
coverage or destructive enforcement testing. Verification covers the documented
settings, not every GitHub setting or inherited policy.

## 🔄 Maintain the baseline

Update the JSON and Markdown together when changing policy, review the preview,
and run the offline tests before applying it. The tests also need Node.js, which
runs the label scripts. Install the pinned configuration parser once; without it
the tests that read YAML and the `codeql-default-branch` check are skipped
locally:

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

**Repository integrity** also runs Ruff and strict mypy on `tools/` and
`tests/` (both configured in `pyproject.toml`, with `warn_unreachable` on top of
strict mode) and actionlint on the workflows, once actionlint has rejected
`.github/actionlint/invalid-workflow.yml`. A test file with a `# mypy:`
comment that turns off the annotation rules need not annotate its functions,
but mypy still checks their code. To run Ruff and mypy locally on
CPython 3.13 for Linux x86_64, the platform the hashes in
`tools/requirements-quality-ci.txt` cover:

```sh
python -m pip install --only-binary=:all: --require-hashes -r tools/requirements-quality-ci.txt
ruff check tools tests
mypy
```

The tests already check every relative link and anchor, and that every
external link uses HTTPS: a plain `http://` link in a Markdown link, a
reference definition, an autolink or an HTML attribute fails **Repository
integrity**; text in code spans and fenced blocks is not checked. To also
request the external links, over HTTPS only, as the weekly **Link check**
workflow does:

```sh
python tools/check_links.py --external
```

Links into this repository itself are skipped, locally as in the workflow: the
repository is read from `GITHUB_REPOSITORY`, else from the `origin` remote.
A 403, 429 or 5xx answer, or no answer, is a warning, because the site may
only be busy or refuse automated requests. To keep such a link while you
review it again later, list it in `.github/link-exceptions.json`, which the
template does not ship; create it with the first exception:

```json
{
  "schema_version": 1,
  "exceptions": [
    {"url": "https://example.org/page", "reason": "Answers 403 to automated requests", "expires": "2027-06-30"}
  ]
}
```

The `url` is the link without its `#fragment`. Up to and including the
`expires` date the warning stays a warning; after it, the weekly check fails
for that URL until you renew or remove the exception. **Repository integrity**
fails while the file is malformed or lists a URL that no Markdown file links.

Once setup tooling has been retired from a generated project, use that project's
replacement validation and settings-maintenance process instead.

---

**References:** [Repository updates](https://docs.github.com/en/rest/repos/repos#update-a-repository)
· [Rulesets](https://docs.github.com/en/rest/repos/rules)
· [Actions permissions](https://docs.github.com/en/rest/actions/permissions)
· [GitHub CLI API](https://cli.github.com/manual/gh_api)

[↑ Back to top](#-repository-setup)
