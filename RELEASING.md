# 🚀 Releasing

How a version is prepared and published. The branch model and the integration
pull requests are in the
[pull request guide](docs/repo_governance/PULL_REQUESTS.md#-release-branches);
the version and changelog rules are in
[CHANGELOG.md](CHANGELOG.md#date-and-version-rules).

> [!IMPORTANT]
> Nothing is versioned, tagged or published unless I decide it. Closing a
> milestone, or integrating a release branch into the default branch,
> publishes nothing. Unreleased work stays under `## [Unreleased]`.

<a id="current-status"></a>

## 🧭 Current status

| Item | State |
| --- | --- |
| Released versions | 1.0.0 |
| `VERSION` | `1.0.0` |
| Active release branch | The branch `.github/dependabot.yml` names in `target-branch`, if any |

A version is released from the preparation pull request that dates its
changelog section ([step 1](#1-finalize-the-changelog-and-version)). It is
published once its tag and its GitHub Release exist (steps 6 and 7).

<a id="versioning"></a>

## 🏷️ Versioning

| Item | Convention |
| --- | --- |
| Version | Semantic Versioning, `MAJOR.MINOR.PATCH`; `0.y.z` during initial development ([rules](CHANGELOG.md#date-and-version-rules)) |
| What the version covers | Calculation results, input and output layouts, and the supported VBA API in [docs/PUBLIC_API.txt](docs/PUBLIC_API.txt) |
| Release branch | `release/X.Y.Z`, without `v`, opened only when I decide to |
| Milestone | `vX.Y.Z - <title>`, on issues only |
| `VERSION` | One `X.Y.Z` line at the root. Created in the release preparation pull request, never before; afterwards it always names the newest released version, published or not ([current status](#current-status)) |
| Changelog | Work stays under `## [Unreleased]`; preparation moves it to `## [X.Y.Z] - YYYY-MM-DD` |
| Tag | Annotated `vX.Y.Z` on the candidate commit on the default branch. The **Protect release tags** ruleset blocks its deletion, updates and force pushes ([settings](docs/repo_governance/SETTINGS.md)) |
| GitHub Release | Created from the tag, with the changelog section and the evidence summary as notes |

**Repository integrity** runs `tools/check_changelog.py`, which ties the two
files together: no `VERSION` before a dated release, and from then on
`VERSION` equal to the newest one.

## 🧱 Preparation and publication

A release has two parts, each decided separately:

| Part | Steps | Effect |
| --- | --- | --- |
| **Preparation** | 1 to 5 | Changes only the repository: the dated changelog section, `VERSION`, the integration into the default branch and the evidence. It can stop at any point and publishes nothing. |
| **Publication** | 6 to 8 | Creates the tag and the GitHub Release, and checks what a user downloads. A pushed tag is never moved or deleted. |

<a id="release-records"></a>

The release certification, the distribution review and the
[milestone closeout](docs/release/CLOSEOUT.md) of version X.Y.Z are issues of
the next open milestone, the one after milestone N, never of N itself.
`verify-published --milestone N` passes only once milestone N is closed with
no open issue, and **Issue triage** rejects an open issue in a closed
milestone; records filed under N could then never be completed or closed.
Open the next milestone first if there is none. Milestone N closes in step 7,
before the GitHub Release is published, once its closeout checklist is
complete for the candidate; the checklist's first criterion is that every
issue of N is closed.

<a id="release-invariants"></a>

## 🔒 Release invariants

A version may be published only when:

1. one exact candidate commit is frozen: the merge commit that integrates the
   release branch into the default branch;
2. the changelog section for the version is dated and matches what the
   candidate contains;
3. **Repository integrity** passes on the candidate, including the
   [static checks](docs/VBA_HOUSE_STYLE.md#static-checks) of its VBA source,
   and every CodeQL alert
   open on the default branch has been fixed or assessed in the
   certification record;
4. the candidate's VBA is imported, compiled and run in Excel on the
   [supported host](INSTALLATION.md#%EF%B8%8F-host-policy), the run is
   [recorded](INSTALLATION.md#recording-an-excel-run) from a workbook built
   from the candidate, or from an earlier commit whose run inputs are
   identical ([step 4](#4-certify-the-workbook)), and passes
   `tools/check_excel_run.py` with a `VERIFIED` binding, never on carried
   evidence, and the result is in
   the [release certification](docs/release/RELEASE_CERTIFICATION.md);
5. the [distribution review](docs/repo_governance/DISTRIBUTION.md) for the
   release is recorded and closed;
6. the [maintenance audit](docs/release/MAINTENANCE_AUDIT.md) was run before
   the candidate was frozen, its report is on the milestone closeout issue,
   and each of its findings is fixed, tracked in an issue or accepted as a
   deviation;
7. the annotated tag targets the candidate; and
8. the checks after publication pass.

If the source changes after steps 3 to 5, the evidence is stale and is
produced again for the new candidate.

## 🪜 Preparing and publishing a version

Start with the [maintenance audit](docs/release/MAINTENANCE_AUDIT.md) of the
repository's issues, pull requests, review discussions and notifications, and
post its report on the milestone closeout issue. Fix its findings through task
pull requests before step 2 freezes the candidate, or track each in an issue
and record it as a deviation.

### 1. Finalize the changelog and `VERSION`

In a task pull request into the release branch:

- move the `[Unreleased]` entries into `## [X.Y.Z] - YYYY-MM-DD`, and add its
  link reference, `[X.Y.Z]: https://github.com/OWNER/REPOSITORY/releases/tag/vX.Y.Z`;
- keep an empty `## [Unreleased]` section, with its link reference changed to
  `https://github.com/OWNER/REPOSITORY/compare/vX.Y.Z...HEAD`;
- create `VERSION`, or update it, to `X.Y.Z`;
- update *Current status* above, and remove the changelog's "No version has
  been released" sentence for the first release.

`tools/check_changelog.py` fails until these agree with the newest release. It
also fails on a placeholder line, such as `TBD`, in a changelog section, and
on a "Release Candidate" mention or an unchecked task-list item in
`README.md`. From step 6 on, the section is frozen: once the tag exists, the
check fails when the section differs from its content at the tag.

### 2. Integrate and freeze the candidate

Integrate the release branch into the default branch and sync it back, as the
[pull request guide](docs/repo_governance/PULL_REQUESTS.md#-release-branches)
describes. The integration merge commit is the candidate: record its full
SHA. Until the tag exists, the release branch takes only fixes for a gate
that failed on the candidate. Each fix records its change in the dated
section being released and is integrated again; the new merge commit becomes
the candidate.

### 3. Run the repository checks

From a clean clone of the candidate:

```sh
git switch --detach FULL_SHA
git status --short
python tools/check_repository.py --strict --ci --repository OWNER/REPOSITORY \
  --default-branch DEFAULT_BRANCH --template false
```

Use `--template true` in the template repository itself. `git status --short`
must print nothing, `--ci` fails the run if the checks leave a file changed,
and `test-results/integrity/report.json` must name
`FULL_SHA` with `"dirty": false` and `"status": "pass"`; keep it with the
certification record. Confirm that **Repository
integrity** is green on the same SHA on the default branch, and review the
open alerts under **Security → Code scanning**.

Complete the [milestone closeout](docs/release/CLOSEOUT.md) for this
candidate: its candidate record, file audit and settings check, and the rest
of its checklist. Records made on an earlier commit do not count. The one
exception is the Excel run of step 4, and only when
`tools/check_excel_run.py --equivalent-to` proves its run inputs identical.

### 4. Certify the workbook

Copy the [release certification](docs/release/RELEASE_CERTIFICATION.md)
into an issue for the release, in the next milestone
([release records](#release-records)), build the workbook from the candidate as
described below, run the regression cases of the
[validation plan](docs/methodology/VALIDATION_PLAN.md), and fill in the
record from the runs themselves.

Build the development workbook from the candidate with
`tools/devbuild/build_workbook.py`
([building a workbook from source](INSTALLATION.md#building-a-workbook-from-source)).
[Importing VBA into Excel](INSTALLATION.md#%EF%B8%8F-importing-vba-into-excel)
by hand is only for a project with forms, which the build does not write, or
when the build fails. Such a workbook has no build stamp, so its run cannot
reach `Binding: VERIFIED`, and while the commit holds
`tests/accuracy/contracts.json` it ends `excel-run: FAIL`, because its results
name no commit ([recording an Excel run](INSTALLATION.md#recording-an-excel-run),
steps 3 and 9 of *By hand*). Such a run cannot certify a release.
Record the run of the [VBA test harness](docs/methodology/TEST_HARNESS.md)
as [recording an Excel run](INSTALLATION.md#recording-an-excel-run)
describes: with `python tools/devbuild/excel_run.py`, which builds, runs,
checks and proves identity by rebuild, or by hand, with the export bound to
the workbook by `tools/devbuild/export_manifest.py`. Then check the evidence
folder against the candidate (the command does this itself):

```sh
python tools/check_excel_run.py test-results/excel-run --commit CANDIDATE_SHA \
  --workbook test-results/build/Workbook-development.xlsm --require-binding
```

It must end with `Binding: VERIFIED` and `excel-run: PASS`: the workbook, or
the export, came from a clean build of the candidate, and the run is
consistent with it. This check is a release gate, not a
**Repository integrity** step, because the evidence exists only for a
candidate. If a case or the check fails, fix it through a task pull request
and start again at step 2.

A run made on an earlier commit certifies the candidate only when every run
input of the candidate has the same Git blob as at the run's commit, for
example when the candidate adds only a fix to a workflow or a guide outside
`docs/methodology/`. Check it from a clean checkout of the candidate, with
`--commit` naming the run's commit
([a run for another commit](INSTALLATION.md#a-run-for-another-commit)):

```sh
python tools/check_excel_run.py test-results/excel-run --commit RUN_SHA \
  --equivalent-to CANDIDATE_SHA \
  --workbook test-results/build/Workbook-development.xlsm --require-binding
```

It must also print
`Equivalent: CANDIDATE_SHA (N run inputs identical to RUN_SHA)`. The
certification record names both commits. Any other earlier run does not
count. In particular, carried evidence does not certify a candidate: a run
that `--carried-to` accepts because the candidate's VBA components differ
from the run's only in comments or layout
([carried evidence](INSTALLATION.md#carried-evidence)), and a report with a
`Carried from` line, are evidence for a pull request only. After such a
change, record a new run of the candidate.

If the repository keeps a [user manual](#user-manual), build its PDF from the
candidate as that section describes, and record it, with the PDF tool and its
version, in *Artifacts* of the certification record.

### 5. Review the distribution

Publishing a GitHub Release is distribution. Record the
[distribution review](docs/repo_governance/DISTRIBUTION.md) for the
candidate and close it.

### 6. Create the annotated tag

Only once I decide to publish, from the command line so that the tag is
annotated:

```sh
git fetch origin
git tag -a vX.Y.Z FULL_SHA -F ../tag-message-vX.Y.Z.txt
git cat-file -t vX.Y.Z        # must print: tag
git push origin vX.Y.Z
```

The tag message is the release's permanent record:

```text
vX.Y.Z - <milestone title>

Candidate <full 40-character SHA>.
Repository integrity: run <run id> PASS.
Excel: <host, build, bitness>; certification #<issue>.
Distribution review: #<issue>.
Deviations: none.
```

`Deviations:` is mandatory. Write `none` after it on the same line, as above,
or list each invariant that was not met on the lines that follow it, one
`- ` or `* ` item per line:

```text
Deviations:
- Invariant 8: <what was not met, and why>.
```

`Deviations:` with nothing after it fails. Copy results from the runs
themselves.

Pushing the tag starts the **Release** workflow
(`.github/workflows/release.yml`). It stops before building anything if any of
these fails:

- the tag is annotated and named `vX.Y.Z`;
- `X.Y.Z` is the newest dated release of `CHANGELOG.md` and the content of
  `VERSION`;
- the tagged commit is on the default branch;
- the message names that commit on the `Candidate` line and has a
  `Deviations:` line in one of the two forms above;
- **Repository integrity** succeeded on the commit.

The workflow then builds `<repository>-X.Y.Z-source.tar.gz` and `SHA256SUMS`
from the tag with `tools/release.py`; the same tag always gives the same bytes.
It attests their build provenance with GitHub's OIDC-signed attestation, and
creates a **draft** release with the tag message and the changelog section as
notes. A private repository attests only on GitHub Enterprise Cloud, with the
repository variable `RELEASE_ATTESTATIONS` set to `enabled`. Without the
variable, the draft is created without an attestation, and its notes say so.
With the variable but without Enterprise Cloud, the attestation fails and no
draft is created: leave the variable unset there. To rehearse before tagging,
run the workflow by hand on the candidate's ref: it checks the changelog and
builds the bundle as an artifact, with no attestation and no release.

### 7. Publish and verify

Close milestone N on GitHub once the
[milestone closeout](docs/release/CLOSEOUT.md) checklist is complete for the
candidate: every one of its issues is closed, as its first criterion
requires, every other row links its evidence or records a deviation, and the
[release records](#release-records) are in the next milestone. Then review
the draft release that the **Release** workflow created. Title it
`vX.Y.Z - <milestone title>`, and add the evidence summary to its notes.
Attach any other artifact the certification record lists, such as the
production workbook built from the candidate or the
[user manual](#user-manual) PDF, with its SHA-256. Then publish it, and
check the published release with `gh` signed in:

```sh
python tools/release.py verify-published vX.Y.Z --repository OWNER/REPOSITORY \
  --milestone N --asset FILE=SHA256
```

Give `--milestone` the number of the release's milestone, and `--asset` once
for each other attached artifact, with the SHA-256 of the certification
record. The command only reads, through `gh`: it prints one line per fact
and ends with `release-verify-published: PASS (vX.Y.Z)` only when:

- the tag is annotated and points at the commit its `Candidate` line names;
- **Repository integrity** succeeded on that commit;
- the release is published, not a draft, and its notes name that commit;
- its assets are exactly `<repository>-X.Y.Z-source.tar.gz`, `SHA256SUMS`
  and each `--asset`; the bundle's SHA-256 is the one `SHA256SUMS` records,
  and each `--asset` has its recorded SHA-256;
- `gh attestation verify` passes on the bundle, as built by the **Release**
  workflow at the tag. When the notes state that the bundle is not attested,
  the command prints a `NOTE` instead: record the absence as a deviation;
- the milestone is closed and has no open issue.

Keep its output with the certification record. Then check that a fresh clone
of the tag passes the commands of step 3.

### 8. Verify the published workbook

When the release attaches a workbook, check it as a user gets it, and record
the results in the *Published workbook* section of the
[release certification](docs/release/RELEASE_CERTIFICATION.md). Give the
workbook and the manual to `verify-published` in step 7 as `--asset`, so that
it compares their published bytes with the record. It cannot download them as
a user does, unblock them or open them in Excel; these checks stay manual:

- **Binding.** The attached workbook is the production build of the
  candidate, as the development workbook is bound in step 4. Its build stamp
  names the candidate's full SHA and `production`, without
  *NOT evidence-grade*. From a clean checkout of the tag, rebuild it:

  ```sh
  python tools/devbuild/build_workbook.py --configuration production
  ```

  The SHA-256 of `test-results/build/Workbook-production.xlsm` must equal the
  attached workbook's. If they differ, unzip both packages and compare their
  parts: another platform's zlib can compress the same parts into other
  bytes, but any difference in a part fails the check.
  `tools/check_excel_run.py --require-binding` proves this binding only for
  the development workbook of an Excel run.
- **Download.** Download the workbook from the release page, as a user
  would. Before unblocking it, check that its SHA-256 matches the record.
  Unblock it as your organization's policy allows, without changing a macro
  setting, and open it on the [supported host](INSTALLATION.md#%EF%B8%8F-host-policy):
  - **Edit Links** (or **Workbook Links**) on the **Data** tab lists no link
    to another workbook;
  - **Formulas → Name Manager** shows the names of the
    [workbook contract](docs/methodology/WORKBOOK_CONTRACT.md#defined-names),
    and none that refers to `#REF!`;
  - on a copy, each operation the workbook contract lists runs on synthetic
    data; the copy is saved, closed and reopened, and opens with no error,
    still without broken links or `#REF!` names.
- **Manual.** If a [user manual](#user-manual) PDF is attached, download it
  too: its SHA-256 matches the record, and it opens and names version
  `X.Y.Z`.

If a check fails, do not replace the asset: follow *After publication* under
[Recovery](#-recovery).

<a id="user-manual"></a>

## 📘 User manual

A recipient who gets only the workbook, not the repository, needs a manual
for the version they hold. It is optional; by convention it lives under
`docs/user-manual/`, one folder per released version,
`docs/user-manual/X.Y.Z/MANUAL.md`, with any image beside it. Create the
folder with the manual itself, never as an empty placeholder. Start a
version's folder from a copy of the previous one, and describe only the
workbook of that version: what it calculates and its limits, the supported
host, how to open it under a macro policy, its inputs, operations, outputs
and failure messages, how to read its build stamp, and how to report a
problem. Use synthetic data in examples and screenshots; an image needs an
entry in `.github/binary-allowlist.json`.

The PDF is built with the certification (step 4), attached to the release
in step 7, and never committed. The template installs no PDF tool: choose one and record it, with its version, in
the certification record. With [Pandoc](https://pandoc.org/) and a PDF engine
it uses, such as a LaTeX distribution, from a clean checkout of the
candidate:

```sh
pandoc docs/user-manual/X.Y.Z/MANUAL.md --resource-path docs/user-manual/X.Y.Z \
  -o test-results/<repository>-X.Y.Z-user-manual.pdf
```

Before attaching it, check that it opens and names version `X.Y.Z`.

## 🗂️ Evidence record

Keep with the release, in its certification issue and the release notes:

- the maintenance audit report;
- the candidate SHA and the integration pull request;
- the **Repository integrity** run and the CodeQL assessment;
- the Excel run, with the host it names;
- the distribution review;
- the tag name and its target;
- the output of `tools/release.py verify-published`;
- the checks of the published workbook and the user manual, if attached;
- every deviation.

Logs and artifacts expire, so copy the decisive lines.

## 🧯 Recovery

- **Before the tag:** fix through a task pull request and repeat from step 2.
- **After the tag, before the Release:** stop and record the problem. A
  pushed tag is never moved or deleted: the **Protect release tags** ruleset
  blocks both and has no bypass. Fix the problem in a new patch version. The
  rulesets are relaxed only to remove sensitive data from the history, as
  [handling a finding](docs/repo_governance/DISTRIBUTION.md#-handling-a-finding)
  describes.
- **After publication:** never move the tag or replace an asset. Publish a
  corrected patch version and record the problem in the changelog; report a
  vulnerability as [SECURITY.md](SECURITY.md) describes. Leave the released
  section as published. Only a demonstrable error in it, such as a wrong file
  name or link, is corrected there, in a reviewed pull request that adds
  `Corrected after release on YYYY-MM-DD: <reason>` to the section;
  `tools/check_changelog.py` then reports the difference as a note instead of
  failing.

## ✅ Checklist

| # | Gate | Done |
| ---: | --- | :---: |
| 1 | I decided to prepare the version | ☐ |
| 2 | Maintenance audit reported on the closeout issue; each finding fixed, tracked or accepted | ☐ |
| 3 | Changelog section dated and `VERSION` created or updated | ☐ |
| 4 | Release branch integrated into the default branch, and synced back | ☐ |
| 5 | Candidate frozen: the integration merge commit's full SHA | ☐ |
| 6 | Repository checks and **Repository integrity** pass on the candidate; CodeQL alerts assessed | ☐ |
| 7 | Release certification recorded for the candidate | ☐ |
| 8 | Distribution review recorded and closed | ☐ |
| 9 | I decided to publish | ☐ |
| 10 | Annotated tag on the candidate | ☐ |
| 11 | Milestone closeout checklist complete for the candidate, then milestone N closed with no open issue; the release records are in the next milestone | ☐ |
| 12 | GitHub Release published from the tag, and `tools/release.py verify-published` passes | ☐ |
| 13 | Published workbook bound to the candidate and checked as downloaded, or none attached | ☐ |
