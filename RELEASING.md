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
| Released versions | None |
| `VERSION` | Does not exist yet; it is created when the first release is prepared |
| Active release branch | The branch `.github/dependabot.yml` names in `target-branch`, if any |

<a id="versioning"></a>

## 🏷️ Versioning

| Item | Convention |
| --- | --- |
| Version | Semantic Versioning, `MAJOR.MINOR.PATCH`; `0.y.z` during initial development ([rules](CHANGELOG.md#date-and-version-rules)) |
| What the version covers | Calculation results, input and output layouts, and the supported VBA API in [docs/PUBLIC_API.txt](docs/PUBLIC_API.txt) |
| Release branch | `release/X.Y.Z`, without `v`, opened only when I decide to |
| Milestone | `vX.Y.Z - <title>`, on issues only |
| `VERSION` | One `X.Y.Z` line at the root. Created in the release preparation pull request, never before; afterwards it always names the newest released version |
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
| **Publication** | 6 and 7 | Creates the tag and the GitHub Release. A tag that may have been fetched is never moved or deleted. |

<a id="release-invariants"></a>

## 🔒 Release invariants

A version may be published only when:

1. one exact candidate commit is frozen: the merge commit that integrates the
   release branch into the default branch;
2. the changelog section for the version is dated and matches what the
   candidate contains;
3. **Repository integrity** passes on the candidate, and every CodeQL alert
   open on the default branch has been fixed or assessed in the
   certification record;
4. the candidate's VBA is imported, compiled and run in Excel on the
   [supported host](INSTALLATION.md#%EF%B8%8F-host-policy), and the run is recorded in
   the [release certification](docs/release/RELEASE_CERTIFICATION.md);
5. the [distribution review](docs/repo_governance/DISTRIBUTION.md) for the
   release is recorded and closed;
6. the annotated tag targets the candidate; and
7. the checks after publication pass.

If the source changes after steps 3 to 5, the evidence is stale and is
produced again for the new candidate.

## 🪜 Preparing and publishing a version

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
`README.md`.

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
python -m unittest discover -s tests
python tools/check_changelog.py
python tools/check_test_cases.py
python tools/check_derivations.py
python tools/check_links.py
```

`git status --short` must print nothing. Confirm that **Repository
integrity** is green on the same SHA on the default branch, and review the
open alerts under **Security → Code scanning**.

### 4. Certify the workbook

Copy the [release certification](docs/release/RELEASE_CERTIFICATION.md)
into an issue for the release, build the workbook from the candidate as
[INSTALLATION.md](INSTALLATION.md#%EF%B8%8F-importing-vba-into-excel)
describes, run the regression cases of the
[validation plan](docs/methodology/VALIDATION_PLAN.md), and fill in the
record from the runs themselves. If a case fails, fix it through a task pull
request and start again at step 2.

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

`Deviations:` is mandatory: write `none`, or list each invariant that was not
met. Copy results from the runs themselves.

### 7. Publish and verify

Create the GitHub Release from the tag, titled `vX.Y.Z - <milestone title>`,
with the changelog section and the evidence summary as notes. Attach only the
artifacts the certification record lists, with their SHA-256. Then check:

- a fresh clone of the tag passes the commands of step 3;
- the release page names the tag and the candidate SHA;
- each attached artifact's SHA-256 matches the record.

## 🗂️ Evidence record

Keep with the release, in its certification issue and the release notes:

- the candidate SHA and the integration pull request;
- the **Repository integrity** run and the CodeQL assessment;
- the Excel run, with the host it names;
- the distribution review;
- the tag name and its target;
- every deviation.

Logs and artifacts expire, so copy the decisive lines.

## 🧯 Recovery

- **Before the tag:** fix through a task pull request and repeat from step 2.
- **After the tag, before the Release:** stop and record the problem. If the
  tag may have been fetched, prefer a new patch version to deleting it.
- **After publication:** never move the tag or replace an asset. Publish a
  corrected patch version and record the problem in the changelog; report a
  vulnerability as [SECURITY.md](SECURITY.md) describes.

## ✅ Checklist

| # | Gate | Done |
| ---: | --- | :---: |
| 1 | I decided to prepare the version | ☐ |
| 2 | Changelog section dated and `VERSION` created or updated | ☐ |
| 3 | Release branch integrated into the default branch, and synced back | ☐ |
| 4 | Candidate frozen: the integration merge commit's full SHA | ☐ |
| 5 | Repository checks and **Repository integrity** pass on the candidate; CodeQL alerts assessed | ☐ |
| 6 | Release certification recorded for the candidate | ☐ |
| 7 | Distribution review recorded and closed | ☐ |
| 8 | I decided to publish | ☐ |
| 9 | Annotated tag on the candidate | ☐ |
| 10 | GitHub Release published from the tag and verified | ☐ |
