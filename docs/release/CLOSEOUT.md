# 🧾 Milestone closeout

The record for closing a milestone: what was delivered, what evidence shows
it, and what closing it does and does not mean. Copy the template below into
an issue filed under the next milestone, not the one it closes, so that the
milestone can close while the record stays open
([release records](../../RELEASING.md#release-records)). Close the milestone
only when every checklist row links its evidence or records a deviation; for
a released milestone that is step 7 of RELEASING.md, before publication.
Close this issue after the last step of its sequence. Closing a milestone
publishes nothing; a release follows [RELEASING.md](../../RELEASING.md)
separately.

When the milestone is released, check the published release as
[RELEASING.md step 7](../../RELEASING.md#7-publish-and-verify) describes:

```sh
python tools/release.py verify-published vX.Y.Z --repository OWNER/REPOSITORY \
  --milestone N --asset FILE=SHA256
```

With `--milestone`, it also fails unless the milestone is closed and has no
open issue. Add its output to the closeout issue. The download, the unblock
and the Excel checks of an attached workbook stay manual
([step 8](../../RELEASING.md#8-verify-the-published-workbook)).

## Template

```markdown
## Closeout: vX.Y.Z - <title>

### Sequence

| Step | Needs | Produces |
| ---: | --- | --- |
| 1 | The milestone's remaining issues | Their changes merged and integrated into the default branch |
| 2 | Step 1 | The file audit below |
| 3 | Step 2 | The candidate record, the evidence index and the scope statement |
| 4 | Step 3, if I decide to release | RELEASING.md, preparation steps |
| 5 | Step 4, after the milestone is closed and the release published | The output of `tools/release.py verify-published --repository OWNER/REPOSITORY --milestone N`, added to this issue |

### Checklist

| # | Criterion | Evidence | Done |
| ---: | --- | --- | :---: |
| 1 | Every issue of the milestone is closed with a Resolution, or moved with a reason | Evidence index | ☐ |
| 2 | The documented setup reproduces from a fresh checkout of the candidate | Candidate record | ☐ |
| 3 | Repository integrity is green on the candidate | Run link | ☐ |
| 4 | Live labels match the catalogue, and Issue triage passes | Live label drift and Issue triage run links | ☐ |
| 5 | Repository settings match the baseline: `setup_repository.py --check` | Dated evidence record | ☐ |
| 6 | The file audit covers every tracked file at the candidate | File audit | ☐ |
| 7 | The scope statement is recorded | This issue | ☐ |

### Candidate record

| Item | Value |
| --- | --- |
| Candidate commit | Full 40-character SHA on the default branch |
| Integration and sync pull requests | #N, #N |
| Fresh checkout | `git status --short` printed nothing: yes or no |
| Repository checks | Commands run and their results |
| Tool versions | Git, Python and Node.js used |
| Hosted CI | Repository integrity run link and result |
| Deviations | None, or each criterion not met and my decision to accept it |

### Evidence index

| Issue | Subject | Evidence | Verified at candidate |
| --- | --- | --- | :---: |
| #N | Title | Its Resolution | ☐ |

### File audit

One row per tracked file, binary files included. A file changed after its
row was recorded is audited again.

| Path | Blob | Needed | Provenance | Findings | Fix |
| --- | --- | --- | --- | --- | --- |
| As `git ls-files` prints it | From `git ls-files -s` | Yes, or removed | Written here, or adapted with its third-party notice entry | None, or each problem | The pull request or issue |

### Scope statement

vX.Y.Z <title> is closed at commit <SHA>. It delivers <what>. It does not
certify <what it does not cover>; no version is released by this closeout.
Limitations: <list>.
```
