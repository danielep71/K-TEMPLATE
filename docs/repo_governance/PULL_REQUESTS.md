# 🔎 Pull request review

Every change reaches the default branch or a release branch through a pull
request that passes **Repository integrity**, the required check. The
data-only `traffic-history` branch is the exception: the optional Traffic
history workflow pushes to it directly ([TRAFFIC.md](TRAFFIC.md)). This guide covers the evidence I
record before merging: which commit was reviewed, by what, and what was not.
The pull request template has a *Final-head review* section for it.

## ✅ Final-head review

A review covers only the commit it read. When Codex reviews pull requests in
this repository, it runs when a pull request is opened or a draft is marked
ready, not after later pushes. It needs an environment for this repository in
its settings; without one it answers with a request to create an environment
instead of reviewing. Before merging a pull request I authored:

1. Push the last fixes and note the current head commit SHA.
2. If no completed review covers that SHA, for example because fixes were
   pushed after the opening review, wait until any running review has
   finished and post `@codex review` as a new top-level comment on the pull
   request. If the opening review already covers the head, go to step 4.
3. Wait until the Codex review summary shows **Completed** for that head SHA.
   A running review, a reaction, or a completed review of an older commit is
   not enough. If the summary shows a shortened SHA, open the commit and
   compare it with the head.
4. Fix or answer every finding and resolve the threads that apply. Any further
   push, a branch update included, makes the review stale: repeat from step 1.
5. Check that **Repository integrity** passes on the current head and that
   GitHub's merge requirements are met. Record the reviewed SHA, the review
   link and the check result in the pull request, then merge.

If the review fails, is unavailable or does not finish, I leave the pull
request open and record what blocks it.

<a id="excel-run"></a>

## 🧪 Excel run

CI cannot run Excel. A pull request that changes VBA source, a workbook layout
or a test case that runs in Excel keeps the template's *Excel run* block, and
records a run on the exact commit of the pull request:

- the commit, the Excel build and bitness, and the Windows version;
- whether **Debug → Compile VBAProject** passed;
- the suite and its mode, and where the log is;
- the completion state:

| State | Meaning |
| --- | --- |
| `PASS` | Every case and assertion ran and passed |
| `FAIL` | At least one case or assertion failed |
| `INCOMPLETE` | Some case did not run, for example because it was skipped or the run stopped. It is not a pass, even if every case that ran passed ([completeness policy](../methodology/TEST_CASES.md#completeness-policy)) |
| `NOT RUN` | No Excel run was made; the *Not validated* section says why |

A run recorded as [recording an Excel run](../../INSTALLATION.md#recording-an-excel-run)
describes, and checked with `tools/check_excel_run.py` against the commit, is
the strongest evidence; cite its report.

A push to VBA source after the run makes it stale, like a review. Record a new
run, or state `NOT RUN` for the new head. A pull request that changes none of
these files deletes the block.

## 🌿 Release branches

While a release branch such as `release/1.2.0` is active, the work for that
release reaches the default branch through it. Without one, task pull requests
target the default branch directly and the integration and sync steps do not
apply.

| Pull request | Base ← head | Issue links | Merge method |
| --- | --- | --- | --- |
| Task | release branch ← task branch created from it | `Resolves #N` | merge commit or squash |
| Integration | default branch ← release branch | `Refs #` for the issues and pull requests it carries | merge commit |
| Sync | release branch ← default branch | `Refs #` for the integration pull request | merge commit |

- GitHub closes an issue from a keyword only when the pull request merges into
  the default branch. After a task pull request merges into the release branch,
  close its issue by hand with the *Resolution* that
  [ISSUE_LABELS.md](ISSUE_LABELS.md#-closing-issues) requires.
- An integration pull request uses `Refs #`, never a closing keyword: a keyword
  there would close every named issue, including one whose acceptance criteria
  still wait for evidence.
- Open the sync pull request right after each integration. It brings the
  integration's merge commit back, changes no file, and lets the next task
  branch start from the same history as the default branch.
- Integration and sync pull requests follow the final-head review above, like
  any other pull request.
- `.github/dependabot.yml` names the active release branch in `target-branch`.
  Update it when a new release branch opens, or remove it when none is active.

## ⚖️ Recorded exceptions

Record these in the *Final-head review* section, in place of the steps above,
or in a comment on the pull request when it has no such section:

- **Identical reviewed content.** A pull request whose head has the same tree
  as a commit already reviewed against the same base may cite that review.
  Record both SHAs and the shared tree, and do not claim a review of the head
  itself.
- **Merge authorized without a final-head review.** I may decide to merge
  without one. Record my decision word for word, what changed after the last
  reviewed commit, and that no review of the final head is claimed.
- **Bot pull requests.** Codex does not review pull requests opened by
  Dependabot or another bot, and their description does not use the pull
  request template. Before merging, record my own review of the final diff and
  head SHA in a comment on the pull request, with **Repository integrity**
  passing, and do not claim a Codex review.

Self-review, an automated review and CI are separate kinds of evidence. This
procedure is a merge rule I apply, not an additional required status check.
