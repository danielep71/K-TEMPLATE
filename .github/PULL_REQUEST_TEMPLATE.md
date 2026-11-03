<!--
  Base: the default branch, or the active release branch when development
  happens on one. State which applies, and the merge method, on the Base line.
  Write NOT RUN or NOT APPLICABLE with a reason; never claim evidence that
  was not produced. Synthetic data only; no secrets or vulnerability details.
-->

Base: <!-- the default branch, merged with a merge commit; or `release/X.Y.Z`, squashed, or merged with a merge commit for a sync PR -->

## 🔗 Linked issues

<!-- Link task issues with Resolves #; integration and sync PRs use Refs #
for the issues and preceding PRs they carry. Release branches section of
docs/repo_governance/PULL_REQUESTS.md: bases, keywords and merge methods.
Milestones belong to issues only. -->

Resolves #

## 🧾 Changes

-

## 💡 Why

## 🧮 Calculation impact

<!-- For a change to a calculation's results, rules or expected values: the
source each changed rule cites (its `SRC-` ID and locator) and the class of
the expected values that test it: `published`, `independent` or
`illustrative`, as the reference classes of docs/methodology/TEST_CASES.md
define them. Otherwise replace both lines with "No calculation change". -->

- Source:
- Class of the expected values:

## ✅ Validated

<!-- Commands and checks actually run, with results. Hosted checks run on this pull request. -->

- [ ] `python -m unittest discover -s tests` (with `.github/requirements.txt` installed and Node.js available)
- [ ] Ruff, strict mypy and actionlint pass in `Repository integrity`
- [ ] Other checks specific to this change:

### 🧪 Excel run

<!-- Keep this block when the pull request changes VBA source, a workbook
layout or a test case that runs in Excel; delete it otherwise. Only PASS with
every case run counts as a pass. Rules: docs/repo_governance/PULL_REQUESTS.md. -->

| Item | Result |
| --- | --- |
| Exact commit | <!-- Full SHA of the source imported --> |
| Excel build and bitness, Windows version | |
| `Debug → Compile VBAProject` | <!-- PASS / FAIL / NOT RUN, with a reason --> |
| Suite and mode | |
| Completion state | <!-- PASS / FAIL / INCOMPLETE / NOT RUN --> |
| Evidence | <!-- Where the log or evidence folder is --> |

## 🔎 Final-head review

<!-- After the last push, compare the reviewed commit with the current head.
Procedure and exceptions: docs/repo_governance/PULL_REQUESTS.md. Bot pull
requests do not use this template; record my review of their final head in a
comment on the pull request. -->

- Final PR head SHA:
- Review type and reviewed commit SHA:
- Review result and link:
- `Repository integrity` result and run link:

- [ ] Review covers the current PR head; all findings are fixed or answered and applicable threads resolved.
- [ ] `Repository integrity` passes for the current revision and GitHub merge requirements are satisfied.
- [ ] Any push after this record is followed by renewed review and CI verification.

## ⚠️ Not validated

<!-- State plainly what was not checked and why. Write "Nothing" only if that is true. -->

## ☑️ Checklist

- [ ] The base is the default branch or the active release branch, as stated above.
- [ ] Each linked issue has the assignee, exactly one priority label, an open milestone and a target end date before the milestone's due date; this PR has no milestone, links the relevant issues and carries appropriate labels.
- [ ] Scope is focused: no unrelated changes, local artifacts or generated output.
- [ ] Documentation affected by the change is updated.
- [ ] `CHANGELOG.md` has an entry under `[Unreleased]` for this change, or the change is one its maintenance policy says needs none.
- [ ] No tokens, credentials, real portfolio or client data.
