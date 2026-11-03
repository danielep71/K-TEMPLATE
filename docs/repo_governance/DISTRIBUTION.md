# 📦 Distribution review

Distribution publishes more than the current files. Before anything leaves
the people who already have access, record the review below in an issue and
close it. Do not distribute while that issue is open.

## 🚪 What counts as distribution

- changing the repository's visibility from private to public, or to
  internal in an organization with more members;
- sharing a clone, an archive, an exported module or a workbook with anyone
  who does not already have access;
- publishing a GitHub Release or attaching a file to one;
- granting a new collaborator or team access to a private repository.

Making a repository public publishes its **whole history**: every branch,
tag and reachable commit, and every issue, pull request, comment, workflow
run log and artifact. A file deleted in a later commit is still published
with the commit that added it.

## ✅ The review

Open an issue for the review, triaged like any other, with the `release` and
`security` labels. Copy this checklist into it, and record for each item what
you checked and what you found:

```markdown
## Distribution review

- What is distributed, to whom, and from which commit:

- [ ] **License.** `LICENSE` is still the intended license for the original
      content.
- [ ] **Provenance.** Every file is original, or carries a compatible license
      with its attribution in `docs/licenses/THIRD_PARTY_NOTICES.txt`.
      The terms of each regulatory source copy under
      `docs/methodology/sources/` allow its redistribution.
- [ ] **Third-party claims.** No employer, client or vendor has a claim on the
      content or on the methods it implements.
- [ ] **History scan.** A scan of every reachable blob found no sensitive
      material. Commit range, pattern categories and redacted findings:
- [ ] **Workbooks.** Each workbook ever committed was inspected as below.
- [ ] **Issues and pull requests.** Their text, comments, edit history and
      attachments contain no confidential material.
- [ ] **Actions.** No workflow run log or artifact holds sensitive output.
```

Close the issue only when every item is checked. If the review finds
something, fix it as *Handling a finding* describes and review again.

## 🔍 History scan

Scan every blob reachable from any branch or tag, not only the current tree.
From a fresh clone:

```sh
git fetch origin '+refs/heads/*:refs/remotes/origin/*' '+refs/tags/*:refs/tags/*'
# Every reachable blob, with the path it was committed under.
git rev-list --objects --all
# Every binary file ever committed: inspect each one.
git log --all --numstat --format= | awk '$1 == "-" {print $3}' | sort -u
# Search the text of every commit for a pattern.
git rev-list --all | xargs git grep -n -I -i -E 'PATTERN'
```

Search for at least:

- credentials: tokens, keys, passwords, connection strings and private
  certificates;
- people: names, email addresses and user names other than your own;
- organizations: employer, client, counterparty and vendor names, internal
  host names, paths and system identifiers;
- real data: portfolio, trade, account or position identifiers, real
  amounts, and any file that is not synthetic.

The review issue becomes public with the repository, so it must not become
another copy of what the scan looks for. Record in it the scanned commit range,
the categories of patterns, such as "client names" or "internal host names",
and each finding redacted to its commit, path and category. Keep the literal
patterns and any sensitive evidence outside the repository, in a private note
or channel; a suspected vulnerability goes through the private reporting
channel in [SECURITY.md](../../SECURITY.md). The `traffic-history` branch,
when it exists, holds aggregate counts only; it is scanned with the rest.

## 📗 Workbooks

A workbook holds content that its exported source and a text diff do not
show. **Repository integrity** already fails a committed Office package,
other than a regulatory source copy, that carries macros or names a person or
organization in `docProps/core.xml` (author, last editor) or
`docProps/app.xml` (company, manager); it checks only the current tree.
Open each workbook ever committed in Excel and run **File → Info → Check for
Issues → Inspect Document**, or unzip a `.xlsx`, `.xlsm` or `.xltx` and read
its parts. Look for:

| Content | Where it is stored |
| --- | --- |
| Author, last editor, company and manager | `docProps/core.xml`, `docProps/app.xml` |
| Hidden and very hidden sheets | `state` attributes in `xl/workbook.xml` |
| Defined names, including hidden ones | `definedNames` in `xl/workbook.xml` |
| Links to other workbooks and data connections | `xl/externalLinks/`, `xl/connections.xml` |
| Comments and threaded comments, with their authors | `xl/comments*.xml`, `xl/threadedComments/`, `xl/persons/` |
| Cached source data of pivot tables | `xl/pivotCache/` |
| Values left in cells, rows or columns outside the visible range | the sheet parts under `xl/worksheets/` |

The VBA project inside a workbook can differ from the exported source in the
repository. Export the modules from the workbook itself and read them too.

## 🧹 Handling a finding

- **A credential:** revoke or rotate it first. Removing it from the history
  does not undo its exposure.
- **A file or text in the history:** remove it from every commit before you
  distribute, for example with `git filter-repo`, then force-push every
  affected branch and tag. That needs the protection rulesets to be relaxed
  temporarily; restore them and run the settings check afterwards. Every
  existing clone keeps the old history. On GitHub, pull request references
  and cached views can keep it too; GitHub's guide to
  [removing sensitive data](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/removing-sensitive-data-from-a-repository)
  explains what needs GitHub Support.
- **Issue or pull request text:** edit it, then delete the earlier revisions
  from the comment's edit history, which readers can open. Removing the link
  to an attachment does not delete the file; ask GitHub Support to remove it.
- **A workflow run log or artifact:** delete the run.

Then run the review again from the start.
