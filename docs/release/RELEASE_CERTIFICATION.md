# 🏅 Release certification

The record that one candidate commit was checked, built, run and accepted
before a version is published. Copy the template below into an issue for the
release, fill it in from the runs themselves, and close the issue only when
every row holds evidence or a recorded deviation. [RELEASING.md](../../RELEASING.md)
says when each part is due.

Write `NOT RUN` with the reason for anything that was not done. A row left
empty or marked `NOT RUN` is never a pass.

## Template

```markdown
## Release certification: vX.Y.Z

### Candidate

| Item | Value |
| --- | --- |
| Version | X.Y.Z |
| Candidate commit | Full 40-character SHA of the integration merge commit |
| Integration and sync pull requests | #N, #N |
| Changelog section | `## [X.Y.Z] - YYYY-MM-DD`, matches the candidate: yes or no |
| Clean checkout | `git status --short` printed nothing: yes or no |

### Environment

| Item | Value |
| --- | --- |
| Excel | Product, version and build, from **File → Account → About Excel** |
| Bitness | 64-bit or 32-bit |
| Windows | Edition, version and build, from `winver` |
| References | The boxes ticked in **Tools → References** |
| Python, Git | Versions used for the repository checks |

### Static gates

| Gate | Result | Evidence |
| --- | --- | --- |
| Repository integrity on the candidate | PASS or FAIL | Run link, and its `integrity-<sha>-…` artifact |
| Repository checks from a clean clone (RELEASING.md step 3) | PASS or FAIL | `report.json`: commit, `dirty: false`, status |
| CodeQL alerts open on the default branch | None, or each with its assessment | Alert links |

### Regression evidence

| Item | Result | Evidence |
| --- | --- | --- |
| Debug → Compile VBAProject | PASS or the error reported | Host record |
| Harness certification ([recording an Excel run](../../INSTALLATION.md#recording-an-excel-run)) | `CERTIFICATION=PASS` or `FAIL`, `COMPLETE` or `INCOMPLETE` | `certification.log` |
| Workbook built from the candidate (`tools/devbuild/build_workbook.py`) | Its `build_commit` stamp, which must be the candidate without *NOT evidence-grade* | The workbook's SHA-256 in `export-manifest.json` |
| `tools/check_excel_run.py --workbook … --require-binding` against the candidate | `Binding: VERIFIED` and `excel-run: PASS`, or each `FAIL` line | Its report, with the Git blob id of every component |
| Regression cases (validation plan) | Number run, passed, failed | Log or evidence folder |
| Failure paths: invalid input, cleanup after an error | PASS or FAIL | Log lines |

### Manual validation

What was checked by hand, by whom, against which independent values. Write
`None` if nothing was.

### Artifacts

| File | SHA-256 | Built from |
| --- | --- | --- |
| `<repository>-X.Y.Z-source.tar.gz` and `SHA256SUMS`, from the **Release** workflow, with their provenance attestation, or a deviation that names its absence | | The tag |
| None, or each other file attached to the GitHub Release | | Candidate commit |

### Deviations

`None`, or each release invariant not met, why, and my decision to accept it.

### Limits

What this certification does not cover, such as hosts not tested,
calculations without independent reference values, sources not yet
verified, or host behaviour still `assumed` in the validation plan's
[host behaviour register](../methodology/VALIDATION_PLAN.md#host-behaviour).

### Sign-off

- Certified by: name, date
- Distribution review: #N, closed
- Decision to publish: recorded here, with the date
```
