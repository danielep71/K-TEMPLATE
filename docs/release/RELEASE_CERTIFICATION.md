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
| Debug → Compile VBAProject | PASS, the error reported, or `ON_DEMAND`, which the check accepts only with identity by rebuild (`--workbook` and no `exported/` folder), as `tools/devbuild/excel_run.py` records it: VBA compiles only the modules the run calls, so list the modules it did not call as a limit | Host record |
| Harness certification ([recording an Excel run](../../INSTALLATION.md#recording-an-excel-run)) | `CERTIFICATION=PASS` or `FAIL`, `COMPLETE` or `INCOMPLETE` | `certification.log` |
| Workbook built from the candidate (`tools/devbuild/build_workbook.py`) | Its `build_commit` stamp, which must be the candidate without *NOT evidence-grade* | The workbook's SHA-256 in `export-manifest.json` |
| `tools/check_excel_run.py --workbook … --require-binding` against the candidate | `Binding: VERIFIED` and `excel-run: PASS`, or each `FAIL` line | Its report, with the Git blob id of every component |
| Accuracy contracts, measured on `results.json` and the results file of each project suite that names one | Each contract's worst calibration and holdout error against its threshold, from the report's `Accuracy:` lines; `frozen` or `draft` | The report and those results files |
| Regression cases (validation plan) | Number run, passed, failed | Log or evidence folder |
| Failure paths: invalid input, cleanup after an error | PASS or FAIL | Log lines |

### Manual validation

What was checked by hand, by whom, against which independent values. Write
`None` if nothing was.

### Artifacts

| File | SHA-256 | Built from |
| --- | --- | --- |
| `<repository>-X.Y.Z-source.tar.gz` and `SHA256SUMS`, from the **Release** workflow, with their provenance attestation, or a deviation that names its absence | | The tag |
| `<repository>-X.Y.Z-user-manual.pdf`, or None when the repository keeps no [user manual](../../RELEASING.md#user-manual) | | `docs/user-manual/X.Y.Z/` at the candidate, with the PDF tool and its version |
| None, or each other file attached to the GitHub Release, such as the production workbook | | Candidate commit |

### Published release

| Check | Result | Evidence |
| --- | --- | --- |
| `tools/release.py verify-published vX.Y.Z --milestone N`, with an `--asset` for each other file of *Artifacts* (RELEASING.md step 7) | `release-verify-published: PASS`, or each `FAIL` line; a `NOTE` on an unattested bundle is a deviation | Its output |

### Published workbook

The checks of RELEASING.md step 8, after publication. Write
`NOT APPLICABLE` when the release attaches no workbook.

| Check | Result | Evidence |
| --- | --- | --- |
| Build stamp of the attached workbook | The candidate's full SHA and `production`, without *NOT evidence-grade* | The `build_commit` and `build_configuration` values |
| Rebuild of the tag (`build_workbook.py --configuration production`) | SHA-256 equal to the attached workbook's, or every part identical | Both SHA-256 values, or the parts compared |
| Downloaded copy: SHA-256 before unblocking | Matches *Artifacts* | The value |
| **Data → Edit Links** | No link to another workbook, or each link listed | Screenshot or list |
| **Formulas → Name Manager** | The workbook contract's names, none referring to `#REF!` | Screenshot or list |
| Run, save and reopen a copy | Each operation run on synthetic data; reopens with no error, links and names unchanged | Operations run, and any message |
| User manual PDF, downloaded | Opens and names version X.Y.Z, or `NOT APPLICABLE` | Its SHA-256 against *Artifacts* |

### Deviations

`None`, or each release invariant not met, why, and my decision to accept it.

### Limits

What this certification does not cover, such as hosts not tested,
calculations without independent reference values, rules that the validation
plan's [traceability table](../methodology/VALIDATION_PLAN.md#traceability)
does not mark `Validated`, sources not yet verified, or host behaviour still
`assumed` in its
[host behaviour register](../methodology/VALIDATION_PLAN.md#host-behaviour).

### Sign-off

- Certified by: name, date
- Distribution review: #N, closed
- Decision to publish: recorded here, with the date
```
