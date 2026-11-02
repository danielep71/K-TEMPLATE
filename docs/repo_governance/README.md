# 🧭 Repository governance

![Governance](https://img.shields.io/badge/Governance-Versioned-00338D?style=flat-square)
![Labels](https://img.shields.io/badge/Issue_labels-20-1D76DB?style=flat-square)

Ongoing repository rules and maintenance guidance live here. One-time initialization instructions and the executable settings baseline live in [setup](../setup/REPOSITORY_SETUP.md).

| Guide | Purpose |
| --- | --- |
| [⚙️ Settings](SETTINGS.md) | Repository settings and branch/tag protection baseline. |
| [🏷️ Issue labels](ISSUE_LABELS.md) | Label catalogue, synchronization, drift detection, and issue triage. |
| [🔎 Pull request review](PULL_REQUESTS.md) | Final-head review before merging, and the exceptions to record. |
| [📈 Traffic history](TRAFFIC.md) | Optional daily traffic collection on the `traffic-history` branch, with alerts. |
| [🚀 Repository setup](../setup/REPOSITORY_SETUP.md) | Preview, apply, verify, and hand over a newly generated repository. |
| [📋 Settings configuration](../setup/repository-settings.json) | Machine-readable baseline used by the setup script. |
| [🔒 Security policy](../../SECURITY.md) | Private vulnerability reporting, response targets, and scope. |
| [📦 Distribution review](DISTRIBUTION.md) | What to check and record before making the repository public, sharing a copy or publishing a release. |
| [🛠️ Contributing](../../CONTRIBUTING.md) | How to open issues and pull requests, branch names, and local checks. |
| [🤝 Code of conduct](../../CODE_OF_CONDUCT.md) | Expected behaviour and how to report a problem. |

## Repository controls

Settings are defined in the [settings configuration](../setup/repository-settings.json). Label configuration, issue forms, and automation live under `.github/`.

Changes follow the protected pull-request process and must pass **Repository integrity**, which also checks every relative link and anchor in the Markdown files. The weekly **Link check** workflow requests the external links. The **CodeQL** workflow scans the Python, JavaScript and workflow code on pushes and pull requests to the default and release branches, and weekly; its availability on private repositories is in the setup guide's [handover table](../setup/REPOSITORY_SETUP.md#recommended-handover). The [setup guide](../setup/REPOSITORY_SETUP.md) explains initialization, verification, and which files to retain after launch.
