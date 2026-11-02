# MyTemplate

[![Clones, including CI](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/danielep71/MyTemplate/traffic-history/data/badges/clones.json)](docs/repo_governance/TRAFFIC.md#badges)
[![Unique visitors, last 14 days](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/danielep71/MyTemplate/traffic-history/data/badges/visitors.json)](docs/repo_governance/TRAFFIC.md#badges)
[![Daily unique visitors since tracking began](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/danielep71/MyTemplate/traffic-history/data/badges/visits.json)](docs/repo_governance/TRAFFIC.md#badges)

I use this template to start repositories for regulatory risk calculation
workbooks, such as SA-CCR, SFT, FRTB, LCR, and NSFR. It provides:

- editing and Git policies for Excel workbooks, exported VBA source, and
  byte-preserved regulatory source copies;
- issue labels, issue forms, a pull request template with a final-head review
  procedure, triage automation, and a workflow that keeps milestones on issues
  only;
- a documented GitHub settings baseline, including security features, applied
  and verified by a script that also writes a dated evidence record;
- quality gates: Ruff, strict mypy and actionlint, plus tests that hold the
  workflows to pinned actions, explicit permissions and the required check, and
  link checks: relative links and anchors on every change, external links weekly;
- a security policy, a review to record before distribution, CodeQL code
  scanning of the Python, JavaScript and workflow code, Dependabot updates for
  the pinned Actions and the hash-locked Python dependencies, and optional
  daily traffic collection.

## 🚀 Start a new repository

GitHub copies files from a template, not its settings. In order:

1. Create the repository with **Use this template**.
2. Delete the template-only section of `AGENTS.md`, with its markers, point
   the two contact links in `.github/ISSUE_TEMPLATE/config.yml` and the
   traffic badges at the top of this README to the new repository, or remove
   the badges, and set `target-branch` in `.github/dependabot.yml` to its
   release branch or remove it. **Repository integrity** fails until you do.
3. Enable Actions and let **Repository integrity** pass on the default branch.
4. Apply and verify the settings baseline, as the
   [repository setup guide](docs/setup/REPOSITORY_SETUP.md) describes. It
   needs Python 3.11+ and GitHub CLI authentication with administrator access:

   ```sh
   python tools/setup_repository.py --repo OWNER/NEW-REPO
   python tools/setup_repository.py --repo OWNER/NEW-REPO --apply
   python tools/setup_repository.py --repo OWNER/NEW-REPO --check \
     --report docs/repo_governance/SETTINGS_EVIDENCE.md
   ```

5. Run **Sync issue labels** and create milestones with due dates
   ([issue labels](docs/repo_governance/ISSUE_LABELS.md)).
6. Optionally, enable traffic collection
   ([traffic history](docs/repo_governance/TRAFFIC.md)).
7. If Codex reviews your pull requests, create its environment for the new
   repository ([pull request review](docs/repo_governance/PULL_REQUESTS.md)).

The [settings reference](docs/repo_governance/SETTINGS.md) and its
[JSON configuration](docs/setup/repository-settings.json) describe the baseline.

## 📐 Methodology

The [methodology guides](docs/methodology/README.md) record the sources, the
data contract, the calculation contracts and the validation plan of the
workbook built from this repository.

## 🧭 Repository governance

See the [governance index](docs/repo_governance/README.md) for repository settings, issue labels, automation, and setup guidance. To take part, read the [contributing guide](CONTRIBUTING.md) and the [code of conduct](CODE_OF_CONDUCT.md).

## ⚖️ License

Repositories created from this template default to the [Mozilla Public License 2.0](LICENSE).
If you choose a different license for a new repository, replace `LICENSE` and keep
[`docs/licenses/`](docs/licenses/THIRD_PARTY_NOTICES.txt): the files listed there keep
the license the notice states for them, MPL-2.0 or MIT.
