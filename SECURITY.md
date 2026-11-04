# 🔒 Security policy

## 📌 Supported versions

The latest release, the active release branch and the default branch are
supported on a best-effort basis. The latest release is the newest version in
[CHANGELOG.md](CHANGELOG.md), published as described in
[RELEASING.md](RELEASING.md); until one exists, only the two branches are
supported. Reports must identify a full commit SHA.
Modified copies are unsupported unless the problem reproduces in this
repository.

## 📨 Reporting a vulnerability

Do **not** disclose a suspected vulnerability in an issue, pull request,
commit message, sample file or screenshot.

Report it privately through **Security → Report a vulnerability** on this
repository's GitHub page. The report stays visible only to you and the people
who manage the repository's security until an advisory is published.

If the page has no **Report a vulnerability** button, for example because the
repository is private, open an issue with the **🔒 Private security contact**
form. It asks only how to reach you; I reply with a private channel. Never put
vulnerability details in the issue.

Include:

- the full commit SHA and the affected file, procedure or workflow;
- the Excel/Office build, bitness and Windows version, where relevant;
- the impact and minimal reproduction steps, with synthetic data only; and
- any workaround you tested.

## ⏱️ Response targets

I maintain this repository alone; the targets are best effort, not a
contractual commitment:

- acknowledgement within 5 business days;
- an initial assessment within 10 business days of receiving enough evidence;
- an update at least every 14 days while the report is investigated.

Remediation and disclosure timing are agreed with the reporter according to
severity.

## 🎯 Scope

**In scope:** this repository's VBA source, workbook templates, tests and
tooling, and its GitHub Actions workflows and their configuration.

**Out of scope:**

- vulnerabilities in Excel, Office, Windows, GitHub, Python or Node.js
  themselves;
- organization-controlled macro and deployment policy;
- modified copies;
- ordinary defects without concrete security impact, such as an incorrect
  calculated figure. Report those as issues.

## 🤝 Safe harbor

I welcome good-faith security research within the scope above. If you report
privately as this policy describes, use only synthetic data and your own
copies and accounts, do not access, change or delete anyone else's data, do
not degrade this repository or the services it runs on, and give me
reasonable time to respond before any disclosure, I will not pursue or
support legal action against you for that research. This covers only claims
I could bring; it cannot authorize testing of GitHub, Microsoft or other
third-party systems, whose own policies apply.

## ⚙️ Runtime risk surface

A workbook's macros run with the rights of the person who opens it. This
section states what the code in this repository may do when it runs, so that
you can tell intended behaviour from a vulnerability.

| Component | Reads | Writes or deletes |
| --- | --- | --- |
| Test harness and example suite (`tests/modules/`), in the development workbook only | The test-case files under the fixtures folder it is given; Excel's `Calculation`, `DisplayAlerts`, `EnableEvents` and `ScreenUpdating`, the number of open workbooks, the active window's caption, and the host's name, version, build and operating system | Its logs, `results.json` and the certification log in the evidence folder it is given, replacing a file of the same name; the example suite changes `Calculation` and `ScreenUpdating` and restores them. It deletes no file |
| `tools/devbuild/excel_run.py`, repository tooling on Windows | The checkout, and the Windows version from the registry | Builds the development workbook under `test-results/build/`; empties the evidence folder, but only one under `test-results/`, and otherwise requires an empty or new folder, and writes `host.txt` in it and its report beside it as `<folder>-report.txt`; opens the workbook read-only without updating links, and closes it without saving |
| Production workbook (`src/`) | Stated for each entry point, as *Other code* below requires | Stated for each entry point, as *Other code* below requires |

- **Excel settings.** The procedures need their workbook's macros to be
  enabled, by your organization's policy or by **Enable Content**, and
  nothing else: no lower macro-security level and no *Trust access to the VBA
  project object model* ([host policy](INSTALLATION.md#%EF%B8%8F-host-policy)).
  `tools/devbuild/excel_run.py` changes no macro setting: where Excel opens
  the workbook with macros disabled, it waits for **Enable Content**. Do not
  change a macro setting to run them.
- **Host-independent code.** The [dependency rules](docs/REPOSITORY_STRUCTURE.md#dependencies)
  forbid `src/core/` and `src/classes/` any Excel object model, file,
  registry, process or network access, external object model, native
  `Declare`, dynamic call, and `CreateObject` or `GetObject` other than
  `CreateObject("Scripting.Dictionary")`. `tools/check_vba_dependencies.py`
  enforces them by name, with the limits the
  [static checks](docs/VBA_HOUSE_STYLE.md#static-checks) list.
- **Other code.** Those rules do not restrict the facade, the workbook and
  form code, or the tests. For each entry point of the production workbook,
  state in the table above which sheets, defined names and files it reads,
  what it writes, saves or deletes, and whether it starts a process
  (`Shell`), creates an external object (`CreateObject`, `GetObject`), calls
  a native library (`Declare`) or reaches the network, and why.

## 🗝️ Data and secrets

- Use synthetic data only, in reports as everywhere else. Never attach a
  workbook or file that holds real portfolio, client or counterparty data.
- A workbook holds more than its visible cells: hidden and very hidden
  sheets, defined names, comments and their authors, links to other workbooks
  and data connections, document properties, cached pivot data and the VBA
  project. Inspect a workbook before you share it, as the
  [distribution review](docs/repo_governance/DISTRIBUTION.md#-workbooks)
  describes.
- Keep secret and key files out of the repository. `.gitignore` ignores them,
  but `git add -f` bypasses it, so **Repository integrity** also fails on a
  tracked file with one of these names, in any directory and in any letter
  case:
  - `.env` and `.env.*`, except `.env.example`;
  - `*.pem`, `*.key`, `*.p12`, `*.pfx`, `*.pvk`, `*.jks` and `*.keystore`;
  - `id_rsa*` and `id_ed25519*`, except their `.pub` public keys;
  - `*.kdbx`.

  The check looks at names only, and secret scanning matches known token
  patterns, so neither proves that a file holds no secret.
- If a secret, such as a token, password or key, has leaked in a file, a
  workbook, the history, an issue or a log, revoke or rotate it first, then
  report it. Deleting it does not undo its exposure
  ([handling a finding](docs/repo_governance/DISTRIBUTION.md#-handling-a-finding)).

## 📦 Before distribution

Before the repository is made public, a copy is shared or a release is
published, I record the
[distribution review](docs/repo_governance/DISTRIBUTION.md): license,
provenance, a scan of the whole history, the workbooks, and the issue and pull
request text.
