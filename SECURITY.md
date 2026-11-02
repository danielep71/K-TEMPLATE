# 🔒 Security policy

## 📌 Supported versions

The latest release, the active release branch and the default branch are
supported on a best-effort basis. Reports must identify a full commit SHA.
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

## 📦 Before distribution

Before the repository is made public, a copy is shared or a release is
published, I record the
[distribution review](docs/repo_governance/DISTRIBUTION.md): license,
provenance, a scan of the whole history, the workbooks, and the issue and pull
request text.
