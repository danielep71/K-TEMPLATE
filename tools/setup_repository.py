#!/usr/bin/env python3
"""Apply the portable baseline through an authenticated GitHub CLI (github.com)."""
# Postponed annotations let the version check in main() run on Python 3.9 and older.
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from collections.abc import Callable
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote

CONFIG = Path(__file__).resolve().parents[1] / 'docs/setup/repository-settings.json'
RULE_KEYS = ('name', 'target', 'enforcement', 'conditions', 'rules', 'bypass_actors')
GENERAL_KEYS = {'delete_branch_on_merge', 'allow_update_branch', 'has_projects', 'has_downloads'}
# Security features and whether GitHub offers them only on public repositories
# (private repositories need a paid licence for secret scanning).
SECURITY_PUBLIC_ONLY = {
    'private_vulnerability_reporting': True,
    'vulnerability_alerts': False,
    'automated_security_fixes': False,
    'secret_scanning': True,
    'secret_scanning_push_protection': True,
}
ANALYSIS_KEYS = ('secret_scanning', 'secret_scanning_push_protection')

# GitHub API values are JSON; a call takes a method, an endpoint and an optional payload.
Json = Any
Call = Callable[..., Json]
Plan = list[tuple[str, str, Json]]


class SetupError(Exception):
    pass


def api(method: str, endpoint: str, payload: Json = None) -> Json:
    command = ['gh', 'api', '--hostname', 'github.com', '--method', method,
               '-H', 'Accept: application/vnd.github+json',
               '-H', 'X-GitHub-Api-Version: 2022-11-28', endpoint]
    if payload is not None:
        command += ['--input', '-']
    # Never inherit HTTP debug logging: it can expose authentication headers.
    env = dict(os.environ)
    env.pop('GH_DEBUG', None)
    # gh writes UTF-8; never decode with the locale code page (cp1252 on Windows).
    result = subprocess.run(command, input=json.dumps(payload) if payload is not None else None,
                            text=True, encoding='utf-8', capture_output=True, env=env, check=False)
    if result.returncode:
        # Do not print CLI stderr (may contain sensitive server or credential detail).
        match = re.search(r'HTTP (\d{3})', result.stderr)
        status = match.group(1) if match else 'unknown'
        raise SetupError(f'{method} {endpoint} failed (HTTP {status}). Check gh authentication, '
                         'repository access, Administration permissions and plan availability. '
                         'No automatic rollback is performed; rerun preview to inspect partial changes.')
    return json.loads(result.stdout) if result.stdout.strip() else None


def pages(call: Call, endpoint: str, key: str | None = None) -> list[Json]:
    result: list[Json] = []
    page = 1
    while True:
        separator = '&' if '?' in endpoint else '?'
        value = call('GET', f'{endpoint}{separator}per_page=100&page={page}')
        batch = value[key] if key else value
        result.extend(batch)
        if len(batch) < 100:
            return result
        page += 1


def canonical(value: Json) -> Json:
    """Rule arrays are sets for this baseline; ignore API ordering, not values."""
    if isinstance(value, dict):
        return {k: canonical(v) for k, v in sorted(value.items())}
    if isinstance(value, list):
        return sorted((canonical(x) for x in value), key=lambda x: json.dumps(x, sort_keys=True))
    return value


def rule_spec(rule: dict[str, Json]) -> Json:
    # Ignore response metadata, but preserve unknown rules/parameters in comparison.
    return canonical({k: rule.get(k, [] if k == 'bypass_actors' else None) for k in RULE_KEYS})


def load_config(path: Path) -> dict[str, Json]:
    value = json.loads(path.read_text(encoding='utf-8'))
    if set(value) != {'schema_version', 'repository', 'security', 'rulesets'} or value['schema_version'] != 3:
        raise SetupError('Unsupported configuration schema.')
    security = value['security']
    if set(security) != set(SECURITY_PUBLIC_ONLY) or any(type(v) is not bool for v in security.values()):
        raise SetupError('Configuration must contain exactly the five documented boolean security features.')
    if security['secret_scanning_push_protection'] and not security['secret_scanning']:
        raise SetupError('Push protection requires secret scanning.')
    if security['automated_security_fixes'] and not security['vulnerability_alerts']:
        raise SetupError('Dependabot security updates require vulnerability alerts.')
    if set(value['repository']) != GENERAL_KEYS or any(type(v) is not bool for v in value['repository'].values()):
        raise SetupError('Configuration must contain exactly the four documented boolean repository settings.')
    names: list[str] = []
    for rule in value['rulesets']:
        if set(rule) != set(RULE_KEYS) or rule['target'] not in ('branch', 'tag') or rule['enforcement'] != 'active':
            raise SetupError('Invalid baseline ruleset.')
        if rule['bypass_actors'] or rule['conditions']['ref_name']['exclude']:
            raise SetupError('Baseline must not contain bypass actors or excluded refs.')
        names.append(rule['name'])
    if len(names) != 5 or len(set(names)) != 5:
        raise SetupError('Expected five uniquely named baseline rulesets.')
    return dict(value)


def inspect(call: Call, repo: str, config: dict[str, Json],
            report: bool = True) -> tuple[dict[str, Json], Plan]:
    root = f'repos/{repo}'
    metadata = call('GET', root)
    if metadata.get('archived') or metadata.get('disabled'):
        raise SetupError('Destination is archived or disabled.')
    if not metadata.get('permissions', {}).get('admin'):
        raise SetupError('Administrative access to the destination is required.')
    current = pages(call, root + '/rulesets?includes_parents=false')
    inherited = pages(call, root + '/rulesets?includes_parents=true')
    print(f'Destination: {metadata["full_name"]}; visibility: {metadata["visibility"]}; '
          f'default branch: {metadata["default_branch"]}')
    if report and any(r.get('source_type') != 'Repository' for r in inherited):
        print('NOTE: inherited rules also apply; they are not modified or certified by this script.')
    if report:
        report_unmanaged(call, root, current, config)
    plan: Plan = []
    for desired in config['rulesets']:
        matches = [r for r in current if r['name'] == desired['name']]
        if len(matches) > 1:
            raise SetupError(f'Ambiguous duplicate ruleset name: {desired["name"]}')
        if not matches:
            plan.append(('POST', root + '/rulesets', desired))
        else:
            endpoint = root + '/rulesets/' + str(matches[0]['id'])
            actual = call('GET', endpoint)
            if rule_spec(actual) != rule_spec(desired):
                plan.append(('PUT', endpoint, desired))
    plan += security_plan(call, root, metadata, config, report)
    changes = {k: v for k, v in config['repository'].items() if metadata.get(k) != v}
    # Install protection first, then enable branch cleanup.
    if changes:
        plan.append(('PATCH', root, changes))
    return metadata, plan


def missing(call: Call, endpoint: str) -> tuple[Json, bool]:
    """GET an endpoint that answers 404 when its feature is off; True when it does."""
    try:
        return call('GET', endpoint), False
    except SetupError as error:
        if '(HTTP 404)' in str(error):
            return None, True
        raise


def security_fixes(call: Call, root: str) -> tuple[bool, bool]:
    """Whether Dependabot security updates are enabled, and whether GitHub has paused them."""
    value, absent = missing(call, root + '/automated-security-fixes')
    if absent:
        return False, False
    value = value or {}
    return bool(value.get('enabled')), bool(value.get('paused'))


def security_state(call: Call, root: str, metadata: dict[str, Json], feature: str) -> bool:
    """Whether a feature is in effect: paused security updates are enabled, not in effect."""
    if feature == 'private_vulnerability_reporting':
        return bool((call('GET', root + '/private-vulnerability-reporting') or {}).get('enabled'))
    if feature == 'vulnerability_alerts':
        return not missing(call, root + '/vulnerability-alerts')[1]
    if feature == 'automated_security_fixes':
        enabled, paused = security_fixes(call, root)
        return enabled and not paused
    analysis = metadata.get('security_and_analysis') or {}
    return bool((analysis.get(feature) or {}).get('status') == 'enabled')


def security_matches(call: Call, root: str, metadata: dict[str, Json], feature: str,
                     desired: bool, report: bool) -> bool:
    if feature != 'automated_security_fixes':
        return security_state(call, root, metadata, feature) == desired
    enabled, paused = security_fixes(call, root)
    if not desired:
        # Paused updates are still enabled: turning the feature off needs the DELETE.
        return not enabled
    if enabled and paused and report:
        print('NOTE: Dependabot security updates are enabled but paused, so they open no pull '
              'requests. GitHub pauses them when Dependabot pull requests go unanswered; they '
              'count as drift, and an apply that does not resume them fails its read-back.')
    return enabled and not paused


def security_plan(call: Call, root: str, metadata: dict[str, Json], config: dict[str, Json],
                  report: bool) -> Plan:
    """Plan the security features; public-only ones are skipped, with a note, elsewhere."""
    public = metadata.get('visibility') == 'public'
    skipped = [name for name, only in SECURITY_PUBLIC_ONLY.items()
               if only and not public and config['security'][name]]
    if skipped and report:
        print(f'NOTE: {", ".join(skipped)} apply only to public repositories here; they are not '
              f'applied while the destination is {metadata.get("visibility")}. Rerun the check '
              'after making the repository public.')
    plan: Plan = []
    analysis: dict[str, dict[str, str]] = {}
    # Alerts before security updates, which depend on them.
    for feature in ('private_vulnerability_reporting', 'vulnerability_alerts',
                    'automated_security_fixes', *ANALYSIS_KEYS):
        desired = config['security'][feature]
        if SECURITY_PUBLIC_ONLY[feature] and not public:
            continue
        if security_matches(call, root, metadata, feature, desired, report):
            continue
        if feature in ANALYSIS_KEYS:
            analysis[feature] = {'status': 'enabled' if desired else 'disabled'}
        else:
            endpoint = root + '/' + feature.replace('_', '-')
            plan.append(('PUT' if desired else 'DELETE', endpoint, None))
    if analysis:
        plan.append(('PATCH', root, {'security_and_analysis': analysis}))
    return plan


def report_unmanaged(call: Call, root: str, current: list[Json], config: dict[str, Json]) -> None:
    """Name repository rulesets outside the baseline; they stay active alongside it."""
    managed = {rule['name'] for rule in config['rulesets']}
    for ruleset in current:
        if ruleset.get('source_type', 'Repository') != 'Repository' or ruleset['name'] in managed:
            continue
        detail = call('GET', f'{root}/rulesets/{ruleset["id"]}')
        refs = ', '.join(((detail.get('conditions') or {}).get('ref_name') or {}).get('include') or [])
        print(f'NOTE: unmanaged ruleset "{ruleset["name"]}" '
              f'({detail.get("target", "unknown target")}: {refs or "no ref conditions"}; '
              f'{detail.get("enforcement", "unknown enforcement")}) stays active. This script '
              'never changes or deletes it; delete it in the repository settings if a managed '
              'ruleset replaces it.')


def evidence(call: Call, repo: str, config: dict[str, Json], metadata: dict[str, Json],
             plan: Plan, mode: str, now: datetime | None = None) -> str:
    """A dated Markdown record of the managed settings as GitHub reports them."""
    root = f'repos/{repo}'
    stamp = (now or datetime.now(timezone.utc)).strftime('%Y-%m-%dT%H:%M:%SZ')
    public = metadata.get('visibility') == 'public'
    differs = {endpoint for method, endpoint, _ in plan if method == 'PUT' and '/rulesets/' in endpoint}
    created = {payload['name'] for method, _, payload in plan if method == 'POST'}
    current = pages(call, root + '/rulesets?includes_parents=false')
    inherited = [r for r in pages(call, root + '/rulesets?includes_parents=true')
                 if r.get('source_type') != 'Repository']
    managed = {rule['name'] for rule in config['rulesets']}

    def cell(value: Json) -> str:
        return 'on' if value is True else 'off' if value is False else str(value)

    lines = ['# Repository settings evidence', '',
             f'- Checked: **{stamp}** with `tools/setup_repository.py` ({mode})',
             f'- Destination: `{metadata["full_name"]}`; visibility: {metadata["visibility"]}; '
             f'default branch: `{metadata["default_branch"]}`',
             f'- Result: **{"VERIFIED" if not plan else "DRIFT"}**', '',
             '## General settings', '', '| Setting | Configured | Observed |', '| --- | --- | --- |']
    lines += [f'| `{key}` | {cell(value)} | {cell(metadata.get(key))} |'
              for key, value in config['repository'].items()]
    lines += ['', '## Security features', '', '| Feature | Configured | Observed |', '| --- | --- | --- |']
    for feature, value in config['security'].items():
        if SECURITY_PUBLIC_ONLY[feature] and not public:
            observed = 'not applied (public repositories only)'
        elif feature == 'automated_security_fixes' and security_fixes(call, root) == (True, True):
            observed = 'on (paused)'
        else:
            observed = cell(security_state(call, root, metadata, feature))
        lines.append(f'| `{feature}` | {cell(value)} | {observed} |')
    lines += ['', '## Managed rulesets', '', '| Ruleset | ID | Status |', '| --- | --- | --- |']
    for rule in config['rulesets']:
        match = next((r for r in current if r['name'] == rule['name']), None)
        if rule['name'] in created or match is None:
            lines.append(f'| {rule["name"]} | — | missing |')
        else:
            status = 'differs' if f'{root}/rulesets/{match["id"]}' in differs else 'matches'
            lines.append(f'| {rule["name"]} | {match["id"]} | {status} |')
    unmanaged = [r for r in current if r.get('source_type', 'Repository') == 'Repository'
                 and r['name'] not in managed]
    lines += ['', '## Other rules', '',
              f'- Unmanaged repository rulesets: {", ".join(r["name"] for r in unmanaged) or "none"}',
              f'- Inherited rulesets: {", ".join(r["name"] for r in inherited) or "none"}',
              '', 'Unmanaged settings are outside this record.']
    return '\n'.join(lines) + '\n'


def write_evidence(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding='utf-8', newline='\n')
    print(f'Evidence written to {path}')


def check_prerequisites(call: Call, repo: str, metadata: dict[str, Json],
                        config: dict[str, Json]) -> None:
    root = f'repos/{repo}'
    branch = call('GET', root + '/branches/' + quote(metadata['default_branch'], safe=''))
    checks = pages(call, root + '/commits/' + branch['commit']['sha'] + '/check-runs?filter=latest', 'check_runs')
    required = [check for r in config['rulesets'] for rule in r['rules']
                if rule['type'] == 'required_status_checks'
                for check in rule['parameters']['required_status_checks']]
    for required_check in required:
        if not any(c['name'] == required_check['context'] and
                   c.get('app', {}).get('id') == required_check['integration_id'] and
                   c.get('conclusion') == 'success' and c.get('status') == 'completed' for c in checks):
            raise SetupError('Required GitHub Actions check has not succeeded on the current default-branch '
                             'HEAD. Enable Actions and run Repository integrity before applying protection.')
    if not metadata.get('allow_merge_commit'):
        raise SetupError('Merge commits must be enabled before applying these rules.')
    if not metadata.get('allow_squash_merge'):
        raise SetupError('Squash merging must be enabled for the release-branch baseline.')


def run(args: argparse.Namespace, call: Call = api) -> int:
    if not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9-]*/[A-Za-z0-9_.-]+', args.repo) or args.repo.split('/')[1] in ('.', '..'):
        raise SetupError('Destination must be an explicit owner/repository on github.com.')
    config = load_config(args.config)
    metadata, plan = inspect(call, args.repo, config)
    mode = 'apply' if args.apply else 'check' if args.check else 'preview'
    report = getattr(args, 'report', None)
    if report and (not args.apply or not plan):
        write_evidence(report, evidence(call, args.repo, config, metadata, plan, mode))
    if not plan:
        print('VERIFIED: all documented settings match. Unmanaged settings are outside this verification.')
        return 0
    for method, endpoint, payload in plan:
        print(f'{method} {endpoint}')
        if payload is not None:
            print(json.dumps(payload, indent=2))
    if args.check:
        print('DRIFT: documented settings do not match.')
        return 1
    if not args.apply:
        print('PREVIEW ONLY: no changes made. Use --apply to execute this plan.')
        return 0
    if any(method == 'PUT' and '/rulesets/' in endpoint for method, endpoint, _ in plan) and not args.update_existing:
        raise SetupError('A named ruleset differs. Review the preview, then use --update-existing '
                         'to replace that entire named ruleset. Additional rules in it will be removed.')
    check_prerequisites(call, args.repo, metadata, config)
    for method, endpoint, payload in plan:
        call(method, endpoint, payload)
        print(f'Applied {method} {endpoint}')
    metadata, remaining = inspect(call, args.repo, config, report=False)
    if report:
        write_evidence(report, evidence(call, args.repo, config, metadata, remaining, mode))
    if remaining:
        raise SetupError('Read-back verification failed; rerun preview to inspect remaining differences.')
    print('VERIFIED: all documented settings match after apply.')
    return 0


def main() -> int:
    if sys.version_info < (3, 11):
        print('ERROR: Python 3.11 or later is required.', file=sys.stderr)
        return 2
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', required=True, help='Explicit destination owner/repository')
    parser.add_argument('--config', type=Path, default=CONFIG)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('--apply', action='store_true', help='Apply changes, then verify')
    mode.add_argument('--check', action='store_true', help='Read-only verification; exit 1 for drift')
    parser.add_argument('--update-existing', action='store_true', help='Allow replacement of differing named rulesets')
    parser.add_argument('--report', type=Path, help='Write a dated Markdown evidence record of the managed settings')
    args = parser.parse_args()
    if args.update_existing and not args.apply:
        parser.error('--update-existing requires --apply')
    try:
        return run(args)
    except (SetupError, OSError, ValueError, KeyError, TypeError) as exc:
        print(f'ERROR: {exc}', file=sys.stderr)
        return 2


if __name__ == '__main__':
    sys.exit(main())
