#!/usr/bin/env python3
"""Preview, check or apply the repository settings baseline through an authenticated GitHub CLI (github.com).

The baseline is docs/setup/repository-settings.json: an optional visibility,
the general repository settings, the security features, the GitHub Actions
settings (reported, never written) and the repository rulesets. Preview (the
default) and --check only read; --apply writes, then reads everything back.
"""
# Postponed annotations let the version check in main() run on older Python versions.
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
from collections.abc import Callable
from datetime import datetime, timezone
from pathlib import Path
from typing import TYPE_CHECKING, Any
from urllib.parse import quote

ROOT = Path(__file__).resolve().parents[1]
CONFIG = ROOT / 'docs/setup/repository-settings.json'
SCHEMA_VERSION = 4
EVIDENCE_SCHEMA_VERSION = 1
TOP_KEYS = ('schema_version', 'visibility', 'repository', 'security', 'actions', 'rulesets')
RULE_KEYS = ('name', 'target', 'enforcement', 'conditions', 'rules', 'bypass_actors')
# The merge methods are written before the rulesets, which may allow only enabled ones.
MERGE_KEYS = ('allow_merge_commit', 'allow_squash_merge', 'allow_rebase_merge')
GENERAL_KEYS = MERGE_KEYS + (
    'allow_auto_merge', 'allow_update_branch', 'delete_branch_on_merge',
    'web_commit_signoff_required', 'has_issues', 'has_projects', 'has_wiki',
    'has_discussions', 'has_downloads',
)
MERGE_METHODS = {'merge': 'allow_merge_commit', 'squash': 'allow_squash_merge',
                 'rebase': 'allow_rebase_merge'}
# Security features in write order (alerts before the security updates that need
# them), and whether GitHub offers them only on public repositories (private
# repositories need a paid licence for secret scanning).
SECURITY_PUBLIC_ONLY = {
    'private_vulnerability_reporting': True,
    'vulnerability_alerts': False,
    'automated_security_fixes': False,
    'secret_scanning': True,
    'secret_scanning_push_protection': True,
}
ANALYSIS_KEYS = ('secret_scanning', 'secret_scanning_push_protection')
# Each Actions setting: the endpoint under actions/permissions that reports it,
# its field there, and whether it applies only to public repositories.
ACTIONS_SOURCES = {
    'sha_pinning_required': ('', 'sha_pinning_required', False),
    'default_workflow_permissions': ('/workflow', 'default_workflow_permissions', False),
    'can_approve_pull_request_reviews': ('/workflow', 'can_approve_pull_request_reviews', False),
    'fork_pr_approval_policy': ('/fork-pr-contributor-approval', 'approval_policy', True),
    'retention_days': ('/artifact-and-log-retention', 'days', False),
}
FORK_POLICIES = ('first_time_contributors_new_to_github', 'first_time_contributors',
                 'all_external_contributors')
REPOSITORY_PATTERN = re.compile(r'[A-Za-z0-9][A-Za-z0-9-]*/[A-Za-z0-9_.-]+')
NOT_APPLIED = 'not applied (public repositories only)'
DISTRIBUTION = 'docs/repo_governance/DISTRIBUTION.md'

# GitHub API values are JSON; a call takes a method, an endpoint and an optional payload.
# The aliases appear only in annotations, which are never evaluated, so Python 3.8
# imports the module and reaches the version check.
Json = Any
if TYPE_CHECKING:
    Call = Callable[..., Json]
    Plan = list[tuple[str, str, Json]]
    Row = dict[str, Json]


class SetupError(Exception):
    """A refusal or failure; the command exits with status 2."""


class ApiError(SetupError):
    """A GitHub API call that failed, with the HTTP status gh reported, if any."""

    def __init__(self, method: str, endpoint: str, status: str) -> None:
        super().__init__(f'{method} {endpoint} failed (HTTP {status}). Check gh authentication, '
                         'repository access, Administration permissions and plan availability. '
                         'No automatic rollback is performed; rerun preview to inspect partial changes.')
        self.status = status


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
        raise ApiError(method, endpoint, match.group(1) if match else 'unknown')
    return json.loads(result.stdout) if result.stdout.strip() else None


def read(call: Call, endpoint: str) -> tuple[Json, str | None]:
    """GET an endpoint: its value, or None and the HTTP status of the failure."""
    try:
        return call('GET', endpoint), None
    except ApiError as error:
        return None, error.status


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


# Configuration


def require_booleans(value: Json, keys: tuple[str, ...], what: str) -> None:
    if not isinstance(value, dict) or set(value) != set(keys) or any(type(v) is not bool for v in value.values()):
        raise SetupError(f'Configuration must contain exactly the {len(keys)} documented boolean '
                         f'{what}: {", ".join(keys)}.')


def check_actions(actions: Json) -> None:
    if not isinstance(actions, dict) or set(actions) != set(ACTIONS_SOURCES):
        raise SetupError('Configuration must contain exactly the documented Actions settings: '
                         + ', '.join(ACTIONS_SOURCES) + '.')
    days = actions['retention_days']
    if (type(actions['sha_pinning_required']) is not bool
            or actions['default_workflow_permissions'] not in ('read', 'write')
            or type(actions['can_approve_pull_request_reviews']) is not bool
            or actions['fork_pr_approval_policy'] not in FORK_POLICIES
            or type(days) is not int or days < 1):
        raise SetupError('Invalid Actions settings: sha_pinning_required and '
                         'can_approve_pull_request_reviews are booleans, default_workflow_permissions '
                         'is read or write, fork_pr_approval_policy is one of '
                         + ', '.join(FORK_POLICIES) + ', and retention_days is a positive integer.')


def check_rule(name: str, rule: Json, settings: dict[str, bool]) -> None:
    """Refuse a rule that the configured repository settings would make GitHub reject."""
    parameters = rule.get('parameters') or {}
    for method in parameters.get('allowed_merge_methods', []):
        setting = MERGE_METHODS.get(method)
        if setting is None:
            raise SetupError(f'Ruleset "{name}" allows an unknown merge method: {method}.')
        if not settings[setting]:
            raise SetupError(f'Ruleset "{name}" allows {method} merges, but the configuration '
                             f'sets {setting} to false.')
    for check in parameters.get('required_status_checks', []):
        if (not isinstance(check, dict) or not isinstance(check.get('context'), str)
                or type(check.get('integration_id')) is not int):
            raise SetupError(f'Ruleset "{name}" has a required status check without a string '
                             'context and an integer integration_id.')


def check_rulesets(rulesets: Json, settings: dict[str, bool]) -> None:
    if not isinstance(rulesets, list) or not rulesets:
        raise SetupError('Configuration must contain at least one ruleset.')
    names: list[str] = []
    for ruleset in rulesets:
        if not isinstance(ruleset, dict) or set(ruleset) != set(RULE_KEYS):
            raise SetupError('Each configured ruleset must have exactly: ' + ', '.join(RULE_KEYS) + '.')
        name = ruleset['name']
        if ruleset['target'] not in ('branch', 'tag') or ruleset['enforcement'] != 'active':
            raise SetupError(f'Ruleset "{name}" must be an active branch or tag ruleset.')
        if ruleset['bypass_actors'] or ruleset['conditions']['ref_name']['exclude']:
            raise SetupError(f'Ruleset "{name}" must not contain bypass actors or excluded refs.')
        types = [rule['type'] for rule in ruleset['rules']]
        if len(types) != len(set(types)):
            raise SetupError(f'Ruleset "{name}" repeats a rule type.')
        for rule in ruleset['rules']:
            check_rule(name, rule, settings)
        names.append(name)
    duplicates = sorted({name for name in names if names.count(name) > 1})
    if duplicates:
        raise SetupError('Configured ruleset names must be unique: ' + ', '.join(duplicates) + '.')


def load_config(path: Path) -> dict[str, Json]:
    value = json.loads(path.read_text(encoding='utf-8'))
    if not isinstance(value, dict) or set(value) != set(TOP_KEYS) or value['schema_version'] != SCHEMA_VERSION:
        raise SetupError('Unsupported configuration schema.')
    if value['visibility'] not in (None, 'public', 'private'):
        raise SetupError('The configured visibility must be null, "public" or "private".')
    require_booleans(value['repository'], GENERAL_KEYS, 'repository settings')
    require_booleans(value['security'], tuple(SECURITY_PUBLIC_ONLY), 'security features')
    security = value['security']
    if security['secret_scanning_push_protection'] and not security['secret_scanning']:
        raise SetupError('Push protection requires secret scanning.')
    if security['automated_security_fixes'] and not security['vulnerability_alerts']:
        raise SetupError('Dependabot security updates require vulnerability alerts.')
    check_actions(value['actions'])
    check_rulesets(value['rulesets'], value['repository'])
    return dict(value)


# Observation: read-only, and tolerant of what a token without administrator access cannot read.


def security_states(call: Call, root: str, metadata: dict[str, Json], public: bool) -> dict[str, Row]:
    """Each security feature as observed: True, False or None when GitHub does not say."""
    states: dict[str, Row] = {}
    if public:
        value, failure = read(call, root + '/private-vulnerability-reporting')
        states['private_vulnerability_reporting'] = (
            {'observed': bool((value or {}).get('enabled')), 'source': 'GET private-vulnerability-reporting enabled'}
            if failure is None else
            {'observed': None, 'source': f'GET private-vulnerability-reporting (HTTP {failure}: unreadable)'})

    value, failure = read(call, root + '/automated-security-fixes')
    if failure is None:
        value = value or {}
        paused = bool(value.get('enabled')) and bool(value.get('paused'))
        fixes: Row = {'observed': bool(value.get('enabled')), 'paused': paused,
                      'source': 'GET automated-security-fixes enabled' + (', paused' if paused else '')}
    elif failure == '404':
        fixes = {'observed': False, 'paused': False, 'source': 'GET automated-security-fixes (HTTP 404: off)'}
    else:
        fixes = {'observed': None, 'paused': False,
                 'source': f'GET automated-security-fixes (HTTP {failure}: unreadable)'}

    # The endpoint answers without content when the alerts are on and 404 when
    # they are off. Enabled security updates prove them on: GitHub requires them.
    _, failure = read(call, root + '/vulnerability-alerts')
    if failure is None:
        states['vulnerability_alerts'] = {'observed': True, 'source': 'GET vulnerability-alerts (on)'}
    elif failure == '404':
        states['vulnerability_alerts'] = {'observed': False, 'source': 'GET vulnerability-alerts (HTTP 404: off)'}
    elif fixes['observed'] is True:
        states['vulnerability_alerts'] = {
            'observed': True, 'source': f'GET vulnerability-alerts (HTTP {failure}: unreadable); on, '
                                        'because Dependabot security updates are enabled and require them'}
    else:
        states['vulnerability_alerts'] = {'observed': None,
                                          'source': f'GET vulnerability-alerts (HTTP {failure}: unreadable)'}
    states['automated_security_fixes'] = fixes

    # GitHub reports this block only to administrators of the repository.
    analysis = metadata.get('security_and_analysis')
    for key in ANALYSIS_KEYS if public else ():
        if not isinstance(analysis, dict):
            states[key] = {'observed': None,
                           'source': f'security_and_analysis.{key} (not reported without administrator access)'}
        else:
            states[key] = {'observed': (analysis.get(key) or {}).get('status') == 'enabled',
                           'source': f'security_and_analysis.{key}'}
    for feature, only in SECURITY_PUBLIC_ONLY.items():
        if only and not public:
            states[feature] = {'observed': None, 'source': NOT_APPLIED, 'applies': False}
    return {feature: states[feature] for feature in SECURITY_PUBLIC_ONLY}


def actions_states(call: Call, root: str, public: bool) -> dict[str, Row]:
    """Each Actions setting as observed, or None with the reason it is not known."""
    answers: dict[str, tuple[Json, str | None]] = {}
    states: dict[str, Row] = {}
    for setting, (suffix, field, public_only) in ACTIONS_SOURCES.items():
        if public_only and not public:
            states[setting] = {'observed': None, 'source': NOT_APPLIED, 'applies': False}
            continue
        endpoint = 'actions/permissions' + suffix
        if endpoint not in answers:
            answers[endpoint] = read(call, f'{root}/{endpoint}')
        value, failure = answers[endpoint]
        if failure is not None:
            states[setting] = {'observed': None, 'source': f'GET {endpoint} (HTTP {failure}: unreadable)'}
        elif not isinstance(value, dict) or field not in value:
            states[setting] = {'observed': None, 'source': f'GET {endpoint} (no {field} field)'}
        else:
            states[setting] = {'observed': value[field], 'source': f'GET {endpoint} {field}'}
    return states


def observe(call: Call, repo: str, config: dict[str, Json]) -> dict[str, Json]:
    """Read every value the configuration manages. Read-only."""
    root = f'repos/{repo}'
    metadata = call('GET', root)
    public = metadata.get('visibility') == 'public'
    listed = pages(call, root + '/rulesets?includes_parents=true')
    own = [r for r in listed if r.get('source_type', 'Repository') == 'Repository']
    managed = {rule['name'] for rule in config['rulesets']}
    names = [r['name'] for r in own if r['name'] in managed]
    duplicates = sorted({name for name in names if names.count(name) > 1})
    if duplicates:
        raise SetupError('Ambiguous duplicate ruleset name: ' + ', '.join(duplicates))
    return {
        'metadata': metadata,
        'public': public,
        'rulesets': [call('GET', f'{root}/rulesets/{r["id"]}') for r in own],
        'inherited': [r['name'] for r in listed if r.get('source_type', 'Repository') != 'Repository'],
        'security': security_states(call, root, metadata, public),
        'actions': actions_states(call, root, public),
    }


# Comparison


_ABSENT = '<absent>'


def keyed(spec: Json) -> Json:
    """A ruleset spec with its rules keyed by type, for readable difference paths."""
    result = dict(spec)
    rules = spec.get('rules') or []
    types = [rule.get('type') for rule in rules]
    if len(types) == len(set(types)):
        result['rules'] = {rule.get('type'): rule.get('parameters', {}) for rule in rules}
    return result


def differences(observed: Json, configured: Json, path: str = '') -> list[str]:
    """Each path where an observed value departs from the configured one."""
    if isinstance(observed, dict) and isinstance(configured, dict):
        found: list[str] = []
        for key in sorted(set(observed) | set(configured), key=str):
            found += differences(observed.get(key, _ABSENT), configured.get(key, _ABSENT),
                                 f'{path}.{key}' if path else str(key))
        return found
    if canonical(observed) == canonical(configured):
        return []
    return [f'{path}: observed {json.dumps(canonical(observed))}, '
            f'configured {json.dumps(canonical(configured))}']


def compare(key: str, name: str, configured: Json, state: Row) -> Row:
    """A row of the record; matches is None for a value that does not apply here."""
    row: Row = {key: name, 'configured': configured, 'observed': state['observed'], 'source': state['source']}
    if 'paused' in state:
        row['paused'] = state['paused']
    if not state.get('applies', True):
        row['matches'] = None
    elif row.get('paused') and configured is True:
        # Paused updates open no pull requests: drift when they should be on, but still
        # enabled, so turning the feature off needs the DELETE.
        row['matches'] = False
    else:
        row['matches'] = state['observed'] is not None and state['observed'] == configured
    return row


def assess(config: dict[str, Json], state: dict[str, Json]) -> dict[str, Json]:
    """Compare every managed value with what GitHub reports."""
    metadata = state['metadata']
    configured_visibility = config['visibility']
    visibility = {'configured': configured_visibility, 'observed': metadata.get('visibility'),
                  'matches': configured_visibility in (None, metadata.get('visibility'))}
    # Rows in the documented order, whatever the order of the configuration file;
    # security features in write order.
    settings = [{'setting': key, 'configured': config['repository'][key], 'observed': metadata.get(key),
                 'matches': metadata.get(key) == config['repository'][key]}
                for key in GENERAL_KEYS]
    security = [compare('feature', key, config['security'][key], state['security'][key])
                for key in SECURITY_PUBLIC_ONLY]
    actions = [compare('setting', key, config['actions'][key], state['actions'][key])
               for key in ACTIONS_SOURCES]
    rulesets: list[Row] = []
    for desired in config['rulesets']:
        actual = next((r for r in state['rulesets'] if r.get('name') == desired['name']), None)
        if actual is None:
            rulesets.append({'name': desired['name'], 'id': None, 'status': 'missing', 'differences': []})
            continue
        found = differences(keyed(rule_spec(actual)), keyed(rule_spec(desired)))
        if rule_spec(actual) != rule_spec(desired) and not found:
            found = ['rules: observed and configured differ']
        rulesets.append({'name': desired['name'], 'id': actual.get('id'),
                         'status': 'differs' if found else 'matches', 'differences': found})
    managed = {r['name'] for r in config['rulesets']}
    drift = (not visibility['matches']
             or any(row['matches'] is False for row in settings + security + actions)
             or any(row['status'] != 'matches' for row in rulesets))
    return {'result': 'DRIFT' if drift else 'VERIFIED', 'visibility': visibility,
            'settings': settings, 'security': security, 'actions': actions, 'rulesets': rulesets,
            'unmanaged_rulesets': [r['name'] for r in state['rulesets'] if r.get('name') not in managed],
            'inherited_rulesets': sorted(state['inherited'])}


def plan(config: dict[str, Json], report: dict[str, Json], repo: str) -> Plan:
    """The writes that would make the repository match. Visibility and Actions are never written.

    Order matters. The merge methods come first, because GitHub may reject a ruleset
    that allows a merge method the repository disables. The other general settings
    come last, after the rulesets and security features: automatic deletion of merged
    branches must not start before the deletion rulesets protect the retained branches.
    """
    root = f'repos/{repo}'
    steps: Plan = []
    differing = {row['setting']: row['configured'] for row in report['settings'] if not row['matches']}
    merge = {key: value for key, value in differing.items() if key in MERGE_KEYS}
    if merge:
        steps.append(('PATCH', root, merge))
    desired = {r['name']: r for r in config['rulesets']}
    for row in report['rulesets']:
        if row['status'] == 'missing':
            steps.append(('POST', root + '/rulesets', desired[row['name']]))
        elif row['status'] == 'differs':
            steps.append(('PUT', f'{root}/rulesets/{row["id"]}', desired[row['name']]))
    analysis: dict[str, dict[str, str]] = {}
    for row in report['security']:
        if row['matches'] is not False:
            continue
        if row['feature'] in ANALYSIS_KEYS:
            analysis[row['feature']] = {'status': 'enabled' if row['configured'] else 'disabled'}
        else:
            endpoint = root + '/' + row['feature'].replace('_', '-')
            steps.append(('PUT' if row['configured'] else 'DELETE', endpoint, None))
    if analysis:
        steps.append(('PATCH', root, {'security_and_analysis': analysis}))
    others = {key: value for key, value in differing.items() if key not in MERGE_KEYS}
    if others:
        steps.append(('PATCH', root, others))
    return steps


# Output


def shown(row: Row) -> str:
    return json.dumps(row['observed']) + (' (paused)' if row.get('paused') else '')


def print_notes(config: dict[str, Json], state: dict[str, Json], report: dict[str, Json]) -> None:
    metadata = state['metadata']
    print(f'Destination: {metadata["full_name"]}; visibility: {metadata["visibility"]}; '
          f'default branch: {metadata["default_branch"]}')
    if state['inherited']:
        print('NOTE: inherited rules also apply; they are not modified or certified by this script.')
    for detail in state['rulesets']:
        if detail.get('name') not in report['unmanaged_rulesets']:
            continue
        refs = ', '.join(((detail.get('conditions') or {}).get('ref_name') or {}).get('include') or [])
        print(f'NOTE: unmanaged ruleset "{detail["name"]}" '
              f'({detail.get("target", "unknown target")}: {refs or "no ref conditions"}; '
              f'{detail.get("enforcement", "unknown enforcement")}) stays active. This script '
              'never changes or deletes it; delete it in the repository settings if a managed '
              'ruleset replaces it.')
    skipped = [row['feature'] for row in report['security'] if row['matches'] is None and row['configured']]
    if skipped:
        print(f'NOTE: {", ".join(skipped)} apply only to public repositories here; they are not '
              f'applied while the destination is {metadata.get("visibility")}. Rerun the check '
              'after making the repository public.')
    if not state['public']:
        print(f'NOTE: fork_pr_approval_policy applies only to public repositories; it is not '
              f'checked while the destination is {metadata.get("visibility")}.')
    fixes = state['security']['automated_security_fixes']
    if fixes.get('paused') and config['security']['automated_security_fixes']:
        print('NOTE: Dependabot security updates are enabled but paused, so they open no pull '
              'requests. GitHub pauses them when Dependabot pull requests go unanswered; they '
              'count as drift, and an apply that does not resume them fails its read-back.')


def print_differences(report: dict[str, Json]) -> None:
    """One line per managed value that differs from the configuration."""
    visibility = report['visibility']
    if not visibility['matches']:
        print(f'DIFF visibility: observed {visibility["observed"]}, configured {visibility["configured"]} '
              f'(never changed by this tool; see {DISTRIBUTION})')
    for row in report['settings']:
        if not row['matches']:
            print(f'DIFF setting {row["setting"]}: observed {json.dumps(row["observed"])}, '
                  f'configured {json.dumps(row["configured"])}')
    for row in report['security']:
        if row['matches'] is False:
            print(f'DIFF security {row["feature"]}: observed {shown(row)}, '
                  f'configured {json.dumps(row["configured"])} [{row["source"]}]')
    for row in report['actions']:
        if row['matches'] is False:
            print(f'DIFF actions {row["setting"]}: observed {shown(row)}, '
                  f'configured {json.dumps(row["configured"])} [{row["source"]}]')
    for row in report['rulesets']:
        if row['status'] == 'missing':
            print(f'DIFF ruleset "{row["name"]}": missing')
        for line in row['differences']:
            print(f'DIFF ruleset "{row["name"]}" ({row["id"]}) {line}')


def print_plan(steps: Plan) -> None:
    for method, endpoint, payload in steps:
        print(f'{method} {endpoint}')
        if payload is not None:
            print(json.dumps(payload, indent=2))


# Records


def markdown_record(state: dict[str, Json], report: dict[str, Json], mode: str, now: datetime) -> str:
    """A dated Markdown record of the managed settings as GitHub reports them."""
    metadata = state['metadata']

    def cell(value: Json) -> str:
        return 'on' if value is True else 'off' if value is False else 'unknown' if value is None else str(value)

    def observed(row: Row) -> str:
        if row['matches'] is None:
            return NOT_APPLIED
        return cell(row['observed']) + (' (paused)' if row.get('paused') else '')

    lines = ['# Repository settings evidence', '',
             f'- Checked: **{now.strftime("%Y-%m-%dT%H:%M:%SZ")}** with `tools/setup_repository.py` ({mode})',
             f'- Destination: `{metadata["full_name"]}`; visibility: {metadata["visibility"]} '
             f'(configured: {report["visibility"]["configured"] or "any"}); '
             f'default branch: `{metadata["default_branch"]}`',
             f'- Result: **{report["result"]}**', '',
             '## General settings', '', '| Setting | Configured | Observed |', '| --- | --- | --- |']
    lines += [f'| `{row["setting"]}` | {cell(row["configured"])} | {cell(row["observed"])} |'
              for row in report['settings']]
    lines += ['', '## Security features', '', '| Feature | Configured | Observed |', '| --- | --- | --- |']
    lines += [f'| `{row["feature"]}` | {cell(row["configured"])} | {observed(row)} |'
              for row in report['security']]
    lines += ['', '## Actions settings', '', '| Setting | Configured | Observed |', '| --- | --- | --- |']
    lines += [f'| `{row["setting"]}` | {cell(row["configured"])} | {observed(row)} |'
              for row in report['actions']]
    lines += ['', '## Managed rulesets', '', '| Ruleset | ID | Status |', '| --- | --- | --- |']
    lines += [f'| {row["name"]} | {"—" if row["id"] is None else row["id"]} | {row["status"]} |'
              for row in report['rulesets']]
    lines += ['', '## Other rules', '',
              f'- Unmanaged repository rulesets: {", ".join(report["unmanaged_rulesets"]) or "none"}',
              f'- Inherited rulesets: {", ".join(report["inherited_rulesets"]) or "none"}',
              '', 'Unmanaged settings are outside this record.']
    return '\n'.join(lines) + '\n'


def evidence_record(config_path: Path, repo: str, state: dict[str, Json], report: dict[str, Json],
                    steps: Plan, mode: str, now: datetime) -> dict[str, Json]:
    """A dated JSON record of the managed settings as GitHub reports them."""
    try:
        shown_path = config_path.resolve().relative_to(ROOT).as_posix()
    except ValueError:
        shown_path = config_path.name
    metadata = state['metadata']
    return {
        'schema_version': EVIDENCE_SCHEMA_VERSION,
        'record': 'repository-settings-evidence',
        'checked_at': now.strftime('%Y-%m-%dT%H:%M:%SZ'),
        'tool': 'tools/setup_repository.py',
        'mode': mode,
        'repository': metadata.get('full_name', repo),
        'default_branch': metadata.get('default_branch'),
        'configuration': {'path': shown_path,
                          'sha256': hashlib.sha256(config_path.read_bytes()).hexdigest()},
        'result': report['result'],
        'visibility': report['visibility'],
        'settings': report['settings'],
        'security': report['security'],
        'actions': report['actions'],
        'rulesets': report['rulesets'],
        'unmanaged_rulesets': report['unmanaged_rulesets'],
        'inherited_rulesets': report['inherited_rulesets'],
        'planned_changes': [{'method': method, 'endpoint': endpoint} for method, endpoint, _ in steps],
        'scope': 'Only the settings that the configuration manages are recorded; '
                 'everything else is outside this record.',
    }


def write_text(path: Path, text: str, what: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding='utf-8', newline='\n')
    print(f'{what} written to {path}')


# Apply


def check_prerequisites(state: dict[str, Json], report: dict[str, Json]) -> None:
    """Refuse to apply anything that cannot be written safely, before any write."""
    metadata = state['metadata']
    if metadata.get('archived') or metadata.get('disabled'):
        raise SetupError('Destination is archived or disabled.')
    if not (metadata.get('permissions') or {}).get('admin'):
        raise SetupError('Administrative access to the destination is required to apply the baseline.')
    visibility = report['visibility']
    if not visibility['matches']:
        raise SetupError(f'The destination is {visibility["observed"]}, but the configuration requires '
                         f'{visibility["configured"]}. This tool never changes visibility; follow '
                         f'{DISTRIBUTION} before changing it, then rerun.')
    unknown = [row['setting'] for row in report['settings'] if row['observed'] is None]
    unknown += [row['feature'] for row in report['security']
                if row['matches'] is not None and row['observed'] is None]
    if unknown:
        raise SetupError('Cannot apply while these managed values are unknown: ' + ', '.join(unknown)
                         + '. Rerun preview to see why GitHub did not report them.')


def check_required_checks(call: Call, repo: str, config: dict[str, Json], metadata: dict[str, Json]) -> None:
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


def readback_error(report: dict[str, Json]) -> SetupError:
    """The read-back failure, naming each Actions setting to set by hand."""
    actions = [row for row in report['actions'] if row['matches'] is False]
    rest = report['result'] == 'DRIFT' and (
        not report['visibility']['matches']
        or any(row['matches'] is False for row in report['settings'] + report['security'])
        or any(row['status'] != 'matches' for row in report['rulesets']))
    message = 'Read-back verification failed'
    message += ('; the differences above remain. Rerun preview to inspect them.' if rest
                else ': the other settings match, but this tool never writes Actions settings.')
    if actions:
        message += (' Set ' + ', '.join(f'{row["setting"]} to {json.dumps(row["configured"])} '
                                        f'(observed {shown(row)})' for row in actions)
                    + " in the repository's Settings > Actions > General, then rerun --check.")
    return SetupError(message)


def run(args: argparse.Namespace, call: Call = api, now: datetime | None = None) -> int:
    if not REPOSITORY_PATTERN.fullmatch(args.repo) or args.repo.split('/')[1] in ('.', '..'):
        raise SetupError('Destination must be an explicit owner/repository on github.com.')
    config = load_config(args.config)
    mode = 'apply' if args.apply else 'check' if args.check else 'preview'
    state = observe(call, args.repo, config)
    report = assess(config, state)
    steps = plan(config, report, args.repo)
    print_notes(config, state, report)
    print_differences(report)
    print_plan(steps)
    if args.apply:
        check_prerequisites(state, report)
    if args.apply and steps:
        if any(method == 'PUT' and '/rulesets/' in endpoint for method, endpoint, _ in steps) \
                and not args.update_existing:
            raise SetupError('A named ruleset differs. Review the preview, then use --update-existing '
                             'to replace that entire named ruleset. Additional rules in it will be removed.')
        check_required_checks(call, args.repo, config, state['metadata'])
        for method, endpoint, payload in steps:
            call(method, endpoint, payload)
            print(f'Applied {method} {endpoint}')
        state = observe(call, args.repo, config)
        report = assess(config, state)
        steps = plan(config, report, args.repo)
    stamp = now or datetime.now(timezone.utc)
    if getattr(args, 'report', None):
        write_text(args.report, markdown_record(state, report, mode, stamp), 'Record')
    if getattr(args, 'evidence', None):
        target = Path(str(args.evidence).replace('{date}', stamp.strftime('%Y-%m-%d')))
        record = evidence_record(args.config, args.repo, state, report, steps, mode, stamp)
        write_text(target, json.dumps(record, indent=2) + '\n', 'Evidence')
    if report['result'] == 'VERIFIED':
        print('VERIFIED: all documented settings match' + (' after apply' if args.apply else '')
              + '. Unmanaged settings are outside this verification.')
        return 0
    if args.apply:
        print_differences(report)
        raise readback_error(report)
    if args.check:
        print('DRIFT: documented settings do not match.')
        return 1
    print('PREVIEW ONLY: no changes made. Use --apply to execute this plan.')
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
    parser.add_argument('--evidence', type=Path,
                        help='Write a JSON evidence record of the managed settings ({date} becomes YYYY-MM-DD, UTC)')
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
