import argparse
import ast
import copy
import hashlib
import io
import json
import shutil
import sys
import tempfile
import unittest
from collections.abc import Callable
from contextlib import redirect_stdout
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'tools'))
import setup_repository as setup  # noqa: E402

REPO = 'repos/example/project'
NOW = datetime(2026, 11, 6, 9, 30, 0, tzinfo=timezone.utc)
# GitHub's documented defaults for a new repository, for the settings the configuration manages.
GITHUB_DEFAULTS = {
    'allow_merge_commit': True, 'allow_squash_merge': True, 'allow_rebase_merge': True,
    'allow_auto_merge': False, 'allow_update_branch': False, 'delete_branch_on_merge': False,
    'web_commit_signoff_required': False, 'has_issues': True, 'has_projects': True,
    'has_wiki': True, 'has_discussions': False, 'has_downloads': True,
}
MERGE_SETTINGS = {'merge': 'allow_merge_commit', 'squash': 'allow_squash_merge', 'rebase': 'allow_rebase_merge'}
ACTIONS: dict[str, dict[str, Any]] = {
    'actions/permissions': {'enabled': True, 'allowed_actions': 'all', 'sha_pinning_required': True},
    'actions/permissions/workflow': {'default_workflow_permissions': 'read',
                                     'can_approve_pull_request_reviews': False},
    'actions/permissions/fork-pr-contributor-approval': {'approval_policy': 'all_external_contributors'},
    'actions/permissions/artifact-and-log-retention': {'days': 90, 'maximum_allowed_days': 400},
}


class FakeAPI:
    """A repository as GitHub's REST API reports it, which applies the writes it receives."""

    def __init__(self) -> None:
        self.repo: dict[str, Any] = {'full_name': 'example/project', 'visibility': 'private',
                                     'default_branch': 'main', 'permissions': {'admin': True},
                                     'security_and_analysis': {}, **GITHUB_DEFAULTS}
        self.rules: list[dict[str, Any]] = []
        self.actions: dict[str, dict[str, Any]] = copy.deepcopy(ACTIONS)
        self.writes: list[tuple[str, str, Any]] = []
        self.reads: list[str] = []
        # Endpoints, relative to the repository, that answer 403.
        self.forbidden: set[str] = set()
        self.green = True
        self.ignore_patch = False
        self.reporting = False
        self.alerts = False
        self.fixes = False
        self.paused = False

    def not_found(self, endpoint: str) -> Any:
        raise setup.ApiError('GET', endpoint, '404')

    def switch_state(self, attribute: str, endpoint: str) -> Any:
        if attribute == 'reporting':
            return {'enabled': self.reporting}
        if attribute == 'alerts':
            return None if self.alerts else self.not_found(endpoint)
        if not self.alerts or not self.fixes:
            return self.not_found(endpoint)
        return {'enabled': True, 'paused': self.paused}

    def check_write_order(self, method: str, endpoint: str, payload: Any) -> None:
        """Refuse what GitHub refuses, and what would expose release branches."""
        if '/rulesets' in endpoint:
            for rule in payload['rules']:
                for merge in (rule.get('parameters') or {}).get('allowed_merge_methods', []):
                    if not self.repo[MERGE_SETTINGS[merge]]:
                        raise setup.ApiError(method, endpoint, '422')
        if method == 'PATCH' and payload.get('delete_branch_on_merge'):
            protected = any('refs/heads/release/**' in r.get('conditions', {}).get('ref_name', {}).get('include', [])
                            and any(rule['type'] == 'deletion' for rule in r.get('rules', []))
                            for r in self.rules)
            if not protected:
                raise AssertionError('branch cleanup enabled before release branches are protected')

    def __call__(self, method: str, endpoint: str, payload: Any = None) -> Any:
        if method == 'GET':
            return self.get(endpoint)
        switches = {'/private-vulnerability-reporting': 'reporting',
                    '/vulnerability-alerts': 'alerts', '/automated-security-fixes': 'fixes'}
        for suffix, attribute in switches.items():
            if endpoint.endswith(suffix):
                self.writes.append((method, endpoint, payload))
                setattr(self, attribute, method == 'PUT')
                return None
        self.check_write_order(method, endpoint, payload)
        self.writes.append((method, endpoint, copy.deepcopy(payload)))
        if method == 'POST':
            self.rules.append(dict(copy.deepcopy(payload), id=len(self.rules) + 1, source_type='Repository'))
        elif method == 'PUT':
            i = int(endpoint.rsplit('/', 1)[1])
            self.rules = [dict(copy.deepcopy(payload), id=i, source_type='Repository') if r['id'] == i else r
                          for r in self.rules]
        elif method == 'PATCH' and not self.ignore_patch:
            payload = copy.deepcopy(payload)
            analysis = payload.pop('security_and_analysis', None)
            if analysis:
                self.repo['security_and_analysis'].update(analysis)
            self.repo.update(payload)
        return {}

    def get(self, endpoint: str) -> Any:
        self.reads.append(endpoint)
        relative = endpoint.removeprefix(REPO + '/')
        if relative in self.forbidden:
            raise setup.ApiError('GET', endpoint, '403')
        switches = {'private-vulnerability-reporting': 'reporting',
                    'vulnerability-alerts': 'alerts', 'automated-security-fixes': 'fixes'}
        if relative in switches:
            return self.switch_state(switches[relative], endpoint)
        if relative.startswith('actions/permissions'):
            return copy.deepcopy(self.actions[relative]) if relative in self.actions else self.not_found(endpoint)
        if '/check-runs?' in endpoint:
            return {'check_runs': [{'name': 'Repository integrity', 'app': {'id': 15368},
                                   'status': 'completed', 'conclusion': 'success' if self.green else 'failure'}]}
        if '/branches/' in endpoint:
            return {'commit': {'sha': 'abc'}}
        if '/rulesets?' in endpoint:
            return copy.deepcopy(self.rules)
        if '/rulesets/' in endpoint:
            ruleset = copy.deepcopy(next(r for r in self.rules if r['id'] == int(endpoint.rsplit('/', 1)[1])))
            if not self.repo['permissions']['admin']:
                # GitHub omits the bypass list unless the token can write the ruleset.
                ruleset.pop('bypass_actors', None)
            return ruleset
        repo = copy.deepcopy(self.repo)
        if not repo['permissions']['admin']:
            # GitHub shows this block to administrators only.
            del repo['security_and_analysis']
        return repo

    def forbid_administration_reads(self) -> None:
        """What a token without administrator access cannot read."""
        self.repo['permissions']['admin'] = False
        self.forbidden |= {'vulnerability-alerts', 'automated-security-fixes', *ACTIONS}


class SetupTests(unittest.TestCase):
    def setUp(self) -> None:
        self.api = FakeAPI()
        self.args = argparse.Namespace(repo='example/project', config=setup.CONFIG, apply=False,
                                       check=False, update_existing=False, report=None, evidence=None)

    def run_setup(self) -> int:
        with redirect_stdout(io.StringIO()):
            return setup.run(self.args, self.api, NOW)

    def run_output(self) -> tuple[int, str]:
        output = io.StringIO()
        with redirect_stdout(output):
            code = setup.run(self.args, self.api, NOW)
        return code, output.getvalue()

    def configure(self) -> None:
        self.args.apply = True
        self.assertEqual(self.run_setup(), 0)
        self.api.writes.clear()
        self.args.apply = False

    def use_config(self, change: Callable[[dict[str, Any]], None]) -> None:
        """Run with a changed copy of the configuration, kept for the rest of the test."""
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        config = json.loads(setup.CONFIG.read_text(encoding='utf-8'))
        change(config)
        self.args.config = Path(directory.name) / 'settings.json'
        self.args.config.write_text(json.dumps(config), encoding='utf-8')

    def test_preview_never_writes(self) -> None:
        self.assertEqual(self.run_setup(), 0)
        self.assertEqual(self.api.writes, [])

    def test_check_returns_drift_without_writes(self) -> None:
        self.args.check = True
        self.assertEqual(self.run_setup(), 1)
        self.assertEqual(self.api.writes, [])

    def test_apply_verify_and_second_run_is_noop(self) -> None:
        self.args.apply = True
        self.assertEqual(self.run_setup(), 0)
        # Five rulesets, vulnerability alerts, security updates, general settings.
        self.assertEqual([w[0] for w in self.api.writes], ['POST'] * 5 + ['PUT', 'PUT', 'PATCH'])
        self.assertEqual([w[1].rsplit('/', 1)[1] for w in self.api.writes[5:7]],
                         ['vulnerability-alerts', 'automated-security-fixes'])
        self.api.writes.clear()
        self.assertEqual(self.run_setup(), 0)
        self.assertEqual(self.api.writes, [])

    def test_failed_check_blocks_all_writes(self) -> None:
        self.args.apply = True
        self.api.green = False
        with self.assertRaisesRegex(setup.SetupError, 'has not succeeded'):
            self.run_setup()
        self.assertEqual(self.api.writes, [])

    # General settings (#332).

    def test_configuration_manages_every_general_setting(self) -> None:
        config = setup.load_config(setup.CONFIG)
        self.assertEqual(list(config), list(setup.TOP_KEYS))
        self.assertEqual(set(config['repository']), set(GITHUB_DEFAULTS))
        changed = {key for key, value in config['repository'].items() if GITHUB_DEFAULTS[key] != value}
        self.assertEqual(changed, {'delete_branch_on_merge', 'allow_update_branch', 'has_projects',
                                   'has_downloads'})
        # The settings reference lists the changed values only.
        settings = (ROOT / 'docs/repo_governance/SETTINGS.md').read_text(encoding='utf-8')
        table = settings.split('## 🏷️ General settings')[1].split('\n## ')[0]
        for key in changed:
            self.assertIn(f'(`{key}: {json.dumps(GITHUB_DEFAULTS[key])}`)', table)
        self.assertEqual(table.count('\n| ') - 2, len(changed))
        self.assertIsNone(config['visibility'])

    def test_each_general_setting_is_reported_and_applied(self) -> None:
        for key in setup.GENERAL_KEYS:
            with self.subTest(setting=key):
                self.api = FakeAPI()
                self.configure()
                configured = setup.load_config(setup.CONFIG)['repository'][key]
                self.api.repo[key] = not configured
                self.args.check = True
                code, output = self.run_output()
                self.assertEqual(code, 1)
                self.assertIn(f'DIFF setting {key}: observed {json.dumps(not configured)}, '
                              f'configured {json.dumps(configured)}', output)
                self.args.check, self.args.apply = False, True
                self.assertEqual(self.run_setup(), 0)
                self.assertEqual(self.api.writes, [('PATCH', REPO, {key: configured})])
                self.args.apply = False

    def test_merge_methods_are_enabled_before_rulesets_and_branch_cleanup_last(self) -> None:
        # The fake refuses a ruleset that allows a disabled merge method, and branch
        # cleanup before the release branches are protected, as the order requires.
        self.api.repo.update(allow_merge_commit=False, allow_squash_merge=False)
        self.args.apply = True
        self.assertEqual(self.run_setup(), 0)
        methods = [(method, endpoint.removeprefix(REPO)) for method, endpoint, _ in self.api.writes]
        self.assertEqual(methods, [('PATCH', '')] + [('POST', '/rulesets')] * 5
                         + [('PUT', '/vulnerability-alerts'), ('PUT', '/automated-security-fixes'),
                            ('PATCH', '')])
        self.assertEqual(self.api.writes[0][2], {'allow_merge_commit': True, 'allow_squash_merge': True})
        self.assertEqual(self.api.writes[-1][2], {'allow_update_branch': True, 'delete_branch_on_merge': True,
                                                  'has_projects': False, 'has_downloads': False})

    def test_plan_puts_merge_methods_first_and_other_settings_last(self) -> None:
        self.api.repo.update(allow_rebase_merge=False, delete_branch_on_merge=False)
        config = setup.load_config(setup.CONFIG)
        state = setup.observe(self.api, 'example/project', config)
        steps = setup.plan(config, setup.assess(config, state), 'example/project')
        self.assertEqual(steps[0], ('PATCH', REPO, {'allow_rebase_merge': True}))
        self.assertEqual(steps[-1][0], 'PATCH')
        self.assertIn('delete_branch_on_merge', steps[-1][2])
        self.assertNotIn('allow_rebase_merge', steps[-1][2])

    def test_visibility_is_optional_and_never_written(self) -> None:
        self.args.check = True
        code, output = self.run_output()
        self.assertNotIn('DIFF visibility', output)
        self.configure()
        self.use_config(lambda config: config.update(visibility='public'))
        self.args.check = True
        code, output = self.run_output()
        self.assertEqual(code, 1)
        self.assertIn('DIFF visibility: observed private, configured public (never changed by this tool; '
                      'see docs/repo_governance/DISTRIBUTION.md)', output)
        self.args.check, self.args.apply = False, True
        with self.assertRaisesRegex(setup.SetupError, 'never changes visibility.*DISTRIBUTION.md'):
            self.run_setup()
        self.assertEqual(self.api.writes, [])
        self.use_config(lambda config: config.update(visibility='private'))
        self.args.apply, self.args.check = False, True
        self.assertEqual(self.run_setup(), 0)

    def broken(self, change: Callable[[dict[str, Any]], Any]) -> None:
        config = json.loads(setup.CONFIG.read_text(encoding='utf-8'))
        change(config)
        with patch.object(Path, 'read_text', return_value=json.dumps(config)):
            setup.load_config(Path('settings.json'))

    def test_invalid_configurations_are_rejected(self) -> None:
        def ruleset(config: dict[str, Any], name: str) -> dict[str, Any]:
            found: dict[str, Any] = next(r for r in config['rulesets'] if r['name'] == name)
            return found

        def rule(config: dict[str, Any], name: str, kind: str) -> dict[str, Any]:
            found: dict[str, Any] = next(r for r in ruleset(config, name)['rules'] if r['type'] == kind)
            return found

        cases: dict[str, Callable[[dict[str, Any]], Any]] = {
            'Push protection requires': lambda c: c['security'].update(secret_scanning=False),
            'security updates require': lambda c: c['security'].update(vulnerability_alerts=False),
            'schema': lambda c: c.update(schema_version=3),
            'visibility must be': lambda c: c.update(visibility='internal'),
            '12 documented boolean repository settings': lambda c: c['repository'].update(has_pages=False),
            'documented Actions settings': lambda c: c['actions'].pop('retention_days'),
            'Invalid Actions settings': lambda c: c['actions'].update(default_workflow_permissions='none'),
            'positive integer': lambda c: c['actions'].update(retention_days=True),
            'Ruleset "Protect main" allows squash merges, but the configuration sets allow_squash_merge':
                lambda c: (c['repository'].update(allow_squash_merge=False),
                           rule(c, 'Protect release branches', 'pull_request')['parameters'].update(
                               allowed_merge_methods=['merge']),
                           rule(c, 'Protect main', 'pull_request')['parameters'].update(
                               allowed_merge_methods=['squash'])),
            'Ruleset "Protect release branches" allows squash merges':
                lambda c: c['repository'].update(allow_squash_merge=False),
            'Ruleset "Protect main" allows an unknown merge method: fast_forward':
                lambda c: rule(c, 'Protect main', 'pull_request')['parameters'].update(
                    allowed_merge_methods=['fast_forward']),
            'Ruleset "Protect main" has a required status check without':
                lambda c: rule(c, 'Protect main', 'required_status_checks')['parameters'][
                    'required_status_checks'].append({'context': 'Other check'}),
            'Ruleset "Protect release tags" repeats a rule type':
                lambda c: ruleset(c, 'Protect release tags')['rules'].append({'type': 'deletion'}),
            'Ruleset "Protect main" must not contain bypass actors':
                lambda c: ruleset(c, 'Protect main').update(
                    bypass_actors=[{'actor_id': 5, 'actor_type': 'RepositoryRole', 'bypass_mode': 'always'}]),
            'Ruleset "Protect main" must be an active': lambda c: ruleset(c, 'Protect main').update(
                enforcement='evaluate'),
            'names must be unique: Protect main': lambda c: c['rulesets'].append(
                copy.deepcopy(ruleset(c, 'Protect main'))),
            'at least one ruleset': lambda c: c.update(rulesets=[]),
        }
        for message, change in cases.items():
            with self.subTest(message=message):
                with self.assertRaisesRegex(setup.SetupError, message):
                    self.broken(change)

    def test_any_number_of_rulesets_is_accepted(self) -> None:
        def keep_one(config: dict[str, Any]) -> None:
            config['rulesets'] = config['rulesets'][:1]
        self.use_config(keep_one)
        self.args.apply = True
        self.assertEqual(self.run_setup(), 0)
        self.assertEqual(len(self.api.rules), 1)

    # Read-only modes and tolerant reads (#333).

    def test_preview_and_check_run_without_administrator_access(self) -> None:
        self.configure()
        self.api.forbid_administration_reads()
        code, output = self.run_output()
        self.assertEqual(code, 0)
        self.assertIn('PREVIEW ONLY', output)
        self.args.check = True
        code, output = self.run_output()
        self.assertEqual(code, 1)
        self.assertIn('DIFF security vulnerability_alerts: observed null, configured true '
                      '[GET vulnerability-alerts (HTTP 403: unreadable)]', output)
        self.assertIn('DIFF security automated_security_fixes: observed null, configured true '
                      '[GET automated-security-fixes (HTTP 403: unreadable)]', output)
        self.assertIn('DIFF actions retention_days: observed null, configured 90 '
                      '[GET actions/permissions/artifact-and-log-retention (HTTP 403: unreadable)]', output)
        self.assertEqual(self.api.writes, [])

    def test_a_hidden_bypass_list_is_unknown_not_verified(self) -> None:
        self.configure()
        # Someone adds a bypass actor that a token without write access cannot see.
        self.api.rules[0]['bypass_actors'] = [{'actor_id': 5, 'actor_type': 'RepositoryRole',
                                               'bypass_mode': 'always'}]
        self.api.repo['permissions']['admin'] = False
        name = self.api.rules[0]['name']
        self.args.check = True
        with tempfile.TemporaryDirectory() as directory:
            self.args.evidence = Path(directory) / 'evidence.json'
            code, output = self.run_output()
            record = json.loads(self.args.evidence.read_text(encoding='utf-8'))
        self.assertEqual(code, 1)
        self.assertIn(f'DIFF ruleset "{name}" (1) bypass_actors: not reported '
                      '(needs write access to the ruleset)', output)
        self.assertEqual(record['result'], 'DRIFT')
        self.assertEqual(next(r for r in record['rulesets'] if r['name'] == name)['status'], 'unknown')
        self.assertEqual(self.api.writes, [])
        # With write access the bypass actor is visible and differs.
        self.api.repo['permissions']['admin'] = True
        self.args.evidence = None
        code, output = self.run_output()
        self.assertEqual(code, 1)
        self.assertIn(f'DIFF ruleset "{name}" (1) bypass_actors: observed', output)

    def test_apply_refuses_a_ruleset_it_cannot_verify(self) -> None:
        report = {'visibility': {'matches': True}, 'settings': [], 'security': [],
                  'rulesets': [{'name': 'Protect main', 'id': 1, 'status': 'unknown', 'differences': []}]}
        state = {'metadata': {'permissions': {'admin': True}}}
        with self.assertRaisesRegex(setup.SetupError, 'unknown: ruleset "Protect main"'):
            setup.check_prerequisites(state, report)

    def test_public_secret_scanning_is_unknown_without_administrator_access(self) -> None:
        self.api.repo['visibility'] = 'public'
        self.configure()
        self.api.repo['permissions']['admin'] = False
        self.args.check = True
        code, output = self.run_output()
        self.assertEqual(code, 1)
        self.assertIn('DIFF security secret_scanning: observed null, configured true '
                      '[security_and_analysis.secret_scanning (not reported without administrator access)]',
                      output)

    def test_non_admin_blocks_all_writes(self) -> None:
        self.args.apply = True
        self.api.repo['permissions']['admin'] = False
        with self.assertRaisesRegex(setup.SetupError, 'Administrative'):
            self.run_setup()
        self.assertEqual(self.api.writes, [])

    def test_archived_destination_is_refused_only_by_apply(self) -> None:
        self.api.repo['archived'] = True
        self.args.check = True
        self.assertEqual(self.run_setup(), 1)
        self.args.check, self.args.apply = False, True
        with self.assertRaisesRegex(setup.SetupError, 'archived'):
            self.run_setup()
        self.assertEqual(self.api.writes, [])

    def test_read_returns_the_status_of_a_failed_get(self) -> None:
        def forbidden(method: str, endpoint: str) -> Any:
            raise setup.ApiError(method, endpoint, '403')

        def broken(method: str, endpoint: str) -> Any:
            raise ValueError('not an API failure')

        self.assertEqual(setup.read(forbidden, 'repos/a/b/vulnerability-alerts'), (None, '403'))
        self.assertEqual(setup.read(lambda method, endpoint: self.api.not_found(endpoint), 'repos/a/b/x'),
                         (None, '404'))
        with self.assertRaises(ValueError):
            setup.read(broken, 'repos/a/b/x')

    def test_unreadable_security_feature_is_unknown_and_blocks_apply(self) -> None:
        self.api.repo['visibility'] = 'public'
        self.configure()
        self.api.forbidden.add('private-vulnerability-reporting')
        self.api.repo['has_wiki'] = False
        self.args.check = True
        code, output = self.run_output()
        self.assertEqual(code, 1)
        self.assertIn('DIFF security private_vulnerability_reporting: observed null, configured true '
                      '[GET private-vulnerability-reporting (HTTP 403: unreadable)]', output)
        self.args.check, self.args.apply = False, True
        with self.assertRaisesRegex(setup.SetupError, 'unknown: private_vulnerability_reporting'):
            self.run_setup()
        self.assertEqual(self.api.writes, [])

    def test_enabled_security_updates_prove_the_alerts_on(self) -> None:
        self.configure()
        self.api.forbidden.add('vulnerability-alerts')
        self.args.check = True
        self.assertEqual(self.run_setup(), 0)
        metadata = self.api('GET', REPO)
        state = setup.security_states(self.api, REPO, metadata, False)['vulnerability_alerts']
        self.assertTrue(state['observed'])
        self.assertIn('HTTP 403', state['source'])
        self.assertIn('Dependabot security updates are enabled and require them', state['source'])

    def test_unreadable_alerts_without_known_security_updates_stay_unknown(self) -> None:
        self.configure()
        self.api.forbidden |= {'vulnerability-alerts', 'automated-security-fixes'}
        self.args.check = True
        code, output = self.run_output()
        self.assertEqual(code, 1)
        self.assertIn('DIFF security vulnerability_alerts: observed null', output)
        self.assertIn('DIFF security automated_security_fixes: observed null', output)

    def test_differences_are_printed_by_path_before_the_plan(self) -> None:
        self.configure()
        main = next(r for r in self.api.rules if r['name'] == 'Protect main')
        pull_request = next(rule for rule in main['rules'] if rule['type'] == 'pull_request')
        pull_request['parameters']['allowed_merge_methods'] = ['merge', 'squash']
        self.api.rules = [r for r in self.api.rules if r['name'] != 'Protect traffic history']
        self.api.repo['has_projects'] = True
        code, output = self.run_output()
        self.assertEqual(code, 0)
        lines = output.splitlines()
        self.assertIn(f'DIFF ruleset "Protect main" ({main["id"]}) '
                      'rules.pull_request.allowed_merge_methods: observed ["merge", "squash"], '
                      'configured ["merge"]', lines)
        self.assertIn('DIFF ruleset "Protect traffic history": missing', lines)
        self.assertIn('DIFF setting has_projects: observed true, configured false', lines)
        last_difference = max(i for i, line in enumerate(lines) if line.startswith('DIFF '))
        self.assertLess(last_difference, lines.index(f'PUT {REPO}/rulesets/{main["id"]}'))
        self.assertLess(last_difference, lines.index(f'POST {REPO}/rulesets'))

    def test_existing_rule_changes_require_explicit_flag(self) -> None:
        rule = copy.deepcopy(setup.load_config(setup.CONFIG)['rulesets'][0])
        rule.update(id=1, source_type='Repository', enforcement='disabled')
        self.api.rules = [rule]
        self.args.apply = True
        with self.assertRaisesRegex(setup.SetupError, 'named ruleset differs'):
            self.run_setup()
        self.assertEqual(self.api.writes, [])
        self.args.update_existing = True
        self.assertEqual(self.run_setup(), 0)

    def test_unrelated_rules_are_preserved(self) -> None:
        self.api.rules = [{'name': 'Other rule', 'id': 10, 'source_type': 'Repository'}]
        self.args.apply = True
        self.run_setup()
        self.assertEqual(self.api.rules[0]['name'], 'Other rule')

    def test_preview_names_unmanaged_rulesets_without_writing(self) -> None:
        self.api.rules = [{'name': 'main', 'id': 10, 'source_type': 'Repository', 'target': 'branch',
                           'enforcement': 'active',
                           'conditions': {'ref_name': {'include': ['~DEFAULT_BRANCH'], 'exclude': []}}}]
        code, output = self.run_output()
        self.assertEqual(code, 0)
        self.assertIn('unmanaged ruleset "main" (branch: ~DEFAULT_BRANCH; active)', output)
        self.assertEqual(self.api.writes, [])

    def test_inherited_rulesets_are_noted_not_managed(self) -> None:
        self.configure()
        self.api.rules.append({'name': 'Organization rule', 'id': 99, 'source_type': 'Organization'})
        self.args.check = True
        code, output = self.run_output()
        self.assertEqual(code, 0)
        self.assertIn('NOTE: inherited rules also apply', output)
        self.assertNotIn('rulesets/99', ' '.join(self.api.reads))

    # Security features.

    def test_public_repository_enables_vulnerability_reporting_once(self) -> None:
        self.api.repo['visibility'] = 'public'
        self.args.apply = True
        self.assertEqual(self.run_setup(), 0)
        self.assertIn(('PUT', 'repos/example/project/private-vulnerability-reporting', None),
                      self.api.writes)
        self.assertEqual(self.api.writes[-1][0], 'PATCH')
        self.api.writes.clear()
        self.assertEqual(self.run_setup(), 0)
        self.assertEqual(self.api.writes, [])

    def test_private_repository_skips_public_only_features_with_a_note(self) -> None:
        self.args.apply = True
        code, output = self.run_output()
        self.assertEqual(code, 0)
        self.assertIn('private_vulnerability_reporting, secret_scanning, '
                      'secret_scanning_push_protection apply only to public repositories', output)
        self.assertFalse(any('private-vulnerability' in endpoint or 'security_and_analysis' in str(payload)
                             for _, endpoint, payload in self.api.writes))

    def test_public_repository_enables_secret_scanning_and_push_protection(self) -> None:
        self.api.repo['visibility'] = 'public'
        self.args.apply = True
        self.assertEqual(self.run_setup(), 0)
        self.assertIn(('PATCH', 'repos/example/project',
                       {'security_and_analysis': {'secret_scanning': {'status': 'enabled'},
                                                  'secret_scanning_push_protection': {'status': 'enabled'}}}),
                      self.api.writes)
        self.assertEqual(self.api.writes[-1][2], {'allow_update_branch': True, 'delete_branch_on_merge': True,
                                                  'has_projects': False, 'has_downloads': False})

    def test_dependent_security_features_are_rejected(self) -> None:
        config = setup.load_config(setup.CONFIG)
        for disabled, needed_by in (('secret_scanning', 'Push protection'),
                                    ('vulnerability_alerts', 'security updates')):
            broken = copy.deepcopy(config)
            broken['security'][disabled] = False
            with patch.object(Path, 'read_text', return_value=json.dumps(broken)):
                with self.assertRaisesRegex(setup.SetupError, needed_by):
                    setup.load_config(Path('settings.json'))

    def test_check_reports_disabled_vulnerability_reporting_as_drift(self) -> None:
        self.api.repo['visibility'] = 'public'
        self.configure()
        self.api.reporting = False
        self.args.check = True
        self.assertEqual(self.run_setup(), 1)

    def test_paused_security_updates_are_drift_not_enabled(self) -> None:
        self.configure()
        self.api.paused = True
        with tempfile.TemporaryDirectory() as directory:
            self.args.report = Path(directory) / 'settings.md'
            self.args.check = True
            code, output = self.run_output()
            record = self.args.report.read_text(encoding='utf-8')
        self.assertEqual(code, 1)
        self.assertIn('enabled but paused', output)
        self.assertIn('DIFF security automated_security_fixes: observed true (paused), configured true', output)
        self.assertIn('PUT repos/example/project/automated-security-fixes', output)
        self.assertIn('- Result: **DRIFT**', record)
        self.assertIn('| `automated_security_fixes` | on | on (paused) |', record)
        self.assertEqual(self.api.writes, [])

    def test_apply_that_leaves_updates_paused_fails_the_readback(self) -> None:
        self.configure()
        self.api.paused = True
        self.args.apply = True
        with self.assertRaisesRegex(setup.SetupError, 'Read-back'):
            self.run_setup()
        self.assertEqual([w[:2] for w in self.api.writes],
                         [('PUT', 'repos/example/project/automated-security-fixes')])

    def test_unpaused_security_updates_verify(self) -> None:
        self.configure()
        self.args.check = True
        self.assertEqual(self.run_setup(), 0)

    def test_paused_updates_still_count_as_enabled_when_configured_off(self) -> None:
        self.configure()
        self.api.paused = True
        config = setup.load_config(setup.CONFIG)
        config['security']['automated_security_fixes'] = False
        state = setup.observe(self.api, 'example/project', config)
        steps = setup.plan(config, setup.assess(config, state), 'example/project')
        self.assertEqual(steps, [('DELETE', 'repos/example/project/automated-security-fixes', None)])

    # Actions settings, reported and never written (#335).

    def test_actions_settings_are_reported_never_written(self) -> None:
        self.configure()
        self.api.actions['actions/permissions/workflow']['default_workflow_permissions'] = 'write'
        self.api.actions['actions/permissions/artifact-and-log-retention']['days'] = 30
        self.args.check = True
        code, output = self.run_output()
        self.assertEqual(code, 1)
        self.assertIn('DIFF actions default_workflow_permissions: observed "write", configured "read" '
                      '[GET actions/permissions/workflow default_workflow_permissions]', output)
        self.assertIn('DIFF actions retention_days: observed 30, configured 90', output)
        self.args.check, self.args.apply = False, True
        with self.assertRaises(setup.SetupError) as caught:
            self.run_setup()
        message = str(caught.exception)
        self.assertIn('Read-back verification failed', message)
        self.assertIn('default_workflow_permissions to "read"', message)
        self.assertIn('retention_days to 90', message)
        self.assertIn("Settings > Actions > General, then rerun --check", message)
        self.assertEqual(self.api.writes, [])

    def test_apply_writes_the_rest_then_names_the_actions_settings(self) -> None:
        self.api.actions['actions/permissions']['sha_pinning_required'] = False
        self.args.apply = True
        with self.assertRaisesRegex(setup.SetupError, 'Read-back.*sha_pinning_required to true'):
            self.run_setup()
        self.assertEqual(len(self.api.writes), 8)
        self.assertFalse(any('actions' in endpoint for _, endpoint, _ in self.api.writes))
        self.args.apply, self.args.check = False, True
        code, output = self.run_output()
        self.assertEqual(code, 1)
        self.assertEqual([line for line in output.splitlines() if line.startswith('DIFF')],
                         ['DIFF actions sha_pinning_required: observed false, configured true '
                          '[GET actions/permissions sha_pinning_required]'])

    def test_unreadable_or_absent_actions_setting_is_unknown(self) -> None:
        self.configure()
        del self.api.actions['actions/permissions']['sha_pinning_required']
        self.api.forbidden.add('actions/permissions/workflow')
        self.api.reads.clear()
        self.args.check = True
        code, output = self.run_output()
        self.assertEqual(code, 1)
        self.assertIn('DIFF actions sha_pinning_required: observed null, configured true '
                      '[GET actions/permissions (no sha_pinning_required field)]', output)
        self.assertIn('DIFF actions can_approve_pull_request_reviews: observed null, configured false '
                      '[GET actions/permissions/workflow (HTTP 403: unreadable)]', output)
        # One read per endpoint, even for two settings.
        self.assertEqual(self.api.reads.count(REPO + '/actions/permissions/workflow'), 1)

    def test_fork_approval_policy_applies_to_public_repositories_only(self) -> None:
        self.configure()
        self.api.actions['actions/permissions/fork-pr-contributor-approval']['approval_policy'] = \
            'first_time_contributors'
        self.args.check = True
        self.api.reads.clear()
        self.assertEqual(self.run_setup(), 0)
        self.assertNotIn(REPO + '/actions/permissions/fork-pr-contributor-approval', self.api.reads)
        self.api.repo['visibility'] = 'public'
        self.api.reporting = True
        self.api.repo['security_and_analysis'] = {'secret_scanning': {'status': 'enabled'},
                                                  'secret_scanning_push_protection': {'status': 'enabled'}}
        code, output = self.run_output()
        self.assertEqual(code, 1)
        self.assertIn('DIFF actions fork_pr_approval_policy: observed "first_time_contributors", '
                      'configured "all_external_contributors"', output)

    # Records: Markdown (--report) and JSON (--evidence, #334).

    def test_report_records_drift_then_verification(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            self.args.report = Path(directory) / 'evidence' / 'settings.md'
            self.args.check = True
            self.assertEqual(self.run_setup(), 1)
            drift = self.args.report.read_text(encoding='utf-8')
            self.assertIn('- Result: **DRIFT**', drift)
            self.assertIn('visibility: private (configured: any)', drift)
            self.assertIn('| Protect main | — | missing |', drift)
            self.assertIn('| `has_projects` | off | on |', drift)
            self.assertIn('| `allow_rebase_merge` | on | on |', drift)
            self.assertIn('| `secret_scanning` | on | not applied (public repositories only) |', drift)
            self.assertIn('## Actions settings', drift)
            self.assertIn('| `default_workflow_permissions` | read | read |', drift)
            self.assertIn('| `fork_pr_approval_policy` | all_external_contributors | '
                          'not applied (public repositories only) |', drift)
            self.args.check, self.args.apply = False, True
            self.assertEqual(self.run_setup(), 0)
            verified = self.args.report.read_text(encoding='utf-8')
            self.assertIn('- Result: **VERIFIED**', verified)
            self.assertIn('| Protect main | 5 | matches |', verified)
            self.assertIn('| `vulnerability_alerts` | on | on |', verified)
            self.assertIn('| `retention_days` | 90 | 90 |', verified)
            self.assertFalse((Path(directory) / 'evidence' / 'settings.json').exists())
        self.assertEqual([w for w in self.api.writes if 'evidence' in w[1]], [])

    def test_evidence_records_drift_then_verification(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            self.args.evidence = Path(directory) / 'records' / 'settings-evidence-{date}.json'
            written = Path(directory) / 'records' / 'settings-evidence-2026-11-06.json'
            self.args.check = True
            self.assertEqual(self.run_setup(), 1)
            raw = written.read_bytes()
            drift = json.loads(raw)
            self.assertNotIn(b'\r', raw)
            self.assertTrue(raw.endswith(b'}\n'))
            self.assertEqual(raw.decode('utf-8'), json.dumps(drift, indent=2) + '\n')
            self.assertEqual(list(drift), [
                'schema_version', 'record', 'checked_at', 'tool', 'mode', 'repository', 'default_branch',
                'configuration', 'result', 'visibility', 'settings', 'security', 'actions', 'rulesets',
                'unmanaged_rulesets', 'inherited_rulesets', 'planned_changes', 'scope'])
            self.assertEqual({key: drift[key] for key in ('schema_version', 'record', 'checked_at', 'tool',
                                                          'mode', 'repository', 'default_branch', 'result')},
                             {'schema_version': 1, 'record': 'repository-settings-evidence',
                              'checked_at': '2026-11-06T09:30:00Z', 'tool': 'tools/setup_repository.py',
                              'mode': 'check', 'repository': 'example/project', 'default_branch': 'main',
                              'result': 'DRIFT'})
            self.assertEqual(drift['configuration'], {
                'path': 'docs/setup/repository-settings.json',
                'sha256': hashlib.sha256(setup.CONFIG.read_bytes()).hexdigest()})
            self.assertEqual(drift['visibility'], {'configured': None, 'observed': 'private', 'matches': True})
            self.assertIn({'setting': 'has_projects', 'configured': False, 'observed': True, 'matches': False},
                          drift['settings'])
            self.assertIn({'feature': 'secret_scanning', 'configured': True, 'observed': None,
                           'source': 'not applied (public repositories only)', 'matches': None},
                          drift['security'])
            self.assertIn({'setting': 'retention_days', 'configured': 90, 'observed': 90,
                           'source': 'GET actions/permissions/artifact-and-log-retention days', 'matches': True},
                          drift['actions'])
            self.assertEqual({r['status'] for r in drift['rulesets']}, {'missing'})
            self.assertEqual(drift['planned_changes'][0], {'method': 'POST', 'endpoint': REPO + '/rulesets'})
            self.assertTrue(all(set(step) == {'method', 'endpoint'} for step in drift['planned_changes']))
            self.assertNotIn('"rules"', json.dumps(drift['planned_changes']))
            self.assertIn('Only the settings', drift['scope'])

            self.args.check, self.args.apply = False, True
            self.assertEqual(self.run_setup(), 0)
            verified = json.loads(written.read_text(encoding='utf-8'))
            self.assertEqual((verified['mode'], verified['result'], verified['planned_changes']),
                             ('apply', 'VERIFIED', []))
            self.assertEqual({r['status'] for r in verified['rulesets']}, {'matches'})
            self.assertTrue(all(row['matches'] is not False
                                for row in verified['settings'] + verified['security'] + verified['actions']))
            self.assertEqual(sorted(p.name for p in written.parent.iterdir()), [written.name])

    def test_evidence_hashes_the_configuration_bytes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            copied = Path(directory) / 'reviewed-settings.json'
            shutil.copyfile(setup.CONFIG, copied)
            self.args.config = copied
            self.args.evidence = Path(directory) / 'evidence.json'
            self.run_setup()
            first = json.loads(self.args.evidence.read_text(encoding='utf-8'))['configuration']
            copied.write_bytes(copied.read_bytes() + b'\n')
            self.run_setup()
            second = json.loads(self.args.evidence.read_text(encoding='utf-8'))['configuration']
        self.assertEqual(first['path'], 'reviewed-settings.json')
        self.assertEqual(first['sha256'], hashlib.sha256(setup.CONFIG.read_bytes()).hexdigest())
        self.assertNotEqual(first['sha256'], second['sha256'])

    def test_required_status_check_is_the_integrity_job(self) -> None:
        from test_workflow_controls import REQUIRED_CHECK
        config = setup.load_config(setup.CONFIG)
        contexts = {check['context'] for ruleset in config['rulesets'] for rule in ruleset['rules']
                    if rule['type'] == 'required_status_checks'
                    for check in rule['parameters']['required_status_checks']}
        self.assertEqual(contexts, {REQUIRED_CHECK})

    def test_duplicate_names_rejected_before_writes(self) -> None:
        name = setup.load_config(setup.CONFIG)['rulesets'][0]['name']
        self.api.rules = [{'name': name, 'id': 1}, {'name': name, 'id': 2}]
        with self.assertRaisesRegex(setup.SetupError, 'Ambiguous'):
            self.run_setup()
        self.assertFalse(self.api.writes)

    def test_readback_detects_ignored_changes(self) -> None:
        self.args.apply = True
        self.api.ignore_patch = True
        with self.assertRaisesRegex(setup.SetupError, 'Read-back verification failed; the differences'):
            self.run_setup()

    def test_pagination_reads_beyond_first_page(self) -> None:
        calls: list[str] = []

        def call(method: str, endpoint: str) -> Any:
            calls.append(endpoint)
            return list(range(100)) if endpoint.endswith('&page=1') else [100]
        self.assertEqual(len(setup.pages(call, 'repos/x/y/rulesets?includes_parents=true')), 101)
        self.assertEqual(len(calls), 2)

    def test_order_is_not_drift_but_extra_rule_is(self) -> None:
        r = copy.deepcopy(setup.load_config(setup.CONFIG)['rulesets'][-1])
        reordered = copy.deepcopy(r)
        reordered['rules'].reverse()
        self.assertEqual(setup.rule_spec(r), setup.rule_spec(reordered))
        reordered['rules'].append({'type': 'creation'})
        self.assertNotEqual(setup.rule_spec(r), setup.rule_spec(reordered))

    def test_cli_uses_explicit_method_host_and_stdin(self) -> None:
        result = argparse.Namespace(returncode=0, stdout='{}', stderr='')
        with patch('setup_repository.subprocess.run', return_value=result) as run, \
                patch.dict('os.environ', {'GH_DEBUG': 'api'}):
            setup.api('PATCH', 'repos/a/b', {'has_projects': False})
        cmd = run.call_args.args[0]
        self.assertIn('github.com', cmd)
        self.assertIn('PATCH', cmd)
        self.assertEqual(cmd[-2:], ['--input', '-'])
        self.assertEqual(run.call_args.kwargs['input'], '{"has_projects": false}')
        # gh writes UTF-8 whatever the locale code page is.
        self.assertEqual(run.call_args.kwargs['encoding'], 'utf-8')
        self.assertNotIn('GH_DEBUG', run.call_args.kwargs['env'])

    def test_api_error_does_not_print_server_detail(self) -> None:
        result = argparse.Namespace(returncode=1, stdout='', stderr='secret token HTTP 403')
        with patch('setup_repository.subprocess.run', return_value=result):
            with self.assertRaises(setup.ApiError) as caught:
                setup.api('GET', 'repos/a/b')
        self.assertEqual(caught.exception.status, '403')
        self.assertIn('HTTP 403', str(caught.exception))
        self.assertNotIn('secret token', str(caught.exception))


class VersionGuardTests(unittest.TestCase):
    def test_the_module_imports_on_older_python_versions(self) -> None:
        # main() refuses Python before 3.11 with a message, but only if the module imports
        # first: a subscripted generic evaluated at import, such as list[str] or
        # Callable[..., Any], stops Python 3.8 with a TypeError before main() runs.
        tree = ast.parse(Path(setup.__file__).read_text(encoding='utf-8'))
        generics = {'list', 'dict', 'tuple', 'set', 'frozenset', 'type'}
        generics.update(alias.asname or alias.name for node in tree.body
                        if isinstance(node, ast.ImportFrom)
                        and node.module in {'collections.abc', 'typing'}
                        for alias in node.names)
        evaluated: list[str] = []
        for node in tree.body:
            if isinstance(node, ast.If) and ast.unparse(node.test) == 'TYPE_CHECKING':
                continue
            values = [node.value] if isinstance(node, (ast.Assign, ast.AnnAssign)) and \
                node.value is not None else []
            evaluated.extend(ast.unparse(item) for value in values for item in ast.walk(value)
                             if isinstance(item, ast.Subscript)
                             and isinstance(item.value, ast.Name) and item.value.id in generics)
        self.assertEqual(evaluated, [])


if __name__ == '__main__':
    unittest.main()
