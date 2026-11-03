import argparse
import copy
import io
import json
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
import setup_repository as setup


class FakeAPI:
    def __init__(self):
        self.repo = {'full_name': 'example/project', 'visibility': 'private',
                     'default_branch': 'main', 'permissions': {'admin': True},
                     'allow_merge_commit': True, 'allow_squash_merge': True,
                     'delete_branch_on_merge': False, 'allow_update_branch': False,
                     'has_projects': True, 'has_downloads': True}
        self.rules = []
        self.writes = []
        self.green = True
        self.ignore_patch = False
        self.reporting = False
        self.alerts = False
        self.fixes = False
        self.paused = False

    def not_found(self, endpoint):
        raise setup.SetupError(f'GET {endpoint} failed (HTTP 404). Check gh authentication.')

    def switch_state(self, attribute, endpoint):
        if attribute == 'reporting':
            return {'enabled': self.reporting}
        if attribute == 'alerts':
            return None if self.alerts else self.not_found(endpoint)
        if not self.alerts or not self.fixes:
            return self.not_found(endpoint)
        return {'enabled': True, 'paused': self.paused}

    def __call__(self, method, endpoint, payload=None):
        switches = {'/private-vulnerability-reporting': 'reporting',
                    '/vulnerability-alerts': 'alerts', '/automated-security-fixes': 'fixes'}
        for suffix, attribute in switches.items():
            if endpoint.endswith(suffix):
                if method == 'GET':
                    return self.switch_state(attribute, endpoint)
                self.writes.append((method, endpoint, payload))
                setattr(self, attribute, method == 'PUT')
                return None
        if method != 'GET':
            self.writes.append((method, endpoint, copy.deepcopy(payload)))
            if method == 'POST':
                self.rules.append(dict(copy.deepcopy(payload), id=len(self.rules) + 1, source_type='Repository'))
            elif method == 'PUT':
                i = int(endpoint.rsplit('/', 1)[1])
                self.rules = [dict(copy.deepcopy(payload), id=i, source_type='Repository') if r['id'] == i else r for r in self.rules]
            elif method == 'PATCH' and not self.ignore_patch:
                payload = copy.deepcopy(payload)
                analysis = payload.pop('security_and_analysis', None)
                if analysis:
                    self.repo.setdefault('security_and_analysis', {}).update(analysis)
                self.repo.update(payload)
            return {}
        if '/check-runs?' in endpoint:
            return {'check_runs': [{'name': 'Repository integrity', 'app': {'id': 15368},
                                   'status': 'completed', 'conclusion': 'success' if self.green else 'failure'}]}
        if '/branches/' in endpoint:
            return {'commit': {'sha': 'abc'}}
        if '/rulesets?' in endpoint:
            return copy.deepcopy(self.rules)
        if '/rulesets/' in endpoint:
            return copy.deepcopy(next(r for r in self.rules if r['id'] == int(endpoint.rsplit('/', 1)[1])))
        return copy.deepcopy(self.repo)


class SetupTests(unittest.TestCase):
    def setUp(self):
        self.api = FakeAPI()
        self.args = argparse.Namespace(repo='example/project', config=setup.CONFIG,
                                       apply=False, check=False, update_existing=False, report=None)

    def run_setup(self):
        with redirect_stdout(io.StringIO()):
            return setup.run(self.args, self.api)

    def test_preview_never_writes(self):
        self.assertEqual(self.run_setup(), 0)
        self.assertEqual(self.api.writes, [])

    def test_check_returns_drift_without_writes(self):
        self.args.check = True
        self.assertEqual(self.run_setup(), 1)
        self.assertEqual(self.api.writes, [])

    def test_apply_verify_and_second_run_is_noop(self):
        self.args.apply = True
        self.assertEqual(self.run_setup(), 0)
        # Five rulesets, vulnerability alerts, security updates, general settings.
        self.assertEqual([w[0] for w in self.api.writes], ['POST'] * 5 + ['PUT', 'PUT', 'PATCH'])
        self.assertEqual([w[1].rsplit('/', 1)[1] for w in self.api.writes[5:7]],
                         ['vulnerability-alerts', 'automated-security-fixes'])
        self.api.writes.clear()
        self.assertEqual(self.run_setup(), 0)
        self.assertEqual(self.api.writes, [])

    def test_failed_check_blocks_all_writes(self):
        self.args.apply = True
        self.api.green = False
        with self.assertRaisesRegex(setup.SetupError, 'has not succeeded'):
            self.run_setup()
        self.assertEqual(self.api.writes, [])

    def test_disabled_merge_methods_block_all_writes(self):
        self.args.apply = True
        for setting, message in (('allow_merge_commit', 'Merge commits'),
                                 ('allow_squash_merge', 'Squash merging')):
            with self.subTest(setting=setting):
                self.api = FakeAPI()
                self.api.repo[setting] = False
                with self.assertRaisesRegex(setup.SetupError, message):
                    self.run_setup()
                self.assertEqual(self.api.writes, [])

    def test_missing_reraises_errors_other_than_404(self):
        def forbidden(method, endpoint, payload=None):
            raise setup.SetupError(f'{method} {endpoint} failed (HTTP 403). Check gh authentication.')
        with self.assertRaisesRegex(setup.SetupError, 'HTTP 403'):
            setup.missing(forbidden, 'repos/a/b/vulnerability-alerts')
        self.assertEqual(setup.missing(lambda method, endpoint: self.api.not_found(endpoint),
                                       'repos/a/b/x'), (None, True))

    def test_non_admin_blocks_all_writes(self):
        self.args.apply = True
        self.api.repo['permissions']['admin'] = False
        with self.assertRaisesRegex(setup.SetupError, 'Administrative'):
            self.run_setup()
        self.assertEqual(self.api.writes, [])

    def test_existing_rule_changes_require_explicit_flag(self):
        rule = copy.deepcopy(setup.load_config(setup.CONFIG)['rulesets'][0])
        rule.update(id=1, source_type='Repository', enforcement='disabled')
        self.api.rules = [rule]
        self.args.apply = True
        with self.assertRaisesRegex(setup.SetupError, 'named ruleset differs'):
            self.run_setup()
        self.assertEqual(self.api.writes, [])
        self.args.update_existing = True
        self.assertEqual(self.run_setup(), 0)

    def test_unrelated_rules_are_preserved(self):
        self.api.rules = [{'name': 'Other rule', 'id': 10, 'source_type': 'Repository'}]
        self.args.apply = True
        self.run_setup()
        self.assertEqual(self.api.rules[0]['name'], 'Other rule')

    def test_preview_names_unmanaged_rulesets_without_writing(self):
        self.api.rules = [{'name': 'main', 'id': 10, 'source_type': 'Repository', 'target': 'branch',
                           'enforcement': 'active',
                           'conditions': {'ref_name': {'include': ['~DEFAULT_BRANCH'], 'exclude': []}}}]
        output = io.StringIO()
        with redirect_stdout(output):
            self.assertEqual(setup.run(self.args, self.api), 0)
        self.assertIn('unmanaged ruleset "main" (branch: ~DEFAULT_BRANCH; active)', output.getvalue())
        self.assertEqual(self.api.writes, [])

    def test_public_repository_enables_vulnerability_reporting_once(self):
        self.api.repo['visibility'] = 'public'
        self.args.apply = True
        self.assertEqual(self.run_setup(), 0)
        self.assertIn(('PUT', 'repos/example/project/private-vulnerability-reporting', None),
                      self.api.writes)
        self.assertEqual(self.api.writes[-1][0], 'PATCH')
        self.api.writes.clear()
        self.assertEqual(self.run_setup(), 0)
        self.assertEqual(self.api.writes, [])

    def test_private_repository_skips_public_only_features_with_a_note(self):
        self.args.apply = True
        output = io.StringIO()
        with redirect_stdout(output):
            self.assertEqual(setup.run(self.args, self.api), 0)
        self.assertIn('private_vulnerability_reporting, secret_scanning, '
                      'secret_scanning_push_protection apply only to public repositories',
                      output.getvalue())
        self.assertFalse(any('private-vulnerability' in endpoint or 'security_and_analysis' in str(payload)
                             for _, endpoint, payload in self.api.writes))

    def test_public_repository_enables_secret_scanning_and_push_protection(self):
        self.api.repo['visibility'] = 'public'
        self.args.apply = True
        self.assertEqual(self.run_setup(), 0)
        self.assertIn(('PATCH', 'repos/example/project',
                       {'security_and_analysis': {'secret_scanning': {'status': 'enabled'},
                                                  'secret_scanning_push_protection': {'status': 'enabled'}}}),
                      self.api.writes)
        self.assertEqual(self.api.writes[-1][2], {'delete_branch_on_merge': True, 'allow_update_branch': True,
                                                  'has_projects': False, 'has_downloads': False})

    def test_dependent_security_features_are_rejected(self):
        config = setup.load_config(setup.CONFIG)
        for disabled, needed_by in (('secret_scanning', 'Push protection'),
                                    ('vulnerability_alerts', 'security updates')):
            broken = copy.deepcopy(config)
            broken['security'][disabled] = False
            with patch.object(Path, 'read_text', return_value=json.dumps(broken)):
                with self.assertRaisesRegex(setup.SetupError, needed_by):
                    setup.load_config(Path('settings.json'))

    def test_check_reports_disabled_vulnerability_reporting_as_drift(self):
        self.api.repo['visibility'] = 'public'
        self.args.apply = True
        self.run_setup()
        self.api.reporting = False
        self.args.apply, self.args.check = False, True
        self.assertEqual(self.run_setup(), 1)

    def test_report_records_drift_then_verification(self):
        with tempfile.TemporaryDirectory() as directory:
            self.args.report = Path(directory) / 'evidence' / 'settings.md'
            self.args.check = True
            self.assertEqual(self.run_setup(), 1)
            drift = self.args.report.read_text(encoding='utf-8')
            self.assertIn('- Result: **DRIFT**', drift)
            self.assertIn('| Protect main | — | missing |', drift)
            self.assertIn('| `has_projects` | off | on |', drift)
            self.assertIn('| `secret_scanning` | on | not applied (public repositories only) |', drift)
            self.args.check, self.args.apply = False, True
            self.assertEqual(self.run_setup(), 0)
            verified = self.args.report.read_text(encoding='utf-8')
            self.assertIn('- Result: **VERIFIED**', verified)
            self.assertIn('| Protect main | 5 | matches |', verified)
            self.assertIn('| `vulnerability_alerts` | on | on |', verified)
        self.assertEqual([w for w in self.api.writes if 'evidence' in w[1]], [])

    def configure(self):
        self.args.apply = True
        self.assertEqual(self.run_setup(), 0)
        self.api.writes.clear()
        self.args.apply = False

    def test_paused_security_updates_are_drift_not_enabled(self):
        self.configure()
        self.api.paused = True
        with tempfile.TemporaryDirectory() as directory:
            self.args.report = Path(directory) / 'settings.md'
            self.args.check = True
            output = io.StringIO()
            with redirect_stdout(output):
                self.assertEqual(setup.run(self.args, self.api), 1)
            record = self.args.report.read_text(encoding='utf-8')
        self.assertIn('enabled but paused', output.getvalue())
        self.assertIn('PUT repos/example/project/automated-security-fixes', output.getvalue())
        self.assertIn('- Result: **DRIFT**', record)
        self.assertIn('| `automated_security_fixes` | on | on (paused) |', record)
        self.assertEqual(self.api.writes, [])

    def test_apply_that_leaves_updates_paused_fails_the_readback(self):
        self.configure()
        self.api.paused = True
        self.args.apply = True
        with self.assertRaisesRegex(setup.SetupError, 'Read-back'):
            self.run_setup()
        self.assertEqual([w[:2] for w in self.api.writes],
                         [('PUT', 'repos/example/project/automated-security-fixes')])

    def test_unpaused_security_updates_verify(self):
        self.configure()
        self.args.check = True
        self.assertEqual(self.run_setup(), 0)

    def test_paused_updates_still_count_as_enabled_when_configured_off(self):
        self.configure()
        self.api.paused = True
        config = setup.load_config(setup.CONFIG)
        config['security']['automated_security_fixes'] = False
        metadata = self.api('GET', 'repos/example/project')
        with redirect_stdout(io.StringIO()):
            plan = setup.security_plan(self.api, 'repos/example/project', metadata, config, True)
        self.assertEqual(plan, [('DELETE', 'repos/example/project/automated-security-fixes', None)])

    def test_required_status_check_is_the_integrity_job(self):
        from test_workflow_controls import REQUIRED_CHECK
        config = setup.load_config(setup.CONFIG)
        contexts = {check['context'] for ruleset in config['rulesets'] for rule in ruleset['rules']
                    if rule['type'] == 'required_status_checks'
                    for check in rule['parameters']['required_status_checks']}
        self.assertEqual(contexts, {REQUIRED_CHECK})

    def test_duplicate_names_rejected_before_writes(self):
        name = setup.load_config(setup.CONFIG)['rulesets'][0]['name']
        self.api.rules = [{'name': name, 'id': 1}, {'name': name, 'id': 2}]
        with self.assertRaisesRegex(setup.SetupError, 'Ambiguous'):
            self.run_setup()
        self.assertFalse(self.api.writes)

    def test_readback_detects_ignored_changes(self):
        self.args.apply = True
        self.api.ignore_patch = True
        with self.assertRaisesRegex(setup.SetupError, 'Read-back'):
            self.run_setup()

    def test_pagination_reads_beyond_first_page(self):
        calls = []
        def call(method, endpoint):
            calls.append(endpoint)
            return list(range(100)) if endpoint.endswith('&page=1') else [100]
        self.assertEqual(len(setup.pages(call, 'repos/x/y/rulesets?includes_parents=false')), 101)
        self.assertEqual(len(calls), 2)

    def test_order_is_not_drift_but_extra_rule_is(self):
        r = copy.deepcopy(setup.load_config(setup.CONFIG)['rulesets'][-1])
        reordered = copy.deepcopy(r)
        reordered['rules'].reverse()
        self.assertEqual(setup.rule_spec(r), setup.rule_spec(reordered))
        reordered['rules'].append({'type': 'creation'})
        self.assertNotEqual(setup.rule_spec(r), setup.rule_spec(reordered))

    def test_cli_uses_explicit_method_host_and_stdin(self):
        result = argparse.Namespace(returncode=0, stdout='{}', stderr='')
        with patch.object(setup.subprocess, 'run', return_value=result) as run:
            setup.api('PATCH', 'repos/a/b', {'has_projects': False})
            cmd = run.call_args.args[0]
            self.assertIn('github.com', cmd)
            self.assertIn('PATCH', cmd)
            self.assertEqual(cmd[-2:], ['--input', '-'])
            self.assertEqual(run.call_args.kwargs['input'], '{"has_projects": false}')
            # gh writes UTF-8 whatever the locale code page is.
            self.assertEqual(run.call_args.kwargs['encoding'], 'utf-8')

    def test_api_error_does_not_print_server_detail(self):
        result = argparse.Namespace(returncode=1, stdout='', stderr='secret token HTTP 403')
        with patch.object(setup.subprocess, 'run', return_value=result):
            with self.assertRaises(setup.SetupError) as caught:
                setup.api('GET', 'repos/a/b')
        self.assertIn('HTTP 403', str(caught.exception))
        self.assertNotIn('secret token', str(caught.exception))


if __name__ == '__main__':
    unittest.main()
