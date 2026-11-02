"""Prove that the Ruff complexity ceiling, mypy strict mode and actionlint are active.

Ruff and mypy are installed from tools/requirements-quality-ci.txt, and
actionlint as a checksum-verified release. The Repository integrity job sets
REQUIRE_QUALITY_TOOLS=1, so these tests fail there if a tool is missing;
elsewhere they are skipped without the tools.
"""

import os
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REQUIRED = os.environ.get('REQUIRE_QUALITY_TOOLS') == '1'


def run(command, stdin=None):
    result = subprocess.run(command, cwd=ROOT, input=stdin, text=True,
                            capture_output=True, check=False)
    return result.returncode, result.stdout + result.stderr


def tool(name):
    return unittest.skipIf(shutil.which(name) is None and not REQUIRED,
                           f'{name} is not installed: python -m pip install --only-binary=:all: '
                           '--require-hashes -r tools/requirements-quality-ci.txt')


class QualityGateTests(unittest.TestCase):
    @tool('ruff')
    def test_complexity_15_passes_and_16_fails(self):
        def source(branches):
            body = ''.join(f'    if x == {i}:\n        return {i}\n' for i in range(branches - 1))
            return f'def probe(x: int) -> int:\n{body}    return -1\n'

        command = ['ruff', 'check', '--no-cache', '--select', 'C901',
                   '--stdin-filename', 'tools/complexity_probe.py', '-']
        ok_code, ok_log = run(command, source(15))
        bad_code, bad_log = run(command, source(16))
        self.assertEqual(ok_code, 0, ok_log)
        self.assertNotEqual(bad_code, 0, bad_log)
        self.assertIn('is too complex (16 > 15)', bad_log)

    @tool('mypy')
    def test_strict_mode_rejects_an_unannotated_function(self):
        with tempfile.TemporaryDirectory() as directory:
            probe = Path(directory) / 'strict_probe.py'
            probe.write_text('def probe(value: int) -> int:\n    return value\n', encoding='utf-8')
            ok_code, ok_log = run(['mypy', '--no-pretty', '--no-incremental', str(probe)])
            probe.write_text('def probe(value):\n    return value\n', encoding='utf-8')
            bad_code, bad_log = run(['mypy', '--no-pretty', '--no-incremental', str(probe)])
        self.assertEqual(ok_code, 0, ok_log)
        self.assertNotEqual(bad_code, 0, bad_log)
        self.assertIn('[no-untyped-def]', bad_log)

    @tool('actionlint')
    def test_actionlint_rejects_an_unknown_key(self):
        workflow = ('name: Probe\non: push\npermissions: {}\njobs:\n  probe:\n'
                    '    runs-on: ubuntu-24.04\n    steps:\n      - run: echo ok\n')
        command = ['actionlint', '-no-color', '-shellcheck=', '-pyflakes=', '-']
        ok_code, ok_log = run(command, workflow)
        bad_code, bad_log = run(command, workflow + '        bogus-key: true\n')
        self.assertEqual(ok_code, 0, ok_log)
        self.assertNotEqual(bad_code, 0, bad_log)
        self.assertIn('unexpected key "bogus-key"', bad_log)


if __name__ == '__main__':
    unittest.main()
