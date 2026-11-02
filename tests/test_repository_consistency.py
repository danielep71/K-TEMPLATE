"""Cross-file contracts: label catalogue, Dependabot, Markdown, API versions and template-only names."""

import json
import re
import subprocess
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def tracked(*patterns):
    """Files Git tracks, so ignored directories such as .venv/ are never scanned."""
    try:
        names = subprocess.run(['git', '-C', str(ROOT), 'ls-files', '-z', '--', *patterns],
                               capture_output=True, text=True, encoding='utf-8',
                               check=True).stdout.split('\0')
    except (OSError, subprocess.CalledProcessError):
        names = [path.relative_to(ROOT).as_posix() for pattern in patterns
                 for path in ROOT.glob(f'**/{pattern}') if '.git' not in path.parts]
    return sorted(ROOT / name for name in names if name and (ROOT / name).is_file())


class CatalogueDocumentationTests(unittest.TestCase):
    def test_guide_matches_label_names_colors_and_descriptions(self):
        manifest = json.loads((ROOT / '.github/labels.json').read_text(encoding='utf-8'))
        guide = (ROOT / 'docs/repo_governance/ISSUE_LABELS.md').read_text(encoding='utf-8')
        rows = re.findall(r'^\| `([^`]+)` \| `#([0-9A-F]+)` \| (.+) \|$', guide, re.M)
        expected = [(label['name'], label['color'], label['description'])
                    for label in manifest['labels']]
        self.assertEqual(rows, expected)

    def test_usage_guide_covers_every_label_once(self):
        manifest = json.loads((ROOT / '.github/labels.json').read_text(encoding='utf-8'))
        guide = (ROOT / 'docs/repo_governance/ISSUE_LABELS.md').read_text(encoding='utf-8')
        section = re.search(r'^## 🧭 Using the labels$(.*?)^## ', guide, re.M | re.S).group(1)
        names = re.findall(r'^\| `([^`]+)` \|', section, re.M)
        self.assertEqual(sorted(names, key=str.lower),
                         sorted((label['name'] for label in manifest['labels']), key=str.lower))

    def test_documented_label_counts_match_the_catalogue(self):
        manifest = json.loads((ROOT / '.github/labels.json').read_text(encoding='utf-8'))
        expected = str(len(manifest['labels']))
        pattern = re.compile(r'labels-(\d+)-|\b(\d+) (?:issue )?labels\b', re.I)
        counts = []
        for path in tracked('*.md'):
            text = path.read_text(encoding='utf-8')
            counts += [(path.relative_to(ROOT).as_posix(), badge or prose)
                       for badge, prose in pattern.findall(text)]
        self.assertTrue(counts, 'no documented label counts found')
        self.assertEqual([item for item in counts if item[1] != expected], [])


def yaml_list(text, key):
    """Return the block-list items under ``key:`` in a simple YAML file."""
    items, indent = [], None
    for line in text.splitlines():
        stripped = line.strip()
        if indent is None:
            if stripped == f'{key}:':
                indent = len(line) - len(line.lstrip())
            continue
        if stripped.startswith('- ') and len(line) - len(line.lstrip()) > indent:
            items.append(stripped[2:].strip().strip('"\''))
        elif stripped:
            indent = None
    return items


class DependabotTests(unittest.TestCase):
    def test_labels_exist_and_assignee_matches_the_issue_forms(self):
        text = (ROOT / '.github/dependabot.yml').read_text(encoding='utf-8')
        manifest = json.loads((ROOT / '.github/labels.json').read_text(encoding='utf-8'))
        catalogue = {label['name'] for label in manifest['labels']}
        labels = yaml_list(text, 'labels')
        self.assertTrue(labels)
        self.assertEqual([label for label in labels if label not in catalogue], [])
        form = (ROOT / '.github/ISSUE_TEMPLATE/bug.yml').read_text(encoding='utf-8')
        assignee = re.search(r'^assignees:\s*\[\s*"?([^",\]]+)"?\s*\]', form, re.M).group(1)
        self.assertEqual(set(yaml_list(text, 'assignees')), {assignee})


class MarkdownTests(unittest.TestCase):
    # tools/check_links.py checks link targets and anchors (tests/test_check_links.py).
    def test_code_fences_close_on_a_line_of_their_own(self):
        # A closing fence followed by text does not close the block on GitHub.
        problems = []
        for path in tracked('*.md'):
            relative = path.relative_to(ROOT)
            opened = None
            for number, line in enumerate(path.read_text(encoding='utf-8').splitlines(), 1):
                if not line.lstrip().startswith('```'):
                    continue
                if opened is None:
                    opened = number
                elif line.strip() == '```':
                    opened = None
                else:
                    problems.append((relative.as_posix(), number))
            if opened is not None:
                problems.append((relative.as_posix(), opened))
        self.assertEqual(problems, [])


class ApiVersionTests(unittest.TestCase):
    def test_every_github_api_call_uses_the_same_version(self):
        found = {}
        for path in [*ROOT.glob('tools/*.py'), *ROOT.glob('.github/workflows/*.yml'),
                     *ROOT.glob('.github/scripts/*.mjs')]:
            text = path.read_text(encoding='utf-8')
            for version in re.findall(r'(?:X-GitHub-Api-Version|API_VERSION)\W+(\d{4}-\d{2}-\d{2})', text):
                found.setdefault(version, []).append(path.relative_to(ROOT).as_posix())
        self.assertEqual(len(found), 1, found)


NOTICE_FILE = re.compile(r'^([.\w][\w./-]*\.[\w]+|\.[\w]+)$', re.M)
LICENSES = {'Mozilla Public License 2.0': 'MPL-2.0', 'MIT License': 'MIT'}


def notice_entries():
    """(repository URL, license, file) for every file of every notice section."""
    notice = (ROOT / 'docs/licenses/THIRD_PARTY_NOTICES.txt').read_text(encoding='utf-8')
    entries = set()
    for section in re.split(r'\n(?=The (?:following|target))', notice)[1:]:
        url = re.search(r'\((https://github\.com/[\w.-]+/[\w.-]+)\)', section).group(1)
        license_id = next(code for name, code in LICENSES.items()
                          if name in ' '.join(section.split()))
        entries |= {(url, license_id, path) for path in NOTICE_FILE.findall(section)}
    return entries


class LicenseNoticeTests(unittest.TestCase):
    def test_every_file_listed_in_the_notice_exists(self):
        notice = (ROOT / 'docs/licenses/THIRD_PARTY_NOTICES.txt').read_text(encoding='utf-8')
        listed = NOTICE_FILE.findall(notice)
        self.assertTrue(listed, 'the notice lists no files')
        self.assertEqual([path for path in listed if not (ROOT / path).is_file()], [])

    def test_provenance_register_agrees_with_the_notice(self):
        register = json.loads((ROOT / 'docs/licenses/provenance.json').read_text(encoding='utf-8'))
        self.assertEqual(register['schema_version'], 1)
        entries = set()
        for source in register['sources']:
            with self.subTest(repository=source['repository']):
                self.assertEqual(set(source), {'repository', 'license', 'revision',
                                               'revision_note', 'files'})
                self.assertRegex(source['repository'], r'^https://github\.com/[\w.-]+/[\w.-]+$')
                revision = source['revision']
                self.assertTrue(revision is None or re.fullmatch(r'[0-9a-f]{40}', revision),
                                f'revision {revision!r} is not a full commit SHA or null')
                self.assertTrue(str(source['revision_note']).strip(),
                                'explain what the revision is, or why it is not recorded')
                self.assertEqual(len(source['files']), len(set(source['files'])))
                entries |= {(source['repository'], source['license'], path)
                            for path in source['files']}
        self.assertEqual(entries, notice_entries())


class TemplateOnlyContentTests(unittest.TestCase):
    def test_reference_repositories_are_named_only_where_allowed(self):
        agents = (ROOT / 'AGENTS.md').read_text(encoding='utf-8')
        block = re.search(r'<!-- template-only:start -->(.*?)<!-- template-only:end -->',
                          agents, re.S)
        if block is None:
            self.skipTest('no template-only section (generated repository)')
        names = set(re.findall(r'`[\w.-]+/([\w.-]+)`', block.group(1)))
        self.assertTrue(names, 'the template-only section names no reference repository')
        outside = agents.replace(block.group(0), '')
        found = [('AGENTS.md', name) for name in sorted(names)
                 if re.search(re.escape(name), outside, re.I)]
        for path in tracked():
            relative = path.relative_to(ROOT)
            if relative.parts[:2] == ('docs', 'licenses') or relative.as_posix() == 'AGENTS.md':
                continue
            text = path.read_text(encoding='utf-8', errors='ignore')
            found += [(relative.as_posix(), name) for name in sorted(names)
                      if re.search(re.escape(name), text, re.I)]
        self.assertEqual(found, [])
