"""Tests for the Markdown link checker (tools/check_links.py)."""

from __future__ import annotations

import contextlib
import io
import json
import subprocess
import sys
import tempfile
import unittest
import unittest.mock
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_links  # noqa: E402


def run_check(files: dict[str, str]) -> tuple[list[str], set[str], int]:
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        for name, text in files.items():
            path = root / name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
        # Not a Git checkout: the checker falls back to every *.md file.
        return check_links.check(root)


class SlugTests(unittest.TestCase):
    def test_github_slugs(self) -> None:
        self.assertEqual(check_links.github_slug("Branches and publication"), "branches-and-publication")
        self.assertEqual(check_links.github_slug("🧭 Before you start"), "-before-you-start")
        # U+FE0F after an emoji is a combining mark that GitHub keeps; "①" is not a word character.
        self.assertEqual(check_links.github_slug("▶\ufe0f Launch sequence"), "\ufe0f-launch-sequence")
        self.assertEqual(check_links.github_slug("① Preview the changes"), "-preview-the-changes")
        self.assertEqual(check_links.github_slug("Café 2026_v1"), "café-2026_v1")
        self.assertEqual(check_links.github_slug("Run `check.py` [now](x.md)"), "run-checkpy-now")
        self.assertEqual(check_links.github_slug("Rock &amp; Roll"), "rock--roll")

    def test_duplicate_and_setext_headings(self) -> None:
        _, anchors = check_links.markdown_links("# Notes\n\n## Notes\n\nTitle\n=====\n")
        self.assertEqual(anchors, {"notes", "notes-1", "title"})
        _, anchors = check_links.markdown_links("# Foo\n# Foo-1\n# Foo\n")
        self.assertEqual(anchors, {"foo", "foo-1", "foo-2"})

    def test_code_is_ignored_and_explicit_anchors_count(self) -> None:
        text = ('<a id="custom"></a>\n\n```sh\n[x](missing.md)\n# not a heading\n```\n'
                "`[y](also-missing.md)` [z](real.md#part)\n")
        links, anchors = check_links.markdown_links(text)
        self.assertEqual(links, ["real.md#part"])
        self.assertEqual(anchors, {"custom"})


class CheckTests(unittest.TestCase):
    def test_valid_links_pass_and_external_links_are_collected(self) -> None:
        errors, external, count = run_check({
            "README.md": "[Guide](docs/guide.md#-install) · [site](https://example.com/page#top)\n"
                         "[self](#readme) <https://example.org>\n\n# README\n",
            "docs/guide.md": "# 🚀 Install\n\n[Back](../README.md)\n",
        })
        self.assertEqual(errors, [])
        self.assertEqual(external, {"https://example.com/page", "https://example.org"})
        self.assertEqual(count, 5)

    def test_missing_files_anchors_and_paths_are_errors(self) -> None:
        errors, _, _ = run_check({
            "README.md": "[a](docs/missing.md) [b](docs/guide.md#nowhere) [c](/abs.md)\n"
                         "[d](../outside.md) [e](empty.md) [f](ftp://example.com)\n",
            "docs/guide.md": "# Guide\n",
            "empty.md": "",
        })
        self.assertEqual(errors, [
            "README.md: missing target: docs/missing.md",
            "README.md: missing anchor: docs/guide.md#nowhere",
            "README.md: root-relative link: /abs.md",
            "README.md: path outside the repository: ../outside.md",
            "README.md: empty target: empty.md",
            "README.md: unsupported URL scheme: ftp://example.com",
        ])

    def test_plain_http_links_are_errors_outside_code(self) -> None:
        errors, external, _ = run_check({
            "README.md": "[a](http://a.example/x) <http://b.example> [c](HTTP://c.example)\n"
                         "[d]: http://d.example/\n<a href=\"http://e.example\">e</a>\n"
                         "`[f](http://f.example)` and `<http://g.example>` are code\n\n"
                         "```text\n[h](http://h.example)\n```\n[ok](https://ok.example)\n",
        })
        self.assertEqual(errors, [
            "README.md: plain HTTP link, use https: http://a.example/x",
            "README.md: plain HTTP link, use https: HTTP://c.example",
            "README.md: plain HTTP link, use https: http://b.example",
            "README.md: plain HTTP link, use https: http://d.example/",
            "README.md: plain HTTP link, use https: http://e.example",
        ])
        self.assertEqual(external, {"https://ok.example"})

    def test_tracked_file_missing_from_the_working_tree_is_reported(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "README.md").write_text("[guide](guide.md)\n", encoding="utf-8")
            (root / "guide.md").write_text("# Guide\n", encoding="utf-8")
            subprocess.run(["git", "init", "-q", str(root)], check=True)
            subprocess.run(["git", "-C", str(root), "add", "."], check=True)
            (root / "guide.md").unlink()
            errors, _, _ = check_links.check(root)
        self.assertEqual(errors, ["guide.md: tracked but missing from the working tree",
                                  "README.md: missing target: guide.md"])

    def test_repository_links_pass(self) -> None:
        errors, _, count = check_links.check(ROOT)
        self.assertGreater(count, 0)
        self.assertEqual(errors, [])


class ExternalTests(unittest.TestCase):
    def test_own_repository_urls_are_skipped(self) -> None:
        own = check_links.own_repository
        self.assertTrue(own("https://github.com/Owner/Repo/blob/HEAD/x.md", "owner/repo"))
        self.assertTrue(own("https://github.com/owner/repo", "owner/repo"))
        self.assertTrue(own("https://raw.githubusercontent.com/owner/repo/main/x", "owner/repo"))
        self.assertFalse(own("https://github.com/owner/repo-other", "owner/repo"))
        self.assertFalse(own("https://github.com/owner/repo", ""))

    def test_only_errors_fail_and_own_urls_are_not_requested(self) -> None:
        results = {"https://a.example/": ("PASS", "https://a.example/"),
                   "https://b.example/": ("WARN", "https://b.example/: unavailable"),
                   "https://c.example/": ("ERROR", "https://c.example/: HTTP 404")}
        with unittest.mock.patch.object(check_links, "check_external",
                                        side_effect=lambda url: results[url]) as fake, \
                contextlib.redirect_stdout(io.StringIO()):
            errors = check_links.external_errors(
                [*results, "https://github.com/owner/repo/issues"], "owner/repo")
        self.assertEqual(errors, ["https://c.example/: HTTP 404"])
        self.assertEqual(sorted(call.args[0] for call in fake.call_args_list), sorted(results))

    def test_a_local_run_skips_the_origin_repository(self) -> None:
        # Without --repository or GITHUB_REPOSITORY, the origin remote names it.
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "README.md").write_text(
                "[own](https://github.com/Owner/Repo/blob/HEAD/x.md) [other](https://a.example/)\n",
                encoding="utf-8")
            subprocess.run(["git", "init", "-q", str(root)], check=True)
            subprocess.run(["git", "-C", str(root), "remote", "add", "origin",
                            "git@github.com:owner/repo.git"], check=True)
            subprocess.run(["git", "-C", str(root), "add", "README.md"], check=True)
            self.assertEqual(check_links.detect_repository(root), "owner/repo")
            with unittest.mock.patch.dict("os.environ", {"GITHUB_REPOSITORY": ""}), \
                    unittest.mock.patch.object(check_links, "check_external",
                                               return_value=("PASS", "ok")) as fake, \
                    contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(check_links.main(["--root", str(root), "--external"]), 0)
        self.assertEqual([call.args[0] for call in fake.call_args_list], ["https://a.example/"])
        with tempfile.TemporaryDirectory() as directory:
            self.assertEqual(check_links.detect_repository(Path(directory)), "")

    def test_http_codes_are_classified(self) -> None:
        def curl(code: str, returncode: int = 0) -> unittest.mock.Mock:
            return unittest.mock.Mock(stdout=code, returncode=returncode)

        for code, returncode, status in (("200", 0, "PASS"), ("404", 0, "ERROR"),
                                         ("429", 0, "WARN"), ("503", 0, "WARN"),
                                         ("", 6, "WARN"), ("301", 1, "ERROR")):
            with self.subTest(code=code), unittest.mock.patch(
                    "check_links.subprocess.run", return_value=curl(code, returncode)):
                self.assertEqual(check_links.check_external("https://x.example/")[0], status)

    def test_requests_use_https_only(self) -> None:
        with unittest.mock.patch("check_links.subprocess.run",
                                 return_value=unittest.mock.Mock(stdout="200", returncode=0)) as run:
            check_links.check_external("https://x.example/")
        command = run.call_args.args[0]
        self.assertEqual(command[command.index("--proto") + 1], "=https")
        self.assertEqual(command[command.index("--proto-redir") + 1], "=https")


def write_exceptions(root: Path, data: object) -> None:
    path = root / check_links.EXCEPTIONS
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(data if isinstance(data, str) else json.dumps(data), encoding="utf-8")


def exceptions_file(*entries: object) -> dict[str, object]:
    return {"schema_version": 1, "exceptions": list(entries)}


ENTRY: dict[str, object] = {"url": "https://slow.example/page",
                            "reason": "Answers 403 to automated requests",
                            "expires": "2027-01-31"}


class ExceptionTests(unittest.TestCase):
    def test_absent_file_means_no_exceptions(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            self.assertEqual(check_links.load_exceptions(Path(directory)), ({}, []))

    def test_valid_file_is_loaded_and_its_urls_must_be_linked(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            write_exceptions(root, exceptions_file(ENTRY, {**ENTRY, "url": "https://gone.example/"}))
            (root / "README.md").write_text("[slow](https://slow.example/page#part)\n",
                                            encoding="utf-8")
            exceptions, problems = check_links.load_exceptions(root)
            errors, _, _ = check_links.check(root)
        self.assertEqual(problems, [])
        self.assertEqual(exceptions["https://slow.example/page"], check_links.LinkException(
            "Answers 403 to automated requests", date(2027, 1, 31)))
        self.assertEqual(errors, [".github/link-exceptions.json: https://gone.example/ is not "
                                  "linked from any Markdown file; remove its exception"])

    def test_malformed_files_fail(self) -> None:
        name = ".github/link-exceptions.json"
        cases: list[tuple[object, str]] = [
            ("{", f"{name}: not valid UTF-8 JSON"),
            ({"exceptions": []}, f"{name}: needs exactly schema_version 1 and an exceptions list"),
            ({"schema_version": True, "exceptions": []}, f"{name}: needs exactly schema_version 1"),
            ({"schema_version": 1, "exceptions": {}}, f"{name}: needs exactly schema_version 1"),
            ({**exceptions_file(), "extra": 1}, f"{name}: needs exactly schema_version 1"),
            (exceptions_file("x"), f"{name}: exception 1: needs exactly url, reason and expires"),
            (exceptions_file({**ENTRY, "note": ""}), f"{name}: exception 1: needs exactly url"),
            (exceptions_file({**ENTRY, "url": "http://slow.example/page"}),
             f"{name}: exception 1: url must be an https URL without a fragment"),
            (exceptions_file({**ENTRY, "url": "https://slow.example/page#x"}),
             f"{name}: exception 1: url must be an https URL without a fragment"),
            (exceptions_file(ENTRY, ENTRY),
             f"{name}: exception 2: https://slow.example/page is listed twice"),
            (exceptions_file({**ENTRY, "reason": " "}),
             f"{name}: exception 1: https://slow.example/page needs a reason"),
            (exceptions_file({**ENTRY, "expires": "31/01/2027"}),
             f"{name}: exception 1: https://slow.example/page needs expires as YYYY-MM-DD"),
            (exceptions_file({**ENTRY, "expires": 20270131}),
             f"{name}: exception 1: https://slow.example/page needs expires as YYYY-MM-DD"),
            (exceptions_file({**ENTRY, "expires": "2027-02-30"}),
             f"{name}: exception 1: https://slow.example/page has an expires date that does not "
             "exist"),
        ]
        for data, expected in cases:
            with self.subTest(data=data), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                write_exceptions(root, data)
                (root / "README.md").write_text("[slow](https://slow.example/page)\n",
                                                encoding="utf-8")
                errors, _, _ = check_links.check(root)
                self.assertEqual(len(errors), 1, errors)
                self.assertTrue(errors[0].startswith(expected), errors[0])

    def test_tolerated_answer_fails_once_its_exception_expires(self) -> None:
        exception = check_links.LinkException("Answers 403", date(2027, 1, 31))
        url = "https://slow.example/page"
        warn = ("WARN", f"{url}: unavailable for verification (HTTP 403)")
        for today, listed, expected in ((date(2027, 1, 30), True, []),
                                        (date(2027, 1, 31), True, []),
                                        (date(2027, 2, 1), True, [
                                            f"{url}: unavailable for verification (HTTP 403); its "
                                            "exception in .github/link-exceptions.json expired "
                                            "on 2027-01-31 (Answers 403)"]),
                                        (date(2099, 1, 1), False, [])):
            output = io.StringIO()
            with self.subTest(today=today, listed=listed), unittest.mock.patch.object(
                    check_links, "check_external", return_value=warn), \
                    contextlib.redirect_stdout(output):
                errors = check_links.external_errors(
                    [url], "", {url: exception} if listed else {}, today)
                self.assertEqual(errors, expected)
                if listed and not expected:
                    self.assertIn("tolerated until 2027-01-31 (Answers 403)", output.getvalue())

    def test_exception_does_not_hide_an_error(self) -> None:
        url = "https://gone.example/"
        with unittest.mock.patch.object(check_links, "check_external",
                                        return_value=("ERROR", f"{url}: HTTP 404")), \
                contextlib.redirect_stdout(io.StringIO()):
            errors = check_links.external_errors(
                [url], "", {url: check_links.LinkException("Gone", date(2099, 1, 1))},
                date(2027, 1, 1))
        self.assertEqual(errors, [f"{url}: HTTP 404"])


if __name__ == "__main__":
    unittest.main()
