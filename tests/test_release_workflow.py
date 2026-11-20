"""릴리스 경로: 사전 릴리스로 게시, 해시 고정 의존성, 게시 단계는 코드를 실행하지 않는다."""
import ast
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]


def read(path):
    return (ROOT / path).read_text(encoding="utf-8")


class ReleaseWorkflowTests(unittest.TestCase):
    def test_release_publishes_pre_release_with_notes_file(self):
        source = read(".github/workflows/release.yml")
        build, publish = source.split("\n  publish:")
        self.assertRegex(publish, r"\n\s+prerelease: true\n")
        self.assertIn("body_path: release/RELEASE_NOTES.md", publish)
        self.assertNotRegex(publish, r"\n\s+body: \|")
        self.assertIn("dist/RELEASE_NOTES.md", build)

    def test_publish_runs_no_build_code_and_build_cannot_write(self):
        source = read(".github/workflows/release.yml")
        build, publish = source.split("\n  publish:")
        self.assertNotRegex(build, r"contents:\s*write")
        self.assertIn("persist-credentials: false", build)
        self.assertRegex(publish, r"needs: build")
        self.assertNotRegex(publish, r"pip install|pyinstaller|actions/checkout")
        self.assertIn("sha256sum --check SHA256SUMS", publish)

    def test_both_workflows_install_hash_locked_dependencies(self):
        for path in (".github/workflows/release.yml", ".github/workflows/test.yml"):
            source = read(path)
            self.assertIn("--require-hashes --no-deps --only-binary=:all: -r requirements-bootstrap.txt", source, path)
            self.assertIn("--require-hashes --no-deps --no-build-isolation -r requirements.txt -r requirements-build.txt", source, path)
            self.assertIn("pip check", source, path)
            # PowerShell 여러 줄 단계는 마지막 종료 코드만 보므로 설치 명령마다 확인한다
            install = source.split("Install hash-locked dependencies")[1].split("- name:")[0]
            self.assertEqual(install.count("if ($LASTEXITCODE) { exit $LASTEXITCODE }"), 3, path)

    def test_pull_request_workflow_tests_and_builds_without_publishing(self):
        source = read(".github/workflows/test.yml")
        self.assertRegex(source, r"\n  pull_request:")
        self.assertNotRegex(source, r"contents:\s*write|action-gh-release|upload-artifact")
        self.assertIn('unittest discover -s tests -p "test_*.py"', source)
        self.assertIn("pyinstaller", source)

    def test_every_requirement_is_pinned_with_hashes(self):
        for path in ("requirements.txt", "requirements-build.txt", "requirements-bootstrap.txt"):
            source = read(path).replace("\\\n", "")
            lines = [line for line in source.splitlines() if line.strip() and not line.lstrip().startswith("#")]
            self.assertGreaterEqual(len(lines), 3, path)
            for line in lines:
                self.assertRegex(line, r"^[A-Za-z0-9_.-]+==[A-Za-z0-9_.-]+(?: +--hash=sha256:[0-9a-f]{64})+$", f"{path}: {line}")

    def test_release_notes_match_source_version(self):
        tree = ast.parse(read("gui.py"))
        version = next(ast.literal_eval(n.value) for n in tree.body if isinstance(n, ast.Assign)
                       and any(isinstance(t, ast.Name) and t.id == "VERSION" for t in n.targets))
        notes = read("RELEASE_NOTES.md")
        self.assertTrue(re.search(r"^## DentWeb Agent " + re.escape(version) + r"\b", notes, re.M), version)


if __name__ == "__main__":
    unittest.main()
