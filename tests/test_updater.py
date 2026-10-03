import hashlib
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import updater

URL = 'https://github.com/denjoy-korea/dentweb-agent/releases/download/v3.6.3/dentweb-agent.exe'


class UpdateVerificationTests(unittest.TestCase):
    def response(self, content, total=None):
        response = MagicMock()
        response.text = content.decode() if isinstance(content, bytes) else content
        response.headers = {'content-length': str(total)} if total is not None else {}
        response.iter_content.return_value = [content] if isinstance(content, bytes) else []
        response.__enter__.return_value = response
        return response

    def download(self, payload, expected=None, total=None):
        with tempfile.TemporaryDirectory() as directory:
            expected = expected or hashlib.sha256(payload).hexdigest()
            with patch.object(updater.requests, 'get', side_effect=[
                    self.response(expected + '  dentweb-agent.exe\n'),
                    self.response(payload, total)]), patch.object(updater.tempfile, 'gettempdir', return_value=directory):
                path = updater.download_update(URL)
            files = list(Path(directory).iterdir())
            return path is not None, [file.read_bytes() for file in files]

    def test_verified_file_can_be_applied(self):
        success, files = self.download(b'not-an-executable-test-payload', total=30)
        self.assertTrue(success)
        self.assertEqual(files, [b'not-an-executable-test-payload'])

    def test_hash_mismatch_removes_download(self):
        success, files = self.download(b'changed-payload', expected='a' * 64)
        self.assertFalse(success)
        self.assertEqual(files, [])

    def test_even_small_size_mismatch_is_rejected(self):
        success, files = self.download(b'123456789', total=10)
        self.assertFalse(success)
        self.assertEqual(files, [])

    def test_missing_or_ambiguous_checksum_is_rejected(self):
        for text in ('', 'bad  dentweb-agent.exe', ('a' * 64 + '  dentweb-agent.exe\n') * 2):
            with self.subTest(text=text), patch.object(updater.requests, 'get', return_value=self.response(text)) as get:
                self.assertIsNone(updater.download_update(URL))
                get.assert_called_once()

    def test_wrong_repository_or_version_never_gets_requested(self):
        for url in (URL.replace('denjoy-korea', 'another-owner'), URL.replace('github.com', 'github.com.attacker.invalid'),
                    URL.replace('/v3.6.3/', '/latest/'), URL + '?token=secret'):
            with self.subTest(url=url), patch.object(updater.requests, 'get') as get:
                self.assertIsNone(updater.download_update(url))
                get.assert_not_called()

    def test_latest_without_checksum_is_not_offered(self):
        response = self.response('')
        response.json.return_value = {'tag_name': 'v3.6.3', 'assets': [{'name': 'dentweb-agent.exe', 'browser_download_url': URL}]}
        with patch.object(updater.requests, 'get', return_value=response):
            self.assertIsNone(updater.check_update('3.5.7'))
        response.json.return_value['assets'].append({'name': 'SHA256SUMS'})
        with patch.object(updater.requests, 'get', return_value=response):
            self.assertEqual(updater.check_update('3.5.7')['download_url'], URL)

    @unittest.skipUnless(os.name == 'nt', 'requires Windows PowerShell parser')
    def test_update_script_parses_and_escapes_real_windows_paths(self):
        script = updater._build_update_script("C:\\Clinic's $tools\\dentweb-agent.exe", "C:\\Clinic's $tools\\update.exe", 123)
        self.assertNotIn('$pid =', script.lower())
        self.assertIn("Clinic''s $tools", script)
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'update.ps1'
            path.write_text(script, encoding='utf-8')
            # Parse only: never execute the updater or any executable.
            command = "$errors = $null; $tokens = $null; [void][System.Management.Automation.Language.Parser]::ParseFile($env:UPDATE_SCRIPT, [ref]$tokens, [ref]$errors); if ($errors.Count) { $errors | Out-String | Write-Error; exit 1 }"
            subprocess.run(['powershell', '-NoProfile', '-Command', command], env={**os.environ, 'UPDATE_SCRIPT': str(path)}, check=True)


if __name__ == '__main__':
    unittest.main()
