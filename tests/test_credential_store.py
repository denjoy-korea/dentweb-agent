import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
import credential_store


class ConfigProtectionTests(unittest.TestCase):
    def test_every_save_strips_plaintext_and_preserves_coordinates(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(config, 'protect_token', return_value='dpapi-v1:opaque'):
            path = str(Path(directory) / 'config.json')
            config.save_config(path, {'agent_token': 'local-test-token', 'coords': {'x': 42}})
            raw = Path(path).read_text()
            self.assertNotIn('local-test-token', raw)
            self.assertNotIn('"agent_token"', raw)
            self.assertEqual(json.loads(raw)['coords'], {'x': 42})

    def test_legacy_config_migrates_before_returning_in_memory_token(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(config, 'protect_token', return_value='dpapi-v1:opaque'):
            path = Path(directory) / 'config.json'
            path.write_text(json.dumps({'agent_token': 'migration-test'}))
            loaded = config.load_config(str(path))
            self.assertEqual(loaded['agent_token'], 'migration-test')
            self.assertNotIn('migration-test', path.read_text())

    def test_encryption_failure_preserves_existing_file(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(config, 'protect_token', side_effect=RuntimeError('denied')):
            path = Path(directory) / 'config.json'
            path.write_text('original')
            with self.assertRaises(RuntimeError):
                config.save_config(str(path), {'agent_token': 'unwritten'})
            self.assertEqual(path.read_text(), 'original')
            self.assertEqual(len(list(path.parent.iterdir())), 1)

    def test_non_windows_never_uses_plaintext_fallback(self):
        if os.name != 'nt':
            with self.assertRaises(RuntimeError):
                credential_store.protect_token('local-test')

    @unittest.skipUnless(os.name == 'nt', 'requires Windows DPAPI')
    def test_native_dpapi_round_trip_and_tamper_rejection(self):
        original = 'dpapi-native-test-not-a-real-credential'
        protected = credential_store.protect_token(original)
        self.assertNotIn(original, protected)
        self.assertEqual(credential_store.unprotect_token(protected), original)
        with self.assertRaises(Exception):
            credential_store.unprotect_token(protected[:-8] + 'AAAAAAAA')


if __name__ == '__main__':
    unittest.main()

class ConfigCompatibilityTests(unittest.TestCase):
    def test_existing_exports_and_coordinates_survive_legacy_endpoint_migration(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(config, 'protect_token', return_value='dpapi-v1:opaque'):
            path = Path(directory) / 'config.json'
            path.write_text(json.dumps({'agent_token': 'migration-test', 'server_url': config.LEGACY_SERVER_URL,
                                        'coords': {'x': 42}, 'poll_interval_seconds': 45}))
            loaded = config.load_config(str(path))
            self.assertEqual(loaded['server_url'], config.SERVER_URL)
            self.assertEqual(loaded['download_dir'], config._default_exports_dir())
            self.assertEqual(loaded['coords'], {'x': 42})
            self.assertEqual(loaded['poll_interval_seconds'], 45)
            self.assertNotIn('migration-test', path.read_text())
            self.assertEqual(json.loads(path.read_text())['server_url'], config.SERVER_URL)

    def test_failed_atomic_replace_preserves_previous_settings_and_removes_tempfile(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(config, 'protect_token', return_value='dpapi-v1:opaque'), patch.object(config.os, 'replace', side_effect=OSError('denied')):
            path = Path(directory) / 'config.json'
            path.write_text('original')
            with self.assertRaises(OSError):
                config.save_config(str(path), {'agent_token': 'unwritten'})
            self.assertEqual(path.read_text(), 'original')
            self.assertEqual(len(list(path.parent.iterdir())), 1)
