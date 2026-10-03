import ast
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config


class GuiCredentialTests(unittest.TestCase):
    def method(self, name):
        # Exercise actual GUI credential methods without launching a GUI or automation.
        tree = ast.parse((Path(__file__).resolve().parents[1] / 'gui.py').read_text(encoding='utf-8'))
        cls = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == 'AgentApp')
        method = next(n for n in cls.body if isinstance(n, ast.FunctionDef) and n.name == name)
        namespace = {'os': os, 'sys': sys, 'DEFAULTS': config.DEFAULTS, 'load_config': config.load_config,
                     'save_config': config.save_config, 'register': MagicMock()}
        exec(compile(ast.Module(body=[method], type_ignores=[]), 'gui.py', 'exec'), namespace)
        return namespace[name]

    def test_gui_migrates_before_returning_config(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(config, 'protect_token', return_value='dpapi-v1:opaque'):
            old = os.getcwd()
            try:
                os.chdir(directory)
                Path('config.json').write_text('{"agent_token":"gui-test-token"}')
                loaded = self.method('_load_config_silent')(MagicMock())
                self.assertEqual(loaded['agent_token'], 'gui-test-token')
                self.assertNotIn('gui-test-token', Path('config.json').read_text(encoding='utf-8'))
            finally:
                os.chdir(old)

    def test_gui_does_not_start_when_token_cannot_be_decrypted(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(config, 'unprotect_token', side_effect=ValueError('wrong user')):
            old = os.getcwd()
            try:
                os.chdir(directory)
                Path('config.json').write_text('{"agent_token_protected":"dpapi-v1:bad"}')
                app = MagicMock()
                self.assertIsNone(self.method('_load_config_silent')(app))
                self.assertTrue(app._config_load_failed)
            finally:
                os.chdir(old)

    def test_save_failure_never_enters_automation_onboarding(self):
        app = MagicMock()
        app._save_config.side_effect = RuntimeError('cannot encrypt')
        self.method('_setup_success')(app, 'test-token')
        app._show_onboarding_step2.assert_not_called()
        app._start_btn.configure.assert_called_with(state='normal', text='다음 →')

    @unittest.skipUnless(os.name == 'nt', 'requires Windows desktop')
    def test_native_setup_screen_masks_token(self):
        import gui
        with tempfile.TemporaryDirectory() as directory:
            old = os.getcwd()
            app = None
            try:
                os.chdir(directory)
                app = gui.AgentApp()
                app.withdraw()
                app._token_entry.insert(0, 'synthetic-mask-test')
                self.assertEqual(app._token_entry.cget('show'), '•')
                self.assertIsNone(app.runner)
            finally:
                if app is not None:
                    app.destroy()
                os.chdir(old)


if __name__ == '__main__':
    unittest.main()
