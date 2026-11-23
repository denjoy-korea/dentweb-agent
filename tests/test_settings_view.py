"""에이전트 설정 창에는 웹에서 저장한 값만 표시된다."""
import ast
from pathlib import Path
from queue import SimpleQueue
import unittest
from unittest.mock import Mock

ROOT = Path(__file__).resolve().parents[1]


class SettingsViewTests(unittest.TestCase):
    def test_loads_server_state_without_writing(self):
        tree = ast.parse((ROOT / "gui.py").read_text(encoding="utf8"))
        window = next(node for node in tree.body if getattr(node, "name", "") == "SettingsWindow")
        load = next(node for node in window.body if getattr(node, "name", "") == "_load_state")
        namespace = {}
        exec(compile(ast.Module(body=[load], type_ignores=[]), "gui.py", "exec"), namespace)
        view, api = Mock(), Mock(spec=["get_state"])
        view._result = SimpleQueue()
        for enabled, label in [(True, "켜짐"), (False, "꺼짐")]:
            api.get_state.return_value = {"state": {"enabled": enabled, "scheduled_time": "19:35"}}
            namespace["_load_state"](view, api)
            self.assertEqual(view._result.get_nowait(), f"자동 실행: {label}\n실행 시간 (KST): 19:35")
        api.get_state.side_effect = ConnectionError()
        namespace["_load_state"](view, api)
        self.assertIn("연결 상태", view._result.get_nowait())
        self.assertFalse(any(isinstance(n, ast.Attribute) and n.attr == "save_settings"
                             for n in ast.walk(window)))
