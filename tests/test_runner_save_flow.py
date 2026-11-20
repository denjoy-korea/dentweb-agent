"""v3.6.5 저장 창 흐름: 고유 이름 입력, 덮어쓰기 확인 클릭 없음, 이번 파일만 기다리고 정리."""
import importlib
import json
import os
from pathlib import Path
import sys
import tempfile
import time
import types
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

# Mac 등 데스크톱 자동화 라이브러리가 없는 곳에서도 흐름만 시험한다(Windows CI는 실제 모듈을 쓴다).
for _name in ("pyautogui", "pyperclip"):
    try:
        importlib.import_module(_name)
    except Exception:
        sys.modules[_name] = types.ModuleType(_name)

import dentweb_runner  # noqa: E402
from export_store import OWNED  # noqa: E402

SAVE_XY = (700, 500)
CONFIRM_XY = (640, 420)


def _steps_file(directory):
    path = Path(directory) / "dentweb_steps.json"
    path.write_text(json.dumps({"data_steps": [
        {"name": "stats_menu", "label": "통계", "x": 1, "y": 1, "wait_after": 0},
        {"name": "export_btn", "label": "엑셀저장", "x": 2, "y": 2, "wait_after": 0},
        {"name": "save_dialog_save_btn", "label": "저장", "x": SAVE_XY[0], "y": SAVE_XY[1], "wait_after": 0},
        {"name": "save_dialog_confirm_yes", "label": "예(Y)", "x": CONFIRM_XY[0], "y": CONFIRM_XY[1], "wait_after": 0},
    ]}, ensure_ascii=False), encoding="utf-8")
    return str(path)


class RunnerSaveFlowTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.exports = self.root / "private-exports"
        self.legacy = self.root / "exports"
        steps = _steps_file(self.root)
        self.runner = dentweb_runner.DentwebRunner({"download_dir": str(self.exports), "download_timeout_seconds": 5})
        self.runner._data = dentweb_runner.load_config_data(steps)
        self.gui = MagicMock()
        self.pasted = []
        self.clicks = []
        self.writes_file = True
        self.saves_to_target = True
        self.patches = [
            patch.object(dentweb_runner, "pyautogui", self.gui),
            patch.object(dentweb_runner, "_paste_text", side_effect=self.pasted.append),
            patch.object(dentweb_runner, "_win32_click", side_effect=self._click),
            patch.object(dentweb_runner, "_find_and_click", return_value=True),
            patch.object(dentweb_runner.DentwebRunner, "_activate_dentweb", return_value=True),
            patch.object(dentweb_runner, "legacy_exports_directory", return_value=self.legacy),
            patch.object(dentweb_runner.time, "sleep", lambda _s: None),
            # Windows CI에는 실제 저장 창이 없으므로 '맨 앞이 저장 창'으로 둔다(창 확인은 따로 시험)
            patch.object(dentweb_runner, "_foreground_is_save_dialog", return_value=True),
        ]
        for p in self.patches:
            p.start()

    def tearDown(self):
        for p in reversed(self.patches):
            p.stop()
        self.tmp.cleanup()

    def _click(self, x, y, activate_first=False):
        self.clicks.append((x, y))
        # 덴트웹이 '저장(S)'을 누르면 파일 이름 칸에 들어간 경로로 저장한다고 흉내 낸다.
        if (x, y) == SAVE_XY and self.writes_file:
            name = self.pasted[-1] if self.saves_to_target else str(self.exports / "임플란트수술통계_20261121.xlsx")
            Path(name).write_bytes(b"xlsx-fixture")

    def test_load_config_drops_retired_overwrite_confirm_step(self):
        names = [s["name"] for s in self.runner._data["data_steps"]]
        self.assertNotIn("save_dialog_confirm_yes", names)
        self.assertNotIn("save_dialog_confirm_yes", [s["name"] for s in dentweb_runner.DATA_STEPS])

    def test_types_unique_path_and_never_clicks_overwrite_confirm(self):
        logs = []
        with patch.object(dentweb_runner.DentwebRunner, "excel_has_data", return_value=True):
            path = self.runner.download_excel(log_callback=logs.append)
        self.assertIsNotNone(path)
        self.assertEqual(Path(path).parent, self.exports)
        self.assertRegex(Path(path).name, OWNED)
        self.assertEqual(self.pasted, [str(self.exports), path])
        self.gui.hotkey.assert_any_call("alt", "n")
        self.gui.hotkey.assert_any_call("ctrl", "a")
        self.assertNotIn(CONFIRM_XY, self.clicks)
        # Enter는 주소창 이동 한 번뿐이다(저장 좌표가 있으므로 저장 폴백·확인 폴백 없음).
        self.assertEqual([c.args for c in self.gui.press.call_args_list], [("enter",)])
        self.runner.cleanup(path)
        self.assertEqual(list(self.exports.iterdir()), [])

    def test_timeout_is_a_failure_and_never_picks_another_workbook(self):
        self.writes_file = False
        self.exports.mkdir(parents=True)
        other = self.exports / "earlier.xlsx"
        other.write_bytes(b"x")
        os.utime(other, (time.time() - 600, time.time() - 600))
        self.runner.download_timeout = 0
        logs = []
        with self.assertRaisesRegex(dentweb_runner.ExportFailed, "같은 Windows 계정"):
            self.runner.download_excel(log_callback=logs.append)
        self.assertEqual([p.name for p in self.exports.iterdir()], ["earlier.xlsx"])

    def test_late_save_of_target_is_removed(self):
        # 시간 초과 직전에 대상 파일이 생겨도(늦은 저장) 실패로 끝나면 지운다
        self.runner.download_timeout = 0
        with self.assertRaises(dentweb_runner.ExportFailed):
            self.runner.download_excel()
        self.assertEqual([p for p in self.exports.iterdir() if p.suffix == ".xlsx"], [])

    def test_default_name_save_is_removed_and_reported_as_failure(self):
        self.saves_to_target = False
        self.runner.download_timeout = 0
        with self.assertRaisesRegex(dentweb_runner.ExportFailed, "다른 이름으로 저장"):
            self.runner.download_excel()
        self.assertEqual(list(self.exports.iterdir()), [])

    def test_no_keys_are_sent_unless_the_save_dialog_is_in_front(self):
        with patch.object(dentweb_runner, "_foreground_is_save_dialog", return_value=False), \
             patch.object(dentweb_runner.time, "monotonic", side_effect=[0, 0, 100, 100, 100]):
            with self.assertRaisesRegex(dentweb_runner.ExportFailed, "저장 창"):
                self.runner.download_excel()
        self.gui.hotkey.assert_not_called()
        self.gui.press.assert_not_called()
        self.assertEqual(self.pasted, [])
        self.assertEqual(list(self.exports.iterdir()), [])

    def test_every_run_prunes_leftovers(self):
        self.exports.mkdir(parents=True)
        leftover = self.exports / ("denjoy-export-" + "b" * 32 + ".xlsx")
        default_name = self.exports / "임플란트수술통계_20261120.xlsx"
        for path in (leftover, default_name):
            path.write_bytes(b"patient")
            os.utime(path, (time.time() - 2 * 3600, time.time() - 2 * 3600))
        self.runner._export_store()  # 첫 정리는 이미 지난 상태(오래 켜 둔 에이전트)
        for path in (leftover, default_name):
            path.write_bytes(b"patient")
            os.utime(path, (time.time() - 2 * 3600, time.time() - 2 * 3600))
        with patch.object(dentweb_runner.DentwebRunner, "excel_has_data", return_value=False):
            self.assertIsNone(self.runner.download_excel())
        self.assertFalse(leftover.exists())
        self.assertFalse(default_name.exists())

    def test_empty_workbook_is_removed(self):
        with patch.object(dentweb_runner.DentwebRunner, "excel_has_data", return_value=False):
            self.assertIsNone(self.runner.download_excel())
        self.assertEqual(list(self.exports.iterdir()), [])

    def test_exception_after_save_still_removes_target(self):
        with patch.object(dentweb_runner.DentwebRunner, "excel_has_data", side_effect=RuntimeError("broken")):
            with self.assertRaises(RuntimeError):
                self.runner.download_excel()
        self.assertEqual(list(self.exports.iterdir()), [])

    def test_legacy_exports_folder_keeps_seven_day_sweep_only(self):
        self.legacy.mkdir()
        old = self.legacy / "old.xlsx"
        new = self.legacy / "new.xlsx"
        note = self.legacy / "note.txt"
        for path in (old, new, note):
            path.write_bytes(b"x")
        stale = time.time() - 8 * 86400
        os.utime(old, (stale, stale))
        os.utime(note, (stale, stale))
        self.runner._cleanup_old_exports(keep_days=7)
        self.assertFalse(old.exists())
        self.assertTrue(new.exists())
        self.assertTrue(note.exists())


if __name__ == "__main__":
    unittest.main()
