"""v3.6.5: 업로드·보고가 실패하거나 예외가 나도 수술 기록 파일을 지운다(예약 실행·지금 실행 공통)."""
import ast
from pathlib import Path
import unittest
from unittest.mock import Mock

SOURCE = Path(__file__).resolve().parents[1] / "gui.py"


def _method(name):
    tree = ast.parse(SOURCE.read_text(encoding="utf-8"))
    cls = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == "AgentApp")
    return next(n for n in cls.body if isinstance(n, ast.FunctionDef) and n.name == name)


class GuiCleanupTests(unittest.TestCase):
    def upload_export(self):
        namespace = {}
        exec(compile(ast.Module(body=[_method("_upload_export")], type_ignores=[]), str(SOURCE), "exec"), namespace)
        return namespace["_upload_export"]

    def test_export_is_removed_whatever_upload_and_report_do(self):
        upload_export = self.upload_export()
        for case in ["success", "failed", "upload_raises", "report_raises"]:
            obj = Mock()
            obj.api.upload_file.return_value = {"success": case != "failed", "error": "거부"}
            if case == "upload_raises":
                obj.api.upload_file.side_effect = RuntimeError("network")
            if case == "report_raises":
                obj.api.report_run.side_effect = RuntimeError("report")
            if case in ("upload_raises", "report_raises"):
                with self.assertRaises(RuntimeError):
                    upload_export(obj, "owned.xlsx", "https://fixture/upload", True)
            else:
                ok, _ = upload_export(obj, "owned.xlsx", "https://fixture/upload", True)
                self.assertEqual(ok, case == "success")
            obj.runner.cleanup.assert_called_once_with("owned.xlsx")

    def test_run_now_does_not_report_and_cleanup_failure_is_logged(self):
        upload_export = self.upload_export()
        obj = Mock()
        obj.api.upload_file.return_value = {"success": True, "inserted": 2, "skipped": 0}
        obj.runner.cleanup.side_effect = PermissionError("locked")
        ok, msg = upload_export(obj, "owned.xlsx", "https://fixture/upload", False)
        self.assertTrue(ok)
        self.assertIn("2건", msg)
        obj.api.report_run.assert_not_called()
        obj.after.assert_called_once()

    def test_cleanup_uses_the_runner_that_produced_the_file(self):
        upload_export = self.upload_export()
        obj = Mock()
        original = obj.runner
        replacement = Mock()

        def swap(*_args):
            obj.runner = replacement  # 업로드 중에 좌표 설정이 끝나 러너가 바뀜
            return {"success": True}

        obj.api.upload_file.side_effect = swap
        original.cleanup.side_effect = RuntimeError("not an OSError")
        upload_export(obj, "owned.xlsx", "https://fixture/upload", False)
        original.cleanup.assert_called_once_with("owned.xlsx")
        replacement.cleanup.assert_not_called()
        obj.after.assert_called_once()

    def test_both_run_paths_use_the_cleanup_helper(self):
        tree = ast.parse(SOURCE.read_text(encoding="utf-8"))
        calls = [n for n in ast.walk(tree) if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)]
        helper_calls = [c for c in calls if c.func.attr == "_upload_export"]
        direct_uploads = [c for c in calls if c.func.attr == "upload_file"]
        self.assertEqual(len(helper_calls), 2)
        self.assertEqual(len(direct_uploads), 1)  # 도우미 안의 한 곳뿐


if __name__ == "__main__":
    unittest.main()
