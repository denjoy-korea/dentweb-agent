"""v3.6.5 수술 기록 파일 보호: 전용 폴더, 고유 이름, 소유 파일만 정리."""
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import export_store
from export_store import ExportStore, OWNED


class ExportStoreTests(unittest.TestCase):
    def test_allocates_unique_owned_names_inside_the_directory(self):
        with tempfile.TemporaryDirectory() as root:
            store = ExportStore(root)
            first, second = store.allocate(), store.allocate()
            self.assertNotEqual(first, second)
            for path in (first, second):
                self.assertEqual(Path(path).parent, Path(root))
                self.assertRegex(Path(path).name, OWNED)
                self.assertTrue(store.owned(path))

    @unittest.skipIf(os.name == "nt", "POSIX permission bits")
    def test_directory_is_private_on_posix(self):
        with tempfile.TemporaryDirectory() as root:
            target = Path(root) / "nested" / "private-exports"
            ExportStore(target)
            self.assertEqual(target.stat().st_mode & 0o777, 0o700)

    def test_windows_restricts_folder_to_current_user_and_system(self):
        with tempfile.TemporaryDirectory() as root:
            run = MagicMock()
            run.side_effect = [
                MagicMock(stdout='"desk\\\\staff","S-1-5-21-1-2-3-1001"\r\n'),
                MagicMock(stdout=""),
            ]
            ExportStore(Path(root) / "private-exports", run=run, windows=True)
            whoami, icacls = run.call_args_list
            # EXE 폴더의 같은 이름 파일이 아니라 System32의 도구를 부른다
            self.assertTrue(os.path.isabs(whoami.args[0][0]) or whoami.args[0][0].startswith("C:"))
            self.assertTrue(whoami.args[0][0].lower().endswith(os.path.join("system32", "whoami.exe").lower()))
            self.assertIs(whoami.kwargs["stdin"], subprocess.DEVNULL)
            command = icacls.args[0]
            self.assertTrue(command[0].lower().endswith(os.path.join("system32", "icacls.exe").lower()))
            self.assertIn("/inheritance:r", command)
            self.assertIn("*S-1-5-21-1-2-3-1001:(OI)(CI)F", command)
            self.assertIn("*S-1-5-18:(OI)(CI)F", command)
            self.assertTrue(icacls.kwargs["check"])

    def test_windows_rejects_unexpected_owner(self):
        with tempfile.TemporaryDirectory() as root:
            run = MagicMock(return_value=MagicMock(stdout='"desk\\\\staff","not-a-sid"\r\n'))
            with self.assertRaisesRegex(RuntimeError, "소유자"):
                ExportStore(Path(root) / "private-exports", run=run, windows=True)
            self.assertEqual(run.call_count, 1)

    def test_rejects_symbolic_link_directory(self):
        with tempfile.TemporaryDirectory() as root:
            real = Path(root) / "real"
            real.mkdir()
            link = Path(root) / "link"
            try:
                link.symlink_to(real, target_is_directory=True)
            except (OSError, NotImplementedError):
                self.skipTest("symbolic links unavailable")
            with self.assertRaisesRegex(RuntimeError, "심볼릭"):
                ExportStore(link)

    def test_prune_removes_old_workbooks_in_the_private_folder_only(self):
        with tempfile.TemporaryDirectory() as root:
            store = ExportStore(root)
            own = Path(store.allocate())
            own.write_bytes(b"fixture")
            fresh = Path(store.allocate())
            fresh.write_bytes(b"fixture")
            default_name = Path(root) / "임플란트수술통계_20261121.xlsx"  # 이름 입력 실패로 덴트웹 기본 이름 저장
            default_name.write_bytes(b"patient")
            note = Path(root) / "readme.txt"
            note.write_bytes(b"keep")
            old = time.time() - 2 * 3600
            for path in (own, default_name, note):
                os.utime(path, (old, old))
            store.prune()
            self.assertFalse(own.exists())
            self.assertFalse(default_name.exists())
            self.assertTrue(fresh.exists())
            self.assertTrue(note.exists())
            store.cleanup(fresh)
            store.cleanup(fresh)
            self.assertFalse(fresh.exists())

    def test_remove_strays_deletes_only_new_non_owned_workbooks(self):
        with tempfile.TemporaryDirectory() as root:
            store = ExportStore(root)
            before = Path(root) / "earlier.xlsx"
            before.write_bytes(b"x")
            os.utime(before, (time.time() - 600, time.time() - 600))
            since = time.time() - 5
            stray = Path(root) / "default-name.xlsx"
            stray.write_bytes(b"patient")
            own = Path(store.allocate())
            own.write_bytes(b"x")
            self.assertEqual(store.remove_strays(since), ["default-name.xlsx"])
            self.assertFalse(stray.exists())
            self.assertTrue(before.exists())
            self.assertTrue(own.exists())

    def test_cleanup_never_touches_files_outside_or_symbolic_links(self):
        with tempfile.TemporaryDirectory() as root, tempfile.TemporaryDirectory() as other:
            store = ExportStore(Path(root) / "private-exports")
            outside = Path(other) / ("denjoy-export-" + "a" * 32 + ".xlsx")
            outside.write_bytes(b"preserve")
            store.cleanup(outside)
            self.assertTrue(outside.exists())
            link = store.directory / ("denjoy-export-" + "0" * 32 + ".xlsx")
            try:
                link.symlink_to(outside)
            except (OSError, NotImplementedError):
                self.skipTest("symbolic links unavailable")
            store.cleanup(link)
            store.prune(now=time.time() + 172800)
            self.assertTrue(outside.exists())
            self.assertTrue(link.is_symlink())

    def test_prune_is_bounded(self):
        with tempfile.TemporaryDirectory() as root:
            store = ExportStore(root)
            with patch.object(export_store, "PRUNE_LIMIT", 3):
                for _ in range(5):
                    path = Path(store.allocate())
                    path.write_bytes(b"x")
                    os.utime(path, (1, 1))
                store.prune()
            self.assertEqual(len(list(Path(root).iterdir())), 2)

    @unittest.skipUnless(os.name == "nt", "real Windows ACL")
    def test_windows_folder_acl_is_protected_and_lists_only_user_and_system(self):
        with tempfile.TemporaryDirectory() as root:
            store = ExportStore(Path(root) / "private-exports")
            script = ("$acl = Get-Acl -LiteralPath $env:EXPORT_DIR; "
                      "'protected=' + $acl.AreAccessRulesProtected; "
                      "$acl.Access | ForEach-Object { $_.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]).Value }")
            out = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
                                 capture_output=True, text=True, check=True,
                                 env={**os.environ, "EXPORT_DIR": str(store.directory)}).stdout.split()
            me = subprocess.run(["whoami", "/user", "/fo", "csv", "/nh"], capture_output=True, text=True,
                                check=True).stdout.strip().split(",")[1].strip('"')
            self.assertIn("protected=True", out)
            self.assertEqual(set(line for line in out if line.startswith("S-1-")), {me, "S-1-5-18"})

    def test_default_directory_is_under_local_app_data(self):
        with patch.dict(os.environ, {"LOCALAPPDATA": "/fixture/local"}):
            self.assertEqual(export_store.default_directory(), Path("/fixture/local/DenJOY/private-exports"))


if __name__ == "__main__":
    unittest.main()
