"""v3.6.5 수술 기록 파일 보호: 전용 폴더, 고유 이름, 소유 파일만 정리."""
import os
from pathlib import Path
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
            self.assertEqual(whoami.args[0][0], "whoami")
            command = icacls.args[0]
            self.assertEqual(command[0], "icacls")
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

    def test_only_owned_old_exports_are_pruned_and_cleanup_is_idempotent(self):
        with tempfile.TemporaryDirectory() as root:
            store = ExportStore(root)
            own = Path(store.allocate())
            own.write_bytes(b"fixture")
            fresh = Path(store.allocate())
            fresh.write_bytes(b"fixture")
            unrelated = Path(root) / "patient.xlsx"
            unrelated.write_bytes(b"preserve")
            old = time.time() - 172800
            os.utime(own, (old, old))
            os.utime(unrelated, (old, old))
            store.prune()
            self.assertFalse(own.exists())
            self.assertTrue(fresh.exists())
            self.assertTrue(unrelated.exists())
            store.cleanup(fresh)
            store.cleanup(fresh)
            self.assertFalse(fresh.exists())

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

    def test_default_directory_is_under_local_app_data(self):
        with patch.dict(os.environ, {"LOCALAPPDATA": "/fixture/local"}):
            self.assertEqual(export_store.default_directory(), Path("/fixture/local/DenJOY/private-exports"))


if __name__ == "__main__":
    unittest.main()
