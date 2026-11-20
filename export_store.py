"""에이전트가 만든 내보내기 파일만 둔다. 사용자의 다른 파일은 모으거나 지우지 않는다.

덴트웹이 저장한 수술 기록 엑셀은 환자 정보다. 3.6.4까지는 EXE 옆 exports/ 폴더(권한 상속)에
저장하고 정상 경로에서만 지웠다. 3.6.5부터는
  - 현재 Windows 사용자와 SYSTEM만 열 수 있는 전용 폴더에
  - 실행마다 고유한 이름으로 저장하고
  - 업로드 결과와 관계없이 그 파일을 지운다(남은 것은 24시간 뒤 정리).
"""
import csv
import io
import os
from pathlib import Path
import re
import subprocess
import sys
import time
import uuid

OWNED = re.compile(r"^denjoy-export-[0-9a-f]{32}\.xlsx$")
PRUNE_AGE_SECONDS = 86400
PRUNE_LIMIT = 1000
SYSTEM_SID = "S-1-5-18"


def default_directory() -> Path:
    """전용 폴더 경로(만들지 않는다)."""
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / ".local" / "share")
    return Path(base) / "DenJOY" / "private-exports"


def legacy_exports_directory() -> Path:
    """3.6.4까지 쓰던 EXE 옆 exports/ 폴더 경로(만들지 않는다)."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return Path(base) / "exports"


def _no_window():
    return getattr(subprocess, "CREATE_NO_WINDOW", 0)


class ExportStore:
    def __init__(self, directory=None, run=subprocess.run, windows=None):
        self.directory = Path(directory) if directory else default_directory()
        if self.directory.is_symlink():
            raise RuntimeError("내보내기 폴더가 바로가기(심볼릭 링크)입니다")
        self.directory.mkdir(parents=True, exist_ok=True, mode=0o700)
        if self.directory.is_symlink() or not self.directory.is_dir():
            raise RuntimeError("내보내기 폴더를 만들 수 없습니다")
        if (os.name == "nt") if windows is None else windows:
            self._restrict_windows(run)
        else:
            self.directory.chmod(0o700)
        self.prune()

    def _restrict_windows(self, run):
        result = run(["whoami", "/user", "/fo", "csv", "/nh"], check=True, capture_output=True,
                     text=True, creationflags=_no_window())
        sid = next(csv.reader(io.StringIO(result.stdout)))[1]
        if not re.fullmatch(r"S-1-[0-9-]+", sid):
            raise RuntimeError("내보내기 폴더의 소유자를 확인할 수 없습니다")
        # 상속을 끊고 현재 사용자와 SYSTEM만 남긴다.
        run(["icacls", str(self.directory), "/inheritance:r", "/grant:r",
             "*" + sid + ":(OI)(CI)F", "*" + SYSTEM_SID + ":(OI)(CI)F"],
            check=True, capture_output=True, creationflags=_no_window())

    def allocate(self) -> str:
        return str(self.directory / ("denjoy-export-" + uuid.uuid4().hex + ".xlsx"))

    def owned(self, path) -> bool:
        p = Path(path)
        return (p.parent.absolute() == self.directory.absolute()
                and bool(OWNED.fullmatch(p.name)) and not p.is_symlink())

    def cleanup(self, path):
        """이 폴더의 소유 패턴 파일만 지운다. 지우지 못하면 OSError를 그대로 낸다."""
        if path and self.owned(path):
            Path(path).unlink(missing_ok=True)

    def prune(self, now=None):
        cutoff = (time.time() if now is None else now) - PRUNE_AGE_SECONDS
        # 비정상 종료로 남은 파일까지, 한 번에 정해진 수만 본다.
        for index, entry in enumerate(self.directory.iterdir()):
            if index >= PRUNE_LIMIT:
                break
            try:
                if self.owned(entry) and entry.is_file() and entry.stat().st_mtime < cutoff:
                    self.cleanup(entry)
            except OSError:
                continue
