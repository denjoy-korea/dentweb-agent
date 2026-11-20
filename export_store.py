"""에이전트 전용 내보내기 폴더. 덴트웹이 저장한 수술 기록 엑셀만 잠깐 둔다.

덴트웹이 저장한 수술 기록 엑셀은 환자 정보다. 3.6.4까지는 EXE 옆 exports/ 폴더(권한 상속)에
저장하고 정상 경로에서만 지웠다. 3.6.5부터는
  - 현재 Windows 사용자와 SYSTEM만 열 수 있는 전용 폴더에
  - 실행마다 고유한 이름으로 저장하고
  - 업로드 결과와 관계없이 그 파일을 지운다.
  - 실행할 때마다 1시간 지난 엑셀을 지운다(늦게 저장됐거나 잠겨서 못 지운 파일, 이름 입력에 실패해
    덴트웹 기본 이름으로 저장된 파일 포함). 이 폴더는 에이전트만 쓰는 곳이다.
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
PRUNE_AGE_SECONDS = 3600
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


def _system_tool(name: str) -> str:
    """관리자 권한으로 돌 때 EXE 폴더에 놓인 같은 이름의 파일을 실행하지 않도록 System32 절대 경로를 쓴다."""
    return os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", name)


def _no_window():
    return getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _is_workbook(entry: Path) -> bool:
    return entry.suffix.lower() == ".xlsx" and not entry.is_symlink() and entry.is_file()


def remove_owned(directory, path):
    """전용 폴더 안의 고유 이름 파일만 지운다. 지우지 못하면 OSError를 그대로 낸다."""
    if not path:
        return
    p = Path(path)
    if (p.parent.absolute() == Path(directory).absolute()
            and OWNED.fullmatch(p.name) and not p.is_symlink()):
        p.unlink(missing_ok=True)


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
        result = run([_system_tool("whoami.exe"), "/user", "/fo", "csv", "/nh"], check=True,
                     capture_output=True, text=True, errors="replace",
                     stdin=subprocess.DEVNULL, creationflags=_no_window())
        sid = next(csv.reader(io.StringIO(result.stdout)))[1]
        if not re.fullmatch(r"S-1-[0-9-]+", sid):
            raise RuntimeError("내보내기 폴더의 소유자를 확인할 수 없습니다")
        # 상속을 끊고 현재 사용자와 SYSTEM만 남긴다.
        run([_system_tool("icacls.exe"), str(self.directory), "/inheritance:r", "/grant:r",
             "*" + sid + ":(OI)(CI)F", "*" + SYSTEM_SID + ":(OI)(CI)F"],
            check=True, capture_output=True, stdin=subprocess.DEVNULL, creationflags=_no_window())

    def allocate(self) -> str:
        return str(self.directory / ("denjoy-export-" + uuid.uuid4().hex + ".xlsx"))

    def owned(self, path) -> bool:
        p = Path(path)
        return (p.parent.absolute() == self.directory.absolute()
                and bool(OWNED.fullmatch(p.name)) and not p.is_symlink())

    def cleanup(self, path):
        """이 폴더의 고유 이름 파일만 지운다. 지우지 못하면 OSError를 그대로 낸다."""
        remove_owned(self.directory, path)

    def prune(self, now=None):
        """1시간 지난 엑셀을 지운다(이름과 관계없이, 바로가기·하위 폴더·다른 형식은 두고)."""
        cutoff = (time.time() if now is None else now) - PRUNE_AGE_SECONDS
        # 비정상 종료로 남은 파일까지, 한 번에 정해진 수만 본다.
        for index, entry in enumerate(self.directory.iterdir()):
            if index >= PRUNE_LIMIT:
                break
            try:
                if _is_workbook(entry) and entry.stat().st_mtime < cutoff:
                    entry.unlink(missing_ok=True)
            except OSError:
                continue

    def remove_strays(self, since: float) -> list:
        """since 이후 이 폴더에 생긴, 고유 이름이 아닌 엑셀을 지우고 그 이름을 돌려준다.

        파일 이름 칸에 고유 이름을 넣지 못하면 덴트웹이 기본 이름으로 이 폴더에 저장한다.
        """
        removed = []
        for index, entry in enumerate(self.directory.iterdir()):
            if index >= PRUNE_LIMIT:
                break
            try:
                if (_is_workbook(entry) and not OWNED.fullmatch(entry.name)
                        and entry.stat().st_mtime >= since):
                    entry.unlink(missing_ok=True)
                    removed.append(entry.name)
            except OSError:
                continue
        return removed
