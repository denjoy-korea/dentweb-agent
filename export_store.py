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


def set_private_dacl(path: str, sid: str):
    """폴더 권한을 '현재 사용자·SYSTEM 모두 허용, 상속 끊음' 두 항목으로 통째로 바꾼다.

    icacls /grant:r은 지정한 주체만 바꾸고 다른 명시적 권한(관리자 권한 실행 때의 Administrators 등)을
    남기므로, SDDL로 만든 보호된 DACL을 SetNamedSecurityInfoW로 넣는다.
    """
    import ctypes
    from ctypes import wintypes
    advapi = ctypes.WinDLL("advapi32", use_last_error=True)
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    advapi.ConvertStringSecurityDescriptorToSecurityDescriptorW.argtypes = [
        wintypes.LPCWSTR, wintypes.DWORD, ctypes.POINTER(ctypes.c_void_p), ctypes.POINTER(wintypes.ULONG)]
    advapi.ConvertStringSecurityDescriptorToSecurityDescriptorW.restype = wintypes.BOOL
    advapi.GetSecurityDescriptorDacl.argtypes = [
        ctypes.c_void_p, ctypes.POINTER(wintypes.BOOL), ctypes.POINTER(ctypes.c_void_p), ctypes.POINTER(wintypes.BOOL)]
    advapi.GetSecurityDescriptorDacl.restype = wintypes.BOOL
    advapi.SetNamedSecurityInfoW.argtypes = [
        wintypes.LPWSTR, ctypes.c_int, wintypes.DWORD, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p]
    advapi.SetNamedSecurityInfoW.restype = wintypes.DWORD
    kernel.LocalFree.argtypes = [ctypes.c_void_p]
    kernel.LocalFree.restype = ctypes.c_void_p

    if not re.fullmatch(r"S-1-[0-9-]+", sid):
        raise RuntimeError("내보내기 폴더의 소유자를 확인할 수 없습니다")
    sddl = f"D:P(A;OICI;FA;;;{sid})(A;OICI;FA;;;SY)"
    descriptor = ctypes.c_void_p()
    if not advapi.ConvertStringSecurityDescriptorToSecurityDescriptorW(sddl, 1, ctypes.byref(descriptor), None):
        raise OSError(ctypes.get_last_error(), "내보내기 폴더 권한을 만들 수 없습니다")
    try:
        present, defaulted, dacl = wintypes.BOOL(), wintypes.BOOL(), ctypes.c_void_p()
        if not advapi.GetSecurityDescriptorDacl(descriptor, ctypes.byref(present), ctypes.byref(dacl), ctypes.byref(defaulted)):
            raise OSError(ctypes.get_last_error(), "내보내기 폴더 권한을 읽을 수 없습니다")
        se_file_object = 1
        dacl_information = 0x00000004
        protected_dacl_information = 0x80000000
        result = advapi.SetNamedSecurityInfoW(str(path), se_file_object, dacl_information | protected_dacl_information,
                                              None, None, dacl, None)
        if result != 0:
            raise OSError(result, "내보내기 폴더 권한을 바꿀 수 없습니다")
    finally:
        kernel.LocalFree(descriptor)


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
    def __init__(self, directory=None, run=subprocess.run, windows=None, set_dacl=None):
        self.directory = Path(directory) if directory else default_directory()
        if self.directory.is_symlink():
            raise RuntimeError("내보내기 폴더가 바로가기(심볼릭 링크)입니다")
        self.directory.mkdir(parents=True, exist_ok=True, mode=0o700)
        if self.directory.is_symlink() or not self.directory.is_dir():
            raise RuntimeError("내보내기 폴더를 만들 수 없습니다")
        if (os.name == "nt") if windows is None else windows:
            self._restrict_windows(run, set_dacl or set_private_dacl)
        else:
            self.directory.chmod(0o700)
        self.prune()

    def _restrict_windows(self, run, set_dacl):
        result = run([_system_tool("whoami.exe"), "/user", "/fo", "csv", "/nh"], check=True,
                     capture_output=True, text=True, errors="replace",
                     stdin=subprocess.DEVNULL, creationflags=_no_window())
        sid = next(csv.reader(io.StringIO(result.stdout)))[1]
        if not re.fullmatch(r"S-1-[0-9-]+", sid):
            raise RuntimeError("내보내기 폴더의 소유자를 확인할 수 없습니다")
        # 상속을 끊고 현재 사용자와 SYSTEM만 남긴다(다른 명시적 권한도 지운다).
        set_dacl(str(self.directory), sid)

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
