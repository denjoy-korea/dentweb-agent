"""Windows 시작프로그램 등록/해제"""

import os
import sys


def _get_startup_folder() -> str:
    """Windows Startup 폴더 경로"""
    return os.path.join(
        os.environ.get("APPDATA", ""),
        "Microsoft", "Windows", "Start Menu", "Programs", "Startup",
    )


def _get_shortcut_path() -> str:
    return os.path.join(_get_startup_folder(), "DenJOY DentWeb Agent.lnk")


def _get_exe_path() -> str:
    """현재 실행 중인 exe 경로 (PyInstaller frozen 또는 .py)"""
    if getattr(sys, "frozen", False):
        return sys.executable
    return os.path.abspath(sys.argv[0])


def is_registered() -> bool:
    """시작프로그램에 등록되어 있는지 확인"""
    return os.path.exists(_get_shortcut_path())


def register() -> bool:
    """시작프로그램에 등록 (Windows 바로가기 생성)"""
    try:
        import winreg
    except ImportError:
        # winreg 없으면 Windows가 아님
        print("[안내] Windows에서만 시작프로그램 등록이 가능합니다.")
        return False

    exe_path = _get_exe_path()
    working_dir = os.path.dirname(exe_path)

    try:
        # PowerShell로 .lnk 바로가기 생성 (COM 없이)
        shortcut_path = _get_shortcut_path()
        ps_script = (
            f'$ws = New-Object -ComObject WScript.Shell; '
            f'$s = $ws.CreateShortcut("{shortcut_path}"); '
            f'$s.TargetPath = "{exe_path}"; '
            f'$s.WorkingDirectory = "{working_dir}"; '
            f'$s.Description = "DenJOY DentWeb Automation Agent"; '
            f'$s.Save()'
        )
        os.system(f'powershell -Command "{ps_script}"')

        if os.path.exists(shortcut_path):
            return True

        # 바로가기 실패 시 레지스트리 방식
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_SET_VALUE,
        )
        winreg.SetValueEx(key, "DenJOY DentWeb Agent", 0, winreg.REG_SZ, exe_path)
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"[ERROR] 시작프로그램 등록 실패: {e}")
        return False


def unregister() -> bool:
    """시작프로그램에서 제거"""
    removed = False

    # 바로가기 제거
    shortcut_path = _get_shortcut_path()
    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)
        removed = True

    # 레지스트리 항목도 제거 (있을 경우)
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_SET_VALUE,
        )
        try:
            winreg.DeleteValue(key, "DenJOY DentWeb Agent")
            removed = True
        except FileNotFoundError:
            pass
        winreg.CloseKey(key)
    except ImportError:
        pass
    except Exception:
        pass

    return removed
