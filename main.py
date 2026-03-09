"""DentWeb Automation Agent - 진입점"""

import sys
import os


def _setup_dpi_awareness():
    """Windows DPI 스케일링 문제 해결 — pyautogui 좌표와 실제 화면 좌표 일치시킴"""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        # Per-Monitor DPI Aware (v2) — 가장 정확한 좌표 매핑
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            import ctypes
            # 폴백: System DPI Aware
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def _cleanup_update_files():
    """이전 업데이트에서 남은 임시 파일 정리"""
    if not getattr(sys, "frozen", False):
        return
    base_dir = os.path.dirname(sys.executable)
    for name in ("dentweb-agent-update.exe", "dentweb-agent-old.exe", "dentweb-agent-update.ps1"):
        path = os.path.join(base_dir, name)
        try:
            if os.path.exists(path):
                os.remove(path)
        except OSError:
            pass


def _ensure_exports_dir():
    """EXE 옆 exports/ 폴더 자동 생성"""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(os.path.join(base, "exports"), exist_ok=True)


def main():
    _setup_dpi_awareness()
    _cleanup_update_files()
    _ensure_exports_dir()

    if len(sys.argv) > 1 and sys.argv[1] == "--startup":
        from config import toggle_startup
        toggle_startup()
        return

    from gui import run_gui
    run_gui()


if __name__ == "__main__":
    main()
