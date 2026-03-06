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


def main():
    _setup_dpi_awareness()

    if len(sys.argv) > 1 and sys.argv[1] == "--startup":
        from config import toggle_startup
        toggle_startup()
        return

    from gui import run_gui
    run_gui()


if __name__ == "__main__":
    main()
