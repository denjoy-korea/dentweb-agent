"""DentWeb Automation Agent - 진입점"""

import sys

def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--startup":
        from config import toggle_startup
        toggle_startup()
        return

    from gui import run_gui
    run_gui()


if __name__ == "__main__":
    main()
