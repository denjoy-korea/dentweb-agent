"""config.json 로드 및 검증 — 첫 실행 시 대화형 설정"""

import json
import os
import sys
from startup import is_registered, register, unregister

SERVER_URL = "https://qhoyaonrkagdngglrcas.supabase.co/functions/v1"

def _default_exports_dir() -> str:
    """EXE 옆 exports/ 폴더 경로"""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "exports")


DEFAULTS = {
    "server_url": SERVER_URL,
    "poll_interval_seconds": 30,
    "download_dir": _default_exports_dir(),
    "dentweb_window_title": "덴트웹",
    "click_delay_ms": 300,
    "download_timeout_seconds": 30,
    "max_retries": 3,
    "log_file": "agent.log",
    "log_max_lines": 1000,
}


def _interactive_setup(path: str) -> dict:
    """첫 실행: 토큰만 입력받아 config.json 자동 생성"""
    print("=" * 50)
    print("  DenJOY 덴트웹 자동화 에이전트 - 초기 설정")
    print("=" * 50)
    print()
    print("앱 설정 화면에서 복사한 에이전트 토큰을 붙여넣어 주세요.")
    print()

    while True:
        token = input("에이전트 토큰: ").strip()
        if token:
            break
        print("토큰을 입력해주세요.")

    cfg = {**DEFAULTS, "agent_token": token}

    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)

    print()
    print(f"설정이 저장되었습니다: {os.path.abspath(path)}")

    # 시작프로그램 등록 여부
    _ask_startup_registration()

    print("에이전트를 시작합니다...")
    print()
    return cfg


def load_config(path: str = "config.json") -> dict:
    if not os.path.exists(path):
        return _interactive_setup(path)

    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    if not cfg.get("agent_token"):
        print("[ERROR] config.json에 agent_token이 없습니다.")
        print("config.json을 삭제하고 다시 실행하면 초기 설정이 시작됩니다.")
        sys.exit(1)

    for key, default in DEFAULTS.items():
        cfg.setdefault(key, default)

    # EXE 위치가 변경되어도 exports 폴더 경로를 항상 최신으로 유지
    cfg["download_dir"] = _default_exports_dir()

    cfg["server_url"] = cfg["server_url"].rstrip("/")
    return cfg


def _ask_startup_registration():
    """시작프로그램 등록 여부를 사용자에게 묻기"""
    if sys.platform != "win32":
        return

    if is_registered():
        print("시작프로그램: 이미 등록되어 있습니다.")
        return

    print()
    print("컴퓨터를 켤 때 에이전트를 자동으로 실행하시겠습니까?")
    answer = input("시작프로그램 등록 (Y/n): ").strip().lower()
    if answer in ("", "y", "yes"):
        if register():
            print("시작프로그램에 등록되었습니다. PC 부팅 시 자동 실행됩니다.")
        else:
            print("시작프로그램 등록에 실패했습니다. 수동으로 등록해주세요.")
    else:
        print("시작프로그램 등록을 건너뜁니다.")
    print()


def toggle_startup():
    """시작프로그램 등록/해제 토글"""
    if sys.platform != "win32":
        print("[안내] Windows에서만 시작프로그램 등록이 가능합니다.")
        return

    if is_registered():
        if unregister():
            print("시작프로그램에서 제거되었습니다.")
        else:
            print("시작프로그램 제거에 실패했습니다.")
    else:
        if register():
            print("시작프로그램에 등록되었습니다.")
        else:
            print("시작프로그램 등록에 실패했습니다.")
