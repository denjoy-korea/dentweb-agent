"""config.json 로드 및 검증 — 첫 실행 시 대화형 설정"""

import json
import os
import sys
import tempfile
from getpass import getpass
from credential_store import protect_token, unprotect_token
from startup import is_registered, register, unregister

SERVER_URL = "https://fcsclwodponmwadyggvl.supabase.co/functions/v1"
LEGACY_SERVER_URL = "https://qhoyaonrkagdngglrcas.supabase.co/functions/v1"

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
        token = getpass("에이전트 토큰: ").strip()
        if token:
            break
        print("토큰을 입력해주세요.")

    cfg = {**DEFAULTS, "agent_token": token}

    save_config(path, cfg)

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

    if not isinstance(cfg, dict):
        raise ValueError("유효하지 않은 에이전트 설정입니다.")
    legacy_token = bool(cfg.get("agent_token"))
    if not legacy_token and cfg.get("agent_token_protected"):
        cfg["agent_token"] = unprotect_token(cfg["agent_token_protected"])

    if not cfg.get("agent_token"):
        print("[ERROR] config.json에 agent_token이 없습니다.")
        print("config.json을 삭제하고 다시 실행하면 초기 설정이 시작됩니다.")
        sys.exit(1)

    for key, default in DEFAULTS.items():
        cfg.setdefault(key, default)

    # EXE 위치가 변경되어도 exports 폴더 경로를 항상 최신으로 유지
    cfg["download_dir"] = _default_exports_dir()

    cfg["server_url"] = cfg["server_url"].rstrip("/")
    legacy_server = cfg["server_url"] == LEGACY_SERVER_URL
    if legacy_server:
        cfg["server_url"] = SERVER_URL
    # Migrate before returning a credential to any HTTP client. Failure stops startup.
    if legacy_token or legacy_server:
        save_config(path, cfg)
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


def save_config(path: str, cfg: dict) -> None:
    """All CLI/GUI saves strip in-memory credentials and atomically replace JSON."""
    stored = {key: value for key, value in cfg.items() if key != "agent_token"}
    if cfg.get("agent_token"):
        stored["agent_token_protected"] = protect_token(cfg["agent_token"])
    elif not stored.get("agent_token_protected"):
        raise ValueError("인증정보가 없습니다.")
    directory = os.path.dirname(os.path.abspath(path))
    fd, temporary = tempfile.mkstemp(prefix=".denjoy-config-", dir=directory, text=True)
    try:
        os.chmod(temporary, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as stream:
            json.dump(stored, stream, indent=2, ensure_ascii=False)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
