"""config.json 로드 및 검증 — 첫 실행 시 대화형 설정"""

import json
import os
import sys

SERVER_URL = "https://qhoyaonrkagdngglrcas.supabase.co/functions/v1"

DEFAULTS = {
    "server_url": SERVER_URL,
    "poll_interval_seconds": 30,
    "download_dir": os.path.join(os.path.expanduser("~"), "Downloads"),
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

    cfg["server_url"] = cfg["server_url"].rstrip("/")
    return cfg
