"""자동 업데이트: GitHub Releases에서 최신 버전 확인 및 다운로드"""

import os
import sys
import tempfile
import requests

REPO = "denjoy-korea/dentweb-agent"
GITHUB_API = f"https://api.github.com/repos/{REPO}/releases/latest"


def parse_version(tag: str) -> tuple:
    """'v2.3.0' → (2, 3, 0)"""
    clean = tag.lstrip("v").split("-")[0]
    parts = clean.split(".")
    return tuple(int(p) for p in parts if p.isdigit())


def check_update(current_version: str) -> dict | None:
    """최신 릴리즈 확인. 업데이트 있으면 {version, download_url} 반환."""
    try:
        resp = requests.get(GITHUB_API, timeout=10)
        resp.raise_for_status()
        data = resp.json()

        latest_tag = data.get("tag_name", "")
        if not latest_tag:
            return None

        current = parse_version(current_version)
        latest = parse_version(latest_tag)

        if latest <= current:
            return None

        # exe 에셋 찾기
        for asset in data.get("assets", []):
            if asset["name"].endswith(".exe"):
                return {
                    "version": latest_tag,
                    "download_url": asset["browser_download_url"],
                    "size": asset.get("size", 0),
                }
        return None
    except Exception:
        return None


def download_update(download_url: str, progress_callback=None) -> str | None:
    """새 exe 다운로드 → 임시 파일 경로 반환."""
    try:
        resp = requests.get(download_url, stream=True, timeout=120)
        resp.raise_for_status()

        total = int(resp.headers.get("content-length", 0))
        downloaded = 0

        tmp = tempfile.NamedTemporaryFile(
            delete=False, suffix=".exe", prefix="dentweb-agent-update-",
        )

        for chunk in resp.iter_content(chunk_size=65536):
            tmp.write(chunk)
            downloaded += len(chunk)
            if progress_callback and total > 0:
                progress_callback(downloaded / total)

        tmp.close()
        return tmp.name
    except Exception:
        return None


def apply_update(new_exe_path: str):
    """현재 exe를 교체하고 재시작 (배치 스크립트 사용)."""
    if not getattr(sys, "frozen", False):
        return

    current_exe = sys.executable
    bat_path = os.path.join(tempfile.gettempdir(), "dentweb-agent-update.bat")

    # 배치 스크립트: 현재 프로세스 종료 대기 → 교체 → 재시작 → 자기 삭제
    bat_content = f'''@echo off
echo 업데이트 적용 중...
timeout /t 2 /nobreak >nul
:retry
del "{current_exe}" >nul 2>&1
if exist "{current_exe}" (
    timeout /t 1 /nobreak >nul
    goto retry
)
move /y "{new_exe_path}" "{current_exe}"
start "" "{current_exe}"
del "%~f0"
'''

    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(bat_content)

    os.startfile(bat_path)
    sys.exit(0)
