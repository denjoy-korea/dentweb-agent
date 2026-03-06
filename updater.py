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
    """새 exe를 현재 exe 옆에 다운로드 → 경로 반환."""
    try:
        # 현재 exe와 같은 디렉토리에 저장 (드라이브 간 이동 문제 방지)
        if getattr(sys, "frozen", False):
            target_dir = os.path.dirname(sys.executable)
        else:
            target_dir = tempfile.gettempdir()

        download_path = os.path.join(target_dir, "dentweb-agent-update.exe")

        # 이전 다운로드 잔여 파일 정리
        for old in ("dentweb-agent-update.exe", "dentweb-agent-old.exe"):
            old_path = os.path.join(target_dir, old)
            if os.path.exists(old_path):
                try:
                    os.remove(old_path)
                except OSError:
                    pass

        resp = requests.get(download_url, stream=True, timeout=120)
        resp.raise_for_status()

        total = int(resp.headers.get("content-length", 0))
        downloaded = 0

        with open(download_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=65536):
                f.write(chunk)
                downloaded += len(chunk)
                if progress_callback and total > 0:
                    progress_callback(downloaded / total)

        # 파일 크기 검증
        actual_size = os.path.getsize(download_path)
        if total > 0 and actual_size != total:
            os.remove(download_path)
            return None

        return download_path
    except Exception:
        return None


def apply_update(new_exe_path: str):
    """현재 exe를 교체하고 재시작 (배치 스크립트 사용)."""
    if not getattr(sys, "frozen", False):
        return

    current_exe = sys.executable
    current_dir = os.path.dirname(current_exe)
    current_name = os.path.basename(current_exe)
    old_name = "dentweb-agent-old.exe"
    new_name = os.path.basename(new_exe_path)
    bat_path = os.path.join(current_dir, "dentweb-agent-update.bat")

    # 배치 스크립트:
    # 1. 현재 프로세스 종료 대기
    # 2. 현재 exe → old로 이름변경 (delete보다 안전)
    # 3. 새 exe → 현재 이름으로 이름변경
    # 4. 새 exe 실행
    # 5. old exe 삭제 + 자기 삭제
    bat_content = f'''@echo off
chcp 65001 >nul
echo 업데이트 적용 중...
timeout /t 3 /nobreak >nul

:wait_exit
tasklist /FI "IMAGENAME eq {current_name}" 2>nul | find /I "{current_name}" >nul
if not errorlevel 1 (
    timeout /t 1 /nobreak >nul
    goto wait_exit
)

if exist "{old_name}" del /f "{old_name}" >nul 2>&1
rename "{current_name}" "{old_name}"
if errorlevel 1 (
    echo 이름 변경 실패. 수동으로 업데이트하세요.
    pause
    exit /b 1
)
rename "{new_name}" "{current_name}"
if errorlevel 1 (
    rename "{old_name}" "{current_name}"
    echo 업데이트 실패. 원래 버전으로 복원했습니다.
    pause
    exit /b 1
)

timeout /t 1 /nobreak >nul
start "" "{current_exe}"
timeout /t 2 /nobreak >nul
del /f "{old_name}" >nul 2>&1
del /f "%~f0"
'''

    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(bat_content)

    # bat를 exe와 같은 디렉토리에서 실행
    os.chdir(current_dir)
    os.startfile(bat_path)
    sys.exit(0)
