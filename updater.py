"""자동 업데이트: GitHub Releases에서 최신 버전 확인 및 다운로드"""

import os
import sys
import subprocess
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


def download_update(download_url: str, progress_callback=None,
                    log_callback=None) -> str | None:
    """새 exe를 현재 exe 옆에 다운로드 → 경로 반환."""
    def _log(msg):
        if log_callback:
            log_callback(msg)

    try:
        if getattr(sys, "frozen", False):
            target_dir = os.path.dirname(sys.executable)
        else:
            target_dir = tempfile.gettempdir()

        download_path = os.path.join(target_dir, "dentweb-agent-update.exe")

        # 이전 잔여 파일 정리
        for old in ("dentweb-agent-update.exe", "dentweb-agent-old.exe"):
            old_path = os.path.join(target_dir, old)
            if os.path.exists(old_path):
                try:
                    os.remove(old_path)
                except OSError:
                    pass

        _log(f"다운로드 시작: {download_url}")
        resp = requests.get(download_url, stream=True, timeout=180,
                            allow_redirects=True)
        resp.raise_for_status()

        total = int(resp.headers.get("content-length", 0))
        downloaded = 0
        _log(f"파일 크기: {total / 1024 / 1024:.1f} MB")

        with open(download_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=65536):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback and total > 0:
                        progress_callback(downloaded / total)

        actual_size = os.path.getsize(download_path)
        _log(f"다운로드 완료: {actual_size / 1024 / 1024:.1f} MB")

        # 크기 검증 (content-length가 있을 때만)
        if total > 0 and actual_size < total * 0.95:
            _log(f"크기 불일치 — 예상 {total}, 실제 {actual_size}")
            os.remove(download_path)
            return None

        return download_path
    except Exception as e:
        _log(f"다운로드 오류: {e}")
        return None


def apply_update(new_exe_path: str):
    """현재 exe를 교체하고 재시작 (PowerShell 사용)."""
    if not getattr(sys, "frozen", False):
        return

    current_exe = sys.executable
    current_dir = os.path.dirname(current_exe)
    ps_path = os.path.join(current_dir, "dentweb-agent-update.ps1")

    # PowerShell 스크립트: 프로세스 종료 대기 → 복사로 교체 → 재시작
    ps_content = f'''
$ErrorActionPreference = "Stop"
$currentExe = "{current_exe}"
$newExe = "{new_exe_path}"
$oldExe = "{os.path.join(current_dir, 'dentweb-agent-old.exe')}"
$pid = {os.getpid()}

# 1. 현재 프로세스 종료 대기
Write-Host "업데이트 적용 중... 프로세스 종료 대기"
try {{
    Wait-Process -Id $pid -Timeout 15 -ErrorAction SilentlyContinue
}} catch {{
    Start-Sleep -Seconds 3
}}
Start-Sleep -Seconds 2

# 2. 현재 exe 백업
if (Test-Path $oldExe) {{ Remove-Item $oldExe -Force -ErrorAction SilentlyContinue }}
if (Test-Path $currentExe) {{
    Move-Item $currentExe $oldExe -Force
}}

# 3. 새 exe로 교체
Copy-Item $newExe $currentExe -Force

# 4. 새 exe 실행
Start-Process $currentExe

# 5. 정리
Start-Sleep -Seconds 3
Remove-Item $newExe -Force -ErrorAction SilentlyContinue
Remove-Item $oldExe -Force -ErrorAction SilentlyContinue
Remove-Item $MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue
'''

    with open(ps_path, "w", encoding="utf-8") as f:
        f.write(ps_content)

    # PowerShell 실행 (창 숨김, 실행 정책 우회)
    subprocess.Popen(
        [
            "powershell", "-ExecutionPolicy", "Bypass",
            "-WindowStyle", "Hidden",
            "-File", ps_path,
        ],
        creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
    )
    sys.exit(0)
