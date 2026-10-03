"""자동 업데이트: GitHub Releases에서 최신 버전 확인 및 다운로드"""

import os
import sys
import subprocess
import tempfile
import requests
import hashlib
import hmac
import re
from urllib.parse import urlparse

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

        assets = data.get("assets", [])
        if not any(asset.get("name") == "SHA256SUMS" for asset in assets):
            return None
        for asset in assets:
            if asset.get("name") == "dentweb-agent.exe":
                return {
                    "version": latest_tag,
                    "download_url": asset["browser_download_url"],
                    "size": asset.get("size", 0),
                }
        return None
    except Exception:
        return None


def _checksum_url(download_url: str) -> str:
    """Only the official repository's versioned release can supply an update."""
    parsed = urlparse(download_url)
    prefix = f"/{REPO}/releases/download/"
    if (parsed.scheme != "https" or parsed.netloc != "github.com"
            or parsed.query or parsed.fragment or not parsed.path.startswith(prefix)):
        raise ValueError("공식 에이전트 릴리즈 주소가 아닙니다.")
    suffix = parsed.path[len(prefix):]
    if not re.fullmatch(r"v[0-9]+\.[0-9]+\.[0-9]+/dentweb-agent\.exe", suffix):
        raise ValueError("지원하지 않는 에이전트 릴리즈 주소입니다.")
    return download_url.rsplit("/", 1)[0] + "/SHA256SUMS"


def _read_checksum(text: str) -> str:
    matches = [line.split() for line in text.splitlines() if line.strip()]
    valid = [parts[0].lower() for parts in matches if len(parts) == 2
             and parts[1] in ("dentweb-agent.exe", "*dentweb-agent.exe")
             and re.fullmatch(r"[0-9a-fA-F]{64}", parts[0])]
    if len(valid) != 1:
        raise ValueError("유효한 에이전트 체크섬이 없습니다.")
    return valid[0]


def download_update(download_url: str, progress_callback=None,
                    log_callback=None) -> str | None:
    """Download into a temporary file; return it only after SHA-256 verification."""
    def _log(msg):
        if log_callback:
            log_callback(msg)

    download_path = None
    try:
        checksum_response = requests.get(_checksum_url(download_url), timeout=30)
        try:
            checksum_response.raise_for_status()
            expected_hash = _read_checksum(checksum_response.text)
        finally:
            checksum_response.close()

        if getattr(sys, "frozen", False):
            target_dir = os.path.dirname(sys.executable)
        else:
            target_dir = tempfile.gettempdir()
        # Do not erase the last backup before a new file has been verified.
        fd, download_path = tempfile.mkstemp(prefix="dentweb-agent-update-", suffix=".exe", dir=target_dir)
        digest = hashlib.sha256()
        _log("공식 업데이트를 다운로드하고 무결성을 확인합니다.")
        with os.fdopen(fd, "wb") as stream:
            with requests.get(download_url, stream=True, timeout=180, allow_redirects=True) as resp:
                resp.raise_for_status()
                total = int(resp.headers.get("content-length", 0))
                downloaded = 0
                for chunk in resp.iter_content(chunk_size=65536):
                    if chunk:
                        stream.write(chunk)
                        digest.update(chunk)
                        downloaded += len(chunk)
                        if progress_callback and total > 0:
                            progress_callback(min(downloaded / total, 1.0))
                if total > 0 and downloaded != total:
                    raise ValueError("업데이트 파일 크기가 일치하지 않습니다.")
        if not hmac.compare_digest(digest.hexdigest(), expected_hash):
            raise ValueError("업데이트 파일 체크섬이 일치하지 않습니다.")
        _log("업데이트 파일 무결성 확인 완료")
        return download_path
    except Exception as error:
        if download_path and os.path.exists(download_path):
            os.remove(download_path)
        _log(f"다운로드 오류: {error}")
        return None


def apply_update(new_exe_path: str):
    """현재 exe를 교체하고 재시작 (PowerShell 사용)."""
    if not getattr(sys, "frozen", False):
        return

    current_exe = sys.executable
    current_dir = os.path.dirname(current_exe)
    ps_path = os.path.join(current_dir, "dentweb-agent-update.ps1")

    ps_content = _build_update_script(current_exe, new_exe_path, os.getpid())

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


def _build_update_script(current_exe: str, new_exe_path: str, process_id: int) -> str:
    def literal(value):
        return "'" + value.replace("'", "''") + "'"
    current_dir = os.path.dirname(current_exe)
    return f'''
$ErrorActionPreference = "Stop"
$currentExe = {literal(current_exe)}
$newExe = {literal(new_exe_path)}
$oldExe = {literal(os.path.join(current_dir, 'dentweb-agent-old.exe'))}
$agentProcessId = {int(process_id)}

# 1. 현재 프로세스 종료 대기
Write-Host "업데이트 적용 중... 프로세스 종료 대기"
try {{
    Wait-Process -Id $agentProcessId -Timeout 15 -ErrorAction SilentlyContinue
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
