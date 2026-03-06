"""좌표 기반 DentWeb 자동화 시퀀스

첫 실행 시 사용자가 덴트웹 화면에서 각 버튼 위치를 직접 클릭하여 좌표를 학습시키고,
이후 자동화에서는 저장된 좌표를 사용합니다.

전체 흐름:
1. 덴트웹 실행 확인 (작업표시줄) → 없으면 프로그램 실행
2. 로그인 (ID/PW)
3. 통계 메뉴 → 임플란트 → 특정기간 → 날짜 입력 → 엑셀 다운로드
"""

import os
import json
import time
import glob
import subprocess
from datetime import datetime, timedelta
import pyautogui
import pyperclip


STEPS_FILE = "dentweb_steps.json"


def _paste_text(text: str):
    """클립보드에 복사 후 Ctrl+V로 붙여넣기 (한글 지원)"""
    pyperclip.copy(text)
    time.sleep(0.05)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.1)

# 클릭 시퀀스: 로그인 + 데이터 추출
LAUNCH_STEPS = [
    {"name": "taskbar_icon", "label": "작업표시줄의 덴트웹 아이콘 (실행 중일 때)", "x": None, "y": None, "wait_after": 1.0, "group": "launch"},
    {"name": "login_id", "label": "로그인 ID 입력 필드", "x": None, "y": None, "wait_after": 0.3, "type": "text_input", "group": "login"},
    {"name": "login_pw", "label": "로그인 비밀번호 입력 필드", "x": None, "y": None, "wait_after": 0.3, "type": "password_input", "group": "login"},
    {"name": "login_btn", "label": "로그인 버튼", "x": None, "y": None, "wait_after": 3.0, "group": "login"},
]

DATA_STEPS = [
    {"name": "stats_menu", "label": "상단 '경영/통계' 아이콘", "x": None, "y": None, "wait_after": 1.5, "group": "data"},
    {"name": "implant_tab", "label": "왼쪽 사이드바 '임플란트 수술 통계'", "x": None, "y": None, "wait_after": 1.5, "group": "data"},
    {"name": "custom_period", "label": "'특정기간' 라디오 버튼", "x": None, "y": None, "wait_after": 0.5, "group": "data"},
    {"name": "date_start", "label": "'부터' 날짜 입력 필드", "x": None, "y": None, "wait_after": 0.3, "type": "date_input", "group": "data"},
    {"name": "date_end", "label": "'까지' 날짜 입력 필드 (입력 후 자동 조회)", "x": None, "y": None, "wait_after": 3.0, "type": "date_input", "group": "data"},
    {"name": "export_btn", "label": "'엑셀저장' 버튼", "x": None, "y": None, "wait_after": 1.0, "group": "data"},
]


def load_config_data(path: str = STEPS_FILE) -> dict | None:
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # v1.1 호환
    if isinstance(data, list):
        data = {"launch_steps": [], "data_steps": data, "meta": {"date_format": "%Y-%m-%d", "period_days": 30}}

    # 필수 좌표 검증 (data_steps는 반드시 있어야 함)
    for step in data.get("data_steps", []):
        if step.get("skip"):
            continue
        if step.get("x") is None or step.get("y") is None:
            return None
    return data


def save_config_data(data: dict, path: str = STEPS_FILE):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def _teach_steps(steps: list[dict], group_label: str) -> list[dict]:
    """공통 좌표 학습 루프"""
    print(f"\n--- {group_label} ---\n")
    result = json.loads(json.dumps(steps))

    i = 0
    while i < len(result):
        step = result[i]
        step_type = step.get("type", "click")
        print(f"[{i + 1}/{len(result)}] {step['label']}")

        if step_type in ("text_input", "password_input"):
            print(f"  → 입력 필드를 클릭할 위치에 마우스를 올려주세요.")
        elif step_type == "date_input":
            print(f"  → 날짜 입력 필드 위에 마우스를 올려주세요.")
        else:
            print(f"  → 해당 버튼 위에 마우스를 올려주세요.")

        choice = input("  → Enter=저장 / s=건너뛰기 / r=처음부터: ").strip().lower()

        if choice == "r":
            result = json.loads(json.dumps(steps))
            i = 0
            print("\n처음부터 다시 시작합니다.\n")
            continue

        if choice == "s":
            print(f"  → '{step['label']}' 건너뜀")
            result[i]["skip"] = True
            i += 1
            continue

        x, y = pyautogui.position()
        result[i]["x"] = x
        result[i]["y"] = y
        result[i]["skip"] = False
        print(f"  → 저장됨: ({x}, {y})")
        print()
        i += 1

    return result


def run_teach_mode() -> dict:
    """전체 좌표 학습 모드"""
    print()
    print("=" * 55)
    print("  덴트웹 자동화 설정")
    print("=" * 55)
    print()
    print("덴트웹 화면의 각 버튼/필드 위치를 학습합니다.")
    print("각 단계에서 해당 위치에 마우스를 올린 뒤 Enter를 누르세요.")
    print()
    print("* 's' = 단계 건너뛰기")
    print("* 'r' = 해당 그룹 처음부터")
    print()

    # --- 덴트웹 실행 경로 ---
    print("=" * 55)
    print("  1. 덴트웹 프로그램 경로")
    print("=" * 55)
    print()
    print("덴트웹 바로가기 또는 exe 경로를 입력하세요.")
    print("예: C:\\Users\\Public\\Desktop\\덴트웹.lnk")
    print("    또는 빈칸 (이미 실행 중인 경우만 자동화)")
    print()
    exe_path = input("경로: ").strip().strip('"')
    dentweb_window_title = input("덴트웹 창 제목 (기본: DentWeb): ").strip() or "DentWeb"
    print()

    # --- 작업표시줄/로그인 ---
    print("=" * 55)
    print("  2. 덴트웹 활성화 + 로그인")
    print("=" * 55)
    print()
    print("덴트웹을 열고 로그인 화면까지 준비해주세요.")
    input("준비되면 Enter...")

    launch_steps = _teach_steps(LAUNCH_STEPS, "프로그램 활성화 + 로그인")

    # 로그인 자격증명
    login_id = ""
    login_pw = ""
    has_login = any(
        not s.get("skip") and s.get("type") in ("text_input", "password_input")
        for s in launch_steps
    )
    if has_login:
        print()
        print("자동 로그인에 사용할 계정 정보를 입력하세요.")
        print("(config.json에 저장됩니다)")
        login_id = input("  덴트웹 ID: ").strip()
        login_pw = input("  덴트웹 PW: ").strip()

    # --- 데이터 추출 시퀀스 ---
    print()
    print("=" * 55)
    print("  3. 데이터 추출 (통계 → 엑셀 다운로드)")
    print("=" * 55)
    print()
    print("로그인 후 메인 화면이 보이는 상태에서 진행합니다.")
    input("준비되면 Enter...")

    data_steps = _teach_steps(DATA_STEPS, "데이터 추출 시퀀스")

    # --- 날짜 형식 ---
    print()
    print("날짜 입력 형식을 선택하세요:")
    print("  1) YYYY-MM-DD  (예: 2026-03-01)")
    print("  2) YYYYMMDD    (예: 20260301)")
    print("  3) YYYY/MM/DD  (예: 2026/03/01)")
    fmt_choice = input("선택 (1/2/3, 기본=1): ").strip()
    date_format = {"2": "%Y%m%d", "3": "%Y/%m/%d"}.get(fmt_choice, "%Y-%m-%d")

    period = input("조회 기간 (일수, 기본=30): ").strip()
    try:
        period_days = int(period) if period else 30
    except ValueError:
        period_days = 30

    result = {
        "launch_steps": launch_steps,
        "data_steps": data_steps,
        "meta": {
            "exe_path": exe_path,
            "window_title": dentweb_window_title,
            "login_id": login_id,
            "login_pw": login_pw,
            "date_format": date_format,
            "period_days": period_days,
        },
    }

    save_config_data(result)
    print()
    print(f"설정 저장 완료: {os.path.abspath(STEPS_FILE)}")
    print(f"  날짜 형식: {date_format} / 조회 기간: {period_days}일")
    print()
    return result


class DentwebRunner:
    def __init__(self, cfg: dict):
        self.download_dir = cfg["download_dir"]
        self.download_timeout = cfg.get("download_timeout_seconds", 30)
        self._data = load_config_data()
        # cfg에서 window_title 기본값
        self._default_window_title = cfg.get("dentweb_window_title", "덴트웹")

    def is_configured(self) -> bool:
        return self._data is not None

    def teach(self):
        self._data = run_teach_mode()

    @property
    def meta(self) -> dict:
        if not self._data:
            return {}
        return self._data.get("meta", {})

    @property
    def window_title(self) -> str:
        return self.meta.get("window_title", self._default_window_title)

    def _is_dentweb_running(self) -> bool:
        """덴트웹 창이 존재하는지 확인"""
        try:
            import pygetwindow as gw
            windows = gw.getWindowsWithTitle(self.window_title)
            return len(windows) > 0
        except Exception:
            return False

    def _activate_dentweb(self) -> bool:
        """덴트웹 창을 포그라운드로 활성화"""
        try:
            import pygetwindow as gw
            windows = gw.getWindowsWithTitle(self.window_title)
            if not windows:
                return False
            win = windows[0]
            if win.isMinimized:
                win.restore()
            win.activate()
            time.sleep(0.5)
            return True
        except Exception:
            return False

    def _launch_dentweb(self) -> bool:
        """덴트웹 프로그램 실행"""
        exe_path = self.meta.get("exe_path", "")
        if not exe_path:
            return False
        try:
            subprocess.Popen(exe_path, shell=True)
            # 프로그램 로딩 대기
            deadline = time.time() + 30
            while time.time() < deadline:
                if self._is_dentweb_running():
                    time.sleep(2)  # 창이 완전히 로드될 때까지 추가 대기
                    return True
                time.sleep(1)
            return False
        except Exception:
            return False

    def _click_taskbar_icon(self) -> bool:
        """작업표시줄 아이콘 클릭으로 덴트웹 활성화"""
        if not self._data:
            return False
        for step in self._data.get("launch_steps", []):
            if step.get("name") == "taskbar_icon" and not step.get("skip"):
                pyautogui.click(step["x"], step["y"])
                time.sleep(step.get("wait_after", 1.0))
                return True
        return False

    def _do_login(self) -> bool:
        """로그인 수행"""
        if not self._data:
            return False

        login_id = self.meta.get("login_id", "")
        login_pw = self.meta.get("login_pw", "")

        for step in self._data.get("launch_steps", []):
            if step.get("skip"):
                continue
            if step.get("group") != "login":
                continue

            if step.get("type") == "text_input" and login_id:
                pyautogui.click(step["x"], step["y"])
                time.sleep(0.2)
                pyautogui.hotkey("ctrl", "a")
                time.sleep(0.1)
                _paste_text(login_id)
            elif step.get("type") == "password_input" and login_pw:
                pyautogui.click(step["x"], step["y"])
                time.sleep(0.2)
                pyautogui.hotkey("ctrl", "a")
                time.sleep(0.1)
                _paste_text(login_pw)
            else:
                # 로그인 버튼 등 일반 클릭
                pyautogui.click(step["x"], step["y"])

            time.sleep(step.get("wait_after", 0.5))

        return True

    def _type_date(self, step: dict, date_str: str):
        """날짜 필드 클릭 후 날짜 타이핑"""
        pyautogui.click(step["x"], step["y"])
        time.sleep(0.2)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.1)
        _paste_text(date_str)
        time.sleep(0.1)
        pyautogui.press("enter")

    def _wait_for_download(self) -> str | None:
        """다운로드 폴더에서 최신 xlsx 파일 감지"""
        deadline = time.time() + self.download_timeout
        before = set(glob.glob(os.path.join(self.download_dir, "*.xlsx")))

        while time.time() < deadline:
            current = set(glob.glob(os.path.join(self.download_dir, "*.xlsx")))
            new_files = current - before
            if new_files:
                return max(new_files, key=os.path.getmtime)
            time.sleep(1)
        return None

    def ensure_dentweb_ready(self) -> bool:
        """덴트웹이 실행 중이고 로그인된 상태인지 확인 + 필요 시 실행/로그인"""
        # 1. 이미 실행 중이면 활성화
        if self._is_dentweb_running():
            if self._activate_dentweb():
                return True
            # activate 실패 시 작업표시줄 클릭 시도
            self._click_taskbar_icon()
            time.sleep(1)
            return self._is_dentweb_running()

        # 2. 실행 중이 아니면 프로그램 실행
        exe_path = self.meta.get("exe_path", "")
        if exe_path:
            if not self._launch_dentweb():
                return False
            # 3. 로그인
            time.sleep(2)
            self._do_login()
            return True

        # 3. 작업표시줄 아이콘 클릭 시도 (exe_path 없을 때)
        self._click_taskbar_icon()
        time.sleep(2)
        if self._is_dentweb_running():
            return True

        return False

    def download_excel(self) -> str | None:
        """전체 자동화: 덴트웹 준비 → 데이터 추출 → Excel 반환"""
        if not self._data:
            return None

        # 1. 덴트웹 실행/활성화
        if not self.ensure_dentweb_ready():
            return None

        # 2. 데이터 추출 시퀀스
        date_fmt = self.meta.get("date_format", "%Y-%m-%d")
        period_days = self.meta.get("period_days", 30)
        end_date = datetime.now()
        start_date = end_date - timedelta(days=period_days)
        start_str = start_date.strftime(date_fmt)
        end_str = end_date.strftime(date_fmt)

        for step in self._data.get("data_steps", []):
            if step.get("skip"):
                continue

            if step.get("type") == "date_input":
                date_str = start_str if step["name"] == "date_start" else end_str
                self._type_date(step, date_str)
            else:
                pyautogui.click(step["x"], step["y"])

            time.sleep(step.get("wait_after", 0.5))

        return self._wait_for_download()

    def cleanup(self, file_path: str):
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
        except OSError:
            pass
