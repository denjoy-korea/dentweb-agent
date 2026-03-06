"""좌표 기반 DentWeb 자동화 시퀀스

첫 실행 시 사용자가 덴트웹 화면에서 각 버튼 위치를 직접 클릭하여 좌표를 학습시키고,
이후 자동화에서는 저장된 좌표를 사용합니다.
"""

import os
import json
import time
import glob
from datetime import datetime, timedelta
import pyautogui


STEPS_FILE = "dentweb_steps.json"

# 기본 클릭 시퀀스 정의 (사용자가 좌표를 가르쳐야 함)
DEFAULT_STEPS = [
    {"name": "stats_menu", "label": "경영/통계 메뉴 버튼", "x": None, "y": None, "wait_after": 1.0},
    {"name": "implant_tab", "label": "임플란트 수술통계 탭", "x": None, "y": None, "wait_after": 1.0},
    {"name": "custom_period", "label": "'특정기간' 선택 버튼", "x": None, "y": None, "wait_after": 0.5},
    {"name": "date_start", "label": "시작일 입력 필드", "x": None, "y": None, "wait_after": 0.3, "type": "date_input"},
    {"name": "date_end", "label": "종료일 입력 필드 (입력 후 자동 조회됨)", "x": None, "y": None, "wait_after": 2.0, "type": "date_input"},
    {"name": "export_btn", "label": "엑셀 다운로드 버튼", "x": None, "y": None, "wait_after": 1.0},
]


def load_steps(path: str = STEPS_FILE) -> list[dict] | None:
    """저장된 클릭 좌표 불러오기"""
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        steps = json.load(f)
    for step in steps:
        if step.get("skip"):
            continue
        if step.get("x") is None or step.get("y") is None:
            return None
    return steps


def save_steps(steps: list[dict], path: str = STEPS_FILE):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(steps, f, indent=2, ensure_ascii=False)


def run_teach_mode() -> list[dict]:
    """좌표 학습 모드: 사용자가 각 단계에서 마우스로 위치를 지정"""
    print()
    print("=" * 55)
    print("  덴트웹 클릭 좌표 학습 모드")
    print("=" * 55)
    print()
    print("덴트웹 프로그램을 열어주세요.")
    print("각 단계에서 해당 버튼/필드 위에 마우스를 올린 뒤")
    print("Enter를 누르면 현재 마우스 위치가 저장됩니다.")
    print()
    print("* 단계를 건너뛰려면 's'를 입력하세요.")
    print("* 처음부터 다시 하려면 'r'을 입력하세요.")
    print()
    input("준비되면 Enter를 누르세요...")
    print()

    steps = json.loads(json.dumps(DEFAULT_STEPS))  # deep copy

    i = 0
    while i < len(steps):
        step = steps[i]
        print(f"[{i + 1}/{len(steps)}] {step['label']}")
        if step.get("type") == "date_input":
            print(f"  → 날짜 입력 필드를 클릭할 위치에 마우스를 올려주세요.")
            print(f"    (에이전트가 자동으로 날짜를 타이핑합니다)")
        else:
            print(f"  → 덴트웹에서 '{step['label']}' 위에 마우스를 올려주세요.")
        choice = input("  → Enter=위치 저장 / s=건너뛰기 / r=처음부터: ").strip().lower()

        if choice == "r":
            steps = json.loads(json.dumps(DEFAULT_STEPS))
            i = 0
            print("\n처음부터 다시 시작합니다.\n")
            continue

        if choice == "s":
            print(f"  → '{step['label']}' 건너뜀")
            steps[i]["skip"] = True
            i += 1
            continue

        x, y = pyautogui.position()
        steps[i]["x"] = x
        steps[i]["y"] = y
        steps[i]["skip"] = False
        print(f"  → 저장됨: ({x}, {y})")
        print()
        i += 1

    # 날짜 형식 확인
    print()
    print("날짜 입력 형식을 선택하세요:")
    print("  1) YYYY-MM-DD  (예: 2026-03-01)")
    print("  2) YYYYMMDD    (예: 20260301)")
    print("  3) YYYY/MM/DD  (예: 2026/03/01)")
    fmt_choice = input("선택 (1/2/3, 기본=1): ").strip()
    date_format = {
        "2": "%Y%m%d",
        "3": "%Y/%m/%d",
    }.get(fmt_choice, "%Y-%m-%d")

    # 조회 기간 (며칠치)
    print()
    period = input("조회 기간 (일수, 기본=30): ").strip()
    try:
        period_days = int(period) if period else 30
    except ValueError:
        period_days = 30

    # 메타 정보 저장
    meta = {
        "date_format": date_format,
        "period_days": period_days,
    }

    result = {"steps": steps, "meta": meta}
    save_steps(result)

    print()
    print(f"설정이 저장되었습니다: {os.path.abspath(STEPS_FILE)}")
    print(f"  날짜 형식: {date_format}")
    print(f"  조회 기간: 최근 {period_days}일")
    print("다음 실행부터 자동화에 사용됩니다.")
    print()
    return result


def load_config_data(path: str = STEPS_FILE) -> dict | None:
    """저장된 설정 전체 불러오기 (steps + meta)"""
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # v1.1.0 호환: list 형태면 새 형식으로 변환
    if isinstance(data, list):
        data = {"steps": data, "meta": {"date_format": "%Y-%m-%d", "period_days": 30}}

    steps = data.get("steps", [])
    for step in steps:
        if step.get("skip"):
            continue
        if step.get("x") is None or step.get("y") is None:
            return None
    return data


class DentwebRunner:
    def __init__(self, cfg: dict):
        self.window_title = cfg.get("dentweb_window_title", "DentWeb")
        self.download_dir = cfg["download_dir"]
        self.download_timeout = cfg.get("download_timeout_seconds", 30)
        self._data = load_config_data()

    def is_configured(self) -> bool:
        return self._data is not None

    def teach(self):
        self._data = run_teach_mode()

    @property
    def steps(self) -> list[dict]:
        if not self._data:
            return []
        return self._data.get("steps", [])

    @property
    def meta(self) -> dict:
        if not self._data:
            return {"date_format": "%Y-%m-%d", "period_days": 30}
        return self._data.get("meta", {"date_format": "%Y-%m-%d", "period_days": 30})

    def _activate_dentweb(self) -> bool:
        """DentWeb 창을 포그라운드로 활성화"""
        try:
            import pygetwindow as gw

            windows = gw.getWindowsWithTitle(self.window_title)
            if not windows:
                return False
            windows[0].activate()
            time.sleep(0.5)
            return True
        except Exception:
            return False

    def _type_date(self, step: dict, date_str: str):
        """날짜 필드 클릭 후 날짜 타이핑"""
        pyautogui.click(step["x"], step["y"])
        time.sleep(0.2)
        # 기존 값 전체 선택 후 덮어쓰기
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.1)
        pyautogui.typewrite(date_str, interval=0.05)
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

    def download_excel(self) -> str | None:
        """DentWeb 자동화 시퀀스 -> Excel 파일 경로 반환"""
        if not self.steps:
            return None

        if not self._activate_dentweb():
            return None

        # 날짜 계산
        date_fmt = self.meta["date_format"]
        period_days = self.meta["period_days"]
        end_date = datetime.now()
        start_date = end_date - timedelta(days=period_days)
        start_str = start_date.strftime(date_fmt)
        end_str = end_date.strftime(date_fmt)

        for step in self.steps:
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
