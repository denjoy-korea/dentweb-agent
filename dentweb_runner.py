"""좌표 기반 DentWeb 자동화 시퀀스

전체 흐름:
1. pygetwindow로 '덴트웹' 창 자동 탐색 → 활성화 (최소화 복원 포함)
2. 경영/통계 → 임플란트 수술 통계
3. 특정기간 → 부터 '오늘' → 까지 '오늘' (당일 조회)
4. 데이터 유무 확인 → 없으면 중단
5. 엑셀저장 → 저장 다이얼로그 → 파일 업로드
"""

import os
import json
import time
import glob
import pyautogui
import pyperclip


STEPS_FILE = "dentweb_steps.json"
WINDOW_TITLE = "덴트웹"


def _paste_text(text: str):
    """클립보드에 복사 후 Ctrl+V로 붙여넣기 (한글 지원)"""
    pyperclip.copy(text)
    time.sleep(0.05)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.1)


# --- 클릭 시퀀스 정의 ---

DATA_STEPS = [
    {"name": "stats_menu", "label": "상단 '경영/통계' 아이콘", "x": None, "y": None, "wait_after": 1.5},
    {"name": "implant_tab", "label": "왼쪽 사이드바 '임플란트 수술 통계'", "x": None, "y": None, "wait_after": 1.5},
    {"name": "custom_period", "label": "'특정기간' 라디오 버튼", "x": None, "y": None, "wait_after": 0.5},
    {"name": "date_start_field", "label": "'부터' 날짜 필드 클릭 (달력 열기)", "x": None, "y": None, "wait_after": 0.5},
    {"name": "date_start_today", "label": "'부터' 달력 하단 '오늘' 버튼", "x": None, "y": None, "wait_after": 0.5},
    {"name": "date_end_field", "label": "'까지' 날짜 필드 클릭 (달력 열기)", "x": None, "y": None, "wait_after": 0.5},
    {"name": "date_end_today", "label": "'까지' 달력 하단 '오늘' 버튼 (자동 조회됨)", "x": None, "y": None, "wait_after": 3.0},
    {"name": "data_check", "label": "수술기록목지 첫 번째 행 위치 (데이터 유무 확인용)", "x": None, "y": None, "wait_after": 0, "group": "check"},
    {"name": "export_btn", "label": "'엑셀저장' 버튼", "x": None, "y": None, "wait_after": 2.0},
]


# --- 설정 파일 관리 ---

def load_config_data(path: str = STEPS_FILE) -> dict | None:
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if isinstance(data, list):
        return None

    for step in data.get("data_steps", []):
        if step.get("skip"):
            continue
        if step.get("x") is None or step.get("y") is None:
            return None
    return data


def save_config_data(data: dict, path: str = STEPS_FILE):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


# --- 학습 모드 ---

CAPTURE_DELAY = 5


def _countdown_capture() -> tuple[int, int]:
    """카운트다운 후 마우스 위치 캡처"""
    for remaining in range(CAPTURE_DELAY, 0, -1):
        print(f"\r  → {remaining}초 후 마우스 위치 저장...", end="", flush=True)
        time.sleep(1)
    x, y = pyautogui.position()
    print(f"\r  → 저장됨: ({x}, {y})              ")
    return x, y


def run_teach_mode() -> dict:
    print()
    print("=" * 55)
    print("  덴트웹 자동화 - 클릭 위치 학습")
    print("=" * 55)
    print()
    print("덴트웹이 열린 상태에서 진행합니다.")
    print()
    print("각 단계에서:")
    print("  1. 이 창에서 Enter를 누르면 5초 카운트다운 시작")
    print("  2. 덴트웹으로 전환하여 해당 위치에 마우스를 올림")
    print("  3. 카운트다운이 끝나면 마우스 위치 자동 저장")
    print()
    print("  's' = 단계 건너뛰기 / 'r' = 처음부터")
    print()
    print("순서: 경영/통계 → 임플란트 수술 통계 → 특정기간")
    print("      → 부터(달력 열기 → 오늘) → 까지(달력 열기 → 오늘)")
    print("      → 데이터 확인 위치 → 엑셀저장")
    print()
    input("준비되면 Enter...")

    steps = json.loads(json.dumps(DATA_STEPS))

    i = 0
    while i < len(steps):
        step = steps[i]
        print(f"\n[{i + 1}/{len(steps)}] {step['label']}")

        if step.get("group") == "check":
            print(f"  → 데이터가 표시되는 영역(첫 번째 행)에 마우스를 올려주세요.")
            print(f"    (비어있으면 엑셀저장을 건너뜁니다)")
        else:
            print(f"  → 해당 위치에 마우스를 올려주세요.")

        print(f"  → Enter=카운트다운 시작 / s=건너뛰기 / r=처음부터")
        choice = input("  → ").strip().lower()

        if choice == "r":
            steps = json.loads(json.dumps(DATA_STEPS))
            i = 0
            print("\n처음부터 다시 시작합니다.")
            continue

        if choice == "s":
            print(f"  → '{step['label']}' 건너뜀")
            steps[i]["skip"] = True
            i += 1
            continue

        print(f"  → 지금 덴트웹에서 '{step['label']}' 위에 마우스를 올려주세요!")
        x, y = _countdown_capture()
        steps[i]["x"] = x
        steps[i]["y"] = y
        steps[i]["skip"] = False
        i += 1

    result = {"data_steps": steps}
    save_config_data(result)
    print()
    print(f"설정 저장 완료: {os.path.abspath(STEPS_FILE)}")
    print("다음 실행부터 자동화에 사용됩니다.")
    print()
    return result


# --- 메인 Runner ---

class DentwebRunner:
    def __init__(self, cfg: dict):
        self.download_dir = cfg["download_dir"]
        self.download_timeout = cfg.get("download_timeout_seconds", 30)
        self._data = load_config_data()

    def is_configured(self) -> bool:
        return self._data is not None

    def teach(self):
        self._data = run_teach_mode()

    def _activate_dentweb(self) -> bool:
        """pygetwindow로 '덴트웹' 창을 찾아서 포그라운드로 활성화"""
        try:
            import pygetwindow as gw
            windows = gw.getWindowsWithTitle(WINDOW_TITLE)
            if not windows:
                return False
            win = windows[0]
            if win.isMinimized:
                win.restore()
                time.sleep(0.5)
            win.activate()
            time.sleep(0.5)
            # 최대화되지 않았으면 최대화
            if not win.isMaximized:
                win.maximize()
                time.sleep(0.5)
            return True
        except Exception:
            return False

    def _has_data(self) -> bool:
        """수술기록목지 첫 행의 픽셀을 검사해서 데이터 유무 판별"""
        if not self._data:
            return False

        check_step = None
        for step in self._data.get("data_steps", []):
            if step.get("name") == "data_check" and not step.get("skip"):
                check_step = step
                break
        if not check_step:
            return True

        cx, cy = check_step["x"], check_step["y"]
        region = (cx - 50, cy - 15, 100, 30)
        screenshot = pyautogui.screenshot(region=region)

        dark_pixel_count = 0
        for x in range(screenshot.width):
            for y in range(screenshot.height):
                r, g, b = screenshot.getpixel((x, y))
                if r < 100 and g < 100 and b < 100:
                    dark_pixel_count += 1

        return dark_pixel_count > 20

    def _wait_for_download(self) -> str | None:
        """다운로드 폴더에서 엑셀 파일 감지"""
        target = os.path.join(self.download_dir, "dentweb_export.xlsx")
        deadline = time.time() + self.download_timeout
        before = set(glob.glob(os.path.join(self.download_dir, "*.xlsx")))

        while time.time() < deadline:
            if os.path.exists(target):
                time.sleep(1)
                return target
            current = set(glob.glob(os.path.join(self.download_dir, "*.xlsx")))
            new_files = current - before
            if new_files:
                time.sleep(1)
                return max(new_files, key=os.path.getmtime)
            time.sleep(1)
        return None

    def download_excel(self) -> str | None:
        """전체 자동화: 덴트웹 활성화 → 데이터 추출 → Excel 반환"""
        if not self._data:
            return None

        # 1. 덴트웹 창 자동 탐색 + 활성화
        if not self._activate_dentweb():
            return None

        # 2. 날짜 선택 시퀀스 (엑셀저장 전까지)
        for step in self._data.get("data_steps", []):
            if step.get("skip"):
                continue
            if step.get("group") == "check":
                continue
            if step.get("name") == "export_btn":
                break
            pyautogui.click(step["x"], step["y"])
            time.sleep(step.get("wait_after", 0.5))

        # 3. 데이터 유무 확인
        if not self._has_data():
            return None

        # 4. 엑셀저장 클릭
        export_step = None
        for step in self._data.get("data_steps", []):
            if step.get("name") == "export_btn" and not step.get("skip"):
                export_step = step
                break
        if not export_step:
            return None

        pyautogui.click(export_step["x"], export_step["y"])
        time.sleep(export_step.get("wait_after", 2.0))

        # 5. "다른 이름으로 저장" 다이얼로그 처리
        save_path = os.path.join(self.download_dir, "dentweb_export.xlsx")
        pyautogui.hotkey("alt", "n")  # 파일 이름 필드 포커스
        time.sleep(0.3)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.1)
        _paste_text(save_path)
        time.sleep(0.3)
        pyautogui.press("enter")

        # 6. 덮어쓰기 확인 → Enter
        time.sleep(1)
        pyautogui.press("enter")

        # 7. 파일 저장 대기
        return self._wait_for_download()

    def cleanup(self, file_path: str):
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
        except OSError:
            pass
