"""좌표 기반 DentWeb 자동화 시퀀스

전체 흐름:
1. pygetwindow로 '덴트웹' 창 자동 탐색 → 활성화 (최소화 복원 포함)
2. 경영/통계 → 임플란트 수술 통계
3. 특정기간 → 부터 '오늘' → 까지 '오늘' (당일 조회)
4. 엑셀저장 → 저장 다이얼로그 → 파일 저장
5. "수술기록지" 시트 2행 이하 데이터 유무로 판별
"""

import os
import json
import time
import glob
import pyautogui
import pyperclip
from openpyxl import load_workbook


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
    {"name": "stats_menu", "label": "상단 '경영/통계' 아이콘", "x": None, "y": None, "wait_after": 3.0},
    {"name": "implant_tab", "label": "왼쪽 사이드바 '임플란트 수술 통계'", "x": None, "y": None, "wait_after": 3.0},
    {"name": "custom_period", "label": "'특정기간' 라디오 버튼", "x": None, "y": None, "wait_after": 1.5},
    {"name": "date_start_field", "label": "'부터' 날짜 필드 클릭 (달력 열기)", "x": None, "y": None, "wait_after": 1.5},
    {"name": "date_start_today", "label": "'부터' 달력 하단 '오늘' 버튼", "x": None, "y": None, "wait_after": 1.5},
    {"name": "date_end_field", "label": "'까지' 날짜 필드 클릭 (달력 열기)", "x": None, "y": None, "wait_after": 1.5},
    {"name": "date_end_today", "label": "'까지' 달력 하단 '오늘' 버튼 (자동 조회됨)", "x": None, "y": None, "wait_after": 5.0},
    {"name": "export_btn", "label": "'엑셀저장' 버튼", "x": None, "y": None, "wait_after": 3.0},
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
        # data_check는 더 이상 사용하지 않으므로 무시
        if step.get("name") == "data_check":
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
    print("      → 엑셀저장")
    print()
    input("준비되면 Enter...")

    steps = json.loads(json.dumps(DATA_STEPS))

    i = 0
    while i < len(steps):
        step = steps[i]
        print(f"\n[{i + 1}/{len(steps)}] {step['label']}")
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

    def _activate_dentweb(self, log_callback=None) -> bool:
        """'덴트웹' 창을 찾아서 포그라운드로 활성화"""
        def _log(msg):
            if log_callback:
                log_callback(msg)

        # 방법 1: pygetwindow
        try:
            import pygetwindow as gw
            all_windows = gw.getAllTitles()
            _log(f"열린 창 {len(all_windows)}개 탐색 중...")

            # 부분 일치로 '덴트웹' 포함된 창 찾기
            matched = [t for t in all_windows if WINDOW_TITLE in t]
            if not matched:
                _log(f"'{WINDOW_TITLE}' 포함된 창 없음")
                # 디버깅: 한글 포함 창 목록 표시
                korean_wins = [t for t in all_windows if t.strip()]
                _log(f"전체 창 목록: {korean_wins[:10]}")
            else:
                _log(f"찾은 창: {matched[0]}")
                windows = gw.getWindowsWithTitle(matched[0])
                if windows:
                    win = windows[0]
                    if win.isMinimized:
                        win.restore()
                        time.sleep(0.5)
                    win.activate()
                    time.sleep(0.5)
                    if not win.isMaximized:
                        win.maximize()
                        time.sleep(0.5)
                    return True
        except Exception as e:
            _log(f"pygetwindow 오류: {e}")

        # 방법 2: win32gui 직접 사용 (fallback)
        try:
            import ctypes
            import ctypes.wintypes

            def _find_window_callback(hwnd, results):
                if ctypes.windll.user32.IsWindowVisible(hwnd):
                    length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
                    if length > 0:
                        buf = ctypes.create_unicode_buffer(length + 1)
                        ctypes.windll.user32.GetWindowTextW(hwnd, buf, length + 1)
                        if WINDOW_TITLE in buf.value:
                            results.append(hwnd)
                return True

            WNDENUMPROC = ctypes.WINFUNCTYPE(
                ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM
            )
            results = []
            ctypes.windll.user32.EnumWindows(
                WNDENUMPROC(lambda hwnd, _: _find_window_callback(hwnd, results) or True),
                0,
            )

            if results:
                hwnd = results[0]
                _log(f"win32gui로 창 발견 (hwnd={hwnd})")
                # SW_RESTORE = 9, SW_MAXIMIZE = 3
                ctypes.windll.user32.ShowWindow(hwnd, 9)
                time.sleep(0.3)
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                time.sleep(0.3)
                ctypes.windll.user32.ShowWindow(hwnd, 3)
                time.sleep(0.5)
                return True
            else:
                _log("win32gui로도 창을 찾을 수 없습니다")
        except Exception as e:
            _log(f"win32gui 오류: {e}")

        return False

    @staticmethod
    def excel_has_data(file_path: str) -> bool:
        """수술기록지 시트의 2행 이하에 데이터가 있는지 확인"""
        try:
            wb = load_workbook(file_path, read_only=True, data_only=True)

            # "수술기록지" 시트 찾기
            sheet = None
            for name in wb.sheetnames:
                if "수술기록지" in name:
                    sheet = wb[name]
                    break

            if sheet is None:
                wb.close()
                return False

            # 2행부터 데이터 확인 (A열 기준)
            has_data = False
            for row in sheet.iter_rows(min_row=2, max_row=2, max_col=1, values_only=True):
                if row[0] is not None and str(row[0]).strip():
                    has_data = True
                    break

            wb.close()
            return has_data
        except Exception:
            return False

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

    def download_excel(self, log_callback=None) -> str | None:
        """전체 자동화: 덴트웹 활성화 → 엑셀 저장 → 데이터 유무 판별

        Args:
            log_callback: 진행 상황 로그 콜백 (GUI 실시간 표시용)

        Returns:
            파일 경로 (데이터 있음) / None (데이터 없음 또는 실패)
        """
        def _log(msg):
            if log_callback:
                log_callback(msg)

        if not self._data:
            _log("설정 데이터 없음")
            return None

        # 1. 덴트웹 창 자동 탐색 + 활성화
        _log("덴트웹 창 탐색 중...")
        if not self._activate_dentweb(log_callback=log_callback):
            _log("덴트웹 창을 찾을 수 없습니다")
            return None
        _log("덴트웹 창 활성화 완료 — 2초 대기")
        time.sleep(2)  # 창 활성화 후 안정화 대기

        # 2. 날짜 선택 시퀀스 (엑셀저장 전까지)
        for step in self._data.get("data_steps", []):
            if step.get("skip"):
                continue
            if step.get("name") == "data_check":
                continue  # 레거시 호환: 이전 설정에 data_check이 있어도 무시
            if step.get("name") == "export_btn":
                break
            _log(f"클릭: {step['label']} ({step['x']}, {step['y']})")
            pyautogui.click(step["x"], step["y"])
            time.sleep(step.get("wait_after", 0.5))

        # 3. 엑셀저장 클릭 (항상 실행)
        export_step = None
        for step in self._data.get("data_steps", []):
            if step.get("name") == "export_btn" and not step.get("skip"):
                export_step = step
                break
        if not export_step:
            _log("엑셀저장 버튼 좌표 없음")
            return None

        _log(f"클릭: {export_step['label']} ({export_step['x']}, {export_step['y']})")
        pyautogui.click(export_step["x"], export_step["y"])
        time.sleep(export_step.get("wait_after", 2.0))

        # 4. "다른 이름으로 저장" 다이얼로그 처리
        save_path = os.path.join(self.download_dir, "dentweb_export.xlsx")
        _log("저장 다이얼로그 처리 중...")
        pyautogui.hotkey("alt", "n")  # 파일 이름 필드 포커스
        time.sleep(0.5)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.3)
        _paste_text(save_path)
        time.sleep(0.5)
        pyautogui.press("enter")

        # 5. 덮어쓰기 확인 → Enter
        time.sleep(2)
        pyautogui.press("enter")

        # 6. 파일 저장 대기
        _log("엑셀 파일 저장 대기 중...")
        excel_path = self._wait_for_download()
        if not excel_path:
            _log("엑셀 파일 저장 시간 초과")
            return None
        _log(f"엑셀 파일 저장 완료: {os.path.basename(excel_path)}")

        # 7. "수술기록지" 시트 2행 데이터 확인
        _log("수술기록지 데이터 확인 중...")
        if not self.excel_has_data(excel_path):
            _log("수술기록지에 데이터 없음 (빈 파일)")
            self.cleanup(excel_path)
            return None

        _log("수술 기록 데이터 확인 완료")
        return excel_path

    def cleanup(self, file_path: str):
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
        except OSError:
            pass
