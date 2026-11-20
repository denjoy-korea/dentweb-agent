"""이미지 인식 기반 DentWeb 자동화 시퀀스

전체 흐름:
1. 덴트웹 창 자동 탐색 → 활성화 (최소화 복원 포함)
2. 경영/통계 → 임플란트 수술 통계
3. 특정기간 → 부터 '오늘' → 까지 '오늘' (당일 조회)
4. 엑셀저장 → 저장 다이얼로그 → 파일 저장
5. "수술기록지" 시트 2행 이하 데이터 유무로 판별

학습 모드에서 각 버튼의 스크린샷을 캡처하고,
자동화 시 pyautogui.locateCenterOnScreen()으로 이미지를 찾아 클릭.
"""

import os
import sys
import json
import time
import glob
import pyautogui
import pyperclip
from PIL import Image
from openpyxl import load_workbook
from export_store import ExportStore, legacy_exports_directory, remove_owned


STEPS_FILE = "dentweb_steps.json"
TEMPLATES_DIR = "templates"
WINDOW_TITLE = "덴트웹"

# 템플릿 캡처 크기 (마우스 위치 중심으로 캡처)
CAPTURE_W = 120
CAPTURE_H = 50


def _paste_text(text: str):
    """클립보드에 복사 후 Ctrl+V로 붙여넣기 (한글 지원)"""
    pyperclip.copy(text)
    time.sleep(0.05)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.1)


def _win32_click(x: int, y: int, activate_first: bool = False):
    """왼쪽 클릭. activate_first=True면 활성화용 클릭 1회 추가"""
    ix, iy = int(x), int(y)
    if activate_first:
        pyautogui.click(ix, iy, button="left")  # 창 활성화용
        time.sleep(0.3)
    pyautogui.click(ix, iy, button="left")  # 실제 클릭
    time.sleep(0.1)


class ExportFailed(RuntimeError):
    """덴트웹 자동화가 실패했다(데이터 없음과 구별해 서버에 'failed'로 보고한다)."""


_DIALOG_EDIT_CLASSES = {"Edit", "ComboBoxEx32"}


def _foreground_is_save_dialog() -> bool:
    """맨 앞 창이 파일 이름 칸이 있는 표준 대화상자(#32770)인지 본다.

    덴트웹 알림 창도 #32770이지만 입력 칸이 없다. 저장 창이 아닌데 Ctrl+A·붙여넣기를 보내면
    덴트웹의 다른 입력란 내용을 바꿀 수 있으므로, 이 확인을 통과할 때만 키를 보낸다.
    """
    if sys.platform != "win32":
        return True
    import ctypes
    from ctypes import wintypes
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return False

    def class_name(handle) -> str:
        buf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(handle, buf, 256)
        return buf.value

    if class_name(hwnd) != "#32770":
        return False
    found = []
    enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def visit(child, _lparam):
        if class_name(child) in _DIALOG_EDIT_CLASSES:
            found.append(child)
            return False
        return True

    user32.EnumChildWindows(hwnd, enum_proc(visit), 0)
    return bool(found)


def _wait_for_save_dialog(timeout: float = 10.0) -> bool:
    deadline = time.monotonic() + timeout
    while True:
        if _foreground_is_save_dialog():
            return True
        if time.monotonic() >= deadline:
            return False
        time.sleep(0.3)


# --- 클릭 시퀀스 정의 ---

DATA_STEPS = [
    {"name": "dentweb_open", "label": "덴트웹 창 열기 — 작업표시줄의 덴트웹 아이콘 클릭", "x": None, "y": None, "wait_after": 2.0},
    {"name": "stats_menu", "label": "상단 '경영/통계' 아이콘", "x": None, "y": None, "wait_after": 3.0},
    {"name": "implant_tab", "label": "왼쪽 사이드바 '임플란트 수술 통계'", "x": None, "y": None, "wait_after": 3.0},
    {"name": "custom_period", "label": "'특정기간' 라디오 버튼", "x": None, "y": None, "wait_after": 1.5},
    {"name": "date_start_field", "label": "'부터' 날짜 필드 클릭 (달력 열기)", "x": None, "y": None, "wait_after": 1.5},
    {"name": "date_start_today", "label": "'부터' 달력 하단 '오늘' 버튼", "x": None, "y": None, "wait_after": 1.5},
    {"name": "date_end_field", "label": "'까지' 날짜 필드 클릭 (달력 열기)", "x": None, "y": None, "wait_after": 1.5},
    {"name": "date_end_today", "label": "'까지' 달력 하단 '오늘' 버튼 (자동 조회됨)", "x": None, "y": None, "wait_after": 5.0},
    {"name": "export_btn", "label": "'엑셀저장' 버튼", "x": None, "y": None, "wait_after": 3.0},
    # ── 저장 다이얼로그 단계 ──────────────────────────────────────
    # 폴더 이동과 파일 이름은 키보드로 자동 처리 (좌표 불필요)
    {"name": "save_dialog_save_btn",    "label": "저장 창 — '저장(S)' 버튼", "x": None, "y": None, "wait_after": 2.0},
]

# 3.6.5부터 쓰지 않는 단계. 실행마다 고유한 파일 이름으로 저장하므로 덮어쓰기 확인 창이 뜨지 않는다.
# 창이 없을 때 저장된 좌표를 누르면 그 아래 덴트웹 화면을 누르게 되므로, 예전 설정에 있어도 읽지 않는다.
RETIRED_STEPS = {"save_dialog_confirm_yes"}


def _get_template_path(step_name: str) -> str:
    return os.path.join(TEMPLATES_DIR, f"{step_name}.png")


def _capture_template(x: int, y: int, step_name: str) -> str:
    """마우스 위치 주변 영역을 스크린샷으로 캡처하여 템플릿 저장"""
    os.makedirs(TEMPLATES_DIR, exist_ok=True)

    # 캡처 영역 계산 (화면 경계 처리)
    screen_w, screen_h = pyautogui.size()
    left = max(0, x - CAPTURE_W // 2)
    top = max(0, y - CAPTURE_H // 2)
    right = min(screen_w, left + CAPTURE_W)
    bottom = min(screen_h, top + CAPTURE_H)
    w = right - left
    h = bottom - top

    screenshot = pyautogui.screenshot(region=(left, top, w, h))
    path = _get_template_path(step_name)
    screenshot.save(path)
    return path


def _find_and_click(step: dict, log_callback=None, confidence: float = 0.8,
                    activate_first: bool = False) -> bool:
    """이미지 인식으로 버튼을 찾아 클릭. 실패 시 좌표 폴백."""
    def _log(msg):
        if log_callback:
            log_callback(msg)

    template_path = _get_template_path(step["name"])

    # 1차: 이미지 인식
    if os.path.exists(template_path):
        try:
            location = pyautogui.locateCenterOnScreen(
                template_path, confidence=confidence
            )
            if location:
                cx, cy = location
                _log(f"이미지 발견: {step['label']} → 클릭 ({cx}, {cy})")
                _win32_click(cx, cy, activate_first=activate_first)
                return True
            else:
                _log(f"이미지 못 찾음: {step['label']} — 좌표 폴백 시도")
        except Exception as e:
            _log(f"이미지 검색 오류: {e} — 좌표 폴백 시도")

    # 2차: 저장된 좌표로 폴백
    if step.get("x") is not None and step.get("y") is not None:
        _log(f"좌표 폴백: {step['label']} ({step['x']}, {step['y']})")
        _win32_click(step["x"], step["y"], activate_first=activate_first)
        return True

    _log(f"클릭 실패: {step['label']} — 이미지도 좌표도 없음")
    return False


# --- 설정 파일 관리 ---

def load_config_data(path: str = STEPS_FILE) -> dict | None:
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if isinstance(data, list):
        return None

    data["data_steps"] = [
        step for step in data.get("data_steps", []) if step.get("name") not in RETIRED_STEPS
    ]

    # 저장 다이얼로그 단계는 선택 사항 — 미설정이어도 run 가능 (폴백 처리)
    OPTIONAL_STEPS = {
        "dentweb_open",
        "save_dialog_save_btn",
    }

    for step in data.get("data_steps", []):
        if step.get("skip"):
            continue
        if step.get("name") == "data_check":
            continue
        if step.get("name") in OPTIONAL_STEPS:
            continue
        # 이미지 템플릿이 있으면 좌표 없어도 OK
        template_path = _get_template_path(step.get("name", ""))
        if os.path.exists(template_path):
            continue
        if step.get("x") is None or step.get("y") is None:
            return None
    return data


def save_config_data(data: dict, path: str = STEPS_FILE):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


# --- 학습 모드 (CLI - 레거시) ---

CAPTURE_DELAY = 5


def _countdown_capture() -> tuple[int, int]:
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
    input("준비되면 Enter...")

    steps = json.loads(json.dumps(DATA_STEPS))
    i = 0
    while i < len(steps):
        step = steps[i]
        print(f"\n[{i + 1}/{len(steps)}] {step['label']}")
        choice = input("  → Enter/s/r: ").strip().lower()
        if choice == "r":
            steps = json.loads(json.dumps(DATA_STEPS))
            i = 0
            continue
        if choice == "s":
            steps[i]["skip"] = True
            i += 1
            continue
        x, y = _countdown_capture()
        steps[i]["x"] = x
        steps[i]["y"] = y
        steps[i]["skip"] = False
        _capture_template(x, y, step["name"])
        i += 1

    result = {"data_steps": steps}
    save_config_data(result)
    return result


# --- 메인 Runner ---

class DentwebRunner:
    def __init__(self, cfg: dict):
        self.download_dir = cfg["download_dir"]
        self.download_timeout = cfg.get("download_timeout_seconds", 30)
        self._data = load_config_data()
        self._exports = None

    def _export_store(self) -> ExportStore:
        """전용 폴더는 실행할 때 만든다(창을 여는 것만으로는 폴더·권한을 바꾸지 않는다)."""
        if self._exports is None:
            self._exports = ExportStore(self.download_dir)
        return self._exports

    def is_configured(self) -> bool:
        return self._data is not None

    def _get_save_step(self, name: str) -> dict | None:
        """저장 다이얼로그용 단계 가져오기 (좌표 설정된 경우만)"""
        for step in self._data.get("data_steps", []):
            if step.get("name") == name and not step.get("skip"):
                if step.get("x") is not None and step.get("y") is not None:
                    return step
        return None

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

            matched = [t for t in all_windows if WINDOW_TITLE in t]
            if not matched:
                _log(f"'{WINDOW_TITLE}' 포함된 창 없음")
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

        # 방법 2: win32gui (fallback)
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
            sheet = None
            for name in wb.sheetnames:
                if "수술기록지" in name:
                    sheet = wb[name]
                    break
            if sheet is None:
                wb.close()
                return False
            has_data = False
            for row in sheet.iter_rows(min_row=2, max_row=2, max_col=1, values_only=True):
                if row[0] is not None and str(row[0]).strip():
                    has_data = True
                    break
            wb.close()
            return has_data
        except Exception:
            return False

    def _wait_for_download(self, target: str) -> str | None:
        """이번 실행에 고른 그 파일만 기다린다. 크기가 두 번 연속 같으면 저장이 끝난 것으로 본다."""
        deadline = time.monotonic() + self.download_timeout
        previous = None
        stable = 0
        while time.monotonic() < deadline:
            if os.path.isfile(target):
                size = os.path.getsize(target)
                stable = stable + 1 if size and size == previous else 0
                previous = size
                if stable >= 2:
                    return target
            time.sleep(1)
        return None

    def _cleanup_old_exports(self, keep_days: int = 7, log_callback=None):
        """3.6.4까지 쓰던 EXE 옆 exports 폴더에서 keep_days일 이상 지난 엑셀 파일 삭제.

        3.6.5는 그 폴더에 쓰지 않는다. 남아 있던 환자 파일이 정해진 기간 안에 비워지게 정리만 계속한다.
        """
        def _log(msg):
            if log_callback:
                log_callback(msg)

        legacy_dir = str(legacy_exports_directory())
        if not os.path.isdir(legacy_dir) or os.path.abspath(legacy_dir) == os.path.abspath(self.download_dir):
            return
        cutoff = time.time() - keep_days * 86400
        removed = 0
        for f in glob.glob(os.path.join(legacy_dir, "*.xlsx")):
            try:
                if os.path.getmtime(f) < cutoff:
                    os.remove(f)
                    removed += 1
                    _log(f"오래된 파일 삭제: {os.path.basename(f)}")
            except OSError:
                pass
        if removed:
            _log(f"총 {removed}개 파일 정리 완료")

    def download_excel(self, log_callback=None) -> str | None:
        """전체 자동화: 덴트웹 활성화 → 엑셀 저장 → 데이터 유무 판별.

        - 성공: 파일 경로. 부르는 쪽이 업로드 뒤 cleanup()으로 지운다.
        - 데이터 없음(빈 파일): None.
        - 실패: ExportFailed(서버에 'failed'로 보고한다).
        실패·데이터 없음·예외로 끝나면 이번 실행의 파일을 여기서 지운다.
        """
        def _log(msg):
            if log_callback:
                log_callback(msg)

        if not self._data:
            _log("설정 데이터 없음")
            raise ExportFailed("클릭 좌표 설정이 없습니다")

        self._cleanup_old_exports(keep_days=7, log_callback=log_callback)
        try:
            store = self._export_store()
            store.prune()
        except Exception as e:
            _log(f"저장 폴더를 준비하지 못했습니다: {e}")
            raise ExportFailed(f"저장 폴더를 준비하지 못했습니다: {e}") from e

        target = store.allocate()
        started_at = time.time() - 2
        completed = False
        try:
            excel_path = self._export_to(target, store, started_at, _log, log_callback)
            completed = excel_path is not None
            return excel_path
        finally:
            if not completed:
                try:
                    store.cleanup(target)
                except OSError as e:
                    _log(f"임시 파일 정리 실패: {e}")

    def _export_to(self, target: str, store: ExportStore, started_at: float, _log, log_callback) -> str | None:
        """성공하면 파일 경로, 데이터가 없으면 None. 실패는 ExportFailed."""
        def fail(message):
            _log(message)
            raise ExportFailed(message)

        # 1. 덴트웹 창 열기 (작업표시줄 아이콘 클릭 — 좌표 설정된 경우)
        open_step = self._get_save_step("dentweb_open")
        if open_step:
            _log("덴트웹 창 열기 (작업표시줄 클릭)")
            _win32_click(int(open_step["x"]), int(open_step["y"]))
            time.sleep(open_step.get("wait_after", 2.0))

        # 2. 덴트웹 창 자동 탐색 + 활성화
        _log("덴트웹 창 탐색 중...")
        if not self._activate_dentweb(log_callback=log_callback):
            fail("덴트웹 창을 찾을 수 없습니다")
        _log("덴트웹 창 활성화 완료 — 2초 대기")
        time.sleep(2)

        # 2. 날짜 선택 시퀀스 (엑셀저장 전까지)
        is_first_click = True
        for step in self._data.get("data_steps", []):
            if step.get("skip"):
                continue
            if step.get("name") == "data_check":
                continue
            if step.get("name") == "export_btn":
                break
            if not _find_and_click(step, log_callback, activate_first=is_first_click):
                fail(f"단계 실패: {step['label']}")
            is_first_click = False  # 이후 단계는 이미 활성화됨
            time.sleep(step.get("wait_after", 1.5))

        # 3. 엑셀저장 클릭
        export_step = None
        for step in self._data.get("data_steps", []):
            if step.get("name") == "export_btn" and not step.get("skip"):
                export_step = step
                break
        if not export_step:
            fail("엑셀저장 버튼 설정 없음")

        if not _find_and_click(export_step, log_callback):
            fail("엑셀저장 버튼 클릭 실패")
        time.sleep(export_step.get("wait_after", 3.0))

        # 4. "다른 이름으로 저장" 다이얼로그 처리 — 저장 창이 맨 앞일 때만 키를 보낸다
        _log("저장 다이얼로그 처리 중...")
        if not _wait_for_save_dialog():
            fail("저장 창이 뜨지 않았습니다(키 입력을 보내지 않았습니다)")

        # 4a. Alt+D로 주소창 포커스 → 전용 폴더 경로 붙여넣기 → Enter로 이동
        _log(f"저장 위치 이동: {self.download_dir}")
        pyautogui.hotkey("alt", "d")
        time.sleep(0.4)
        _paste_text(self.download_dir)
        time.sleep(0.3)
        pyautogui.press("enter")
        time.sleep(1.2)  # 폴더 탐색 완료 대기

        # 4b. Alt+N으로 '파일 이름' 칸 → 이번 실행의 고유 이름(전체 경로)으로 바꾼다
        if not _foreground_is_save_dialog():
            fail("저장 창이 앞에 없습니다(파일 이름을 입력하지 않았습니다)")
        _log(f"파일 이름 입력: {os.path.basename(target)}")
        pyautogui.hotkey("alt", "n")
        time.sleep(0.3)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.1)
        _paste_text(target)
        time.sleep(0.3)

        # 4c. '저장(S)' 버튼 클릭
        save_btn_step = self._get_save_step("save_dialog_save_btn")
        if save_btn_step:
            _log("'저장(S)' 버튼 클릭")
            _win32_click(int(save_btn_step["x"]), int(save_btn_step["y"]))
            time.sleep(save_btn_step.get("wait_after", 2.0))
        else:
            _log("[경고] 저장 버튼 좌표 미설정 — Enter로 폴백")
            pyautogui.press("enter")
            time.sleep(2)

        # 고유 이름이라 덮어쓰기 확인 창이 뜨지 않는다. 확인 창 좌표·Enter를 누르지 않는다.

        # 5. 파일 저장 대기
        _log("엑셀 파일 저장 대기 중...")
        excel_path = self._wait_for_download(target)
        if not excel_path:
            strays = store.remove_strays(since=started_at)
            if strays:
                fail("파일 이름을 넣지 못해 다른 이름으로 저장됐습니다. 그 파일은 지웠습니다")
            _log("저장 폴더에 쓰지 못했을 수 있습니다. 덴트웹과 같은 Windows 계정으로 에이전트를 실행하세요")
            fail("엑셀 파일 저장 시간 초과(덴트웹과 같은 Windows 계정으로 실행하는지 확인하세요)")
        _log(f"엑셀 파일 저장 완료: {os.path.basename(excel_path)}")

        # 6. "수술기록지" 시트 2행 데이터 확인
        _log("수술기록지 데이터 확인 중...")
        if not self.excel_has_data(excel_path):
            _log("수술기록지에 데이터 없음 (빈 파일)")
            return None

        _log("수술 기록 데이터 확인 완료")
        return excel_path

    def cleanup(self, file_path: str):
        """이번 실행이 만든 파일만 지운다. 폴더·권한을 다시 만들지 않는다.
        지우지 못하면 OSError를 낸다(부르는 쪽이 기록하고, 다음 실행의 정리가 지운다)."""
        remove_owned(self.download_dir, file_path)
