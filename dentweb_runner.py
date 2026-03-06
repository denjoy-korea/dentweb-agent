"""좌표 기반 DentWeb 자동화 시퀀스

전체 흐름:
1. 덴트웹 실행 확인 → 없으면 프로그램 실행 + 로그인
2. 경영/통계 → 임플란트 수술 통계
3. 특정기간 → 부터 '오늘' → 까지 '오늘' (당일 조회)
4. 자동 조회 → 엑셀저장 → 저장 다이얼로그에서 Enter
5. 저장된 xlsx 파일 경로 반환
"""

import os
import json
import time
import glob
import subprocess
import pyautogui
import pyperclip


STEPS_FILE = "dentweb_steps.json"


def _paste_text(text: str):
    """클립보드에 복사 후 Ctrl+V로 붙여넣기 (한글 지원)"""
    pyperclip.copy(text)
    time.sleep(0.05)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.1)


# --- 클릭 시퀀스 정의 ---

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
    {"name": "date_start_field", "label": "'부터' 날짜 필드 클릭 (달력 열기)", "x": None, "y": None, "wait_after": 0.5, "group": "data"},
    {"name": "date_start_today", "label": "'부터' 달력 하단 '오늘' 버튼", "x": None, "y": None, "wait_after": 0.5, "group": "data"},
    {"name": "date_end_field", "label": "'까지' 날짜 필드 클릭 (달력 열기)", "x": None, "y": None, "wait_after": 0.5, "group": "data"},
    {"name": "date_end_today", "label": "'까지' 달력 하단 '오늘' 버튼 (자동 조회됨)", "x": None, "y": None, "wait_after": 3.0, "group": "data"},
    {"name": "data_check", "label": "수술기록목지 첫 번째 행 위치 (데이터 유무 확인용)", "x": None, "y": None, "wait_after": 0, "group": "check"},
    {"name": "export_btn", "label": "'엑셀저장' 버튼", "x": None, "y": None, "wait_after": 2.0, "group": "data"},
]


# --- 설정 파일 관리 ---

def load_config_data(path: str = STEPS_FILE) -> dict | None:
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # v1.1~1.2 호환
    if isinstance(data, list):
        return None  # 이전 형식은 재설정 필요

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

CAPTURE_DELAY = 5  # 마우스 위치 캡처까지 대기 시간(초)


def _countdown_capture() -> tuple[int, int]:
    """카운트다운 후 마우스 위치 캡처 (덴트웹 창이 포커스 상태여도 동작)"""
    for remaining in range(CAPTURE_DELAY, 0, -1):
        print(f"\r  → {remaining}초 후 마우스 위치 저장...", end="", flush=True)
        time.sleep(1)
    x, y = pyautogui.position()
    print(f"\r  → 저장됨: ({x}, {y})              ")
    return x, y


def _teach_steps(steps: list[dict], group_label: str) -> list[dict]:
    print(f"\n--- {group_label} ---\n")
    result = json.loads(json.dumps(steps))

    i = 0
    while i < len(result):
        step = result[i]
        step_type = step.get("type", "click")
        print(f"[{i + 1}/{len(result)}] {step['label']}")

        if step_type in ("text_input", "password_input"):
            print(f"  → 입력 필드 위치에 마우스를 올려주세요.")
        elif step.get("group") == "check":
            print(f"  → 데이터가 표시되는 영역(첫 번째 행)에 마우스를 올려주세요.")
            print(f"    (비어있으면 엑셀저장을 건너뜁니다)")
        else:
            print(f"  → 해당 위치에 마우스를 올려주세요.")

        print(f"  → 옵션: Enter={CAPTURE_DELAY}초 카운트다운 시작 / s=건너뛰기 / r=처음부터")
        choice = input("  → ").strip().lower()

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

        # Enter 누른 후 카운트다운 → 덴트웹으로 전환할 시간
        print(f"  → 지금 덴트웹에서 '{step['label']}' 위에 마우스를 올려주세요!")
        x, y = _countdown_capture()
        result[i]["x"] = x
        result[i]["y"] = y
        result[i]["skip"] = False
        print()
        i += 1

    return result


def run_teach_mode() -> dict:
    print()
    print("=" * 55)
    print("  덴트웹 자동화 설정")
    print("=" * 55)
    print()
    print("덴트웹 화면의 각 버튼/필드 위치를 학습합니다.")
    print("각 단계에서 해당 위치에 마우스를 올린 뒤 Enter를 누르세요.")
    print()
    print("* 's' = 단계 건너뛰기 (이미 로그인 상태 등)")
    print("* 'r' = 해당 그룹 처음부터")
    print()

    # --- 덴트웹 실행 경로 ---
    print("=" * 55)
    print("  1. 덴트웹 프로그램 경로")
    print("=" * 55)
    print()
    print("덴트웹 바로가기 또는 exe 경로를 입력하세요.")
    print("예: C:\\Users\\Public\\Desktop\\덴트웹.lnk")
    print("    빈칸 = 이미 실행 중인 경우만 자동화")
    print()
    exe_path = input("경로: ").strip().strip('"')
    print()

    # --- 작업표시줄/로그인 ---
    print("=" * 55)
    print("  2. 프로그램 활성화 + 로그인")
    print("=" * 55)
    print()
    print("덴트웹을 열고 로그인 화면까지 준비해주세요.")
    input("준비되면 Enter...")

    launch_steps = _teach_steps(LAUNCH_STEPS, "프로그램 활성화 + 로그인")

    login_id = ""
    login_pw = ""
    has_login = any(
        not s.get("skip") and s.get("type") in ("text_input", "password_input")
        for s in launch_steps
    )
    if has_login:
        print()
        print("자동 로그인 계정:")
        login_id = input("  덴트웹 ID: ").strip()
        login_pw = input("  덴트웹 PW: ").strip()

    # --- 데이터 추출 ---
    print()
    print("=" * 55)
    print("  3. 데이터 추출 (통계 → 엑셀저장)")
    print("=" * 55)
    print()
    print("로그인 후 메인 화면이 보이는 상태에서 진행합니다.")
    print()
    print("순서: 경영/통계 → 임플란트 수술 통계 → 특정기간")
    print("      → 부터(오늘) → 까지(오늘) → 엑셀저장")
    print()
    print("TIP: '부터' 날짜 필드와 '오늘' 버튼을 각각 가리켜주세요.")
    print("     '까지'도 마찬가지입니다.")
    print()
    input("준비되면 Enter...")

    data_steps = _teach_steps(DATA_STEPS, "데이터 추출 시퀀스")

    result = {
        "launch_steps": launch_steps,
        "data_steps": data_steps,
        "meta": {
            "exe_path": exe_path,
            "login_id": login_id,
            "login_pw": login_pw,
        },
    }

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
        return self._default_window_title

    def _is_dentweb_running(self) -> bool:
        try:
            import pygetwindow as gw
            windows = gw.getWindowsWithTitle(self.window_title)
            return len(windows) > 0
        except Exception:
            return False

    def _activate_dentweb(self) -> bool:
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
        exe_path = self.meta.get("exe_path", "")
        if not exe_path:
            return False
        try:
            subprocess.Popen(exe_path, shell=True)
            deadline = time.time() + 30
            while time.time() < deadline:
                if self._is_dentweb_running():
                    time.sleep(2)
                    return True
                time.sleep(1)
            return False
        except Exception:
            return False

    def _click_taskbar_icon(self) -> bool:
        if not self._data:
            return False
        for step in self._data.get("launch_steps", []):
            if step.get("name") == "taskbar_icon" and not step.get("skip"):
                pyautogui.click(step["x"], step["y"])
                time.sleep(step.get("wait_after", 1.0))
                return True
        return False

    def _do_login(self) -> bool:
        if not self._data:
            return False
        login_id = self.meta.get("login_id", "")
        login_pw = self.meta.get("login_pw", "")

        for step in self._data.get("launch_steps", []):
            if step.get("skip") or step.get("group") != "login":
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
                pyautogui.click(step["x"], step["y"])
            time.sleep(step.get("wait_after", 0.5))
        return True

    def ensure_dentweb_ready(self) -> bool:
        """덴트웹 실행 확인 → 활성화 또는 실행+로그인"""
        # 1. 이미 실행 중이면 활성화
        if self._is_dentweb_running():
            if self._activate_dentweb():
                return True
            self._click_taskbar_icon()
            time.sleep(1)
            return self._is_dentweb_running()

        # 2. 프로그램 실행
        exe_path = self.meta.get("exe_path", "")
        if exe_path:
            if not self._launch_dentweb():
                return False
            time.sleep(2)
            self._do_login()
            return True

        # 3. 작업표시줄 아이콘 클릭
        self._click_taskbar_icon()
        time.sleep(2)
        return self._is_dentweb_running()

    def _wait_for_download(self) -> str | None:
        """다운로드 폴더에서 엑셀 파일 감지"""
        target = os.path.join(self.download_dir, "dentweb_export.xlsx")
        deadline = time.time() + self.download_timeout

        # 기존 파일이 있으면 삭제 (새 파일 감지를 위해)
        # 이미 download_excel() 호출 전 시점의 상태
        before = set(glob.glob(os.path.join(self.download_dir, "*.xlsx")))

        while time.time() < deadline:
            # 지정 파일명 또는 새 xlsx 파일
            if os.path.exists(target):
                time.sleep(1)  # 쓰기 완료 대기
                return target
            current = set(glob.glob(os.path.join(self.download_dir, "*.xlsx")))
            new_files = current - before
            if new_files:
                time.sleep(1)
                return max(new_files, key=os.path.getmtime)
            time.sleep(1)
        return None

    def _has_data(self) -> bool:
        """수술기록목지 첫 행 위치의 픽셀을 검사해서 데이터 유무 판별.
        빈 화면은 배경색(흰색 계열)만 있고, 데이터가 있으면 텍스트 픽셀이 존재."""
        if not self._data:
            return False

        check_step = None
        for step in self._data.get("data_steps", []):
            if step.get("name") == "data_check" and not step.get("skip"):
                check_step = step
                break
        if not check_step:
            return True  # 체크 좌표가 없으면 항상 데이터 있다고 간주

        cx, cy = check_step["x"], check_step["y"]
        # 해당 좌표 주변 100x30 영역 캡처
        region = (cx - 50, cy - 15, 100, 30)
        screenshot = pyautogui.screenshot(region=region)

        # 픽셀 색상 분석: 어두운 픽셀(텍스트)이 있으면 데이터 존재
        dark_pixel_count = 0
        for x in range(screenshot.width):
            for y in range(screenshot.height):
                r, g, b = screenshot.getpixel((x, y))
                # 텍스트는 어두운 색 (RGB 모두 < 100)
                if r < 100 and g < 100 and b < 100:
                    dark_pixel_count += 1

        # 어두운 픽셀이 일정 수 이상이면 텍스트(데이터) 존재
        return dark_pixel_count > 20

    def download_excel(self) -> str | None:
        """전체 자동화: 덴트웹 준비 → 데이터 추출 → Excel 반환
        데이터가 없으면 None 반환 (엑셀저장 생략)"""
        if not self._data:
            return None

        # 1. 덴트웹 실행/활성화
        if not self.ensure_dentweb_ready():
            return None

        # 2. 날짜 선택 시퀀스 (엑셀저장 전까지)
        for step in self._data.get("data_steps", []):
            if step.get("skip"):
                continue
            if step.get("group") == "check":
                continue  # 체크 단계는 별도 처리
            if step.get("name") == "export_btn":
                break  # 엑셀저장은 데이터 확인 후
            pyautogui.click(step["x"], step["y"])
            time.sleep(step.get("wait_after", 0.5))

        # 3. 데이터 유무 확인
        if not self._has_data():
            return None  # 수술 기록 없음 → 중단

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
        pyautogui.hotkey("alt", "n")  # 파일 이름 필드로 포커스
        time.sleep(0.3)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.1)
        _paste_text(save_path)
        time.sleep(0.3)
        pyautogui.press("enter")

        # 6. 덮어쓰기 확인 다이얼로그가 뜰 수 있음 → Enter
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
