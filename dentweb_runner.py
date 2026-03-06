"""pyautogui 기반 DentWeb 자동화 시퀀스"""

import os
import time
import glob
import pyautogui


class DentwebRunner:
    def __init__(self, cfg: dict):
        self.window_title = cfg.get("dentweb_window_title", "DentWeb")
        self.download_dir = cfg["download_dir"]
        self.click_delay = cfg.get("click_delay_ms", 300) / 1000
        self.download_timeout = cfg.get("download_timeout_seconds", 30)

    def _find_and_click(self, image_name: str, fallback_xy: tuple = None) -> bool:
        """이미지 인식 우선, 실패 시 fallback 좌표 사용"""
        image_path = os.path.join(os.path.dirname(__file__), "images", image_name)
        try:
            location = pyautogui.locateOnScreen(image_path, confidence=0.8)
            if location:
                pyautogui.click(pyautogui.center(location))
                time.sleep(self.click_delay)
                return True
        except Exception:
            pass

        if fallback_xy:
            pyautogui.click(*fallback_xy)
            time.sleep(self.click_delay)
            return True
        return False

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
        if not self._activate_dentweb():
            return None

        # 클릭 시퀀스 (이미지 인식 + fallback 좌표)
        # 실제 좌표는 병원별 config 또는 이미지 매칭으로 결정
        steps = [
            ("btn_stats.png", None),    # 경영/통계 메뉴
            ("btn_implant.png", None),   # 임플란트 수술통계
            ("btn_export.png", None),    # 엑셀 다운로드
        ]

        for image, fallback in steps:
            if not self._find_and_click(image, fallback):
                return None

        return self._wait_for_download()

    def cleanup(self, file_path: str):
        """임시 파일 삭제"""
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
        except OSError:
            pass
