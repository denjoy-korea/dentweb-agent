"""DentWeb Agent GUI — customtkinter 기반 모던 UI"""

import customtkinter as ctk
import threading
import time
import traceback
import json
import os
import sys
import datetime

from config import DEFAULTS, SERVER_URL
from api_client import ApiClient
from dentweb_runner import (
    DentwebRunner, DATA_STEPS, save_config_data, load_config_data, STEPS_FILE,
    _capture_template,
)
from logger import AgentLogger
from startup import is_registered, register, unregister
from updater import check_update, download_update, apply_update
import pyautogui

VERSION = "3.1.0"

# --- 색상 ---
EMERALD_600 = "#059669"
EMERALD_500 = "#10b981"
EMERALD_100 = "#d1fae5"
EMERALD_50 = "#ecfdf5"
SLATE_900 = "#0f172a"
SLATE_700 = "#334155"
SLATE_500 = "#64748b"
SLATE_400 = "#94a3b8"
SLATE_200 = "#e2e8f0"
SLATE_100 = "#f1f5f9"
SLATE_50 = "#f8fafc"
WHITE = "#ffffff"
ROSE_500 = "#f43f5e"
AMBER_500 = "#f59e0b"
BLUE_500 = "#3b82f6"


def _now_str():
    return datetime.datetime.now().strftime("%H:%M:%S")


class AgentApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("green")

        self.title("DenJOY 덴트웹 에이전트")
        self.geometry("440x620")
        self.resizable(False, False)
        self.configure(fg_color=SLATE_50)
        self.attributes("-topmost", True)

        # 아이콘 (있으면)
        icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)

        # 상태
        self.cfg = None
        self.api = None
        self.runner = None
        self.log = None
        self.polling = False
        self._poll_thread = None
        self._last_status = ""
        self._last_message = ""
        self._last_time = ""

        # config 로드 시도
        self.cfg = self._load_config_silent()
        if self.cfg and self.cfg.get("agent_token"):
            self._show_main_screen()
        else:
            self._show_setup_screen()

    # ─── Config ───

    def _load_config_silent(self):
        path = "config.json"
        if not os.path.exists(path):
            return None
        try:
            with open(path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            if not cfg.get("agent_token"):
                return None
            for key, default in DEFAULTS.items():
                cfg.setdefault(key, default)
            cfg["server_url"] = cfg["server_url"].rstrip("/")
            return cfg
        except Exception:
            return None

    def _save_config(self, token: str):
        cfg = {**DEFAULTS, "agent_token": token}
        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
        self.cfg = cfg

    # ─── Setup Screen ───

    def _show_setup_screen(self):
        self._clear()
        self.geometry("440x460")

        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=30, pady=30)

        # 로고 영역
        logo_frame = ctk.CTkFrame(frame, fg_color=EMERALD_100, corner_radius=16,
                                  height=64, width=64)
        logo_frame.pack(pady=(20, 0))
        logo_label = ctk.CTkLabel(logo_frame, text="🦷", font=ctk.CTkFont(size=28))
        logo_label.pack(padx=16, pady=12)

        ctk.CTkLabel(
            frame, text="DenJOY 덴트웹 에이전트",
            font=ctk.CTkFont(size=20, weight="bold"), text_color=SLATE_900,
        ).pack(pady=(16, 4))

        ctk.CTkLabel(
            frame, text="병원 PC에서 덴트웹 데이터를 자동 수집합니다",
            font=ctk.CTkFont(size=12), text_color=SLATE_500,
        ).pack(pady=(0, 24))

        # 토큰 입력
        ctk.CTkLabel(
            frame, text="에이전트 토큰",
            font=ctk.CTkFont(size=12, weight="bold"), text_color=SLATE_700,
            anchor="w",
        ).pack(fill="x")

        self._token_entry = ctk.CTkEntry(
            frame, placeholder_text="앱 설정에서 복사한 토큰 붙여넣기",
            height=42, corner_radius=10,
            font=ctk.CTkFont(size=13),
        )
        self._token_entry.pack(fill="x", pady=(4, 16))

        # 시작프로그램 체크박스
        self._startup_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            frame, text="컴퓨터 시작 시 자동 실행",
            variable=self._startup_var,
            font=ctk.CTkFont(size=12), text_color=SLATE_700,
            fg_color=EMERALD_500, hover_color=EMERALD_600,
        ).pack(anchor="w", pady=(0, 20))

        # 에러 메시지
        self._setup_error = ctk.CTkLabel(
            frame, text="", font=ctk.CTkFont(size=11),
            text_color=ROSE_500, wraplength=360,
        )
        self._setup_error.pack()

        # 시작 버튼
        self._start_btn = ctk.CTkButton(
            frame, text="시작하기", height=44, corner_radius=10,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=EMERALD_600, hover_color=EMERALD_500,
            command=self._on_setup_submit,
        )
        self._start_btn.pack(fill="x", pady=(8, 0))

    def _on_setup_submit(self):
        token = self._token_entry.get().strip()
        if not token:
            self._setup_error.configure(text="토큰을 입력해주세요.")
            return

        self._start_btn.configure(state="disabled", text="연결 확인 중...")
        self._setup_error.configure(text="")
        self.update()

        # 백그라운드에서 서버 연결 확인
        def check():
            try:
                cfg = {**DEFAULTS, "agent_token": token}
                api = ApiClient(cfg["server_url"], token)
                api.ping()
                # 성공
                self.after(0, lambda: self._setup_success(token))
            except Exception as e:
                self.after(0, lambda: self._setup_fail(str(e)))

        threading.Thread(target=check, daemon=True).start()

    def _setup_success(self, token: str):
        self._save_config(token)

        # 시작프로그램 등록
        if self._startup_var.get() and sys.platform == "win32":
            register()

        self._show_main_screen()

    def _setup_fail(self, error: str):
        self._setup_error.configure(text=f"서버 연결 실패: {error}\n토큰을 확인해주세요.")
        self._start_btn.configure(state="normal", text="시작하기")

    # ─── Main Screen ───

    def _show_main_screen(self):
        self._clear()
        self.geometry("440x620")

        self.api = ApiClient(self.cfg["server_url"], self.cfg["agent_token"])
        self.runner = DentwebRunner(self.cfg)
        self.log = AgentLogger(
            self.cfg.get("log_file", "agent.log"),
            self.cfg.get("log_max_lines", 1000),
        )

        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=20, pady=16)

        # 헤더
        header = ctk.CTkFrame(container, fg_color="transparent")
        header.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(
            header, text="DenJOY 덴트웹 에이전트",
            font=ctk.CTkFont(size=16, weight="bold"), text_color=SLATE_900,
        ).pack(side="left")

        ctk.CTkLabel(
            header, text=f"v{VERSION}",
            font=ctk.CTkFont(size=11), text_color=SLATE_400,
        ).pack(side="right")

        # 업데이트 배너 (처음엔 숨김)
        self._update_banner = ctk.CTkFrame(container, fg_color=EMERALD_50,
                                           corner_radius=10, border_width=1,
                                           border_color=EMERALD_100)
        # pack 하지 않음 — 업데이트 발견 시 표시

        banner_inner = ctk.CTkFrame(self._update_banner, fg_color="transparent")
        banner_inner.pack(fill="x", padx=12, pady=8)

        self._update_text = ctk.CTkLabel(
            banner_inner, text="",
            font=ctk.CTkFont(size=11), text_color=EMERALD_600,
        )
        self._update_text.pack(side="left", fill="x", expand=True)

        self._update_btn = ctk.CTkButton(
            banner_inner, text="업데이트", width=80, height=28, corner_radius=6,
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=EMERALD_600, hover_color=EMERALD_500,
            command=self._on_update_click,
        )
        self._update_btn.pack(side="right")

        self._update_progress = ctk.CTkProgressBar(
            self._update_banner, height=3, corner_radius=2,
            fg_color=SLATE_200, progress_color=EMERALD_500,
        )
        # 다운로드 시에만 표시

        self._pending_update = None  # {version, download_url}

        # 연결 상태 카드
        status_card = ctk.CTkFrame(container, fg_color=WHITE, corner_radius=12)
        status_card.pack(fill="x", pady=(0, 8))

        status_inner = ctk.CTkFrame(status_card, fg_color="transparent")
        status_inner.pack(fill="x", padx=16, pady=12)

        self._status_dot = ctk.CTkLabel(
            status_inner, text="●", font=ctk.CTkFont(size=14),
            text_color=SLATE_400,
        )
        self._status_dot.pack(side="left")

        self._status_label = ctk.CTkLabel(
            status_inner, text="연결 확인 중...",
            font=ctk.CTkFont(size=13, weight="bold"), text_color=SLATE_700,
        )
        self._status_label.pack(side="left", padx=(6, 0))

        self._status_sub = ctk.CTkLabel(
            status_inner, text="",
            font=ctk.CTkFont(size=11), text_color=SLATE_400,
        )
        self._status_sub.pack(side="right")

        # 최근 실행 카드
        run_card = ctk.CTkFrame(container, fg_color=WHITE, corner_radius=12)
        run_card.pack(fill="x", pady=(0, 8))

        run_inner = ctk.CTkFrame(run_card, fg_color="transparent")
        run_inner.pack(fill="x", padx=16, pady=12)

        ctk.CTkLabel(
            run_inner, text="최근 실행",
            font=ctk.CTkFont(size=11, weight="bold"), text_color=SLATE_500,
        ).pack(anchor="w")

        self._run_time_label = ctk.CTkLabel(
            run_inner, text="기록 없음",
            font=ctk.CTkFont(size=12), text_color=SLATE_700,
        )
        self._run_time_label.pack(anchor="w", pady=(4, 0))

        self._run_result_label = ctk.CTkLabel(
            run_inner, text="",
            font=ctk.CTkFont(size=12), text_color=SLATE_500,
        )
        self._run_result_label.pack(anchor="w")

        # 로그 카드
        log_card = ctk.CTkFrame(container, fg_color=WHITE, corner_radius=12)
        log_card.pack(fill="both", expand=True, pady=(0, 8))

        log_header = ctk.CTkFrame(log_card, fg_color="transparent")
        log_header.pack(fill="x", padx=16, pady=(12, 4))

        ctk.CTkLabel(
            log_header, text="활동 로그",
            font=ctk.CTkFont(size=11, weight="bold"), text_color=SLATE_500,
        ).pack(side="left")

        self._log_text = ctk.CTkTextbox(
            log_card, height=180, corner_radius=0,
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=WHITE, text_color=SLATE_700,
            state="disabled",
            border_width=0,
        )
        self._log_text.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        # 하단 버튼 영역
        btn_frame = ctk.CTkFrame(container, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(0, 4))

        teach_btn = ctk.CTkButton(
            btn_frame, text="좌표 설정", height=36, corner_radius=8,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=SLATE_100, hover_color=SLATE_200,
            text_color=SLATE_700, border_width=1, border_color=SLATE_200,
            command=self._on_teach_click,
        )
        teach_btn.pack(side="left", expand=True, fill="x", padx=(0, 3))

        test_btn = ctk.CTkButton(
            btn_frame, text="단계별 테스트", height=36, corner_radius=8,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=EMERALD_50, hover_color=EMERALD_100,
            text_color=EMERALD_600, border_width=1, border_color=EMERALD_100,
            command=self._on_test_click,
        )
        test_btn.pack(side="left", expand=True, fill="x", padx=(3, 3))

        startup_text = "시작프로그램 해제" if is_registered() else "시작프로그램 등록"
        self._startup_btn = ctk.CTkButton(
            btn_frame, text=startup_text, height=36, corner_radius=8,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=SLATE_100, hover_color=SLATE_200,
            text_color=SLATE_700, border_width=1, border_color=SLATE_200,
            command=self._on_startup_toggle,
        )
        self._startup_btn.pack(side="left", expand=True, fill="x", padx=(3, 0))

        # 초기화 완료 → 좌표 체크 → 폴링 시작
        self._gui_log("에이전트 시작")

        if not self.runner.is_configured():
            self._gui_log("클릭 좌표 미설정 — 좌표 학습이 필요합니다")
            self._update_status("설정 필요", AMBER_500, "좌표 재설정을 눌러주세요")
        else:
            self._start_polling()

        # 백그라운드 업데이트 확인
        threading.Thread(target=self._check_for_updates, daemon=True).start()

    # ─── GUI 로그 ───

    def _gui_log(self, message: str, level: str = "INFO"):
        timestamp = _now_str()
        line = f"[{timestamp}] {message}\n"

        self._log_text.configure(state="normal")
        self._log_text.insert("end", line)
        self._log_text.see("end")
        self._log_text.configure(state="disabled")

        if self.log:
            if level == "ERROR":
                self.log.error(message)
            else:
                self.log.info(message)

    def _update_status(self, text: str, color: str = EMERALD_500, sub: str = ""):
        self._status_dot.configure(text_color=color)
        self._status_label.configure(text=text)
        self._status_sub.configure(text=sub)

    def _update_last_run(self, time_str: str, result: str, color: str = SLATE_700):
        self._run_time_label.configure(text=time_str)
        self._run_result_label.configure(text=result, text_color=color)

    # ─── Polling ───

    def _start_polling(self):
        self.polling = True
        self._poll_thread = threading.Thread(target=self._poll_loop, daemon=True)
        self._poll_thread.start()
        self.after(0, lambda: self._update_status("서버 연결됨", EMERALD_500, "대기 중"))
        self.after(0, lambda: self._gui_log("서버 폴링 시작"))

    def _poll_loop(self):
        poll_interval = self.cfg.get("poll_interval_seconds", 30)

        # 초기 연결 확인 (claim 없이 상태만 조회)
        try:
            self.api.ping()
            self.after(0, lambda: self._update_status("서버 연결됨", EMERALD_500, "대기 중"))
            self.after(0, lambda: self._gui_log("서버 연결 성공"))
        except Exception as e:
            self.after(0, lambda: self._update_status("연결 실패", ROSE_500, str(e)[:30]))
            self.after(0, lambda: self._gui_log(f"서버 연결 실패: {e}", "ERROR"))

        while self.polling:
            try:
                result = self.api.claim_run()
                if not result.get("should_run"):
                    time.sleep(poll_interval)
                    continue

                reason = result.get("reason", "")
                self.after(0, lambda r=reason: self._gui_log(f"실행 시작: {r}"))
                self.after(0, lambda: self._update_status("실행 중", BLUE_500, "자동화 진행 중..."))

                if not self.runner.is_configured():
                    self.api.report_run("failed", "클릭 좌표 미설정")
                    self.after(0, lambda: self._gui_log("클릭 좌표 미설정", "ERROR"))
                    self.after(0, lambda: self._update_status("설정 필요", AMBER_500))
                    time.sleep(poll_interval)
                    continue

                # 자동화 실행 — topmost 해제 + 화면 우하단으로 이동 (덴트웹 클릭 방해 방지)
                self.after(0, lambda: self._prepare_for_automation())
                time.sleep(0.5)  # 이동 완료 대기
                self.after(0, lambda: self._gui_log("덴트웹 자동화 시작..."))
                excel_path = self.runner.download_excel(
                    log_callback=lambda msg: self.after(0, lambda m=msg: self._gui_log(m))
                )
                self.after(0, lambda: self._restore_after_automation())

                now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

                if not excel_path:
                    self.api.report_run("no_data", "오늘 수술 기록이 없습니다")
                    self.after(0, lambda: self._gui_log("오늘 수술 기록 없음"))
                    self.after(0, lambda t=now_str: self._update_last_run(
                        t, "데이터 없음", SLATE_500))
                    self.after(0, lambda: self._update_status("서버 연결됨", EMERALD_500, "대기 중"))
                    time.sleep(poll_interval)
                    continue

                # 업로드
                upload_url = result.get("upload_url", f"{self.cfg['server_url']}/dentweb-upload")
                upload_result = self.api.upload_file(upload_url, excel_path)

                if upload_result.get("success"):
                    inserted = upload_result.get("inserted", 0)
                    skipped = upload_result.get("skipped", 0)
                    msg = f"{inserted}건 업로드, {skipped}건 스킵"
                    self.api.report_run("success", msg)
                    self.after(0, lambda m=msg: self._gui_log(f"완료: {m}"))
                    self.after(0, lambda t=now_str, m=msg: self._update_last_run(
                        t, f"성공 — {m}", EMERALD_600))
                else:
                    error_msg = upload_result.get("error", "업로드 실패")
                    self.api.report_run("failed", error_msg)
                    self.after(0, lambda e=error_msg: self._gui_log(f"업로드 실패: {e}", "ERROR"))
                    self.after(0, lambda t=now_str, e=error_msg: self._update_last_run(
                        t, f"실패 — {e}", ROSE_500))

                self.runner.cleanup(excel_path)
                self.after(0, lambda: self._update_status("서버 연결됨", EMERALD_500, "대기 중"))

            except Exception as e:
                self.after(0, lambda e=e: self._gui_log(f"에러: {e}", "ERROR"))
                try:
                    self.api.report_run("failed", str(e)[:1000])
                except Exception:
                    pass

            time.sleep(poll_interval)

    # ─── Automation Window Management ───

    def _prepare_for_automation(self):
        """자동화 시작: topmost 해제 + 화면 우하단 이동 (로그 계속 표시)"""
        self._saved_geometry = self.geometry()  # 원래 위치 저장
        self.attributes("-topmost", False)
        # 화면 우하단으로 이동 (작게)
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        w, h = 440, 620
        x = screen_w - w - 10
        y = screen_h - h - 60  # 태스크바 위
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _restore_after_automation(self):
        """자동화 끝: topmost 복원 + 원래 위치"""
        self.attributes("-topmost", True)
        if hasattr(self, "_saved_geometry") and self._saved_geometry:
            self.geometry(self._saved_geometry)
        self.lift()

    # ─── Teach Mode (GUI) ───

    def _on_teach_click(self):
        self.polling = False
        time.sleep(0.3)
        TeachWindow(self, self._on_teach_complete)

    def _on_teach_complete(self, success: bool):
        if success:
            self.runner = DentwebRunner(self.cfg)
            self._gui_log("클릭 좌표 설정 완료")
            self._update_status("서버 연결됨", EMERALD_500, "대기 중")
            self._start_polling()
        else:
            if self.runner.is_configured():
                self._start_polling()
            else:
                self._update_status("설정 필요", AMBER_500, "좌표 재설정을 눌러주세요")

    # ─── Test Mode (GUI) ───

    def _on_test_click(self):
        if not self.runner or not self.runner.is_configured():
            self._gui_log("좌표 설정을 먼저 해주세요", "ERROR")
            return
        self.polling = False
        time.sleep(0.3)
        TestWindow(self, self._on_test_complete)

    def _on_test_complete(self):
        self.runner = DentwebRunner(self.cfg)
        if self.runner.is_configured():
            self._start_polling()

    # ─── Update ───

    def _check_for_updates(self):
        update_info = check_update(VERSION)
        if update_info:
            self.after(0, lambda: self._show_update_banner(update_info))

    def _show_update_banner(self, update_info: dict):
        self._pending_update = update_info
        version = update_info["version"]
        size_mb = update_info.get("size", 0) / 1024 / 1024
        self._update_text.configure(
            text=f"새 버전 {version} ({size_mb:.1f}MB)")
        # 연결 상태 카드 위에 삽입
        try:
            status_card = self._status_dot.master.master  # status_inner → status_card
            self._update_banner.pack(fill="x", pady=(0, 8), before=status_card)
        except Exception:
            self._update_banner.pack(fill="x", pady=(0, 8))
        self._gui_log(f"업데이트 발견: {version}")

    def _on_update_click(self):
        if not self._pending_update:
            return

        self._update_btn.configure(state="disabled", text="다운로드 중...")
        self._update_progress.set(0)
        self._update_progress.pack(fill="x", padx=12, pady=(0, 8))

        def do_download():
            def on_progress(p):
                self.after(0, lambda p=p: self._update_progress.set(p))

            new_path = download_update(
                self._pending_update["download_url"],
                progress_callback=on_progress,
            )

            if new_path:
                self.after(0, lambda: self._apply_downloaded_update(new_path))
            else:
                self.after(0, lambda: self._update_download_failed())

        threading.Thread(target=do_download, daemon=True).start()

    def _apply_downloaded_update(self, new_exe_path: str):
        self._update_btn.configure(text="재시작 중...")
        self._gui_log("업데이트 다운로드 완료 — 재시작합니다")
        self.polling = False
        self.update()
        apply_update(new_exe_path)

    def _update_download_failed(self):
        self._update_btn.configure(state="normal", text="재시도")
        self._update_progress.pack_forget()
        self._gui_log("업데이트 다운로드 실패", "ERROR")

    # ─── Startup Toggle ───

    def _on_startup_toggle(self):
        if is_registered():
            unregister()
            self._startup_btn.configure(text="시작프로그램 등록")
            self._gui_log("시작프로그램에서 제거됨")
        else:
            register()
            self._startup_btn.configure(text="시작프로그램 해제")
            self._gui_log("시작프로그램에 등록됨")

    # ─── Util ───

    def _clear(self):
        for widget in self.winfo_children():
            widget.destroy()

    def on_closing(self):
        self.polling = False
        self.destroy()


# ─── Teach Mode Window ───

class TeachWindow(ctk.CTkToplevel):
    """클릭 좌표 학습 전용 창 (항상 최상위, 작은 사이즈)"""

    def __init__(self, parent: AgentApp, callback):
        super().__init__(parent)
        self.parent_app = parent
        self.callback = callback

        self.title("클릭 위치 학습")
        self.geometry("360x340")
        self.resizable(False, False)
        self.configure(fg_color=WHITE)
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

        # 위치: 화면 우측 하단
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        self.geometry(f"360x340+{screen_w - 400}+{screen_h - 420}")

        self.steps = json.loads(json.dumps(DATA_STEPS))
        self.current_step = 0
        self.capturing = False

        self._build_ui()
        self._show_step()

    def _build_ui(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=20, pady=16)

        # 진행률
        self._progress_label = ctk.CTkLabel(
            container, text="",
            font=ctk.CTkFont(size=11, weight="bold"), text_color=SLATE_500,
        )
        self._progress_label.pack(anchor="w")

        self._progress_bar = ctk.CTkProgressBar(
            container, height=4, corner_radius=2,
            fg_color=SLATE_200, progress_color=EMERALD_500,
        )
        self._progress_bar.pack(fill="x", pady=(4, 12))

        # 지시사항
        self._instruction = ctk.CTkLabel(
            container, text="",
            font=ctk.CTkFont(size=14, weight="bold"), text_color=SLATE_900,
            wraplength=300,
        )
        self._instruction.pack(pady=(0, 4))

        self._hint = ctk.CTkLabel(
            container, text="",
            font=ctk.CTkFont(size=11), text_color=SLATE_500,
            wraplength=300,
        )
        self._hint.pack(pady=(0, 16))

        # 카운트다운 / 결과
        self._countdown_label = ctk.CTkLabel(
            container, text="",
            font=ctk.CTkFont(size=36, weight="bold"), text_color=EMERALD_600,
        )
        self._countdown_label.pack(pady=(0, 16))

        # 버튼 영역
        btn_frame = ctk.CTkFrame(container, fg_color="transparent")
        btn_frame.pack(fill="x")

        self._capture_btn = ctk.CTkButton(
            btn_frame, text="캡처 시작", height=38, corner_radius=8,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=EMERALD_600, hover_color=EMERALD_500,
            command=self._on_capture,
        )
        self._capture_btn.pack(side="left", expand=True, fill="x", padx=(0, 4))

        self._skip_btn = ctk.CTkButton(
            btn_frame, text="건너뛰기", height=38, corner_radius=8,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=SLATE_100, hover_color=SLATE_200,
            text_color=SLATE_700, border_width=1, border_color=SLATE_200,
            command=self._on_skip,
        )
        self._skip_btn.pack(side="left", expand=True, fill="x", padx=(4, 4))

        self._reset_btn = ctk.CTkButton(
            btn_frame, text="처음부터", height=38, corner_radius=8,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=SLATE_100, hover_color=SLATE_200,
            text_color=SLATE_700, border_width=1, border_color=SLATE_200,
            command=self._on_reset,
        )
        self._reset_btn.pack(side="left", expand=True, fill="x", padx=(4, 0))

    def _show_step(self):
        if self.current_step >= len(self.steps):
            self._finish()
            return

        step = self.steps[self.current_step]
        total = len(self.steps)
        idx = self.current_step + 1

        self._progress_label.configure(text=f"단계 {idx}/{total}")
        self._progress_bar.set(idx / total)
        self._instruction.configure(text=step["label"])

        self._hint.configure(text="해당 위치에 마우스를 올린 뒤\n'캡처 시작'을 클릭하세요.")

        self._countdown_label.configure(text="")
        self._capture_btn.configure(state="normal", text="캡처 시작")
        self._skip_btn.configure(state="normal")
        self._reset_btn.configure(state="normal")

    def _on_capture(self):
        if self.capturing:
            return
        self.capturing = True
        self._capture_btn.configure(state="disabled", text="카운트다운...")
        self._skip_btn.configure(state="disabled")
        self._reset_btn.configure(state="disabled")
        self._hint.configure(text="지금 덴트웹에서 해당 위치에 마우스를 올려주세요!")

        threading.Thread(target=self._countdown_and_capture, daemon=True).start()

    def _countdown_and_capture(self):
        for remaining in range(5, 0, -1):
            self.after(0, lambda r=remaining: self._countdown_label.configure(text=str(r)))
            time.sleep(1)

        x, y = pyautogui.position()
        step_name = self.steps[self.current_step]["name"]
        self.steps[self.current_step]["x"] = x
        self.steps[self.current_step]["y"] = y
        self.steps[self.current_step]["skip"] = False

        # 마우스 위치 주변 영역을 템플릿 이미지로 캡처
        _capture_template(x, y, step_name)

        self.after(0, lambda: self._countdown_label.configure(
            text=f"✓ ({x}, {y})", text_color=EMERALD_600))
        self.after(0, lambda: self._hint.configure(
            text="좌표 + 이미지 템플릿 저장 완료"))

        self.capturing = False
        self.current_step += 1
        self.after(800, self._show_step)

    def _on_skip(self):
        self.steps[self.current_step]["skip"] = True
        self.current_step += 1
        self._show_step()

    def _on_reset(self):
        self.steps = json.loads(json.dumps(DATA_STEPS))
        self.current_step = 0
        self._show_step()

    def _finish(self):
        result = {"data_steps": self.steps}
        save_config_data(result)
        self.callback(True)
        self.destroy()

    def _on_cancel(self):
        self.capturing = False
        self.callback(False)
        self.destroy()


# ─── Test Mode Window ───

class TestWindow(ctk.CTkToplevel):
    """단계별 테스트: 한 단계씩 클릭하고 결과 확인, 재캡처 가능"""

    def __init__(self, parent: AgentApp, callback):
        super().__init__(parent)
        self.parent_app = parent
        self.callback = callback

        self.title("단계별 테스트")
        self.geometry("400x480")
        self.resizable(False, False)
        self.configure(fg_color=WHITE)
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # 화면 우하단
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        self.geometry(f"400x480+{screen_w - 440}+{screen_h - 560}")

        data = load_config_data()
        self.steps = data.get("data_steps", []) if data else []
        self.current_step = 0

        self._build_ui()
        self._show_step()

    def _build_ui(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=20, pady=16)

        # 진행률
        self._progress_label = ctk.CTkLabel(
            container, text="",
            font=ctk.CTkFont(size=11, weight="bold"), text_color=SLATE_500,
        )
        self._progress_label.pack(anchor="w")

        self._progress_bar = ctk.CTkProgressBar(
            container, height=4, corner_radius=2,
            fg_color=SLATE_200, progress_color=EMERALD_500,
        )
        self._progress_bar.pack(fill="x", pady=(4, 12))

        # 단계 이름
        self._step_label = ctk.CTkLabel(
            container, text="",
            font=ctk.CTkFont(size=14, weight="bold"), text_color=SLATE_900,
            wraplength=340,
        )
        self._step_label.pack(pady=(0, 4))

        # 템플릿 이미지 미리보기
        self._image_frame = ctk.CTkFrame(container, fg_color=SLATE_100, corner_radius=8,
                                          height=80, width=200)
        self._image_frame.pack(pady=(4, 4))
        self._image_frame.pack_propagate(False)

        self._image_label = ctk.CTkLabel(self._image_frame, text="템플릿 없음",
                                          font=ctk.CTkFont(size=11), text_color=SLATE_400)
        self._image_label.pack(expand=True)

        # 결과 표시
        self._result_label = ctk.CTkLabel(
            container, text="",
            font=ctk.CTkFont(size=12), text_color=SLATE_700,
            wraplength=340,
        )
        self._result_label.pack(pady=(4, 8))

        # 좌표 조정 영역
        adjust_frame = ctk.CTkFrame(container, fg_color="transparent")
        adjust_frame.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(adjust_frame, text="클릭 위치 미세 조정:",
                     font=ctk.CTkFont(size=11), text_color=SLATE_500).pack(anchor="w")

        offset_frame = ctk.CTkFrame(adjust_frame, fg_color="transparent")
        offset_frame.pack(fill="x", pady=(4, 0))

        ctk.CTkLabel(offset_frame, text="X:", font=ctk.CTkFont(size=11),
                     text_color=SLATE_700).pack(side="left")
        self._offset_x = ctk.CTkEntry(offset_frame, width=60, height=28,
                                        font=ctk.CTkFont(size=11), placeholder_text="0")
        self._offset_x.pack(side="left", padx=(4, 12))
        self._offset_x.insert(0, "0")

        ctk.CTkLabel(offset_frame, text="Y:", font=ctk.CTkFont(size=11),
                     text_color=SLATE_700).pack(side="left")
        self._offset_y = ctk.CTkEntry(offset_frame, width=60, height=28,
                                        font=ctk.CTkFont(size=11), placeholder_text="0")
        self._offset_y.pack(side="left", padx=(4, 0))
        self._offset_y.insert(0, "0")

        # 버튼 영역
        btn_frame1 = ctk.CTkFrame(container, fg_color="transparent")
        btn_frame1.pack(fill="x", pady=(0, 4))

        self._click_btn = ctk.CTkButton(
            btn_frame1, text="이 단계 클릭 실행", height=38, corner_radius=8,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=EMERALD_600, hover_color=EMERALD_500,
            command=self._on_click_test,
        )
        self._click_btn.pack(fill="x")

        btn_frame2 = ctk.CTkFrame(container, fg_color="transparent")
        btn_frame2.pack(fill="x", pady=(0, 4))

        self._recapture_btn = ctk.CTkButton(
            btn_frame2, text="재캡처 (5초)", height=34, corner_radius=8,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=AMBER_500, hover_color="#d97706",
            command=self._on_recapture,
        )
        self._recapture_btn.pack(side="left", expand=True, fill="x", padx=(0, 3))

        self._next_btn = ctk.CTkButton(
            btn_frame2, text="다음 단계 →", height=34, corner_radius=8,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=BLUE_500, hover_color="#2563eb",
            command=self._on_next,
        )
        self._next_btn.pack(side="left", expand=True, fill="x", padx=(3, 0))

    def _show_step(self):
        # 건너뛸 단계 넘기기
        while self.current_step < len(self.steps):
            step = self.steps[self.current_step]
            if step.get("skip") or step.get("name") == "data_check":
                self.current_step += 1
                continue
            break

        if self.current_step >= len(self.steps):
            self._result_label.configure(text="모든 단계 테스트 완료!", text_color=EMERALD_600)
            self._click_btn.configure(state="disabled")
            self._recapture_btn.configure(state="disabled")
            self._next_btn.configure(text="닫기", command=self._on_close)
            return

        step = self.steps[self.current_step]
        total = len([s for s in self.steps if not s.get("skip") and s.get("name") != "data_check"])
        active_idx = len([s for s in self.steps[:self.current_step]
                         if not s.get("skip") and s.get("name") != "data_check"]) + 1

        self._progress_label.configure(text=f"단계 {active_idx}/{total}")
        self._progress_bar.set(active_idx / total)
        self._step_label.configure(text=step["label"])
        self._result_label.configure(text=f"좌표: ({step.get('x', '?')}, {step.get('y', '?')})",
                                      text_color=SLATE_500)

        # 오프셋 초기화
        self._offset_x.delete(0, "end")
        self._offset_x.insert(0, "0")
        self._offset_y.delete(0, "end")
        self._offset_y.insert(0, "0")

        # 템플릿 이미지 표시
        from dentweb_runner import _get_template_path
        template_path = _get_template_path(step["name"])
        if os.path.exists(template_path):
            try:
                from PIL import Image, ImageTk
                img = Image.open(template_path)
                # 프레임 크기에 맞게 리사이즈
                img.thumbnail((200, 70))
                self._tk_image = ImageTk.PhotoImage(img)
                self._image_label.configure(image=self._tk_image, text="")
            except Exception:
                self._image_label.configure(image=None, text="이미지 로드 실패")
        else:
            self._image_label.configure(image=None, text="템플릿 없음 (재캡처 필요)")

        self._click_btn.configure(state="normal")
        self._recapture_btn.configure(state="normal")

    def _on_click_test(self):
        """현재 단계의 이미지를 찾아 클릭"""
        step = self.steps[self.current_step]
        from dentweb_runner import _find_and_click

        # 오프셋 적용
        try:
            ox = int(self._offset_x.get() or 0)
            oy = int(self._offset_y.get() or 0)
        except ValueError:
            ox, oy = 0, 0

        if ox != 0 or oy != 0:
            # 오프셋이 있으면 좌표에 반영하고 저장
            if step.get("x") is not None:
                step["x"] += ox
                step["y"] += oy
                self._result_label.configure(
                    text=f"오프셋 적용: ({step['x']}, {step['y']})", text_color=BLUE_500)
                # 설정 저장
                save_config_data({"data_steps": self.steps})

        # topmost 해제하고 클릭
        self.attributes("-topmost", False)
        time.sleep(0.3)

        result_msgs = []
        def log_cb(msg):
            result_msgs.append(msg)

        success = _find_and_click(step, log_callback=log_cb)
        time.sleep(0.5)

        self.attributes("-topmost", True)
        self.lift()

        msg = "\n".join(result_msgs) if result_msgs else ("클릭 성공" if success else "클릭 실패")
        color = EMERALD_600 if success else ROSE_500
        self._result_label.configure(text=msg, text_color=color)

    def _on_recapture(self):
        """현재 단계 재캡처: 5초 카운트다운 후 마우스 위치 + 이미지 저장"""
        self._recapture_btn.configure(state="disabled", text="카운트다운...")
        self._click_btn.configure(state="disabled")
        self._result_label.configure(text="5초 안에 해당 위치에 마우스를 올려주세요!",
                                      text_color=AMBER_500)
        threading.Thread(target=self._do_recapture, daemon=True).start()

    def _do_recapture(self):
        for remaining in range(5, 0, -1):
            self.after(0, lambda r=remaining: self._result_label.configure(
                text=f"{r}초... 해당 위치에 마우스를 올려주세요!", text_color=AMBER_500))
            time.sleep(1)

        x, y = pyautogui.position()
        step = self.steps[self.current_step]
        step["x"] = x
        step["y"] = y

        # 이미지 템플릿 캡처
        _capture_template(x, y, step["name"])

        # 설정 저장
        save_config_data({"data_steps": self.steps})

        self.after(0, lambda: self._result_label.configure(
            text=f"재캡처 완료: ({x}, {y}) + 이미지 저장", text_color=EMERALD_600))
        self.after(0, lambda: self._recapture_btn.configure(state="normal", text="재캡처 (5초)"))
        self.after(0, lambda: self._click_btn.configure(state="normal"))
        self.after(0, self._show_step)

    def _on_next(self):
        self.current_step += 1
        self._show_step()

    def _on_close(self):
        self.callback()
        self.destroy()


def run_gui():
    app = AgentApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()
