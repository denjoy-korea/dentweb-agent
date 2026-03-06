"""DentWeb Automation Agent - 진입점 및 메인 루프"""

import sys
import time
import traceback
from config import load_config
from api_client import ApiClient
from dentweb_runner import DentwebRunner
from logger import AgentLogger


def main():
    cfg = load_config("config.json")
    api = ApiClient(cfg["server_url"], cfg["agent_token"])
    runner = DentwebRunner(cfg)
    log = AgentLogger(cfg.get("log_file", "agent.log"), cfg.get("log_max_lines", 1000))

    # 서버 연결 확인
    try:
        api.claim_run()
        log.info("서버 연결 성공")
    except Exception as e:
        log.error(f"서버 연결 실패: {e}")
        print(f"\n[ERROR] 서버 연결 실패: {e}")
        print("토큰이 올바른지 확인해주세요.")
        print("config.json을 삭제하고 다시 실행하면 토큰을 재입력할 수 있습니다.")
        input("\nEnter를 누르면 종료합니다...")
        sys.exit(1)

    # 클릭 좌표가 없으면 학습 모드 실행
    if not runner.is_configured():
        print()
        print("[안내] 덴트웹 클릭 좌표가 설정되지 않았습니다.")
        print("자동화를 위해 덴트웹 화면의 버튼 위치를 학습해야 합니다.")
        print()
        runner.teach()

    log.info("에이전트 시작 - 서버 폴링 중...")
    print()
    print("=" * 50)
    print("  에이전트가 실행 중입니다")
    print("  이 창을 닫지 마세요")
    print("=" * 50)
    print()
    print("  서버에서 실행 요청을 대기 중입니다...")
    print("  앱에서 '지금 실행 요청' 버튼을 누르면 자동으로 시작됩니다.")
    print()
    print("  [단축키]")
    print("  Ctrl+C  = 에이전트 종료")
    print("  좌표 재설정 = config.json 옆의 dentweb_steps.json 삭제 후 재시작")
    print()

    while True:
        try:
            result = api.claim_run()
            if not result.get("should_run"):
                time.sleep(cfg["poll_interval_seconds"])
                continue

            log.info(f"실행 시작: reason={result.get('reason')}")

            if not runner.is_configured():
                api.report_run("failed", "클릭 좌표가 설정되지 않았습니다. 에이전트를 재시작하세요.")
                log.error("클릭 좌표 미설정")
                continue

            # 1. DentWeb 자동화 -> Excel 다운로드
            log.info("덴트웹 자동화 시작...")
            excel_path = runner.download_excel()
            if not excel_path:
                api.report_run("no_data", "오늘 수술 기록이 없습니다")
                log.info("오늘 수술 기록 없음 - 작업 완료")
                continue

            # 2. 서버로 업로드
            upload_url = result.get("upload_url", f"{cfg['server_url']}/dentweb-upload")
            upload_result = api.upload_file(upload_url, excel_path)
            if upload_result.get("success"):
                inserted = upload_result.get("inserted", 0)
                skipped = upload_result.get("skipped", 0)
                api.report_run("success", f"{inserted}건 업로드, {skipped}건 스킵")
                log.info(f"완료: {inserted}건 업로드, {skipped}건 스킵")
            else:
                error_msg = upload_result.get("error", "업로드 실패")
                api.report_run("failed", error_msg)
                log.error(f"업로드 실패: {error_msg}")

            # 3. 임시 파일 정리
            runner.cleanup(excel_path)

        except KeyboardInterrupt:
            log.info("에이전트 종료 (사용자 중단)")
            print("\n에이전트를 종료합니다.")
            break
        except Exception as e:
            log.error(f"에러: {e}\n{traceback.format_exc()}")
            try:
                api.report_run("failed", str(e)[:1000])
            except Exception:
                pass

        time.sleep(cfg["poll_interval_seconds"])


if __name__ == "__main__":
    main()
