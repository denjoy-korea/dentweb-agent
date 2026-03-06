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
        result = api.claim_run()
        log.info("서버 연결 성공")
    except Exception as e:
        log.error(f"서버 연결 실패: {e}")
        print(f"\n[ERROR] 서버 연결 실패: {e}")
        print("토큰이 올바른지 확인해주세요.")
        print("config.json을 삭제하고 다시 실행하면 토큰을 재입력할 수 있습니다.")
        input("\nEnter를 누르면 종료합니다...")
        sys.exit(1)

    log.info("에이전트 시작 - 서버 폴링 중...")
    print("에이전트가 실행 중입니다. 이 창을 닫지 마세요.")
    print("중지하려면 Ctrl+C를 누르세요.\n")

    while True:
        try:
            result = api.claim_run()
            if not result.get("should_run"):
                time.sleep(cfg["poll_interval_seconds"])
                continue

            log.info(f"실행 시작: reason={result.get('reason')}")
            upload_url = result.get("upload_url", f"{cfg['server_url']}/dentweb-upload")

            # 1. DentWeb 자동화 -> Excel 다운로드
            excel_path = runner.download_excel()
            if not excel_path:
                api.report_run("no_data", "Excel 파일을 찾을 수 없습니다")
                log.warn("Excel 다운로드 실패")
                continue

            # 2. 서버로 업로드
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
