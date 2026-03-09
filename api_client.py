"""HTTP 클라이언트: claim_run, report_run, upload"""

import os
import requests


class ApiClient:
    def __init__(self, server_url: str, agent_token: str):
        self.automation_url = f"{server_url}/dentweb-automation"
        self.agent_token = agent_token
        self.headers = {
            "Authorization": f"Bearer {agent_token}",
            "Content-Type": "application/json",
        }

    def ping(self) -> dict:
        """연결 확인 전용 (claim 없이 상태만 조회)"""
        resp = requests.post(
            self.automation_url,
            json={"action": "get_state"},
            headers=self.headers,
            timeout=30,
        )
        resp.raise_for_status()
        return resp.json()

    def claim_run(self) -> dict:
        resp = requests.post(
            self.automation_url,
            json={"action": "claim_run"},
            headers=self.headers,
            timeout=30,
        )
        resp.raise_for_status()
        return resp.json()

    def report_run(self, status: str, message: str = "") -> dict:
        resp = requests.post(
            self.automation_url,
            json={"action": "report_run", "status": status, "message": message},
            headers=self.headers,
            timeout=30,
        )
        resp.raise_for_status()
        return resp.json()

    def get_state(self) -> dict:
        """현재 설정/상태 조회"""
        resp = requests.post(
            self.automation_url,
            json={"action": "get_state"},
            headers=self.headers,
            timeout=30,
        )
        resp.raise_for_status()
        return resp.json()

    def save_settings(self, enabled: bool, scheduled_time: str) -> dict:
        """에이전트에서 자동 실행 설정 저장"""
        resp = requests.post(
            self.automation_url,
            json={"action": "agent_save_settings", "enabled": enabled, "scheduled_time": scheduled_time},
            headers=self.headers,
            timeout=30,
        )
        resp.raise_for_status()
        return resp.json()

    def upload_file(self, upload_url: str, file_path: str) -> dict:
        filename = os.path.basename(file_path)
        with open(file_path, "rb") as f:
            resp = requests.post(
                upload_url,
                headers={"Authorization": f"Bearer {self.agent_token}"},
                files={"file": (filename, f)},
                timeout=60,
            )
        resp.raise_for_status()
        return resp.json()
