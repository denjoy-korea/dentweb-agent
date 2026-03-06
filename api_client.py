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

    def claim_run(self) -> dict:
        resp = requests.post(
            self.automation_url,
            json={"action": "claim_run"},
            headers=self.headers,
            timeout=10,
        )
        resp.raise_for_status()
        return resp.json()

    def report_run(self, status: str, message: str = "") -> dict:
        resp = requests.post(
            self.automation_url,
            json={"action": "report_run", "status": status, "message": message},
            headers=self.headers,
            timeout=10,
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
