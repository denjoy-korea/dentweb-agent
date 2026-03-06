"""로컬 파일 로깅"""

import os
import datetime


class AgentLogger:
    def __init__(self, log_file: str = "agent.log", max_lines: int = 1000):
        self.log_file = log_file
        self.max_lines = max_lines

    def _write(self, level: str, message: str):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{timestamp}] [{level}] {message}"
        print(line)

        try:
            with open(self.log_file, "a", encoding="utf-8") as f:
                f.write(line + "\n")
            self._truncate_if_needed()
        except OSError:
            pass

    def _truncate_if_needed(self):
        try:
            with open(self.log_file, "r", encoding="utf-8") as f:
                lines = f.readlines()
            if len(lines) > self.max_lines:
                with open(self.log_file, "w", encoding="utf-8") as f:
                    f.writelines(lines[-self.max_lines :])
        except OSError:
            pass

    def info(self, message: str):
        self._write("INFO", message)

    def error(self, message: str):
        self._write("ERROR", message)

    def warn(self, message: str):
        self._write("WARN", message)
