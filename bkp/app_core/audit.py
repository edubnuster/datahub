# -*- coding: utf-8 -*-
from datetime import datetime
from .constants import AUDIT_PATH

class AuditLogger:
    @staticmethod
    def write(username: str, action: str, detail: str = ""):
        try:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            line = f"{ts} | usuario={username or '-'} | acao={action}"
            if detail:
                line += f" | detalhe={detail}"
            line += "\n"
            with open(AUDIT_PATH, "a", encoding="utf-8") as f:
                f.write(line)
        except Exception:
            pass
