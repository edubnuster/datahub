# -*- coding: utf-8 -*-
from datetime import datetime
import logging
from .constants import AUDIT_PATH

class AuditLogger:
    @staticmethod
    def write(username: str, action: str, detail: str = ""):
        try:
            msg = f"usuario={username or '-'} | acao={action}"
            if detail:
                msg += f" | detalhe={detail}"
            logger = logging.getLogger("audit")
            if logger.handlers or (logger.propagate and logging.getLogger().handlers):
                logger.info(msg)
                return
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            line = f"{ts} | {msg}\n"
            try:
                AUDIT_PATH.parent.mkdir(parents=True, exist_ok=True)
            except Exception:
                pass
            with open(AUDIT_PATH, "a", encoding="utf-8") as f:
                f.write(line)
        except Exception:
            pass
