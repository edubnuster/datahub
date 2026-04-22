from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from .documents_history import DocumentsHistory


def _parse_dt(value: str) -> Optional[datetime]:
    value = str(value or "").strip()
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except Exception:
        return None


def _normalize_level(value: Any) -> str:
    raw = str(value or "").strip().lower()
    if raw in ("info", "informacao", "informação"):
        return "info"
    if raw in ("warn", "warning", "alerta", "atencao", "atenção"):
        return "warn"
    if raw in ("error", "erro", "err", "falha", "failed", "exception"):
        return "error"
    return "info"


def _as_dt_key(value: str) -> Tuple[int, float]:
    dt = _parse_dt(value)
    if dt is None:
        return (0, 0.0)
    try:
        return (1, dt.timestamp())
    except Exception:
        return (1, 0.0)


@dataclass(frozen=True)
class CentralAlert:
    created_at: str
    level: str
    source: str
    title: str
    message: str


def _auto_docs_run_to_alert(run_status: str, result: Dict[str, Any], finished_at: str) -> CentralAlert:
    run_status = str(run_status or "").strip().lower()
    docs_failed = int(result.get("docs_failed") or 0) if isinstance(result, dict) else 0
    failed_emails = int(result.get("failed_emails") or 0) if isinstance(result, dict) else 0
    docs_no_email = int(result.get("docs_no_email") or 0) if isinstance(result, dict) else 0

    if run_status in ("error", "failed"):
        level = "error"
    elif docs_failed > 0 or failed_emails > 0:
        level = "error"
    elif docs_no_email > 0:
        level = "warn"
    else:
        level = "info"

    title = "Envio automático de documentos"
    if run_status == "dry_run":
        title = "Envio automático de documentos (simulação)"
    elif run_status and run_status not in ("ok", "running"):
        title = f"Envio automático de documentos ({run_status})"

    message = ""
    if isinstance(result, dict) and result:
        message = (
            f"Encontrados: {int(result.get('discovered') or 0)} | "
            f"Pendentes antes: {int(result.get('pending_before') or 0)} | "
            f"E-mails enviados: {int(result.get('emails_sent') or 0)} (falhas: {int(result.get('failed_emails') or 0)}) | "
            f"Docs enviados: {int(result.get('docs_sent') or 0)} (falhas: {int(result.get('docs_failed') or 0)}) | "
            f"Sem e-mail: {int(result.get('docs_no_email') or 0)}"
        )
        err = str(result.get("error") or "").strip()
        if err:
            message = (message + " | " if message else "") + f"Erro: {err}"
    else:
        message = "Execução registrada."

    return CentralAlert(
        created_at=str(finished_at or "").strip(),
        level=level,
        source="envio_auto_docs",
        title=title,
        message=message,
    )


def list_central_alerts(cfg: Dict[str, Any], *, limit: int = 500) -> List[CentralAlert]:
    limit = max(50, min(2000, int(limit or 500)))
    cfg = cfg if isinstance(cfg, dict) else {}
    history = DocumentsHistory()

    out: List[CentralAlert] = []

    try:
        for ev in history.list_events(limit=limit):
            out.append(
                CentralAlert(
                    created_at=str(getattr(ev, "created_at", "") or "").strip(),
                    level=_normalize_level(getattr(ev, "level", "info")),
                    source=str(getattr(ev, "source", "") or "").strip(),
                    title=str(getattr(ev, "title", "") or "").strip(),
                    message=str(getattr(ev, "message", "") or "").strip(),
                )
            )
    except Exception:
        pass

    try:
        for p in history.list_problems():
            status = str(getattr(p, "status", "") or "").strip().lower()
            level = "error" if status == "failed" else "warn"
            created_at = str(getattr(p, "last_attempt_at", "") or "").strip() or str(getattr(p, "generated_at", "") or "").strip()
            title = "Problema no envio"
            if status == "no_email":
                title = "Sem e-mail do cliente"
            message = str(getattr(p, "error", "") or "").strip()
            if not message and status == "no_email":
                message = "Não foi possível enviar porque o cliente está sem e-mail."
            msg2 = f"{message}".strip()
            extra = []
            doc = str(getattr(p, "documento", "") or "").strip()
            bg = str(getattr(p, "boleto_grid", "") or "").strip()
            to_email = str(getattr(p, "customer_email", "") or "").strip()
            if doc:
                extra.append(f"Documento: {doc}")
            if bg:
                extra.append(f"ID Boleto: {bg}")
            if to_email:
                extra.append(f"Para: {to_email}")
            if extra:
                msg2 = (msg2 + " | " if msg2 else "") + " | ".join(extra)
            out.append(CentralAlert(created_at=created_at, level=level, source="docs", title=title, message=msg2))
    except Exception:
        pass

    try:
        for r in history.list_runs(limit=200):
            status = str(getattr(r, "status", "") or "").strip()
            finished_at = str(getattr(r, "finished_at", "") or "").strip() or str(getattr(r, "started_at", "") or "").strip()
            result = getattr(r, "result", None)
            result = result if isinstance(result, dict) else {}
            out.append(_auto_docs_run_to_alert(status, result, finished_at))
    except Exception:
        pass

    try:
        for a in (cfg.get("financeiro_agendas", []) or []):
            if not isinstance(a, dict):
                continue
            last_run_at = str(a.get("last_run_at") or "").strip()
            if not last_run_at:
                continue
            name = str(a.get("name") or "Alerta de vencimento").strip()
            result = a.get("last_result") if isinstance(a.get("last_result"), dict) else {}
            emails_sent = int(result.get("emails_sent") or 0)
            skipped = int(result.get("skipped_no_email") or 0)
            failed = int(result.get("failed") or 0)
            late = int(a.get("last_late_minutes") or 0)
            out_of_time = bool(a.get("last_out_of_time") or False)
            level = "info"
            if failed > 0:
                level = "error"
            elif skipped > 0:
                level = "warn"
            title = f"Alerta de vencimento: {name}"
            message = f"Enviados: {emails_sent} | Sem e-mail: {skipped} | Falhas: {failed}"
            if out_of_time:
                message = message + f" | Atraso: {late} min"
            due_dates = str(a.get("last_due_date") or "").strip()
            if due_dates:
                message = message + f" | Datas: {due_dates}"
            out.append(CentralAlert(created_at=last_run_at, level=level, source="alertas_vencimento", title=title, message=message))
    except Exception:
        pass

    out.sort(key=lambda a: _as_dt_key(getattr(a, "created_at", "")), reverse=True)
    if len(out) > limit:
        out = out[:limit]
    return out


def count_unseen_central_alerts(cfg: Dict[str, Any], *, limit: int = 500) -> int:
    cfg = cfg if isinstance(cfg, dict) else {}
    last_seen = _parse_dt(str(cfg.get("central_alerts_last_seen_at") or "").strip() or "")
    if last_seen is None:
        last_seen = datetime.fromtimestamp(0)
    alerts = list_central_alerts(cfg, limit=limit)
    n = 0
    for a in alerts:
        dt = _parse_dt(getattr(a, "created_at", ""))
        if dt is not None and dt > last_seen:
            n += 1
    return int(n)


def count_problem_central_alerts(cfg: Dict[str, Any]) -> int:
    try:
        history = DocumentsHistory()
        return len(history.list_problems() or [])
    except Exception:
        return 0
