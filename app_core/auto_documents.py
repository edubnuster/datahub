# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import time
from datetime import datetime, timedelta
from email.message import EmailMessage
import logging
import smtplib
import ssl
from typing import Any, Dict, List, Optional, Tuple

from .database import Database
from .documents_history import DocumentsHistory
from .helpers import AppError
from .models import InvoiceRow
from .logging_setup import get_docs_generated_logger, get_docs_sent_logger, get_system_logger


def _parse_iso_datetime(value: str) -> Optional[datetime]:
    value = str(value or "").strip()
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except Exception:
        return None


def _mime_parts_from_filename(filename: str) -> Tuple[str, str]:
    name = str(filename or "").strip().lower()
    ext = name.rsplit(".", 1)[-1] if "." in name else ""
    if ext == "pdf":
        return "application", "pdf"
    if ext == "png":
        return "image", "png"
    if ext in ("jpg", "jpeg"):
        return "image", "jpeg"
    if ext in ("tif", "tiff"):
        return "image", "tiff"
    return "application", "octet-stream"


def _smtp_send_message(cfg: Dict[str, Any], msg: EmailMessage) -> None:
    smtp_cfg = cfg.get("smtp", {})
    smtp_email = str(smtp_cfg.get("email", "")).strip()
    smtp_host = str(smtp_cfg.get("host", "")).strip()
    smtp_password = str(smtp_cfg.get("password", "")).strip()
    smtp_port = int(smtp_cfg.get("port", 587) or 587)
    if not smtp_email or not smtp_host or not smtp_password or not smtp_port:
        raise AppError("SMTP não configurado.")
    if smtp_port == 465:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_host, smtp_port, context=context, timeout=30) as server:
            server.login(smtp_email, smtp_password)
            server.send_message(msg)
    else:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as server:
            server.ehlo()
            try:
                server.starttls(context=ssl.create_default_context())
                server.ehlo()
            except Exception:
                pass
            server.login(smtp_email, smtp_password)
            server.send_message(msg)


def _generate_boleto_pdf_if_needed(boleto_data: Dict[str, Any], inv: InvoiceRow) -> Tuple[Optional[bytes], str]:
    if not (boleto_data or {}).get("exists"):
        return None, ""
    attachment_data = boleto_data.get("attachment_data")
    filename = boleto_data.get("filename") or f"boleto_{inv.invoice_id}.pdf"
    if attachment_data:
        return attachment_data, filename
    from ui import build_boleto_pdf_bytes
    return build_boleto_pdf_bytes(boleto_data, inv), filename


def _subject(customer_name: str, window_text: str) -> str:
    return f"Documentos gerados - {customer_name} - {window_text}"


class _AutoDocsRunLock:
    def __init__(self, lock_path: str, *, stale_seconds: int = 3 * 60 * 60):
        self.lock_path = str(lock_path or "").strip()
        self.stale_seconds = max(60, int(stale_seconds or 0))
        self._acquired = False

    def _is_stale(self) -> bool:
        try:
            mtime = os.path.getmtime(self.lock_path)
            return (time.time() - float(mtime)) > float(self.stale_seconds)
        except Exception:
            return False

    def acquire(self) -> bool:
        if not self.lock_path:
            return True
        flags = os.O_CREAT | os.O_EXCL | os.O_WRONLY
        try:
            fd = os.open(self.lock_path, flags)
            try:
                os.write(fd, str(datetime.now().isoformat(timespec="seconds")).encode("utf-8"))
            finally:
                os.close(fd)
            self._acquired = True
            return True
        except FileExistsError:
            if self._is_stale():
                try:
                    os.remove(self.lock_path)
                except Exception:
                    return False
                return self.acquire()
            return False
        except OSError as e:
            try:
                if getattr(e, "winerror", None) == 183:
                    if self._is_stale():
                        try:
                            os.remove(self.lock_path)
                        except Exception:
                            return False
                        return self.acquire()
                    return False
            except Exception:
                pass
            raise

    def release(self) -> None:
        if not self._acquired:
            return
        try:
            os.remove(self.lock_path)
        except Exception:
            pass
        self._acquired = False

    def __enter__(self):
        ok = self.acquire()
        if not ok:
            raise AppError("Envio automático já está em execução.")
        return self

    def __exit__(self, exc_type, exc, tb):
        self.release()
        return False


def run_auto_documents(
    cfg: Dict[str, Any],
    *,
    dry_run: bool,
    user_label: str,
    now: Optional[datetime] = None,
) -> Dict[str, Any]:
    system_logger = get_system_logger()
    docs_generated_logger = get_docs_generated_logger()
    docs_sent_logger = get_docs_sent_logger()

    now = now or datetime.now()

    auto_cfg = cfg.get("financeiro_envio_auto_documentos")
    if not isinstance(auto_cfg, dict):
        auto_cfg = {}
        cfg["financeiro_envio_auto_documentos"] = auto_cfg

    enabled = bool(auto_cfg.get("enabled", False))
    try:
        interval_hours = int(auto_cfg.get("interval_hours") or 4)
    except Exception:
        interval_hours = 4
    interval_hours = max(1, min(72, interval_hours))

    if not enabled and not dry_run:
        return {"enabled": False, "skipped": True, "reason": "disabled"}

    last_scan_end = _parse_iso_datetime(auto_cfg.get("last_scan_end") or "")
    if not last_scan_end:
        last_scan_end = now - timedelta(hours=interval_hours)
    overlap = timedelta(minutes=5)
    window_start = last_scan_end - overlap
    window_end = now
    window_text = f"{window_start.strftime('%d/%m/%Y %H:%M')} até {window_end.strftime('%d/%m/%Y %H:%M')}"

    history = DocumentsHistory()
    lock_path = str(getattr(history, "db_path", "")).strip()
    if lock_path:
        lock_path = lock_path + ".lock"

    if not dry_run:
        try:
            run_lock = _AutoDocsRunLock(lock_path)
        except Exception:
            run_lock = None
    else:
        run_lock = None

    if run_lock is not None and not dry_run:
        if not run_lock.acquire():
            return {"enabled": bool(enabled), "dry_run": bool(dry_run), "skipped": True, "reason": "already_running"}

    run_id = history.start_run(window_start=window_start, window_end=window_end)

    discovered_rows: List[Dict[str, Any]] = []
    try:
        discovered_rows = Database(cfg).list_generated_boletos(window_start, window_end)
    except Exception as e:
        history.finish_run(run_id, "error", {"error": str(e)})
        raise

    for r in discovered_rows:
        history.upsert_generated(r)
        docs_generated_logger.info(
            "gerado boleto_grid=%s movto_id=%s customer_id=%s customer_email=%s documento=%s generated_at=%s",
            str(r.get("boleto_grid") or ""),
            str(r.get("movto_id") or ""),
            str(r.get("customer_id") or ""),
            str(r.get("customer_email") or ""),
            str(r.get("documento") or ""),
            str(r.get("generated_at") or ""),
        )

    smtp_cfg = cfg.get("smtp", {})
    smtp_email = str(smtp_cfg.get("email", "")).strip()

    emails_sent = 0
    failed_emails = 0
    docs_sent = 0
    docs_failed = 0
    docs_no_email = 0
    attachments_total = 0
    missing_total = 0

    extra_body = str(auto_cfg.get("extra_body") or "").strip()

    delay_seconds = int(smtp_cfg.get("delay_seconds", 5))
    first_email = True
    emails_planned = 0
    pending_snapshot = history.list_pending(limit=5000)
    pending_before = len(pending_snapshot)

    try:
        batch_size = int(auto_cfg.get("pending_batch_size") or 2000)
    except Exception:
        batch_size = 2000
    batch_size = max(50, min(5000, batch_size))

    try:
        no_email_retry_hours = int(auto_cfg.get("no_email_retry_hours") or 24)
    except Exception:
        no_email_retry_hours = 24
    no_email_retry_hours = max(1, min(24 * 30, no_email_retry_hours))

    retryable = history.list_retryable(limit=batch_size, no_email_retry_hours=no_email_retry_hours)
    retryable_grids = [p.boleto_grid for p in retryable if p.boleto_grid]
    if retryable_grids:

        rows_by_grid: Dict[str, Dict[str, Any]] = {}
        try:
            pending_rows = Database(cfg).list_boletos_by_grids(retryable_grids, include_closed=True)
            for r in pending_rows:
                bg = str(r.get("boleto_grid") or "").strip()
                if bg:
                    rows_by_grid[bg] = dict(r)
        except Exception:
            rows_by_grid = {}

        to_process: List[Dict[str, Any]] = []
        closed_grids: List[str] = []
        missing_grids: List[str] = []
        for p in retryable:
            bg = str(p.boleto_grid or "").strip()
            if not bg:
                continue
            r = rows_by_grid.get(bg)
            if not r:
                missing_grids.append(bg)
                continue
            try:
                saldo = float(r.get("saldo_em_aberto") or 0)
            except Exception:
                saldo = 0.0
            if saldo <= 0:
                closed_grids.append(bg)
                continue
            to_process.append(r)

        if not dry_run:
            if closed_grids:
                history.mark_closed(closed_grids)
            if missing_grids:
                history.mark_failed(missing_grids, error="Boleto não localizado no banco ou título não elegível para envio.")

        boleto_grids = [str(r.get("boleto_grid") or "").strip() for r in to_process if str(r.get("boleto_grid") or "").strip()]
        payload_map: Dict[Any, Dict[str, Any]] = {}
        if boleto_grids:
            try:
                payload_map = Database(cfg).get_boletos_email_payload_by_boleto_grids(boleto_grids)
            except Exception:
                payload_map = {}

        grouped: Dict[str, Dict[str, Any]] = {}
        no_email_grids: List[str] = []
        for r in to_process:
            bg = str(r.get("boleto_grid") or "").strip()
            if not bg:
                continue
            to_email = str(r.get("customer_email") or "").strip()
            cid = r.get("customer_id")
            key = str(cid) if cid not in (None, "", 0, "0") else (to_email or str(r.get("cliente") or "") or bg)
            if not to_email:
                no_email_grids.append(bg)
                continue
            item = grouped.get(key)
            if not item:
                item = {"customer_name": str(r.get("cliente") or "").strip(), "to_email": to_email, "rows": []}
                grouped[key] = item
            item["rows"].append(r)

        if no_email_grids and not dry_run:
            history.mark_no_email(no_email_grids)
        docs_no_email += len(no_email_grids)

        if grouped:
            emails_planned += len([g for g in grouped.values() if (g.get("to_email") or "").strip()])

        for g in grouped.values():
            rows = g.get("rows") or []
            if not rows:
                continue
            to_email = str(g.get("to_email") or "").strip()
            if not to_email:
                continue

            invoices: List[InvoiceRow] = []
            items: List[Tuple[str, InvoiceRow]] = []
            attachments: List[Tuple[bytes, str, str]] = []
            missing = 0

            for r in rows:
                bg = str(r.get("boleto_grid") or "").strip()
                inv = InvoiceRow(
                    invoice_id=(r.get("documento") or r.get("movto_id")),
                    company=str(r.get("empresa") or "").strip(),
                    customer_id=r.get("customer_id"),
                    customer_code=r.get("codigo_cliente"),
                    customer_name=str(r.get("cliente") or "").strip(),
                    account_code=str(r.get("conta") or "").strip(),
                    account_name=str(r.get("conta_nome") or "").strip(),
                    issue_date=r.get("data"),
                    due_date=r.get("vencto"),
                    amount=r.get("valor"),
                    discount_amount=r.get("valor_desconto"),
                    paid_amount=r.get("valor_baixado"),
                    open_balance=r.get("saldo_em_aberto"),
                    customer_email=str(r.get("customer_email") or "").strip(),
                    movto_id=r.get("movto_id"),
                )
                invoices.append(inv)
                if bg:
                    items.append((bg, inv))

            signature_map: Dict[Any, Dict[str, Any]] = {}
            try:
                movto_ids = [inv.movto_id for inv in invoices if getattr(inv, "movto_id", None) not in (None, "", 0, "0")]
                signature_map = Database(cfg).get_sale_signatures_pdf_bulk(movto_ids)
            except Exception:
                signature_map = {}

            for bg, inv in items:
                boleto_data = payload_map.get(bg) or {}
                try:
                    if boleto_data.get("exists"):
                        attachment_data = boleto_data.get("attachment_data")
                        filename = boleto_data.get("filename") or f"boleto_{inv.invoice_id}.pdf"
                        if attachment_data:
                            attachments.append((attachment_data, filename, bg))
                        else:
                            generated, fname = _generate_boleto_pdf_if_needed(boleto_data, inv)
                            if generated:
                                attachments.append((generated, fname or filename, bg))
                            else:
                                missing += 1
                    else:
                        missing += 1
                except Exception:
                    missing += 1

                sig = signature_map.get(getattr(inv, "movto_id", None)) or {}
                sig_added = False
                for a in (sig.get("attachments") or []):
                    data = a.get("data")
                    name = a.get("filename")
                    if data and name:
                        attachments.append((data, name, ""))
                        sig_added = True
                sig_bytes = sig.get("attachment_data")
                if not sig_added and sig.get("exists") and sig_bytes:
                    attachments.append((sig_bytes, sig.get("filename") or f"assinatura_{inv.movto_id}", ""))

            attachments_total += len(attachments)
            missing_total += missing
            if dry_run:
                continue

            from ui import build_agenda_email_body

            subject = _subject(invoices[0].customer_name if invoices else "", window_text)

            purchase_map = {}
            try:
                movto_ids = [inv.movto_id for inv in invoices if getattr(inv, "movto_id", None) not in (None, "", 0, "0")]
                purchase_map = Database(cfg).get_purchase_info_bulk(movto_ids)
            except Exception:
                purchase_map = {}

            text_body, html_body = build_agenda_email_body(
                invoices[0].customer_name if invoices else "",
                window_text,
                invoices,
                missing,
                extra_body,
                context_label="Período de geração",
                purchase_info_map=purchase_map,
            )

            max_attachments = 18
            max_bytes = 15 * 1024 * 1024
            batches: List[List[Tuple[bytes, str, str]]] = []
            current: List[Tuple[bytes, str, str]] = []
            current_bytes = 0
            for data, name, bg in attachments:
                size = len(data) if data else 0
                if current and (len(current) >= max_attachments or (current_bytes + size) > max_bytes):
                    batches.append(current)
                    current = []
                    current_bytes = 0
                current.append((data, name, bg))
                current_bytes += size
            if current or not attachments:
                batches.append(current)

            for idx, batch in enumerate(batches, start=1):
                if not first_email and delay_seconds > 0:
                    time.sleep(delay_seconds)
                first_email = False

                msg = EmailMessage()
                msg["From"] = smtp_email
                msg["To"] = to_email
                msg["Subject"] = subject if len(batches) == 1 else f"{subject} ({idx}/{len(batches)})"
                msg.set_content(text_body)
                msg.add_alternative(html_body, subtype="html")
                grids_in_batch: List[str] = []
                for data, name, bg in batch:
                    if not data:
                        continue
                    maintype, subtype = _mime_parts_from_filename(name)
                    msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=name)
                    if bg:
                        grids_in_batch.append(bg)
                try:
                    _smtp_send_message(cfg, msg)
                    emails_sent += 1
                    if grids_in_batch:
                        history.mark_sent(grids_in_batch, to_email=to_email)
                        docs_sent += len(grids_in_batch)
                        for bg in grids_in_batch:
                            docs_sent_logger.info(
                                "enviado boleto_grid=%s para=%s cliente=%s periodo=%s",
                                bg,
                                to_email,
                                invoices[0].customer_name if invoices else "",
                                window_text,
                            )
                except Exception as e:
                    failed_emails += 1
                    if grids_in_batch:
                        history.mark_failed(grids_in_batch, error=str(e))
                        docs_failed += len(grids_in_batch)
                    system_logger.error("Falha ao enviar docs para=%s erro=%s", to_email, e)

    result = {
        "enabled": bool(enabled),
        "dry_run": bool(dry_run),
        "window_start": window_start.isoformat(timespec="seconds"),
        "window_end": window_end.isoformat(timespec="seconds"),
        "discovered": len(discovered_rows),
        "pending_before": int(pending_before),
        "emails_planned": int(emails_planned),
        "emails_sent": int(emails_sent),
        "failed_emails": int(failed_emails),
        "docs_sent": int(docs_sent),
        "docs_failed": int(docs_failed),
        "docs_no_email": int(docs_no_email),
        "attachments_total": int(attachments_total),
        "missing_total": int(missing_total),
        "run_id": int(run_id),
    }

    history.finish_run(run_id, ("dry_run" if dry_run else "ok"), result)
    try:
        history.vacuum()
    except Exception:
        pass

    auto_cfg["last_scan_end"] = window_end.isoformat(timespec="seconds")
    auto_cfg["last_run_at"] = datetime.now().isoformat(timespec="seconds")
    auto_cfg["last_result"] = dict(result)

    if not dry_run:
        system_logger.info(
            "auto_docs concluido discovered=%s pending_before=%s planned=%s sent=%s failed_emails=%s docs_sent=%s docs_failed=%s no_email=%s",
            result["discovered"],
            result["pending_before"],
            result["emails_planned"],
            result["emails_sent"],
            result["failed_emails"],
            result["docs_sent"],
            result["docs_failed"],
            result["docs_no_email"],
        )

    if run_lock is not None and not dry_run:
        try:
            run_lock.release()
        except Exception:
            pass

    return result

