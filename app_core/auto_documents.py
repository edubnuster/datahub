# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import time
from datetime import datetime, timedelta
from email.message import EmailMessage
import logging
import smtplib
import ssl
from typing import Any, Callable, Dict, List, Optional, Tuple

from .database import Database
from .danfe import danfe_pdf_from_nfe_xml
from .documents_history import DocumentsHistory
from .email_utils import attach_email_logo, zip_named_files
from .helpers import AppError, format_smtp_from
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
    if ext == "xml":
        return "application", "xml"
    if ext == "txt":
        return "text", "plain"
    if ext == "zip":
        return "application", "zip"
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


def _generate_boleto_pdf_if_needed(boleto_data: Dict[str, Any], inv: InvoiceRow, *, include_pix_qrcode: bool) -> Tuple[Optional[bytes], str]:
    if not (boleto_data or {}).get("exists"):
        return None, ""
    attachment_data = boleto_data.get("attachment_data")
    filename = boleto_data.get("filename") or f"boleto_{inv.invoice_id}.pdf"
    if attachment_data and include_pix_qrcode:
        return attachment_data, filename
    from ui import build_boleto_pdf_bytes
    return build_boleto_pdf_bytes(boleto_data, inv, include_pix_qrcode=bool(include_pix_qrcode)), filename


def _subject(company_name: str, customer_name: str) -> str:
    customer_name = str(customer_name or "").strip()
    parts = ["Documentos gerados"]
    if customer_name:
        parts.append(customer_name)
    return " - ".join(parts)


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


def _effective_auto_docs_cfg(cfg: Dict[str, Any]) -> Dict[str, Any]:
    schedules = cfg.get("financeiro_envio_auto_documentos_agendamentos")
    active_id = str(cfg.get("financeiro_envio_auto_documentos_active_id") or "").strip()
    if isinstance(schedules, list):
        if not schedules:
            cfg["financeiro_envio_auto_documentos"] = {"enabled": False}
            return cfg["financeiro_envio_auto_documentos"]
        cleaned: List[Dict[str, Any]] = []
        enabled: List[Dict[str, Any]] = []
        for s in schedules:
            if not isinstance(s, dict):
                continue
            sid = str(s.get("id") or "").strip()
            if not sid:
                continue
            cleaned.append(s)
            if bool(s.get("enabled")):
                enabled.append(s)
        if cleaned != schedules:
            cfg["financeiro_envio_auto_documentos_agendamentos"] = cleaned
            schedules = cleaned
        if enabled:
            active = None
            for s in enabled:
                if str(s.get("id") or "").strip() == active_id:
                    active = s
                    break
            if active is None:
                active = enabled[0]
            keep_id = str(active.get("id") or "").strip()
            if len(enabled) > 1 and keep_id:
                for s in schedules:
                    if not isinstance(s, dict):
                        continue
                    sid = str(s.get("id") or "").strip()
                    if sid and sid != keep_id:
                        s["enabled"] = False
                cfg["financeiro_envio_auto_documentos_active_id"] = keep_id
            cfg["financeiro_envio_auto_documentos"] = active
            return active
        if not active_id and schedules:
            cfg["financeiro_envio_auto_documentos_active_id"] = str(schedules[0].get("id") or "").strip()

    auto_cfg = cfg.get("financeiro_envio_auto_documentos")
    if not isinstance(auto_cfg, dict):
        auto_cfg = {}
        cfg["financeiro_envio_auto_documentos"] = auto_cfg
    return auto_cfg


def run_auto_documents(
    cfg: Dict[str, Any],
    *,
    dry_run: bool,
    user_label: str,
    now: Optional[datetime] = None,
    force: bool = False,
    allow_resend: bool = False,
    progress_cb: Optional[Callable[[Dict[str, Any]], None]] = None,
) -> Dict[str, Any]:
    system_logger = get_system_logger()
    docs_generated_logger = get_docs_generated_logger()
    docs_sent_logger = get_docs_sent_logger()

    now = now or datetime.now()
    
    def _emit(payload: Dict[str, Any]):
        cb = progress_cb
        if cb is None:
            return
        try:
            cb(dict(payload or {}))
        except Exception:
            return

    auto_cfg = _effective_auto_docs_cfg(cfg)

    enabled = bool(auto_cfg.get("enabled", False))
    try:
        interval_hours = int(auto_cfg.get("interval_hours") or 4)
    except Exception:
        interval_hours = 4
    max_interval = 720 if force else 72
    interval_hours = max(1, min(max_interval, interval_hours))

    if not enabled and not (dry_run or force):
        return {"enabled": False, "skipped": True, "reason": "disabled"}

    window_end = now
    window_start = now - timedelta(hours=interval_hours)
    window_text = f"{window_start.strftime('%d/%m/%Y %H:%M')} até {window_end.strftime('%d/%m/%Y %H:%M')}"
    _emit({"stage": "start", "message": "Iniciando...", "window_text": window_text, "dry_run": bool(dry_run)})

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
        _emit({"stage": "scan", "message": "Buscando documentos gerados...", "window_text": window_text})
        discovered_rows = Database(cfg).list_generated_boletos(window_start, window_end)
    except Exception as e:
        history.finish_run(run_id, "error", {"error": str(e)})
        raise
    _emit({"stage": "scan_done", "message": f"Encontrados: {len(discovered_rows)}", "discovered": len(discovered_rows), "window_text": window_text})

    skipped_duplicates = 0
    kept_sent = 0
    for r in discovered_rows:
        st = history.upsert_generated(r, allow_duplicate=bool(allow_resend))
        if st == "skipped_duplicate":
            skipped_duplicates += 1
        elif st == "kept_sent":
            kept_sent += 1
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

    planned_to_set: set[str] = set()
    emails_sent = 0
    failed_emails = 0
    sent_to_set: set[str] = set()
    docs_sent = 0
    docs_failed = 0
    docs_no_email = 0
    attachments_total = 0
    missing_total = 0

    extra_body = str(auto_cfg.get("extra_body") or "").strip()
    include_pix_qrcode = bool(auto_cfg.get("send_pix_qrcode", False))

    delay_seconds = int(smtp_cfg.get("delay_seconds", 5))
    first_email = True
    emails_planned = 0

    try:
        batch_size = int(auto_cfg.get("pending_batch_size") or 2000)
    except Exception:
        batch_size = 2000
    batch_size = max(50, min(5000, batch_size))

    discovered_grids = [str(r.get("boleto_grid") or "").strip() for r in discovered_rows if str(r.get("boleto_grid") or "").strip()]
    pending_records = history.list_pending_by_grids(discovered_grids)
    pending_before = len(pending_records)
    _emit({"stage": "pending", "message": "Separando pendências...", "pending_before": int(pending_before)})
    pending_grids = [p.boleto_grid for p in pending_records if str(p.boleto_grid or "").strip()]
    if len(pending_grids) > batch_size:
        pending_grids = pending_grids[:batch_size]
        pending_records = pending_records[:batch_size]

    if dry_run:
        grouped_keys = set()
        planned_emails = set()
        for p in pending_records:
            to_email = str(p.customer_email or "").strip()
            if not to_email:
                continue
            planned_emails.add(to_email)
            cid = str(p.customer_id or "").strip()
            grouped_keys.add(cid if cid not in ("", "0") else to_email)
        docs_no_email = len([p for p in pending_records if not str(p.customer_email or "").strip()])
        emails_planned = len(grouped_keys)
        pending_grids = []
        planned_to_set = planned_emails
        _emit({"stage": "planned", "message": f"Planejados: {int(emails_planned)}", "emails_planned": int(emails_planned), "docs_no_email": int(docs_no_email)})

    if pending_grids:

        rows_by_grid: Dict[str, Dict[str, Any]] = {}
        try:
            pending_rows = Database(cfg).list_boletos_by_grids(pending_grids, include_closed=True)
            for r in pending_rows:
                bg = str(r.get("boleto_grid") or "").strip()
                if bg:
                    rows_by_grid[bg] = dict(r)
        except Exception:
            rows_by_grid = {}

        to_process: List[Dict[str, Any]] = []
        closed_grids: List[str] = []
        missing_grids: List[str] = []
        for bg in pending_grids:
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
            if cid not in (None, "", 0, "0"):
                key = str(cid)
            else:
                cliente = str(r.get("cliente") or "").strip()
                key = f"{to_email}|{cliente}" if (to_email and cliente) else (to_email or cliente or bg)
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
        total_targets = len([g for g in grouped.values() if (g.get("to_email") or "").strip()])
        _emit(
            {
                "stage": "planning_done",
                "message": f"Planejados: {int(emails_planned)}",
                "emails_planned": int(emails_planned),
                "total_targets": int(total_targets),
                "docs_no_email": int(docs_no_email),
                "pending_before": int(pending_before),
            }
        )

        for target_i, g in enumerate([g for g in grouped.values() if (g.get("to_email") or "").strip()], start=1):
            rows = g.get("rows") or []
            if not rows:
                continue
            to_email = str(g.get("to_email") or "").strip()
            if not to_email:
                continue
            planned_to_set.add(str(to_email))
            _emit(
                {
                    "stage": "target_start",
                    "message": "Preparando anexos...",
                    "current": int(target_i),
                    "total": int(total_targets),
                    "customer_name": str(g.get("customer_name") or "").strip(),
                    "to_email": str(to_email),
                    "emails_sent": int(emails_sent),
                    "failed_emails": int(failed_emails),
                    "docs_sent": int(docs_sent),
                    "docs_failed": int(docs_failed),
                    "docs_no_email": int(docs_no_email),
                }
            )

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

            nfe_map: Dict[Any, Dict[str, Any]] = {}
            try:
                movto_ids = [inv.movto_id for inv in invoices if getattr(inv, "movto_id", None) not in (None, "", 0, "0")]
                nfe_map = Database(cfg).get_nfe_attachments_bulk(movto_ids)
            except Exception:
                nfe_map = {}

            purchase_map = {}
            try:
                movto_ids = [inv.movto_id for inv in invoices if getattr(inv, "movto_id", None) not in (None, "", 0, "0")]
                purchase_map = Database(cfg).get_purchase_info_bulk(movto_ids)
            except Exception:
                purchase_map = {}

            signature_files: List[Tuple[bytes, str]] = []
            for bg, inv in items:
                boleto_data = payload_map.get(bg) or {}
                try:
                    if boleto_data.get("exists"):
                        filename = boleto_data.get("filename") or f"boleto_{inv.invoice_id}.pdf"
                        generated, fname = _generate_boleto_pdf_if_needed(boleto_data, inv, include_pix_qrcode=include_pix_qrcode)
                        if generated:
                            attachments.append((generated, fname or filename, bg))
                        else:
                            missing += 1
                    else:
                        missing += 1
                except Exception:
                    missing += 1

                sig = signature_map.get(getattr(inv, "movto_id", None)) or {}
                added_for_inv = False
                for a in (sig.get("attachments") or []):
                    data = a.get("data")
                    name = a.get("filename")
                    if data and name:
                        signature_files.append((data, name))
                        added_for_inv = True
                sig_bytes = sig.get("attachment_data")
                if not added_for_inv and sig.get("exists") and sig_bytes:
                    signature_files.append((sig_bytes, sig.get("filename") or f"assinatura_{inv.movto_id}"))

                nfe = nfe_map.get(getattr(inv, "movto_id", None)) or nfe_map.get(str(getattr(inv, "movto_id", None) or "")) or {}
                nfe_atts = list((nfe.get("attachments") or []))
                has_pdf = bool([a for a in nfe_atts if str(a.get("filename") or "").lower().endswith(".pdf") and a.get("data")])
                for a in nfe_atts:
                    if has_pdf:
                        break
                    data = a.get("data")
                    name = str(a.get("filename") or "").lower()
                    if not data or not name.endswith(".xml"):
                        continue
                    extra_inf = ""
                    try:
                        extra_inf = Database(cfg).get_placa_km_text_bulk([getattr(inv, "movto_id", None)]).get(getattr(inv, "movto_id", None)) or ""
                    except Exception:
                        extra_inf = ""
                    pdf_bytes, pdf_name = danfe_pdf_from_nfe_xml(
                        data,
                        fallback_suffix=str(getattr(inv, "invoice_id", "") or getattr(inv, "movto_id", "") or ""),
                        extra_inf_cpl_text=extra_inf,
                    )
                    if pdf_bytes and pdf_name:
                        nfe_atts.append({"data": pdf_bytes, "filename": pdf_name, "mime_type": "application/pdf"})
                        has_pdf = True
                for a in nfe_atts:
                    data = a.get("data")
                    name = a.get("filename")
                    if data and name:
                        attachments.append((data, name, ""))

            if signature_files:
                cust_part = "".join([c if c.isalnum() or c in ("_", "-") else "_" for c in str(invoices[0].customer_name if invoices else "").strip()])[:40]
                zip_name = f"assinaturas_{cust_part}.zip" if cust_part else "assinaturas.zip"
                zip_bytes, _ = zip_named_files(signature_files, zip_filename=zip_name)
                attachments.append((zip_bytes, zip_name, ""))
            attachments_total += len(attachments)
            missing_total += missing
            if dry_run:
                continue

            from ui import build_agenda_email_body

            subject = _subject(invoices[0].company if invoices else "", invoices[0].customer_name if invoices else "")

            fatura_txt: Optional[Tuple[bytes, str]] = None
            try:
                from ui import build_faturas_detalhamento_txt_bytes

                txt_bytes, txt_name = build_faturas_detalhamento_txt_bytes(invoices, purchase_info_map=purchase_map)
                if txt_bytes and txt_name:
                    fatura_txt = (txt_bytes, txt_name)
            except Exception:
                fatura_txt = None
            if fatura_txt:
                attachments_total += 1

            max_attachments = 18
            max_bytes = 15 * 1024 * 1024
            reserved_count = 1 if fatura_txt else 0
            reserved_bytes = len(fatura_txt[0]) if fatura_txt else 0
            usable_max_attachments = max(1, max_attachments - reserved_count)
            usable_max_bytes = max(1024, max_bytes - reserved_bytes)
            batches: List[List[Tuple[bytes, str, str]]] = []
            current: List[Tuple[bytes, str, str]] = []
            current_bytes = 0
            for data, name, bg in attachments:
                size = len(data) if data else 0
                if current and (len(current) >= usable_max_attachments or (current_bytes + size) > usable_max_bytes):
                    batches.append(current)
                    current = []
                    current_bytes = 0
                current.append((data, name, bg))
                current_bytes += size
            if current or not attachments:
                batches.append(current)
            _emit(
                {
                    "stage": "target_ready",
                    "message": "Enviando...",
                    "current": int(target_i),
                    "total": int(total_targets),
                    "customer_name": invoices[0].customer_name if invoices else str(g.get("customer_name") or "").strip(),
                    "to_email": str(to_email),
                    "batch_total": int(len(batches)),
                    "emails_sent": int(emails_sent),
                    "failed_emails": int(failed_emails),
                    "docs_sent": int(docs_sent),
                    "docs_failed": int(docs_failed),
                    "docs_no_email": int(docs_no_email),
                }
            )

            for idx, batch in enumerate(batches, start=1):
                _emit(
                    {
                        "stage": "batch_sending",
                        "message": "Enviando e-mail...",
                        "current": int(target_i),
                        "total": int(total_targets),
                        "customer_name": invoices[0].customer_name if invoices else str(g.get("customer_name") or "").strip(),
                        "to_email": str(to_email),
                        "batch_idx": int(idx),
                        "batch_total": int(len(batches)),
                        "emails_sent": int(emails_sent),
                        "failed_emails": int(failed_emails),
                        "docs_sent": int(docs_sent),
                        "docs_failed": int(docs_failed),
                        "docs_no_email": int(docs_no_email),
                    }
                )
                if not first_email and delay_seconds > 0:
                    time.sleep(delay_seconds)
                first_email = False
                flags = {"boleto": False, "fatura_pdf": False, "fatura_txt": False, "xml": False, "danfe": False, "assinatura": False}
                try:
                    from ui import _detect_email_attachment_flags

                    flag_names = [name for data, name, _bg in batch if data]
                    if fatura_txt:
                        flag_names.append(fatura_txt[1])
                    flags = _detect_email_attachment_flags(flag_names)
                except Exception:
                    flags = flags

                text_body, html_body = build_agenda_email_body(
                    invoices[0].customer_name if invoices else "",
                    window_text,
                    invoices,
                    missing,
                    extra_body,
                    context_label="Período de geração",
                    purchase_info_map=purchase_map,
                    attachment_flags=flags,
                )

                msg = EmailMessage()
                msg["From"] = format_smtp_from(smtp_cfg) or smtp_email
                msg["To"] = to_email
                msg["Subject"] = subject if len(batches) == 1 else f"{subject} ({idx}/{len(batches)})"
                msg.set_content(text_body)
                msg.add_alternative(html_body, subtype="html")
                try:
                    attach_email_logo(msg)
                except Exception:
                    pass
                grids_in_batch: List[str] = []
                if fatura_txt:
                    maintype, subtype = _mime_parts_from_filename(fatura_txt[1])
                    msg.add_attachment(fatura_txt[0], maintype=maintype, subtype=subtype, filename=fatura_txt[1])
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
                    if to_email:
                        sent_to_set.add(str(to_email))
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
                    _emit(
                        {
                            "stage": "batch_sent",
                            "message": "Enviado",
                            "current": int(target_i),
                            "total": int(total_targets),
                            "customer_name": invoices[0].customer_name if invoices else str(g.get("customer_name") or "").strip(),
                            "to_email": str(to_email),
                            "batch_idx": int(idx),
                            "batch_total": int(len(batches)),
                            "emails_sent": int(emails_sent),
                            "failed_emails": int(failed_emails),
                            "docs_sent": int(docs_sent),
                            "docs_failed": int(docs_failed),
                            "docs_no_email": int(docs_no_email),
                        }
                    )
                except Exception as e:
                    failed_emails += 1
                    if grids_in_batch:
                        history.mark_failed(grids_in_batch, error=str(e))
                        docs_failed += len(grids_in_batch)
                    system_logger.error("Falha ao enviar docs para=%s erro=%s", to_email, e)
                    _emit(
                        {
                            "stage": "batch_failed",
                            "message": "Falha",
                            "current": int(target_i),
                            "total": int(total_targets),
                            "customer_name": invoices[0].customer_name if invoices else str(g.get("customer_name") or "").strip(),
                            "to_email": str(to_email),
                            "batch_idx": int(idx),
                            "batch_total": int(len(batches)),
                            "emails_sent": int(emails_sent),
                            "failed_emails": int(failed_emails),
                            "docs_sent": int(docs_sent),
                            "docs_failed": int(docs_failed),
                            "docs_no_email": int(docs_no_email),
                            "error": str(e),
                        }
                    )

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
        "sent_to": sorted(sent_to_set),
        "planned_to": sorted(planned_to_set),
        "skipped_duplicates": int(skipped_duplicates),
        "already_sent": int(kept_sent),
    }

    history.finish_run(run_id, ("dry_run" if dry_run else "ok"), result)
    _emit(
        {
            "stage": "done",
            "message": "Concluído.",
            "result": dict(result),
            "discovered": int(result.get("discovered") or 0),
            "pending_before": int(result.get("pending_before") or 0),
            "emails_planned": int(result.get("emails_planned") or 0),
            "emails_sent": int(result.get("emails_sent") or 0),
            "failed_emails": int(result.get("failed_emails") or 0),
            "docs_sent": int(result.get("docs_sent") or 0),
            "docs_failed": int(result.get("docs_failed") or 0),
            "docs_no_email": int(result.get("docs_no_email") or 0),
        }
    )
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

