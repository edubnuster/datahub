# -*- coding: utf-8 -*-
import argparse
import time
from datetime import date, datetime, time as dtime, timedelta
from email.message import EmailMessage
import smtplib
import ssl
from typing import Any, Dict, List, Optional, Tuple
import re

from app_core.audit import AuditLogger
from app_core.config_manager import ConfigManager
from app_core.database import Database
from app_core.danfe import danfe_pdf_from_nfe_xml
from app_core.helpers import AppError, format_smtp_from
from app_core.logging_setup import get_system_logger, init_logging
from app_core.models import InvoiceRow


def _parse_time(value: str) -> dtime:
    value = str(value or "").strip()
    m = re.match(r"^(\d{1,2}):(\d{2})$", value)
    if not m:
        return dtime(6, 0)
    hh = int(m.group(1))
    mm = int(m.group(2))
    if hh < 0 or hh > 23 or mm < 0 or mm > 59:
        return dtime(6, 0)
    return dtime(hh, mm)


def _money_br(value: Any) -> str:
    if value in (None, ""):
        return "0,00"
    try:
        num = float(value)
        return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(value)


def _mime_parts_from_filename(filename: str) -> Tuple[str, str]:
    name = str(filename or "").strip().lower()
    ext = name.rsplit(".", 1)[-1] if "." in name else ""
    if ext == "pdf":
        return "application", "pdf"
    if ext == "xml":
        return "application", "xml"
    if ext == "png":
        return "image", "png"
    if ext in ("jpg", "jpeg"):
        return "image", "jpeg"
    if ext in ("tif", "tiff"):
        return "image", "tiff"
    return "application", "octet-stream"


def _load_agendas(cfg: Dict[str, Any]) -> List[Dict[str, Any]]:
    agendas = cfg.get("financeiro_agendas", []) or []
    if not isinstance(agendas, list):
        return []
    out = []
    for a in agendas:
        if not isinstance(a, dict):
            continue
        out.append(dict(a))
    return out


def _save_agendas(cfg: Dict[str, Any], agendas: List[Dict[str, Any]]) -> None:
    cfg["financeiro_agendas"] = agendas
    ConfigManager.save(cfg)


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


def _invoice_from_row(r: Dict[str, Any]) -> InvoiceRow:
    return InvoiceRow(
        invoice_id=r.get("movto_id"),
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


def _default_subject(inv: InvoiceRow) -> str:
    return f"Fatura a receber - {inv.customer_name} - vencimento {inv.due_date_display()}"

def _subject_group(customer_name: str) -> str:
    return f"Alerta de vencimento de boleto - {customer_name}"

def _default_body(inv: InvoiceRow, boleto_data: Dict[str, Any], attachment_bytes: Optional[bytes]) -> str:
    account_display = (f"{inv.account_code or ''} - {inv.account_name or ''}").strip(" -")
    note = str((boleto_data or {}).get("email_note", "") or "").strip()
    if not note:
        if (boleto_data or {}).get("exists"):
            note = "Observação: o boleto segue em anexo." if attachment_bytes else "Observação: foi localizado um boleto, mas não foi possível anexá-lo automaticamente."
        else:
            note = "Observação: o boleto ainda não foi gerado."
    return (
        f"Prezado(a),\n\n"
        f"Segue abaixo os dados da fatura para conferência e programação do pagamento.\n\n"
        f"Empresa: {inv.company}\n"
        f"Cliente: {inv.customer_name}\n"
        f"Código do cliente: {inv.customer_code}\n"
        f"Conta: {account_display}\n"
        f"Emissão: {inv.issue_date_display()}\n"
        f"Vencimento: {inv.due_date_display()}\n"
        f"Valor original: {inv.amount_display()}\n"
        f"Desconto: {inv.discount_amount_display()}\n"
        f"Saldo em aberto: {inv.open_balance_display()}\n\n"
        f"{note}\n\n"
        f"Em caso de dúvidas, ficamos à disposição.\n\n"
        f"Atenciosamente,\n"
        f"{inv.company}"
    )

def _group_body(
    customer_name: str,
    base_date: date,
    invoices: List[InvoiceRow],
    total: float,
    missing_count: int,
    extra_body: str,
    purchase_map: Dict[Any, Dict[str, Any]],
    attachment_flags: Optional[Dict[str, bool]] = None,
) -> tuple[str, str]:
    from ui import build_due_alert_email_body
    return build_due_alert_email_body(
        customer_name,
        base_date,
        invoices,
        missing_count,
        extra_body,
        purchase_info_map=purchase_map,
        attachment_flags=attachment_flags,
    )


def _generate_boleto_pdf_if_needed(boleto_data: Dict[str, Any], inv: InvoiceRow, include_pix_qrcode: bool) -> Tuple[Optional[bytes], str]:
    if not (boleto_data or {}).get("exists"):
        return None, ""
    attachment_data = boleto_data.get("attachment_data")
    filename = boleto_data.get("filename") or f"boleto_{inv.invoice_id}.pdf"
    from ui import build_boleto_pdf_bytes
    try:
        return build_boleto_pdf_bytes(boleto_data, inv, include_pix_qrcode=include_pix_qrcode), filename
    except Exception:
        return (attachment_data if include_pix_qrcode and attachment_data else None), filename


def _should_run_agenda(
    agenda: Dict[str, Any],
    now: datetime,
    today_key: str,
    respect_time: bool,
    force: bool,
) -> bool:
    if force:
        return True
    if not agenda.get("enabled"):
        return False
    if str(agenda.get("last_run_date") or "").strip() == today_key:
        return False
    if respect_time and now.time() < _parse_time(agenda.get("send_time")):
        return False
    return True


def run_agenda(
    cfg: Dict[str, Any],
    agenda: Dict[str, Any],
    now: datetime,
    respect_time: bool,
    force: bool,
    dry_run: bool,
    user_label: str,
    verbose: bool,
) -> Dict[str, Any]:
    agenda_id = str(agenda.get("id") or "")
    agenda_name = str(agenda.get("name") or "").strip() or agenda_id or "Alerta de vencimento"
    today_key = now.date().isoformat()
    if not _should_run_agenda(agenda, now, today_key, respect_time, force):
        return {"agenda_id": agenda_id, "agenda_name": agenda_name, "skipped": True}

    base_date = now.date()
    late_minutes = 0
    out_of_time = False
    try:
        days_before = int(agenda.get("days_before_due") or 0)
    except Exception:
        days_before = 0
    try:
        days_after = int(agenda.get("days_after_due") or 0)
    except Exception:
        days_after = 0
    days_before = max(0, min(365, days_before))
    days_after = max(0, min(365, days_after))
    try:
        late_minutes = int((now - datetime.combine(now.date(), _parse_time(agenda.get("send_time")))).total_seconds() // 60)
        if late_minutes < 0:
            late_minutes = 0
    except Exception:
        late_minutes = 0
    out_of_time = late_minutes > 15

    due_dates = []
    if days_before > 0:
        due_dates.append(base_date + timedelta(days=days_before))
    if days_after > 0:
        due_dates.append(base_date - timedelta(days=days_after))
    due_dates = sorted({d for d in due_dates})

    rows: List[Dict[str, Any]] = []
    db = Database(cfg)
    for d in due_dates:
        rows.extend(
            db.list_agenda_invoices(
                d,
                group_id=agenda.get("group_id"),
                portador_id=agenda.get("portador_id"),
                customer_id=agenda.get("customer_id"),
            )
        )

    invoices = [_invoice_from_row(r) for r in rows]
    invoice_ids = [i.invoice_id for i in invoices if i.invoice_id not in (None, "", 0, "0")]
    boleto_map = {}
    try:
        boleto_map = Database(cfg).get_boletos_email_payload_bulk(invoice_ids)
    except Exception:
        boleto_map = {}

    signature_map = {}
    try:
        signature_map = Database(cfg).get_sale_signatures_pdf_bulk(invoice_ids)
    except Exception:
        signature_map = {}

    nfe_map = {}
    try:
        nfe_map = Database(cfg).get_nfe_attachments_bulk(invoice_ids)
    except Exception:
        nfe_map = {}

    purchase_map = {}
    try:
        purchase_map = Database(cfg).get_purchase_info_bulk(invoice_ids)
    except Exception:
        purchase_map = {}

    grouped: Dict[str, Dict[str, Any]] = {}
    for inv in invoices:
        cid = inv.customer_id
        key = str(cid) if cid not in (None, "", 0, "0") else (inv.customer_email or inv.customer_name or str(inv.customer_code))
        item = grouped.get(key)
        if not item:
            item = {"customer_id": cid, "customer_name": inv.customer_name, "customer_email": inv.customer_email, "invoices": []}
            grouped[key] = item
        item["invoices"].append(inv)

    emails_planned = 0
    emails_sent = 0
    skipped_no_email = 0
    failed = 0
    attachments_total = 0
    missing_total = 0
    total_sum = 0.0
    for inv in invoices:
        try:
            total_sum += float(inv.open_balance or 0)
        except Exception:
            pass
    due_dates_str = ",".join([d.isoformat() for d in due_dates])

    smtp_cfg = cfg.get("smtp", {})
    smtp_email = str(smtp_cfg.get("email", "")).strip()
    include_pix_qrcode = bool(agenda.get("send_pix_qrcode", False))

    delay_seconds = int(smtp_cfg.get("delay_seconds", 5))
    first_email = True

    for g in grouped.values():
        invs: List[InvoiceRow] = g.get("invoices") or []
        if not invs:
            continue
        to_email = (g.get("customer_email") or "").strip()
        if not to_email and g.get("customer_id") not in (None, "", 0, "0"):
            try:
                to_email = Database(cfg).get_customer_email(g.get("customer_id"))
            except Exception:
                to_email = ""
        if not to_email:
            skipped_no_email += 1
            continue

        emails_planned += 1

        attachments: List[Tuple[bytes, str]] = []
        missing = 0
        total = 0.0
        for inv in invs:
            try:
                total += float(inv.open_balance or 0)
            except Exception:
                pass
            boleto_data = boleto_map.get(inv.invoice_id) or {}
            try:
                if boleto_data.get("exists"):
                    data, filename = _generate_boleto_pdf_if_needed(boleto_data, inv, include_pix_qrcode=include_pix_qrcode)
                    if data:
                        attachments.append((data, filename))
                    else:
                        missing += 1
                else:
                    missing += 1
            except Exception:
                missing += 1

            sig = signature_map.get(inv.invoice_id) or {}
            sig_added = False
            for a in (sig.get("attachments") or []):
                data = a.get("data")
                name = a.get("filename")
                if data and name:
                    attachments.append((data, name))
                    sig_added = True
            sig_bytes = sig.get("attachment_data")
            if not sig_added and sig.get("exists") and sig_bytes:
                attachments.append((sig_bytes, sig.get("filename") or f"assinatura_{inv.invoice_id}"))

            nfe = nfe_map.get(inv.invoice_id) or nfe_map.get(str(inv.invoice_id)) or {}
            nfe_atts = list((nfe.get("attachments") or []))
            has_pdf = bool([a for a in nfe_atts if str(a.get("filename") or "").lower().endswith(".pdf") and a.get("data")])
            for a in nfe_atts:
                if has_pdf:
                    break
                ndata = a.get("data")
                nname = str(a.get("filename") or "").lower()
                if not ndata or not nname.endswith(".xml"):
                    continue
                pdf_bytes, pdf_name = danfe_pdf_from_nfe_xml(ndata, fallback_suffix=str(inv.invoice_id))
                if pdf_bytes and pdf_name:
                    nfe_atts.append({"data": pdf_bytes, "filename": pdf_name, "mime_type": "application/pdf"})
                    has_pdf = True
            for a in nfe_atts:
                ndata = a.get("data")
                nname = a.get("filename")
                if ndata and nname:
                    attachments.append((ndata, nname))

        attachments_total += len(attachments)
        missing_total += missing
        if dry_run:
            continue

        subject = _subject_group(invs[0].customer_name)

        max_attachments = 18
        max_bytes = 15 * 1024 * 1024
        batches: List[List[Tuple[bytes, str]]] = []
        current: List[Tuple[bytes, str]] = []
        current_bytes = 0
        for data, name in attachments:
            size = len(data) if data else 0
            if current and (len(current) >= max_attachments or (current_bytes + size) > max_bytes):
                batches.append(current)
                current = []
                current_bytes = 0
            current.append((data, name))
            current_bytes += size
        if current or not attachments:
            batches.append(current)

        for idx, batch in enumerate(batches, start=1):
            if not first_email and delay_seconds > 0:
                time.sleep(delay_seconds)
            first_email = False
            try:
                from ui import _detect_email_attachment_flags

                flags = _detect_email_attachment_flags([name for data, name in batch if data])
            except Exception:
                flags = {"boleto": False, "fatura_pdf": False, "xml": False, "danfe": False, "assinatura": False}

            text_body, html_body = _group_body(
                invs[0].customer_name,
                base_date,
                invs,
                total,
                missing,
                agenda.get("extra_body"),
                purchase_map=purchase_map,
                attachment_flags=flags,
            )

            msg = EmailMessage()
            msg["From"] = format_smtp_from(smtp_cfg) or smtp_email
            msg["To"] = to_email
            msg["Subject"] = subject if len(batches) == 1 else f"{subject} ({idx}/{len(batches)})"
            
            msg.set_content(text_body)
            msg.add_alternative(html_body, subtype='html')
            
            for data, name in batch:
                if not data:
                    continue
                maintype, subtype = _mime_parts_from_filename(name)
                msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=name)
            try:
                _smtp_send_message(cfg, msg)
                emails_sent += 1
                AuditLogger.write(user_label, "agenda_envio_email_cli", f"agenda_id={agenda_id};cliente={invs[0].customer_name};para={to_email};titulos={len(invs)};anexos={len(batch)};pix_incluido_no_boleto={'sim' if include_pix_qrcode else 'nao'}")
            except Exception as e:
                failed += 1
                AuditLogger.write(user_label, "agenda_envio_email_cli_erro", f"agenda_id={agenda_id};cliente={invs[0].customer_name};para={to_email};erro={e}")
                if verbose:
                    print(f"[{agenda_name}] erro ao enviar para {to_email}: {e}")

    if dry_run:
        return {
            "agenda_id": agenda_id,
            "agenda_name": agenda_name,
            "due_dates": due_dates_str,
            "count": len(rows),
            "emails_planned": emails_planned,
            "total": total_sum,
            "sent": 0,
            "skipped_no_email": skipped_no_email,
            "failed": 0,
            "attachments_total": attachments_total,
            "missing_total": missing_total,
            "late_minutes": late_minutes,
            "out_of_time": bool(out_of_time),
            "dry_run": True,
        }

    return {
        "agenda_id": agenda_id,
        "agenda_name": agenda_name,
        "due_dates": due_dates_str,
        "count": len(rows),
        "emails_planned": emails_planned,
        "sent": emails_sent,
        "skipped_no_email": skipped_no_email,
        "failed": failed,
        "attachments_total": attachments_total,
        "missing_total": missing_total,
        "late_minutes": late_minutes,
        "out_of_time": bool(out_of_time),
        "dry_run": False,
    }


def _update_last_run(cfg: Dict[str, Any], agenda_id: str, today_key: str, result: Dict[str, Any], now: datetime) -> bool:
    agendas = _load_agendas(cfg)
    updated = False
    for i, a in enumerate(agendas):
        if str(a.get("id") or "") == str(agenda_id):
            merged = dict(a)
            merged["last_run_date"] = today_key
            merged["last_run_at"] = now.isoformat(timespec="seconds")
            merged["last_due_date"] = str(result.get("due_dates") or "")
            merged.pop("last_window_start", None)
            merged.pop("last_window_end", None)
            merged["last_late_minutes"] = int(result.get("late_minutes") or 0)
            merged["last_out_of_time"] = bool(result.get("out_of_time"))
            merged["last_result"] = {
                "emails_sent": int(result.get("sent") or 0),
                "skipped_no_email": int(result.get("skipped_no_email") or 0),
                "failed": int(result.get("failed") or 0),
                "attachments_total": int(result.get("attachments_total") or 0),
                "missing_total": int(result.get("missing_total") or 0),
                "emails_planned": int(result.get("emails_planned") or 0),
            }
            agendas[i] = merged
            updated = True
            break
    if updated:
        _save_agendas(cfg, agendas)
    return updated


def main(argv=None) -> int:
    init_logging()
    syslog = get_system_logger()
    parser = argparse.ArgumentParser(prog="datahub-agenda", description="Executa os agendamentos de envio de faturas/boletos por e-mail.")
    parser.add_argument("--agenda-id", default="", help="Executa apenas um agendamento específico (id).")
    parser.add_argument("--dry-run", action="store_true", help="Não envia e-mails; apenas lista quantidades/valores.")
    parser.add_argument("--respect-time", action="store_true", help="Só executa agendamentos cujo send_time <= agora.")
    parser.add_argument("--force", action="store_true", help="Executa mesmo que já tenha rodado hoje (ignora last_run_date e horário).")
    parser.add_argument("--verbose", action="store_true", help="Mostra detalhes de falhas no console.")
    args = parser.parse_args(argv)

    cfg = ConfigManager.load()
    agendas = _load_agendas(cfg)

    now = datetime.now()
    today_key = now.date().isoformat()

    if args.agenda_id:
        agendas = [a for a in agendas if str(a.get("id") or "") == str(args.agenda_id)]

    if not agendas:
        print("Nenhum agendamento encontrado.")
        return 2

    user_label = "agenda_cli"
    try:
        syslog.info("agenda_cli inicio agendas=%s dry_run=%s respect_time=%s force=%s agenda_id=%s", len(agendas), bool(args.dry_run), bool(args.respect_time), bool(args.force), str(args.agenda_id or ""))
    except Exception:
        pass

    ran_any = False
    total_sent = 0
    total_skipped = 0
    total_failed = 0

    for a in agendas:
        res = run_agenda(
            cfg=cfg,
            agenda=a,
            now=now,
            respect_time=args.respect_time,
            force=args.force,
            dry_run=args.dry_run,
            user_label=user_label,
            verbose=args.verbose,
        )

        if res.get("skipped"):
            continue
        ran_any = True

        agenda_name = res.get("agenda_name") or res.get("agenda_id") or "Agendamento"
        if args.dry_run:
            print(f"[{agenda_name}] vencimento={res.get('due_date')} titulos={res.get('count')} emails={res.get('emails_planned')} sem_email={res.get('skipped_no_email')} boletos_sem_anexo={res.get('missing_total')}")
        else:
            print(f"[{agenda_name}] vencimento={res.get('due_date')} titulos={res.get('count')} emails={res.get('emails_planned')} enviados={res.get('sent')} sem_email={res.get('skipped_no_email')} falhas={res.get('failed')}")
            total_sent += int(res.get("sent") or 0)
            total_skipped += int(res.get("skipped_no_email") or 0)
            total_failed += int(res.get("failed") or 0)
            if not args.force:
                _update_last_run(cfg, res.get("agenda_id") or "", today_key, res, now)

    if not ran_any:
        try:
            syslog.info("agenda_cli fim nenhum_agendamento")
        except Exception:
            pass
        print("Nenhum agendamento elegível para execução no momento.")
        return 3

    if args.dry_run:
        try:
            syslog.info("agenda_cli fim dry_run")
        except Exception:
            pass
        return 0

    if total_failed > 0:
        try:
            syslog.info("agenda_cli fim status=erro enviados=%s sem_email=%s falhas=%s", total_sent, total_skipped, total_failed)
        except Exception:
            pass
        return 1
    try:
        syslog.info("agenda_cli fim status=ok enviados=%s sem_email=%s falhas=%s", total_sent, total_skipped, total_failed)
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

