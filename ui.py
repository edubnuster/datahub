# -*- coding: utf-8 -*-
from copy import deepcopy
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import html
import os
import smtplib
import ssl
import sys
from datetime import date, datetime, time, timedelta
from typing import Any, Dict, List, Optional, Tuple
import re
import threading
import time as time_module
import tkinter as tk
import zlib
from tkinter import ttk, messagebox
from email.message import EmailMessage
from app_core.audit import AuditLogger
from app_core.auth import UserManager
from app_core.config_manager import ConfigManager
from app_core.constants import APP_TITLE, CONFIG_PATH, LICENSE_FILENAME, MASTER_PASSWORD, MASTER_USERNAME, app_dir
from app_core.danfe import danfe_pdf_from_nfe_xml
from app_core.database import Database
from app_core.email_utils import EMAIL_LOGO_CID, attach_email_logo, zip_named_files
from app_core.helpers import AppError, format_smtp_from
from app_core.license_manager import LicenseManager
from app_core.models import CustomerRow, InvoiceRow


DATE_INPUT_FORMAT = "%d/%m/%Y"


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


def _detect_email_attachment_flags(filenames: List[str]) -> Dict[str, bool]:
    names = [str(n or "").strip().lower() for n in (filenames or []) if str(n or "").strip()]
    has_xml = any(n.endswith(".xml") and ("nfe" in n or "xml" in n) for n in names)
    has_danfe = any(n.endswith(".pdf") and "danfe" in n for n in names)
    has_boleto = any(n.endswith(".pdf") and "boleto" in n and "danfe" not in n for n in names)
    has_signature = any(("assinatura" in n or "assinaturas" in n or "cupom" in n) and (n.endswith(".pdf") or n.endswith(".png") or n.endswith(".jpg") or n.endswith(".jpeg") or n.endswith(".zip")) for n in names)
    has_invoice_pdf = any(n.endswith(".pdf") and ("fatura" in n or "invoice" in n) for n in names)
    has_invoice_txt = any(n.endswith(".txt") and ("fatura" in n or "invoice" in n) for n in names)
    return {"boleto": has_boleto, "fatura_pdf": has_invoice_pdf, "fatura_txt": has_invoice_txt, "xml": has_xml, "danfe": has_danfe, "assinatura": has_signature}


def _format_attachment_status(value: bool, reason: str = "") -> str:
    if value:
        return "Sim"
    r = str(reason or "").strip()
    if r:
        return f"Não ({r})"
    return "Não"


def build_attachments_note_text(
    *,
    has_boleto: bool,
    has_fatura_pdf: bool,
    has_fatura_txt: bool = False,
    has_xml: bool,
    has_danfe: bool,
    has_assinatura: bool,
    boleto_reason: str = "",
    fatura_reason: str = "a fatura está no corpo do e-mail",
    xml_reason: str = "",
    danfe_reason: str = "",
    assinatura_reason: str = "",
    missing_boleto_count: int = 0,
) -> str:
    found_labels: List[str] = []
    if has_boleto:
        found_labels.append("boleto")
    if has_fatura_pdf:
        found_labels.append("fatura (PDF)")
    elif has_fatura_txt:
        found_labels.append("fatura")
    if has_xml:
        found_labels.append("xml")
    if has_danfe:
        found_labels.append("danfe")
    if has_assinatura:
        found_labels.append("assinaturas")

    if not found_labels:
        return ""

    if len(found_labels) == 1:
        found_text = found_labels[0]
    elif len(found_labels) == 2:
        found_text = f"{found_labels[0]} e {found_labels[1]}"
    else:
        found_text = f"{', '.join(found_labels[:-1])} e {found_labels[-1]}"

    return f"Observação: Anexo: {found_text}"


def build_attachments_note_html(
    *,
    has_boleto: bool,
    has_fatura_pdf: bool,
    has_fatura_txt: bool = False,
    has_xml: bool,
    has_danfe: bool,
    has_assinatura: bool,
    boleto_reason: str = "",
    fatura_reason: str = "a fatura está no corpo do e-mail",
    xml_reason: str = "",
    danfe_reason: str = "",
    assinatura_reason: str = "",
    missing_boleto_count: int = 0,
) -> str:
    found_labels: List[str] = []
    if has_boleto:
        found_labels.append("boleto")
    if has_fatura_pdf:
        found_labels.append("fatura (PDF)")
    elif has_fatura_txt:
        found_labels.append("fatura")
    if has_xml:
        found_labels.append("xml")
    if has_danfe:
        found_labels.append("danfe")
    if has_assinatura:
        found_labels.append("assinaturas")

    if not found_labels:
        return ""

    if len(found_labels) == 1:
        found_text = found_labels[0]
    elif len(found_labels) == 2:
        found_text = f"{found_labels[0]} e {found_labels[1]}"
    else:
        found_text = f"{', '.join(found_labels[:-1])} e {found_labels[-1]}"

    return (
        "<div style='margin:12px 0;padding:10px 12px;background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;color:#9a3412;'>"
        f"<b>Observação:</b> Anexo: {html.escape(found_text)}"
        "</div>"
    )


def build_faturas_detalhamento_txt_bytes(
    invoices: List["InvoiceRow"],
    purchase_info_map: Optional[Dict[Any, Dict[str, Any]]] = None,
    generated_at: Optional[datetime] = None,
) -> tuple[bytes, str]:
    invs = [i for i in (invoices or []) if i is not None]
    if not invs:
        return b"", ""

    generated_at = generated_at or datetime.now()
    dt_txt = generated_at.strftime("%d/%m/%Y %H:%M")
    company = str(getattr(invs[0], "company", "") or "").strip()
    customer = str(getattr(invs[0], "customer_name", "") or "").strip()
    safe_customer = re.sub(r"[^0-9A-Za-z]+", "_", customer).strip("_") or "cliente"
    filename = f"faturas_{safe_customer}_{generated_at.strftime('%Y%m%d_%H%M')}.txt"

    width = 160
    inner = width - 2
    content_inner = inner - 1

    def _line_lr(left: str, right: str) -> str:
        left = str(left or "")
        right = str(right or "")
        if len(left) + len(right) >= width:
            return (left + " " + right)[:width]
        return left + (" " * (width - len(left) - len(right))) + right

    def _box_line(text: str = "") -> str:
        text = str(text or "")
        return "|" + text[:content_inner].ljust(content_inner) + " " + "|"

    def _box_lr(left: str, right: str) -> str:
        left = str(left or "")
        right = str(right or "")
        if len(left) + 1 + len(right) > content_inner:
            left = left[: max(0, content_inner - len(right) - 1)]
        return _box_line(left + (" " * (content_inner - len(left) - len(right))) + right)

    def _border(ch: str = "-") -> str:
        return "+" + (ch * inner) + "+"

    def _dash() -> str:
        return "-" * width

    def _date_br(value: Any) -> str:
        if isinstance(value, datetime):
            return value.strftime("%d/%m/%Y")
        if isinstance(value, date):
            return value.strftime("%d/%m/%Y")
        return str(value or "").strip()

    def _num_br(value: Any, decimals: int = 2) -> str:
        if value in (None, ""):
            value = 0
        try:
            num = float(value)
            fmt = f"{{:,.{int(decimals)}f}}"
            return fmt.format(num).replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return str(value)

    def _safe_float(value: Any) -> float:
        try:
            return float(value or 0)
        except Exception:
            return 0.0

    def _normalize_placa(value: Any) -> str:
        return re.sub(r"[^A-Za-z0-9]+", "", str(value or "").strip()).upper()

    def _to_float_br(value: Any) -> Optional[float]:
        s = str(value or "").strip()
        if not s:
            return None
        cleaned = re.sub(r"[^\d,\.]", "", s)
        if not cleaned or not re.search(r"\d", cleaned):
            return None
        try:
            if "," in cleaned and "." in cleaned:
                return float(cleaned.replace(".", "").replace(",", "."))
            if "," in cleaned:
                return float(cleaned.replace(".", "").replace(",", "."))
            return float(cleaned)
        except Exception:
            return None

    def _km_fields(doc: Dict[str, Any]) -> tuple[str, str, str, str]:
        placa_txt = _normalize_placa(doc.get("placa"))
        km_ini_num = doc.get("km_ini", None)
        km_fin_num = doc.get("km_fin", None)
        km_lt_num = doc.get("km_lt", None)

        def _num_plain(value: Any, decimals: int) -> str:
            try:
                num = float(value)
            except Exception:
                num = 0.0
            return f"{num:.{int(decimals)}f}"

        if km_ini_num is not None or km_fin_num is not None or km_lt_num is not None:
            km_ini_txt = _num_plain(km_ini_num, 1) if km_ini_num is not None else "0.0"
            km_fin_txt = _num_plain(km_fin_num, 1) if km_fin_num is not None else "0.0"
            km_lt_txt = _num_plain(km_lt_num, 2) if km_lt_num is not None else "0.0"
            return placa_txt, km_ini_txt, km_fin_txt, km_lt_txt

        km_raw = str(doc.get("km") or "").strip()
        media_raw = str(doc.get("km_media") or "").strip()
        low = km_raw.lower()

        nums: List[float] = []
        for t in re.findall(r"\d[\d\.,]*", km_raw):
            v = _to_float_br(t)
            if v is not None:
                nums.append(v)

        km_ini: Optional[float] = None
        km_fin: Optional[float] = None
        km_lt: Optional[float] = None

        has_media_hint = any(k in low for k in ("km/l", "km\\l", "km/lt", "kml", "media"))

        if len(nums) >= 3:
            km_ini, km_fin, km_lt = nums[0], nums[1], nums[2]
        elif len(nums) == 2:
            if has_media_hint:
                km_fin, km_lt = nums[0], nums[1]
            else:
                km_ini, km_fin = nums[0], nums[1]
        elif len(nums) == 1:
            if has_media_hint:
                km_lt = nums[0]
            else:
                km_fin = nums[0]

        if km_lt is None and media_raw:
            v = _to_float_br(media_raw)
            if v is not None:
                km_lt = v

        km_ini_txt = _num_plain(km_ini, 1) if km_ini is not None else "0.0"
        km_fin_txt = _num_plain(km_fin, 1) if km_fin is not None else "0.0"
        km_lt_txt = _num_plain(km_lt, 2) if km_lt is not None else "0.0"
        return placa_txt, km_ini_txt, km_fin_txt, km_lt_txt

    def _sanitize_inline(value: Any) -> str:
        s = str(value or "")
        s = s.replace("\r", " ").replace("\n", " ")
        s = re.sub(r"\s+", " ", s).strip()
        return s

    def _format_doc_br(value: Any) -> str:
        raw = _sanitize_inline(value)
        if not raw:
            return ""
        digits = re.sub(r"\D", "", raw)
        if len(digits) == 11:
            return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"
        if len(digits) == 14:
            return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"
        return raw

    def _wrap_text(value: Any, max_len: int) -> List[str]:
        text = _sanitize_inline(value)
        if not text:
            return []
        max_len = int(max_len or 0)
        if max_len <= 0:
            return [text]
        out: List[str] = []
        s = text
        while s:
            if len(s) <= max_len:
                out.append(s)
                break
            cut = s.rfind(" ", 0, max_len + 1)
            if cut < max(8, int(max_len * 0.4)):
                cut = max_len
            out.append(s[:cut].rstrip())
            s = s[cut:].lstrip()
        return out

    purchase_map = purchase_info_map or {}
    resumo: Dict[str, Dict[str, float]] = {}

    lines: List[str] = []
    lines.append(_line_lr(company, dt_txt) if (company or dt_txt) else dt_txt)
    lines.append(_dash())
    lines.append("")
    lines.append("DETALHAMENTO DE FATURAS".center(width))
    lines.append("")
    lines.append(_border("-"))
    lines.append(_box_line(f" Número de faturas: {len(invs)}"))
    lines.append(_border("-"))
    lines.append("")

    for inv in invs:
        inv_id = str(getattr(inv, "invoice_id", "") or "").strip() or "N/A"
        cust_code = str(getattr(inv, "customer_code", "") or "").strip()
        cust_name = str(getattr(inv, "customer_name", "") or "").strip()
        issue = str(getattr(inv, "issue_date_display", lambda: "")() or "").strip()
        due = str(getattr(inv, "due_date_display", lambda: "")() or "").strip()
        valor_fatura = str(getattr(inv, "open_balance_display", lambda: "")() or "").strip()
        desconto = _num_br(getattr(inv, "discount_amount", 0) or 0, 2)

        key = getattr(inv, "movto_id", None)
        if key in (None, "", 0, "0"):
            key = getattr(inv, "invoice_id", None)
        slot = purchase_map.get(key) or purchase_map.get(str(key)) or {}
        documents = slot.get("documents") or []

        cust_doc = (
            slot.get("customer_doc")
            or slot.get("cpf_cnpj")
            or slot.get("cnpj")
            or slot.get("cpf")
            or getattr(inv, "customer_doc", None)
            or getattr(inv, "cpf_cnpj", None)
            or getattr(inv, "cnpj", None)
            or getattr(inv, "cpf", None)
        )
        cust_doc = _format_doc_br(cust_doc)

        inv_obs = (
            slot.get("invoice_obs")
            or slot.get("obs")
            or slot.get("observacao")
            or getattr(inv, "invoice_obs", None)
            or getattr(inv, "obs", None)
            or getattr(inv, "observacao", None)
        )
        inv_obs = _sanitize_inline(inv_obs)

        lines.append(_border("-"))
        right_doc = f"CPF/CNPJ: {cust_doc}".rstrip() if cust_doc else "CPF/CNPJ:"
        lines.append(_box_lr(f" Cliente...: {cust_code} - {cust_name}".strip(), right_doc))
        lines.append(_box_line(f" Fatura nr.: {inv_id}    Emissão: {issue}    Vencimento: {due}".strip()))
        obs_prefix = " Obs....: "
        obs_max = max(10, content_inner - len(obs_prefix))
        obs_lines = _wrap_text(inv_obs, obs_max)
        if obs_lines:
            for idx, chunk in enumerate(obs_lines):
                prefix = obs_prefix if idx == 0 else (" " * len(obs_prefix))
                lines.append(_box_line(prefix + chunk))
        else:
            lines.append(_box_line(" Obs....:"))
        lines.append(_border("-"))

        doc_w = 9
        nota_w = 6
        dt_w = 10
        placa_w = 7
        qty_w = 7
        km_w = 7
        kmlt_w = 5

        unit_w = 8
        val_w = 8
        if documents:
            max_unit_len = 0
            max_val_len = 0
            for doc in documents:
                items = doc.get("items") or []
                if not items:
                    max_val_len = max(max_val_len, len(str(money_br(doc.get("total")) or "")))
                    continue
                for it in items:
                    qty = _safe_float(it.get("quantity"))
                    val = _safe_float(it.get("item_total"))
                    unit = (val / qty) if qty else 0.0
                    max_unit_len = max(max_unit_len, len(str(_num_br(unit, 3) or "")))
                    max_val_len = max(max_val_len, len(str(_num_br(val, 2) or "")))

            unit_w = max(unit_w, max_unit_len, len("Pr. unit."), len("Pr. tab."))
            val_w = max(val_w, max_val_len, len("Vl. item"), len("Vl. tab."))

        header_prefix = (
            f" {'Cupom':<{doc_w}}  "
            f"{'N.nota':<{nota_w}}  "
            f"{'Dt emissão':<{dt_w}}  "
            f"{'Placa':<{placa_w}}  "
            f"{'Qtde':>{qty_w}}  "
            f"{'Km inic.':>{km_w}}  "
            f"{'Km final':>{km_w}} "
            f"{'Km/Lt':>{kmlt_w}}  "
        )
        header_suffix = (
            f"  {'Pr. unit.':>{unit_w}}"
            f"  {'Vl. item':>{val_w}}"
            f"  {'Pr. tab.':>{unit_w}}"
            f"  {'Vl. tab.':>{val_w}}"
        )
        product_w = max(0, content_inner - len(header_prefix) - len(header_suffix))
        header = header_prefix + f"{'Produto':<{product_w}}" + header_suffix
        lines.append(_box_line(header))
        lines.append(_border("-"))

        subtotal_qty = 0.0
        subtotal_val = 0.0

        if documents:
            for doc in documents:
                doc_num = str(doc.get("documento") or "").strip() or "N/A"
                doc_dt = _date_br(doc.get("dt"))
                placa_txt, km_ini_txt, km_fin_txt, km_lt_txt = _km_fields(doc)
                items = doc.get("items") or []
                if not items:
                    total_txt = money_br(doc.get("total"))
                    row_prefix = (
                        f" {doc_num:<{doc_w}}  "
                        f"{'':<{nota_w}}  "
                        f"{doc_dt:<{dt_w}}  "
                        f"{placa_txt[:placa_w]:<{placa_w}}  "
                        f"{'':>{qty_w}}  "
                        f"{'':>{km_w}}  "
                        f"{'':>{km_w}} "
                        f"{'':>{kmlt_w}}  "
                    )
                    row_suffix = f"  {'':>{unit_w}}  {total_txt:>{val_w}}  {'':>{unit_w}}  {total_txt:>{val_w}}"
                    product_w_row = max(0, content_inner - len(row_prefix) - len(row_suffix))
                    row = row_prefix + f"{'(sem itens)':<{product_w_row}}" + row_suffix
                    lines.append(_box_line(row))
                    subtotal_val += _safe_float(doc.get("total"))
                    continue
                for it in items:
                    prod = html.unescape(str(it.get("product") or "").strip()) or "Item"
                    prod_type = str(it.get("product_type") or it.get("tipo") or "").strip().upper()
                    show_km = prod_type == "C"
                    qty = _safe_float(it.get("quantity"))
                    val = _safe_float(it.get("item_total"))
                    unit = (val / qty) if qty else 0.0
                    qty_txt = _num_br(qty, 2)
                    unit_txt = _num_br(unit, 3)
                    val_txt = _num_br(val, 2)
                    km_ini_out = km_ini_txt if show_km else ""
                    km_fin_out = km_fin_txt if show_km else ""
                    km_lt_out = km_lt_txt if show_km else ""
                    row_prefix = (
                        f" {doc_num:<{doc_w}}  "
                        f"{'':<{nota_w}}  "
                        f"{doc_dt:<{dt_w}}  "
                        f"{placa_txt[:placa_w]:<{placa_w}}  "
                        f"{qty_txt:>{qty_w}}  "
                        f"{km_ini_out:>{km_w}}  "
                        f"{km_fin_out:>{km_w}} "
                        f"{km_lt_out:>{kmlt_w}}  "
                    )
                    row_suffix = f"  {unit_txt:>{unit_w}}  {val_txt:>{val_w}}  {unit_txt:>{unit_w}}  {val_txt:>{val_w}}"
                    product_w_row = max(0, content_inner - len(row_prefix) - len(row_suffix))
                    row = row_prefix + f"{prod[:product_w_row]:<{product_w_row}}" + row_suffix
                    lines.append(_box_line(row))
                    if show_km:
                        subtotal_qty += qty
                    subtotal_val += val

                    s = resumo.get(prod)
                    if not s:
                        s = {"quantity": 0.0, "total": 0.0}
                        resumo[prod] = s
                    s["quantity"] += qty
                    s["total"] += val
        else:
            lines.append(_box_line(" (sem itens)"))

        lines.append(_box_lr(f" Subtotal litros combustivel: {_num_br(subtotal_qty, 2)}", f"Subtotal....: {_num_br(subtotal_val, 2)}"))
        lines.append("")
        total_line = (
            f" Total litros combustivel: {_num_br(subtotal_qty, 2)}"
            f"     Faturado: {_num_br(subtotal_val, 2)}"
            f"     Desconto: {desconto}"
            f"     Valor fatura: {valor_fatura or _num_br(subtotal_val, 2)}"
        )
        lines.append(_box_line(total_line))
        lines.append(_border("-"))
        lines.append("")

    if resumo:
        sum_width = width
        sum_inner = sum_width - 2
        sum_content_inner = sum_inner - 1
        qty_w = 10
        val_w = 10
        tab_w = 12
        desc_w = max(
            0,
            sum_content_inner
            - (1 + len(" | ") + qty_w + len(" | ") + val_w + len(" | ") + tab_w),
        )

        def _s_border() -> str:
            return "+" + ("-" * sum_inner) + "+"

        def _s_line(text: str = "") -> str:
            text = str(text or "")
            return "|" + text[:sum_content_inner].ljust(sum_content_inner) + " " + "|"

        lines.append(_s_border())
        lines.append(_s_line(f" Resumo cliente: {customer}".strip()))
        lines.append(_s_border())
        lines.append(_s_border())
        lines.append(
            _s_line(
                f" {'Descrição':<{desc_w}} | {'Quantidade':>{qty_w}} | {'Valor':>{val_w}} | {'Valor tabela':>{tab_w}}"
            )
        )
        lines.append(_s_border())
        for prod, vals in sorted(resumo.items(), key=lambda kv: kv[0].lower()):
            q_txt = _num_br(vals.get("quantity") or 0, 2)
            v_txt = _num_br(vals.get("total") or 0, 2)
            row_suffix = f" | {q_txt:>{qty_w}} | {v_txt:>{val_w}} | {v_txt:>{tab_w}}"
            desc_w_row = max(0, sum_content_inner - len(row_suffix) - 1)
            row = f" {prod[:desc_w_row]:<{desc_w_row}}" + row_suffix
            lines.append(_s_line(row))
        lines.append(_s_border())

    content = "\n".join([str(l).rstrip() for l in lines]).rstrip() + "\n"
    return content.encode("utf-8"), filename


def format_date_br(value: date) -> str:
    return value.strftime(DATE_INPUT_FORMAT)


def parse_flexible_date(value: str):
    value = (value or "").strip()
    if not value:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def bind_date_entry_shortcuts(entry):
    def _set_entry_text(value: str):
        entry.delete(0, "end")
        entry.insert(0, value)
        entry.icursor("end")

    def _apply_shortcut(action: str):
        if action == "today":
            new_value = date.today()
        else:
            current_value = parse_flexible_date(entry.get()) or date.today()
            delta_days = -1 if action == "minus" else 1
            new_value = current_value + timedelta(days=delta_days)

        _set_entry_text(format_date_br(new_value))

    def _format_typed_date(event=None):
        keysym = getattr(event, "keysym", "")
        if keysym in (
            "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R",
            "Left", "Right", "Up", "Down", "Home", "End", "Tab", "Return",
            "Escape"
        ):
            return None

        digits = re.sub(r"\D", "", entry.get())[:8]
        if not digits:
            if entry.get():
                _set_entry_text("")
            return None

        if len(digits) <= 2:
            formatted = digits
        elif len(digits) <= 4:
            formatted = f"{digits[:2]}/{digits[2:]}"
        else:
            formatted = f"{digits[:2]}/{digits[2:4]}/{digits[4:]}"

        if entry.get() != formatted:
            _set_entry_text(formatted)
        return None

    def _handle_keypress(event=None):
        keysym = getattr(event, "keysym", "")
        char = getattr(event, "char", "")

        if keysym in ("equal", "KP_Equal") or char == "=":
            _apply_shortcut("today")
            return "break"

        if keysym in ("minus", "KP_Subtract") or char == "-":
            _apply_shortcut("minus")
            return "break"

        if keysym in ("plus", "KP_Add") or char == "+":
            _apply_shortcut("plus")
            return "break"

        return None

    entry.bind("<KeyPress>", _handle_keypress, add="+")
    entry.bind("<KeyRelease>", _format_typed_date, add="+")


def _parse_flexible_time_hhmm(value: str) -> Optional[time]:
    raw = str(value or "").strip().lower().replace(" ", "")
    if not raw:
        return None
    raw = raw.replace("h", ":")
    if raw.endswith(":"):
        raw = raw + "00"
    if ":" not in raw:
        raw = raw + ":00"
    parts = raw.split(":")
    if len(parts) >= 2:
        raw = f"{parts[0]}:{parts[1]}"
    try:
        return datetime.strptime(raw, "%H:%M").time()
    except Exception:
        return None


def bind_time_entry_shortcuts(entry):
    def _set_entry_text(value: str):
        entry.delete(0, "end")
        entry.insert(0, value)
        entry.icursor("end")

    def _apply_shortcut(action: str):
        if action == "now":
            new_dt = datetime.now().replace(second=0, microsecond=0)
        else:
            current_t = _parse_flexible_time_hhmm(entry.get())
            if current_t is None:
                new_dt = datetime.now().replace(second=0, microsecond=0)
            else:
                new_dt = datetime.combine(date.today(), current_t)
            new_dt = new_dt + timedelta(minutes=(-1 if action == "minus" else 1))
        _set_entry_text(new_dt.strftime("%H:%M"))

    def _format_typed_time(event=None):
        keysym = getattr(event, "keysym", "")
        if keysym in (
            "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R",
            "Left", "Right", "Up", "Down", "Home", "End", "Tab", "Return",
            "Escape"
        ):
            return None

        digits = re.sub(r"\D", "", entry.get())[:4]
        if not digits:
            if entry.get():
                _set_entry_text("")
            return None

        if len(digits) <= 2:
            formatted = digits
        else:
            formatted = f"{digits[:2]}:{digits[2:]}"

        if entry.get() != formatted:
            _set_entry_text(formatted)
        return None

    def _handle_keypress(event=None):
        keysym = getattr(event, "keysym", "")
        char = getattr(event, "char", "")

        if keysym in ("equal", "KP_Equal") or char == "=":
            _apply_shortcut("now")
            return "break"

        if keysym in ("minus", "KP_Subtract") or char == "-":
            _apply_shortcut("minus")
            return "break"

        if keysym in ("plus", "KP_Add") or char == "+":
            _apply_shortcut("plus")
            return "break"

        return None

    entry.bind("<KeyPress>", _handle_keypress, add="+")
    entry.bind("<KeyRelease>", _format_typed_time, add="+")


def money_br(value: Any) -> str:
    if value in (None, ""):
        return "0,00"
    try:
        num = float(value)
        return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(value)


def pix_amount_str(value: Any) -> Optional[str]:
    if value in (None, "", 0, "0"):
        return None
    try:
        s = str(value).strip()
        if not s:
            return None
        s = re.sub(r"[^\d,.\-]", "", s)
        if not s:
            return None
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
        d = Decimal(s).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if d <= 0:
            return None
        return f"{d:.2f}"
    except (InvalidOperation, ValueError):
        return None


def build_pix_brcode_payload(chave, nome, cidade, valor, txid: str = "***") -> str:
    def f(field_id: int, val: Any) -> str:
        val = str(val or "")
        return f"{field_id:02d}{len(val):02d}{val}"

    chave_limpa = "".join(filter(str.isdigit, str(chave or "")))
    if not chave_limpa:
        chave_limpa = str(chave or "")

    account_info = f(0, "br.gov.bcb.pix") + f(1, chave_limpa)
    payload = f(0, "01")
    payload += f(26, account_info)
    payload += f(52, "0000")
    payload += f(53, "986")
    amount = pix_amount_str(valor)
    if amount:
        payload += f(54, amount)
    payload += f(58, "BR")
    payload += f(59, (nome or "RECEBEDOR")[:25])
    payload += f(60, (cidade or "CIDADE")[:15])
    payload += f(62, f(5, txid))
    payload += "6304"

    crc = 0xFFFF
    for b in payload.encode("utf-8"):
        crc ^= b << 8
        for _ in range(8):
            if crc & 0x8000:
                crc = (crc << 1) ^ 0x1021
            else:
                crc <<= 1
            crc &= 0xFFFF
    return payload + f"{crc:04X}"


def _pix_payload_for_boleto(boleto_data: Dict[str, Any], invoice_row: "InvoiceRow") -> str:
    cedente_doc = str((boleto_data or {}).get("cedente_documento") or "").strip()
    cedente_nome = str((boleto_data or {}).get("cedente_nome") or "").strip()
    if not cedente_doc and not cedente_nome:
        return ""
    valor_src = (boleto_data or {}).get("valor")
    if valor_src in (None, "", 0, "0"):
        valor_src = getattr(invoice_row, "open_balance", None)
    if valor_src in (None, "", 0, "0"):
        valor_src = getattr(invoice_row, "amount", None)
    if valor_src in (None, "", 0, "0"):
        valor_src = str((boleto_data or {}).get("valor_display") or "").strip()
    try:
        return build_pix_brcode_payload(cedente_doc, cedente_nome, "CIDADE", valor_src, txid="***")
    except Exception:
        return ""

def qty_br(value: Any) -> str:
    if value in (None, ""):
        return ""
    try:
        num = float(value)
        if abs(num - round(num)) < 1e-9:
            return str(int(round(num)))
        s = f"{num:.3f}".rstrip("0").rstrip(".")
        return s.replace(".", ",")
    except Exception:
        return str(value)

def datetime_br(value: Any) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y %H:%M")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    return str(value)


def format_iso_datetime_br(value: Any, *, with_seconds: bool = True) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    try:
        dt = datetime.fromisoformat(raw)
        return dt.strftime("%d/%m/%Y %H:%M:%S" if with_seconds else "%d/%m/%Y %H:%M")
    except Exception:
        try:
            dt = datetime.fromisoformat(raw.replace(" ", "T", 1))
            return dt.strftime("%d/%m/%Y %H:%M:%S" if with_seconds else "%d/%m/%Y %H:%M")
        except Exception:
            return raw.replace("T", " ")

_invoice_mismatch_event_keys: set[str] = set()

def build_purchase_info_blocks(invoice_row: InvoiceRow, purchase_info_map: Optional[Dict[Any, Dict[str, Any]]]) -> tuple[str, str]:
    purchase_info_map = purchase_info_map or {}
    key = getattr(invoice_row, "movto_id", None)
    if key in (None, "", 0, "0"):
        key = getattr(invoice_row, "invoice_id", None)
    info = purchase_info_map.get(key) or purchase_info_map.get(str(key)) or {}
    documents = info.get("documents") or []
    items_total = info.get("items_total")
    items_total_gross = info.get("items_total_gross")
    invoice_amount = info.get("invoice_amount")
    invoice_discount = info.get("invoice_discount")
    if not documents and items_total in (None, ""):
        return "", ""

    try:
        items_total_num = float(items_total) if items_total not in (None, "") else None
    except Exception:
        items_total_num = None
    try:
        items_total_gross_num = float(items_total_gross) if items_total_gross not in (None, "") else None
    except Exception:
        items_total_gross_num = None
    if items_total_gross_num is None:
        items_total_gross_num = items_total_num
    try:
        invoice_amount_num = float(invoice_amount) if invoice_amount not in (None, "") else None
    except Exception:
        invoice_amount_num = None

    lines: List[str] = []
    lines.append("Informações da compra:")
    for d in documents:
        doc_no = str(d.get("documento") or "").strip() or "N/A"
        doc_dt = d.get("dt")
        doc_total = d.get("total")
        doc_items_total = d.get("items_total")
        header = f"- Documento: {doc_no}"
        if doc_dt not in (None, ""):
            header += f" | Data/hora: {datetime_br(doc_dt)}"
        if doc_total not in (None, ""):
            header += f" | Total: {money_br(doc_total)}"
        if doc_total not in (None, "") and doc_items_total not in (None, ""):
            try:
                adj = float(doc_total) - float(doc_items_total)
                if abs(adj) > 0.01:
                    header += f" | Itens: {money_br(doc_items_total)} | Ajuste: {money_br(adj)}"
            except Exception:
                pass
        lines.append(header)
        for it in (d.get("items") or []):
            prod = html.unescape(str(it.get("product") or "").strip()) or "Item"
            q = it.get("quantity")
            t = it.get("item_total")
            q_txt = qty_br(q) if q not in (None, "") else ""
            t_txt = money_br(t) if t not in (None, "") else ""
            lines.append(f"  - {prod} | Qtd: {q_txt} | Valor: {t_txt}".rstrip(" |"))
    has_gross_diff = False
    if items_total_num is not None and items_total_gross_num is not None:
        try:
            has_gross_diff = abs(float(items_total_gross_num) - float(items_total_num)) > 0.01
        except Exception:
            has_gross_diff = False
    if items_total_num is not None:
        if has_gross_diff and items_total_gross_num is not None:
            lines.append(f"Total produtos (bruto): {money_br(items_total_gross_num)}")
            lines.append(f"Total faturado: {money_br(items_total_num)}")
            if invoice_amount_num is not None:
                try:
                    if abs(float(invoice_amount_num) - float(items_total_num)) <= 0.01:
                        lines.append("As diferenças entre a soma dos valores dos produtos e o valor faturado ocorrem devido às formas de pagamento a prazo.")
                except Exception:
                    pass
        else:
            lines.append(f"Total produtos: {money_br(items_total_num)}")
    if invoice_amount_num is not None and items_total_num is not None:
        try:
            if abs(float(invoice_amount_num) - float(items_total_num)) > 0.01:
                diff = float(items_total_num) - float(invoice_amount_num)
                lines.append(
                    f"Atenção: soma dos produtos ({money_br(items_total_num)}) difere do valor da fatura ({money_br(invoice_amount_num)}). Diferença: {money_br(diff)}."
                )
                try:
                    inv_id = getattr(invoice_row, "invoice_id", None)
                    inv_id_txt = str(inv_id or "").strip()
                    k = f"{inv_id_txt}|{money_br(invoice_amount_num)}|{money_br(items_total_num)}|{money_br(items_total_gross_num)}"
                    if inv_id_txt and k not in _invoice_mismatch_event_keys:
                        _invoice_mismatch_event_keys.add(k)
                        from app_core.documents_history import DocumentsHistory

                        DocumentsHistory().add_event(
                            kind="invoice_items_total_mismatch",
                            source="conferencia_fatura",
                            level="error",
                            title="Divergência: soma dos produtos x fatura",
                            message=(
                                f"Fatura: {inv_id_txt} | Cliente: {str(getattr(invoice_row, 'customer_name', '') or '').strip()} | "
                                f"Valor fatura: {money_br(invoice_amount_num)} | Total produtos: {money_br(items_total_num)} | "
                                f"Total bruto: {money_br(items_total_gross_num)} | Diferença: {money_br(diff)}"
                            ),
                        )
                except Exception:
                    pass
        except Exception:
            pass
    text_block = "\n".join(lines).rstrip() + "\n"

    html_rows: List[str] = []
    for d in documents:
        doc_no = str(d.get("documento") or "").strip() or "N/A"
        doc_dt = d.get("dt")
        doc_total = d.get("total")
        doc_items_total = d.get("items_total")
        doc_header = f"<b>Documento:</b> {html.escape(doc_no)}"
        if doc_dt not in (None, ""):
            doc_header += f" | <b>Data/hora:</b> {html.escape(datetime_br(doc_dt))}"
        if doc_total not in (None, ""):
            doc_header += f" | <b>Total:</b> {html.escape(money_br(doc_total))}"
        if doc_total not in (None, "") and doc_items_total not in (None, ""):
            try:
                adj = float(doc_total) - float(doc_items_total)
                if abs(adj) > 0.01:
                    doc_header += f" | <b>Itens:</b> {html.escape(money_br(doc_items_total))} | <b>Ajuste:</b> {html.escape(money_br(adj))}"
            except Exception:
                pass
        html_rows.append(f"<tr style='background-color: #e9ecef;'><td colspan='3'>{doc_header}</td></tr>")
        for it in (d.get("items") or []):
            prod = html.unescape(str(it.get("product") or "").strip()) or "Item"
            q = it.get("quantity")
            t = it.get("item_total")
            q_txt = qty_br(q) if q not in (None, "") else ""
            t_txt = money_br(t) if t not in (None, "") else ""
            html_rows.append(
                "<tr>"
                f"<td>{html.escape(prod)}</td>"
                f"<td>{html.escape(q_txt)}</td>"
                f"<td>{html.escape(t_txt)}</td>"
                "</tr>"
            )
    has_gross_diff_html = False
    if items_total_num is not None and items_total_gross_num is not None:
        try:
            has_gross_diff_html = abs(float(items_total_gross_num) - float(items_total_num)) > 0.01
        except Exception:
            has_gross_diff_html = False
    if items_total_num is not None:
        if has_gross_diff_html and items_total_gross_num is not None:
            html_rows.append(
                "<tr style='background-color: #f8f9fa;'>"
                "<td colspan='2' style='text-align: right;'>Total produtos (bruto)</td>"
                f"<td>{html.escape(money_br(items_total_gross_num))}</td>"
                "</tr>"
            )
            html_rows.append(
                "<tr style='background-color: #f8f9fa; font-weight: bold;'>"
                "<td colspan='2' style='text-align: right;'>Total faturado</td>"
                f"<td>{html.escape(money_br(items_total_num))}</td>"
                "</tr>"
            )
        else:
            html_rows.append(
                "<tr style='background-color: #f8f9fa; font-weight: bold;'>"
                "<td colspan='2' style='text-align: right;'>Total produtos</td>"
                f"<td>{html.escape(money_br(items_total_num))}</td>"
                "</tr>"
            )
    html_warn = ""
    html_note = ""
    if items_total_num is not None and items_total_gross_num is not None and invoice_amount_num is not None:
        try:
            if abs(float(items_total_gross_num) - float(items_total_num)) > 0.01 and abs(float(invoice_amount_num) - float(items_total_num)) <= 0.01:
                html_note = (
                    "<div class='note'>"
                    "<b>Observação:</b> As diferenças entre a soma dos valores dos produtos e o valor faturado ocorrem devido às formas de pagamento a prazo."
                    "</div>"
                )
        except Exception:
            html_note = ""
    if invoice_amount_num is not None and items_total_num is not None:
        try:
            if abs(float(invoice_amount_num) - float(items_total_num)) > 0.01:
                diff = float(items_total_num) - float(invoice_amount_num)
                html_warn = (
                    "<div class='note'>"
                    f"<b>Atenção:</b> soma dos produtos ({html.escape(money_br(items_total_num))}) difere do valor da fatura ({html.escape(money_br(invoice_amount_num))}). "
                    f"Diferença: {html.escape(money_br(diff))}."
                    "</div>"
                )
        except Exception:
            pass
    html_block = (
        "<h3>Informações da compra</h3>"
        "<table>"
        "<thead><tr><th>Produto</th><th>Qtd</th><th>Valor</th></tr></thead>"
        f"<tbody>{''.join(html_rows)}</tbody>"
        "</table>"
        f"{html_warn}{html_note}"
    )
    return "\n" + text_block + "\n", html_block


def build_due_alert_email_body(
    customer_name: str,
    base_date: date,
    invoices: List["InvoiceRow"],
    missing_count: int,
    extra_body: str,
    purchase_info_map: Optional[Dict[Any, Dict[str, Any]]] = None,
    attachment_flags: Optional[Dict[str, bool]] = None,
) -> tuple[str, str]:
    def _status_text(vencto: Any) -> str:
        if not isinstance(vencto, date):
            return ""
        diff = (vencto - base_date).days
        if diff == 0:
            return "Seu boleto vence hoje."
        if diff > 0:
            return f"Seu boleto vencerá em {diff} dia(s)."
        return f"Seu boleto está vencido há {abs(diff)} dia(s)."

    total = 0.0
    lines = []
    html_rows = []
    for inv in invoices:
        try:
            total += float(getattr(inv, "open_balance", 0) or 0)
        except Exception:
            pass
        status = _status_text(getattr(inv, "due_date", None))
        lines.append(f"- Fatura {inv.invoice_id} | Emissão: {inv.issue_date_display()} | Venc.: {inv.due_date_display()} | Saldo: {inv.open_balance_display()} | {status}")
        html_rows.append(
            "<tr>"
            f"<td>{html.escape(str(inv.invoice_id))}</td>"
            f"<td>{html.escape(inv.issue_date_display())}</td>"
            f"<td>{html.escape(inv.due_date_display())}</td>"
            f"<td>{html.escape(inv.open_balance_display())}</td>"
            f"<td>{html.escape(status)}</td>"
            "</tr>"
        )

    note = ""
    note_html = ""
    if attachment_flags is not None:
        note_text = build_attachments_note_text(
            has_boleto=bool(attachment_flags.get("boleto")),
            has_fatura_pdf=bool(attachment_flags.get("fatura_pdf")),
            has_fatura_txt=bool(attachment_flags.get("fatura_txt")),
            has_xml=bool(attachment_flags.get("xml")),
            has_danfe=bool(attachment_flags.get("danfe")),
            has_assinatura=bool(attachment_flags.get("assinatura")),
            missing_boleto_count=int(missing_count or 0),
        )
        if note_text:
            note = "\n" + note_text + "\n"
        note_html = build_attachments_note_html(
            has_boleto=bool(attachment_flags.get("boleto")),
            has_fatura_pdf=bool(attachment_flags.get("fatura_pdf")),
            has_fatura_txt=bool(attachment_flags.get("fatura_txt")),
            has_xml=bool(attachment_flags.get("xml")),
            has_danfe=bool(attachment_flags.get("danfe")),
            has_assinatura=bool(attachment_flags.get("assinatura")),
            missing_boleto_count=int(missing_count or 0),
        )

    extra = html.unescape(str(extra_body or "").strip())
    extra_text = f"{extra}\n\n" if extra else ""
    extra_html = ""
    if extra:
        chunks = re.split(r"\n\s*\n", extra.strip())
        extra_html = "".join([f"<p>{html.escape(c.strip()).replace(chr(10), '<br>')}</p>" for c in chunks if c.strip()])

    company = str((invoices[0].company if invoices else "") or "").strip()
    base_txt = base_date.strftime("%d/%m/%Y") if isinstance(base_date, date) else ""

    purchase_text_parts: List[str] = []
    purchase_html_parts: List[str] = []
    for inv in invoices:
        t, h = build_purchase_info_blocks(inv, purchase_info_map)
        if str(t or "").strip():
            purchase_text_parts.append(f"Fatura {inv.invoice_id}\n{str(t).strip()}")
        if str(h or "").strip():
            purchase_html_parts.append(f"<h4>Fatura {html.escape(str(inv.invoice_id))}</h4>{h}")
    purchase_text = ("\n\n" + "\n\n".join(purchase_text_parts) + "\n") if purchase_text_parts else ""
    purchase_html = ("<hr>" + "<hr>".join(purchase_html_parts)) if purchase_html_parts else ""

    text_body = (
        f"Prezado(a),\n\n"
        f"Este é um alerta de vencimento do boleto.\n\n"
        f"Cliente: {customer_name}\n"
        f"Data de referência: {base_txt}\n"
        f"Quantidade de títulos: {len(invoices)}\n"
        f"Total em aberto: {money_br(total)}\n\n"
        + "\n".join(lines)
        + f"{purchase_text}\n"
        + f"{note}\n\n"
        + f"{extra_text}"
        f"Em caso de dúvidas, estamos à disposição.\n\n"
        f"Atenciosamente,\n"
        f"{company}"
    )

    logo_html = f"<div class='logo'><img src='cid:{EMAIL_LOGO_CID}' alt='Logo' width='240' height='60' style='width:240px;height:60px;display:block;border:0;outline:none;text-decoration:none;'></div>"
    html_body = f"""<html>
<head>
<style>
    body {{ font-family: Arial, sans-serif; font-size: 14px; color: #333; }}
    .card {{ max-width: 780px; border: 1px solid #e5e7eb; border-radius: 10px; padding: 16px; background: #ffffff; }}
    .logo {{ margin: 0 0 12px 0; }}
    .logo img {{ max-height: 60px; max-width: 240px; display: block; }}
    .title {{ font-size: 18px; font-weight: 700; color: #2563eb; margin: 0 0 10px 0; }}
    .muted {{ color: #6b7280; margin: 0 0 14px 0; }}
    table {{ border-collapse: collapse; width: 100%; margin-top: 10px; margin-bottom: 12px; }}
    th {{ background-color: #f8fafc; text-align: left; padding: 10px; border: 1px solid #e5e7eb; }}
    td {{ padding: 10px; border: 1px solid #e5e7eb; }}
</style>
</head>
<body>
  <div class="card">
    {logo_html}
    <p class="title">Alerta de vencimento de boleto</p>
    <p class="muted">Cliente: <b>{html.escape(str(customer_name))}</b> &nbsp;|&nbsp; Data de referência: <b>{html.escape(base_txt)}</b></p>
    <p><b>Quantidade de títulos:</b> {len(invoices)}<br><b>Total em aberto:</b> {html.escape(money_br(total))}</p>
    <table>
      <thead>
        <tr>
          <th>Documento / Fatura</th>
          <th>Emissão</th>
          <th>Vencimento</th>
          <th>Valor da fatura</th>
          <th>Situação</th>
        </tr>
      </thead>
      <tbody>
        {''.join(html_rows)}
      </tbody>
    </table>
    {purchase_html}
    {note_html}
    {extra_html}
    <p>Em caso de dúvidas, estamos à disposição.</p>
    <p>Atenciosamente,<br>{html.escape(company)}</p>
  </div>
</body>
</html>"""
    return text_body, html_body


def build_agenda_email_body(
    customer_name: str,
    due_text: str,
    invoices: List[InvoiceRow],
    missing_count: int,
    extra_body: str,
    context_label: str = "Vencimento",
    purchase_info_map: Optional[Dict[Any, Dict[str, Any]]] = None,
    attachment_flags: Optional[Dict[str, bool]] = None,
) -> tuple[str, str]:
    total = 0.0
    lines = []
    html_rows = []
    for inv in invoices:
        try:
            total += float(getattr(inv, "open_balance", 0) or 0)
        except Exception:
            pass
        lines.append(
            f"- Fatura {inv.invoice_id} | Emissão: {inv.issue_date_display()} | Venc.: {inv.due_date_display()} | Total fatura: {inv.open_balance_display()}"
        )
        html_rows.append(
            "<tr>"
            f"<td>{html.escape(str(inv.invoice_id))}</td>"
            f"<td>{html.escape(inv.issue_date_display())}</td>"
            f"<td>{html.escape(inv.due_date_display())}</td>"
            f"<td>{html.escape(inv.open_balance_display())}</td>"
            "</tr>"
        )

    note = ""
    note_html = ""
    if attachment_flags is not None:
        note_text = build_attachments_note_text(
            has_boleto=bool(attachment_flags.get("boleto")),
            has_fatura_pdf=bool(attachment_flags.get("fatura_pdf")),
            has_fatura_txt=bool(attachment_flags.get("fatura_txt")),
            has_xml=bool(attachment_flags.get("xml")),
            has_danfe=bool(attachment_flags.get("danfe")),
            has_assinatura=bool(attachment_flags.get("assinatura")),
            missing_boleto_count=int(missing_count or 0),
        )
        if note_text:
            note = "\n" + note_text + "\n"
        note_html = build_attachments_note_html(
            has_boleto=bool(attachment_flags.get("boleto")),
            has_fatura_pdf=bool(attachment_flags.get("fatura_pdf")),
            has_fatura_txt=bool(attachment_flags.get("fatura_txt")),
            has_xml=bool(attachment_flags.get("xml")),
            has_danfe=bool(attachment_flags.get("danfe")),
            has_assinatura=bool(attachment_flags.get("assinatura")),
            missing_boleto_count=int(missing_count or 0),
        )

    extra = html.unescape(str(extra_body or "").strip())
    extra_text = f"{extra}\n\n" if extra else ""
    extra_html = ""
    if extra:
        chunks = re.split(r"\n\s*\n", extra.strip())
        extra_html = "".join([f"<p>{html.escape(c.strip()).replace(chr(10), '<br>')}</p>" for c in chunks if c.strip()])

    company = str((invoices[0].company if invoices else "") or "").strip()

    purchase_text_parts: List[str] = []
    purchase_html_parts: List[str] = []
    for inv in invoices:
        t, h = build_purchase_info_blocks(inv, purchase_info_map)
        if str(t or "").strip():
            purchase_text_parts.append(f"Fatura {inv.invoice_id}\n{str(t).strip()}")
        if str(h or "").strip():
            purchase_html_parts.append(f"<h4>Fatura {html.escape(str(inv.invoice_id))}</h4>{h}")
    purchase_text = ("\n\n" + "\n\n".join(purchase_text_parts) + "\n") if purchase_text_parts else ""
    purchase_html = ("<hr>" + "<hr>".join(purchase_html_parts)) if purchase_html_parts else ""

    text_body = (
        f"Prezado(a),\n\n"
        f"Segue(m) fatura(s) para conferência e programação do pagamento.\n\n"
        f"Cliente: {customer_name}\n"
        f"{context_label}: {due_text}\n"
        f"Quantidade de títulos: {len(invoices)}\n"
        f"Total previsto: {money_br(total)}\n\n"
        + "\n".join(lines)
        + f"{purchase_text}\n"
        + f"{note}\n\n"
        + f"{extra_text}"
        f"Em caso de dúvidas, ficamos à disposição.\n\n"
        f"Atenciosamente,\n"
        f"{company}"
    )

    logo_html = f"<div class='logo'><img src='cid:{EMAIL_LOGO_CID}' alt='Logo' width='240' height='60' style='width:240px;height:60px;display:block;border:0;outline:none;text-decoration:none;'></div>"
    html_body = f"""<html>
<head>
<style>
    body {{ font-family: Arial, sans-serif; font-size: 14px; color: #333; }}
    .card {{ max-width: 780px; border: 1px solid #e5e7eb; border-radius: 10px; padding: 16px; background: #ffffff; }}
    .logo {{ margin: 0 0 12px 0; }}
    .logo img {{ max-height: 60px; max-width: 240px; display: block; }}
    .title {{ font-size: 18px; font-weight: 700; color: #2563eb; margin: 0 0 10px 0; }}
    .muted {{ color: #6b7280; margin: 0 0 14px 0; }}
    table {{ border-collapse: collapse; width: 100%; margin-top: 10px; margin-bottom: 12px; }}
    th {{ background-color: #f8fafc; text-align: left; padding: 10px; border: 1px solid #e5e7eb; }}
    td {{ padding: 10px; border: 1px solid #e5e7eb; }}
</style>
</head>
<body>
  <div class="card">
    {logo_html}
    <p class="title">Faturas para programação de pagamento</p>
    <p class="muted">Cliente: <b>{html.escape(str(customer_name))}</b> &nbsp;|&nbsp; {html.escape(str(context_label))}: <b>{html.escape(str(due_text))}</b></p>
    <p><b>Quantidade de títulos:</b> {len(invoices)}<br><b>Total previsto:</b> {html.escape(money_br(total))}</p>
    <table>
      <thead>
        <tr>
          <th>Documento / Fatura</th>
          <th>Emissão</th>
          <th>Vencimento</th>
          <th>Total fatura</th>
        </tr>
      </thead>
      <tbody>
        {''.join(html_rows)}
      </tbody>
    </table>
    {purchase_html}
    {note_html}
    {extra_html}
    <p>Em caso de dúvidas, ficamos à disposição.</p>
    <p>Atenciosamente,<br>{html.escape(company)}</p>
  </div>
</body>
</html>"""
    return text_body, html_body


class BusyOverlay:
    def __init__(self, parent):
        self.parent = parent
        self._visible = False
        self._prev_grab = None
        self._pb_running = False
        self._pb_mode = "indeterminate"

        top = tk.Toplevel(parent)
        top.withdraw()
        top.overrideredirect(True)
        top.transient(parent)
        trans_color = "#abcdef"
        try:
            top.configure(background=trans_color)
            top.attributes("-transparentcolor", trans_color)
        except Exception:
            pass

        outer = tk.Frame(top, bg=trans_color, padx=18, pady=18)
        outer.pack(fill="both", expand=True)

        box = ttk.Frame(outer, padding=18, relief="ridge")
        box.place(relx=0.5, rely=0.5, anchor="center")

        self._message_var = tk.StringVar(value="Carregando...")
        ttk.Label(box, textvariable=self._message_var, font=("Segoe UI", 11, "bold")).pack(anchor="center")
        self._detail_var = tk.StringVar(value="")
        ttk.Label(box, textvariable=self._detail_var, font=("Segoe UI", 9), foreground="#374151", wraplength=520).pack(anchor="center", pady=(6, 0))
        self._pb = ttk.Progressbar(box, mode="indeterminate", length=220)
        self._pb.pack(anchor="center", pady=(10, 0))

        self.top = top
        try:
            self.parent.bind("<Configure>", self._on_parent_configure, add="+")
            self.parent.bind("<Map>", self._on_parent_configure, add="+")
        except Exception:
            pass

    def _on_parent_configure(self, event=None):
        if self._visible:
            self._sync_geometry()

    def set_message(self, message: str):
        self._message_var.set(message or "Carregando...")

    def set_detail(self, detail: str):
        self._detail_var.set(detail or "")

    def set_progress(self, current: Optional[int] = None, total: Optional[int] = None):
        if total is not None:
            try:
                total = int(total)
            except Exception:
                total = None
        if current is not None:
            try:
                current = int(current)
            except Exception:
                current = None

        determinate = bool(total and total > 0 and current is not None)
        if determinate:
            if self._pb_mode != "determinate":
                try:
                    self._pb.stop()
                except Exception:
                    pass
                self._pb_running = False
                self._pb_mode = "determinate"
                try:
                    self._pb.configure(mode="determinate")
                except Exception:
                    pass
            try:
                self._pb.configure(maximum=total)
            except Exception:
                pass
            try:
                self._pb.configure(value=max(0, min(current, total)))
            except Exception:
                pass
            return

        if self._pb_mode != "indeterminate":
            try:
                self._pb.configure(mode="indeterminate")
            except Exception:
                pass
            self._pb_mode = "indeterminate"
        try:
            self._pb.configure(value=0)
        except Exception:
            pass
        if self._visible and not self._pb_running:
            try:
                self._pb.start(10)
                self._pb_running = True
            except Exception:
                self._pb_running = False

    def _sync_geometry(self):
        try:
            self.parent.update_idletasks()
            w = max(1, int(self.parent.winfo_width() or 1))
            h = max(1, int(self.parent.winfo_height() or 1))
            x = int(self.parent.winfo_rootx() or 0)
            y = int(self.parent.winfo_rooty() or 0)
            self.top.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass

    def show(self, message: str = "Carregando..."):
        if not self.parent.winfo_exists():
            return
        self.set_message(message)
        self.set_detail("")
        if self._visible:
            self._sync_geometry()
            try:
                self.top.lift()
                self.top.update_idletasks()
            except Exception:
                pass
            return

        self._prev_grab = None
        try:
            self._prev_grab = self.parent.grab_current()
        except Exception:
            self._prev_grab = None

        try:
            self.parent.configure(cursor="watch")
        except Exception:
            pass

        self._sync_geometry()
        try:
            self.top.deiconify()
            self.top.lift()
        except Exception:
            pass
        try:
            self.top.grab_set()
        except Exception:
            pass
        self._visible = True
        self.set_progress(None, None)
        try:
            self.top.update_idletasks()
        except Exception:
            pass


    def hide(self):
        if not self._visible:
            return
        try:
            self._pb.stop()
        except Exception:
            pass
        self._pb_running = False
        self.set_detail("")
        try:
            self.top.grab_release()
        except Exception:
            pass
        try:
            self.top.withdraw()
        except Exception:
            pass
        try:
            self.parent.configure(cursor="")
        except Exception:
            pass
        prev = self._prev_grab
        self._prev_grab = None
        if prev is not None:
            try:
                if prev.winfo_exists():
                    prev.grab_set()
            except Exception:
                pass
        self._visible = False


def _ensure_busy_overlay(window: tk.Toplevel | tk.Tk) -> BusyOverlay:
    overlay = getattr(window, "_busy_overlay", None)
    if overlay is None:
        overlay = BusyOverlay(window)
        setattr(window, "_busy_overlay", overlay)
    return overlay


def run_with_busy(
    window: tk.Toplevel | tk.Tk,
    message: str,
    work,
    on_success,
    on_error=None,
):
    if threading.current_thread() is not threading.main_thread():
        try:
            window.after(0, lambda: run_with_busy(window, message, work, on_success, on_error))
        except Exception:
            pass
        return None
    overlay = _ensure_busy_overlay(window)
    overlay.show(message)
    cancelled = {"value": False}

    def _finish_success(result):
        if cancelled["value"] or not window.winfo_exists():
            return
        overlay.hide()
        on_success(result)

    def _finish_error(err: Exception):
        if cancelled["value"] or not window.winfo_exists():
            return
        overlay.hide()
        if on_error is not None:
            on_error(err)
        else:
            messagebox.showerror(APP_TITLE, str(err), parent=window)

    def _run():
        try:
            result = work()
        except Exception as e:
            try:
                window.after(0, lambda: _finish_error(e))
            except Exception:
                cancelled["value"] = True
            return
        try:
            window.after(0, lambda: _finish_success(result))
        except Exception:
            cancelled["value"] = True

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    return t


def _pdf_escape(text: str) -> str:
    if not text:
        return ""
    text = str(text)
    
    # Standard PDF WinAnsi octal codes for common Portuguese characters
    # (escaping the backslash for Python string)
    rep = {
        'á': r'\341', 'à': r'\340', 'â': r'\342', 'ã': r'\343',
        'é': r'\351', 'ê': r'\352', 'í': r'\355',
        'ó': r'\363', 'ô': r'\364', 'õ': r'\365', 'ú': r'\372',
        'ç': r'\347',
        'Á': r'\301', 'À': r'\300', 'Â': r'\302', 'Ã': r'\303',
        'É': r'\311', 'Ê': r'\312', 'Í': r'\315',
        'Ó': r'\323', 'Ô': r'\324', 'Õ': r'\325', 'Ú': r'\332',
        'Ç': r'\307',
        'º': r'\272', 'ª': r'\252'
    }
    
    # First escape special PDF characters: \, (, )
    text = text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    
    # Then replace accented characters with their octal escape sequences
    # These sequences use the literal backslash in the PDF stream.
    for char, escape in rep.items():
        text = text.replace(char, escape)
        
    return text


def _load_pdf_png_image(path: str) -> Optional[Dict[str, Any]]:
    p = str(path or "").strip()
    if not p or not os.path.isfile(p):
        return None

    def read_u32(b: bytes, off: int) -> int:
        return (b[off] << 24) | (b[off + 1] << 16) | (b[off + 2] << 8) | b[off + 3]

    def paeth(a: int, b: int, c: int) -> int:
        p = a + b - c
        pa = abs(p - a)
        pb = abs(p - b)
        pc = abs(p - c)
        if pa <= pb and pa <= pc:
            return a
        if pb <= pc:
            return b
        return c

    def unfilter(ft: int, scan: bytes, prev: Optional[bytes], bpp: int) -> bytes:
        out = bytearray(len(scan))
        if ft == 0:
            out[:] = scan
            return bytes(out)
        if ft == 1:
            for i in range(len(scan)):
                left = out[i - bpp] if i >= bpp else 0
                out[i] = (scan[i] + left) & 0xFF
            return bytes(out)
        if ft == 2:
            if not prev:
                out[:] = scan
                return bytes(out)
            for i in range(len(scan)):
                out[i] = (scan[i] + prev[i]) & 0xFF
            return bytes(out)
        if ft == 3:
            for i in range(len(scan)):
                left = out[i - bpp] if i >= bpp else 0
                up = prev[i] if prev else 0
                out[i] = (scan[i] + ((left + up) // 2)) & 0xFF
            return bytes(out)
        if ft == 4:
            for i in range(len(scan)):
                left = out[i - bpp] if i >= bpp else 0
                up = prev[i] if prev else 0
                up_left = prev[i - bpp] if (prev and i >= bpp) else 0
                out[i] = (scan[i] + paeth(left, up, up_left)) & 0xFF
            return bytes(out)
        return b""

    try:
        with open(p, "rb") as f:
            data = f.read()
        if len(data) < 16 or data[:8] != b"\x89PNG\r\n\x1a\n":
            return None
        off = 8
        width = height = 0
        bit_depth = 0
        color_type = 0
        interlace = 0
        idat = bytearray()
        while off + 8 <= len(data):
            ln = read_u32(data, off)
            ctype = data[off + 4 : off + 8]
            off += 8
            if off + ln + 4 > len(data):
                break
            chunk = data[off : off + ln]
            off += ln + 4
            if ctype == b"IHDR":
                if ln < 13:
                    return None
                width = read_u32(chunk, 0)
                height = read_u32(chunk, 4)
                bit_depth = chunk[8]
                color_type = chunk[9]
                interlace = chunk[12]
            elif ctype == b"IDAT":
                idat.extend(chunk)
            elif ctype == b"IEND":
                break
        if not width or not height:
            return None
        if bit_depth != 8:
            return None
        if interlace != 0:
            return None
        if color_type not in (2, 6):
            return None

        raw = zlib.decompress(bytes(idat))
        bpp = 3 if color_type == 2 else 4
        stride = width * bpp
        expected = height * (stride + 1)
        if len(raw) < expected:
            return None

        out_rgb = bytearray(width * height * 3)
        prev = None
        in_off = 0
        out_off = 0
        for _ in range(height):
            ft = raw[in_off]
            in_off += 1
            scan = raw[in_off : in_off + stride]
            in_off += stride
            recon = unfilter(ft, scan, prev, bpp)
            if not recon:
                return None
            prev = recon
            if bpp == 3:
                out_rgb[out_off : out_off + (width * 3)] = recon
                out_off += width * 3
            else:
                for i in range(0, len(recon), 4):
                    r = recon[i]
                    g = recon[i + 1]
                    b = recon[i + 2]
                    a = recon[i + 3]
                    if a == 255:
                        out_rgb[out_off] = r
                        out_rgb[out_off + 1] = g
                        out_rgb[out_off + 2] = b
                    elif a == 0:
                        out_rgb[out_off] = 255
                        out_rgb[out_off + 1] = 255
                        out_rgb[out_off + 2] = 255
                    else:
                        inv = 255 - a
                        out_rgb[out_off] = (r * a + 255 * inv) // 255
                        out_rgb[out_off + 1] = (g * a + 255 * inv) // 255
                        out_rgb[out_off + 2] = (b * a + 255 * inv) // 255
                    out_off += 3

        return {
            "width": int(width),
            "height": int(height),
            "stream": zlib.compress(bytes(out_rgb), level=9),
        }
    except Exception:
        return None


def _resolve_sicredi_logo_path(boleto_data: Dict[str, Any]) -> str:
    custom_logo = str((boleto_data or {}).get("logo_path") or "").strip()
    if custom_logo and os.path.isfile(custom_logo):
        return custom_logo

    meipass = ""
    try:
        meipass = str(getattr(sys, "_MEIPASS", "") or "").strip()
    except Exception:
        meipass = ""

    base_dir = os.path.dirname(os.path.abspath(__file__))
    exe_dir = str(app_dir())
    roots = [p for p in [meipass, exe_dir, base_dir, os.path.join(base_dir, "assets")] if str(p or "").strip()]
    names = ["sicredi_logo.png", "logo_sicredi.png", "danfe_emitente_logo.png"]

    candidates: List[str] = []
    for root in roots:
        for name in names:
            candidates.append(os.path.join(root, name))

    for cand in candidates:
        if os.path.isfile(cand):
            return cand
    return ""


def build_text_pdf_bytes(ops: List[str], images: Optional[List[Dict[str, Any]]] = None) -> bytes:
    page_width = 595
    page_height = 842
    
    stream_content = "\n".join(ops)
    # The octal escapes work best with latin-1 (iso-8859-1) encoding
    # which is the default expected for PDF content streams
    stream = stream_content.encode("latin-1", errors="replace")

    img_defs = [i for i in (images or []) if isinstance(i, dict) and i.get("name") and i.get("width") and i.get("height") and i.get("stream")]

    first_img_obj = 7
    xobj_entries = []
    for idx, img in enumerate(img_defs):
        obj_num = first_img_obj + idx
        xobj_entries.append(f"/{img['name']} {obj_num} 0 R")
    xobj_block = f" /XObject << {' '.join(xobj_entries)} >>" if xobj_entries else ""

    objects = []
    objects.append(b"1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n")
    objects.append(b"2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n")
    objects.append(
        f"3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {page_width} {page_height}] /Resources << /Font << /F1 5 0 R /F2 6 0 R >>{xobj_block} >> /Contents 4 0 R >>\nendobj\n".encode("latin-1")
    )
    objects.append(b"4 0 obj\n<< /Length " + str(len(stream)).encode("ascii") + b" >>\nstream\n" + stream + b"\nendstream\nendobj\n")
    objects.append(b"5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj\n")
    objects.append(b"6 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>\nendobj\n")
    for img in img_defs:
        width = int(img["width"])
        height = int(img["height"])
        payload = bytes(img["stream"])
        objects.append(
            b"%d 0 obj\n<< /Type /XObject /Subtype /Image /Width %d /Height %d /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length %d >>\nstream\n"
            % (first_img_obj + len(objects) - 6, width, height, len(payload))
            + payload
            + b"\nendstream\nendobj\n"
        )

    pdf = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]
    for obj in objects:
        offsets.append(len(pdf))
        pdf.extend(obj)

    xref_start = len(pdf)
    pdf.extend(f"xref\n0 {len(objects)+1}\n".encode("ascii"))
    pdf.extend(b"0000000000 65535 f \n")
    for off in offsets[1:]:
        pdf.extend(f"{off:010d} 00000 n \n".encode("ascii"))
    pdf.extend(f"trailer\n<< /Size {len(objects)+1} /Root 1 0 R >>\nstartxref\n{xref_start}\n%%EOF".encode("ascii"))
    return bytes(pdf)


def build_boleto_pdf_bytes(boleto_data: Dict[str, Any], invoice_row: "InvoiceRow", include_pix_qrcode: bool = True) -> bytes:
    def get_qr_matrix(data: str):
        try:
            import qrcode
            qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=1, border=0)
            qr.add_data(data)
            qr.make(fit=True)
            return qr.get_matrix()
        except Exception:
            import hashlib
            size = 25
            matrix = [[0 for _ in range(size)] for _ in range(size)]
            for r in range(7):
                for c in range(7):
                    if r == 0 or r == 6 or c == 0 or c == 6 or (2 <= r <= 4 and 2 <= c <= 4):
                        matrix[r][c] = 1
                        matrix[size - 1 - r][c] = 1
                        matrix[r][size - 1 - c] = 1
            h = hashlib.md5(data.encode()).digest()
            for i in range(len(h) * 8):
                r = 8 + (i // (size - 8))
                c = 8 + (i % (size - 8))
                if r < size and c < size:
                    if (h[i // 8] >> (i % 8)) & 1:
                        matrix[r][c] = 1
            return matrix

    bank_code = str(boleto_data.get("banco_codigo") or "000").strip()
    bank_name = str(boleto_data.get("banco_nome") or "BANCO").strip()
    portador_name = str(boleto_data.get("portador_nome") or "").strip()
    is_sicredi = ("SICREDI" in portador_name.upper()) if portador_name else ((bank_code == "748") and ("SICREDI" in bank_name.upper()))
    nosso_numero = str(boleto_data.get("nosso_numero") or "").strip()
    documento = str(boleto_data.get("documento") or "").strip()
    vencto = str(boleto_data.get("vencto_display") or "").strip()
    valor = str(boleto_data.get("valor_display") or "0,00").strip()
    linha_digitavel = str(boleto_data.get("linha_digitavel") or "").strip()
    codigo_barra = str(boleto_data.get("codigo_barra") or "").strip()

    cedente_nome = str(boleto_data.get("cedente_nome") or "").strip()
    cedente_doc = str(boleto_data.get("cedente_documento") or "").strip()

    sacado_nome = str(boleto_data.get("sacado_nome") or "").strip()
    sacado_doc = str(boleto_data.get("sacado_inscricao") or "").strip()
    sacado_end = str(boleto_data.get("sacado_endereco") or "").strip()
    sacado_cep = str(boleto_data.get("sacado_cep") or "").strip()
    cidade_uf = str(boleto_data.get("sacado_cidade_uf") or "").strip()

    if sacado_cep:
        digits = "".join([c for c in sacado_cep if c.isdigit()])
        cep_fmt = f"{digits[:5]}-{digits[5:]}" if len(digits) == 8 else sacado_cep
        if cep_fmt and ("CEP" not in sacado_end.upper()) and (cep_fmt not in sacado_end):
            sacado_end = (sacado_end + f" - CEP {cep_fmt}").strip(" -")

    agencia = str(boleto_data.get("agencia") or "").strip()
    agencia_dv = str(boleto_data.get("agencia_digito") or "").strip()
    conta = str(boleto_data.get("nr_conta") or "").strip()
    conta_dv = str(boleto_data.get("conta_digito") or "").strip()
    carteira = str(boleto_data.get("portador_carteira") or "").strip()

    agencia_conta = f"{agencia}{'-'+agencia_dv if agencia_dv else ''} / {conta}{'-'+conta_dv if conta_dv else ''}"
    mensagem = str(boleto_data.get("mensagem") or "").strip()

    pix_code = _pix_payload_for_boleto(boleto_data, invoice_row) if include_pix_qrcode else ""

    ops: List[str] = []
    pdf_images: List[Dict[str, Any]] = []
    png_cache: Dict[str, str] = {}

    def draw_text(x, y, text, size=9, bold=False, max_len=None, color: Optional[Tuple[float, float, float]] = None, align: str = "left"):
        raw_text = str(text or "").strip()
        if max_len and len(raw_text) > max_len:
            raw_text = raw_text[: max_len - 3] + "..."

        def _text_width_pt(s: str, font_size: float, is_bold: bool) -> float:
            widths = {
                "0": 0.556,
                "1": 0.556,
                "2": 0.556,
                "3": 0.556,
                "4": 0.556,
                "5": 0.556,
                "6": 0.556,
                "7": 0.556,
                "8": 0.556,
                "9": 0.556,
                ".": 0.278,
                ",": 0.278,
                "/": 0.278,
                "-": 0.333,
                " ": 0.278,
                ":": 0.278,
                "(": 0.333,
                ")": 0.333,
            }
            w = 0.0
            for ch in s:
                w += widths.get(ch, 0.556)
            if is_bold:
                w *= 1.03
            return w * float(font_size)

        text = _pdf_escape(raw_text)
        approx_width = _text_width_pt(raw_text, float(size), bool(bold))
        if align == "center":
            x = float(x) - (approx_width / 2.0)
        elif align == "right":
            x = float(x) - approx_width
        font = "/F2" if bold else "/F1"
        ops.append("BT")
        if color:
            r, g, b = color
            ops.append(f"{max(0.0, min(1.0, float(r))):.3f} {max(0.0, min(1.0, float(g))):.3f} {max(0.0, min(1.0, float(b))):.3f} rg")
        ops.append(f"{font} {size} Tf")
        ops.append(f"{x} {y} Td")
        ops.append(f"({text}) Tj")
        ops.append("ET")
        if color:
            ops.append("0 0 0 rg")

    def draw_line(x1, y1, x2, y2, width=0.5):
        ops.append(f"{width} w")
        ops.append(f"{x1} {y1} m")
        ops.append(f"{x2} {y2} l")
        ops.append("S")

    def draw_rect(x, y, w, h, fill=False):
        ops.append(f"{x} {y} {w} {h} re")
        ops.append("f" if fill else "S")

    def register_png(path: str) -> str:
        p = str(path or "").strip()
        if not p:
            return ""
        if p in png_cache:
            return png_cache[p]
        img = _load_pdf_png_image(p)
        if not img:
            return ""
        name = f"Im{len(pdf_images) + 1}"
        pdf_images.append({"name": name, **img})
        png_cache[p] = name
        return name

    def draw_image(name: str, x: float, y: float, w: float, h: float):
        if not name or w <= 0 or h <= 0:
            return
        ops.append("q")
        ops.append(f"{w:.3f} 0 0 {h:.3f} {x:.3f} {y:.3f} cm")
        ops.append(f"/{name} Do")
        ops.append("Q")

    def wrap_text(text, max_chars):
        if not text:
            return []
        lines = []
        for p in text.split("\n"):
            p = p.strip()
            if not p:
                lines.append("")
                continue
            while len(p) > max_chars:
                idx = p.rfind(" ", 0, max_chars)
                if idx == -1:
                    idx = max_chars
                lines.append(p[:idx].strip())
                p = p[idx:].strip()
            if p:
                lines.append(p)
        return lines

    if is_sicredi:
        left = 40
        width = 515
        right = left + width
        logo_path = _resolve_sicredi_logo_path(boleto_data)

        slot_raw = str((boleto_data or {}).get("page_slot") or "").strip().lower()
        if slot_raw in ("2", "bottom", "inferior", "baixo", "lower"):
            y_offset = 0.0
        else:
            y_offset = 421.0

        def _y(v: float) -> float:
            return float(y_offset) + float(v)

        def _h(v: float) -> float:
            return float(v)

        draw_text_raw = draw_text
        draw_line_raw = draw_line
        draw_rect_raw = draw_rect
        draw_image_raw = draw_image
        font_scale = 0.92

        def draw_text(x, y, text, size=9, bold=False, max_len=None, color: Optional[Tuple[float, float, float]] = None, align: str = "left"):
            return draw_text_raw(x, _y(y), text, size=(float(size) * float(font_scale)), bold=bold, max_len=max_len, color=color, align=align)

        def draw_line(x1, y1, x2, y2, width=0.5):
            return draw_line_raw(x1, _y(y1), x2, _y(y2), width=width)

        def draw_rect(x, y, w, h, fill=False):
            return draw_rect_raw(x, _y(y), w, _h(h), fill=fill)

        def draw_image(name: str, x: float, y: float, w: float, h: float):
            return draw_image_raw(name, x, _y(y), w, _h(h))

        def draw_box(x: float, y: float, w: float, h: float, lw: float = 0.8):
            draw_line(x, y, x + w, y, width=lw)
            draw_line(x, y + h, x + w, y + h, width=lw)
            draw_line(x, y, x, y + h, width=lw)
            draw_line(x + w, y, x + w, y + h, width=lw)

        def draw_logo(x: float, y: float, max_w: float, max_h: float):
            name = register_png(logo_path)
            if not name:
                return
            meta = next((m for m in pdf_images if m.get("name") == name), None)
            lw = float((meta or {}).get("width") or 0.0)
            lh = float((meta or {}).get("height") or 0.0)
            if lw <= 0 or lh <= 0:
                return
            s = min(max_w / lw, max_h / lh)
            draw_image(name, x, y, lw * s, lh * s)

        bank_dvs = {"001": "9", "237": "2", "341": "7", "104": "0", "033": "7", "748": "X", "756": "0", "041": "8"}
        bank_dv = bank_dvs.get(bank_code, "X")

        def set_dash(on: bool):
            ops.append("[2 2] 0 d" if on else "[] 0 d")

        set_dash(True)
        draw_line_raw(left, 421.0, right, 421.0, width=0.6)
        set_dash(False)

        r_top = 410
        r_bottom = 320
        mid = left + (width / 2)

        def draw_receipt_half(x0: float, title: str, *, is_delivery: bool):
            half_w = (width / 2) - 2
            x_end = x0 + half_w
            draw_line(x0, r_top, x_end, r_top, width=1)
            draw_line(x0, r_bottom, x_end, r_bottom, width=1)

            header_h = 16
            header_bottom = r_top - header_h
            draw_line(x0, header_bottom, x_end, header_bottom, width=0.8)

            x_logo_end2 = x0 + 70
            x_bank_end2 = x_logo_end2 + 45
            draw_line(x_logo_end2, header_bottom, x_logo_end2, r_top, width=0.8)
            draw_line(x_bank_end2, header_bottom, x_bank_end2, r_top, width=0.8)

            draw_logo(x0 + 4, header_bottom + 3, 62, 11)
            draw_text(x_logo_end2 + 6, header_bottom + 5, f"{bank_code}-{bank_dv}", size=11.0, bold=True)
            draw_text(x_end - 6, header_bottom + 5, title, size=8.0, bold=True, align="right")

            line_h = 14.0
            table_top = header_bottom
            y1 = table_top - line_h
            y2 = y1 - line_h
            y3 = y2 - line_h
            draw_line(x0, y1, x_end, y1, width=0.8)
            draw_line(x0, y2, x_end, y2, width=0.8)
            draw_line(x0, y3, x_end, y3, width=0.8)

            c1 = x0 + 70
            c2 = x0 + 155
            c3 = x0 + 195
            for cx in (c1, c2, c3):
                draw_line(cx, y1, cx, table_top, width=0.8)

            def _fit_chars(cell_w: float, font_size: float) -> int:
                try:
                    return max(4, int(float(cell_w) / (float(font_size) * 0.55)))
                except Exception:
                    return 20

            def cell_label(xl: float, xr: float, yt: float, text: str, *, size: float = 5.0):
                draw_text(xl + 3, yt - 6, text, size=size, max_len=_fit_chars(xr - xl - 6, size))

            def cell_value(xl: float, xr: float, yb: float, text: str, *, bold: bool = True, align_right: bool = False, size: float = 8.0):
                if align_right:
                    draw_text(xr - 3, yb + 2, text, size=size, bold=bold, align="right", max_len=_fit_chars(xr - xl - 6, size))
                else:
                    draw_text(xl + 3, yb + 2, text, size=size, bold=bold, max_len=_fit_chars(xr - xl - 6, size))

            cell_label(x0, c1, table_top, "Vencimento", size=4.8)
            cell_value(x0, c1, y1, vencto, bold=True, size=7.8)
            cell_label(c1, c2, table_top, "Agência/Código do Beneficiário", size=4.6)
            cell_value(c1, c2, y1, agencia_conta, bold=True, size=7.2)
            cell_label(c2, c3, table_top, "Espécie", size=4.8)
            cell_value(c2, c3, y1, "REAL", bold=True, size=7.8)
            cell_label(c3, x_end, table_top, "Quantidade")

            c21 = x0 + 85
            c22 = x0 + 170
            for cx in (c21, c22):
                draw_line(cx, y2, cx, y1, width=0.8)

            cell_label(x0, c21, y1, "(=) Valor Documento", size=4.6)
            cell_value(x0, c21, y2, valor, bold=True, align_right=True, size=7.8)
            cell_label(c21, c22, y1, "(-) Desconto / Abatimento", size=4.4)
            cell_label(c22, x_end, y1, "(+) Juros / Multa", size=4.4)

            for cx in (c21, c22):
                draw_line(cx, y3, cx, y2, width=0.8)

            cell_label(x0, c21, y2, "(=) Valor Cobrado", size=4.6)
            cell_value(x0, c21, y3, valor, bold=True, align_right=True, size=7.8)
            if is_delivery:
                cell_label(c21, c22, y2, "Nº do Documento", size=4.8)
                cell_value(c21, c22, y3, documento, bold=True, align_right=True, size=7.4)
                cell_label(c22, x_end, y2, "Nosso Número", size=4.8)
                cell_value(c22, x_end, y3, nosso_numero, bold=True, align_right=True, size=7.4)
            else:
                cell_label(c21, c22, y2, "Nosso Número", size=4.8)
                cell_value(c21, c22, y3, nosso_numero, bold=True, align_right=True, size=7.4)
                cell_label(c22, x_end, y2, "Nº do Documento", size=4.8)
                cell_value(c22, x_end, y3, documento, bold=True, align_right=True, size=7.4)

            y_info_top = y3
            y_info_sep = y_info_top - line_h
            draw_text(x0 + 4, y_info_top - 5, "Pagador:", size=5.6)
            payer_max = _fit_chars((x_end - (x0 + 4) - 6), 7.0)
            draw_text(x0 + 4, y_info_top - 12, f"{sacado_nome} - {sacado_doc}", size=7.0, bold=True, max_len=payer_max)
            if not is_delivery:
                draw_line(x0, y_info_sep, x_end, y_info_sep, width=0.6)
                draw_text(x0 + 4, y_info_sep - 4, "Beneficiário:", size=5.6)
                bene_max = _fit_chars((x_end - (x0 + 4) - 6), 6.6)
                draw_text(x0 + 4, y_info_sep - 10, f"{cedente_nome} - {cedente_doc}", size=6.6, bold=False, max_len=bene_max)
                ced_addr = str((boleto_data or {}).get("cedente_endereco") or (boleto_data or {}).get("cedente_endereco_full") or "").strip()
                if not ced_addr:
                    c_log = str((boleto_data or {}).get("cedente_logradouro") or "").strip()
                    c_nro = str((boleto_data or {}).get("cedente_numero") or "").strip()
                    c_bai = str((boleto_data or {}).get("cedente_bairro") or "").strip()
                    c_cid = str((boleto_data or {}).get("cedente_cidade") or "").strip()
                    c_uf = str((boleto_data or {}).get("cedente_estado") or (boleto_data or {}).get("cedente_uf") or "").strip()
                    c_cep = str((boleto_data or {}).get("cedente_cep") or "").strip()
                    parts = [p for p in [c_log, c_nro, c_bai, c_cid, c_uf, c_cep] if p]
                    if parts:
                        ced_addr = " - ".join(parts[:3])
                        if len(parts) > 3:
                            ced_addr = ced_addr + " - " + " / ".join(parts[3:])
                if ced_addr:
                    addr_max = _fit_chars((x_end - (x0 + 4) - 6), 6.4)
                    draw_text(x0 + 4, y_info_sep - 16, ced_addr, size=6.4, bold=False, max_len=addr_max)
            else:
                draw_line(x0, y_info_sep, x_end, y_info_sep, width=0.6)
                draw_text(x0 + 4, y_info_sep - 5, "Assinatura", size=5.6)
                split_x = x0 + (half_w * 0.70)
                draw_text(split_x + 4, y_info_sep - 5, "Data de Entrega", size=5.6)

        set_dash(True)
        draw_line(mid, r_bottom, mid, r_top, width=0.8)
        set_dash(False)
        draw_receipt_half(left, "Recibo do Pagador", is_delivery=False)
        draw_receipt_half(mid + 2, "Recibo de Entrega", is_delivery=True)

        y_bar_bottom = 20
        y_bar_top = 65
        y_pag_bottom = y_bar_top
        y_pag_top = 105
        y_instr_bottom = y_pag_top
        y_instr_top = 195
        y_uso_bottom = y_instr_top
        y_uso_top = 220
        y_doc_bottom = y_uso_top
        y_doc_top = 245
        y_ben_bottom = y_doc_top
        y_ben_top = 270
        y_loc_bottom = y_ben_top
        y_loc_top = 290
        y_head_bottom = y_loc_top
        y_head_top = y_head_bottom + 25

        draw_line(left, y_head_top, right, y_head_top, width=1)
        draw_line(left, y_bar_bottom, right, y_bar_bottom, width=1)
        for yy in (y_bar_top, y_pag_top, y_instr_top, y_uso_top, y_doc_top, y_ben_top, y_loc_top, y_head_bottom):
            draw_line(left, yy, right, yy, width=0.8)

        x_right = left + 380
        draw_line(x_right, y_instr_bottom, x_right, y_loc_top, width=0.8)

        x_logo_end = left + 155
        x_bank_end = x_logo_end + 60
        draw_line(x_logo_end, y_head_bottom, x_logo_end, y_head_top, width=0.8)
        draw_line(x_bank_end, y_head_bottom, x_bank_end, y_head_top, width=0.8)

        draw_logo(left + 6, y_head_bottom + 5, 125, 16)
        draw_text(x_logo_end + 12, y_head_bottom + 9, f"{bank_code}-{bank_dv}", size=13, bold=True)
        draw_text(x_bank_end + 8, y_head_bottom + 10, linha_digitavel, size=8.6, bold=True)

        draw_text(left + 6, y_loc_top - 7, "Local de Pagamento", size=6)
        draw_text(left + 6, y_loc_bottom + 5, "PAGÁVEL PREFERENCIALMENTE NAS COOPERATIVAS DE CRÉDITO DO SICREDI", size=8, bold=True)
        draw_text(x_right + 6, y_loc_top - 7, "Vencimento", size=6)
        draw_text(right - 6, y_loc_bottom + 5, vencto, size=10, bold=True, align="right")

        draw_text(left + 6, y_ben_top - 7, "Beneficiário", size=6)
        ben_lines = wrap_text(f"{cedente_nome} - CNPJ: {cedente_doc}", 60)
        if ben_lines:
            draw_text(left + 6, y_ben_top - 16, ben_lines[0], size=8, bold=True)
        if len(ben_lines) > 1:
            draw_text(left + 6, y_ben_top - 26, ben_lines[1], size=8, bold=True)
        draw_text(x_right + 6, y_ben_top - 7, "Agência / Código do Beneficiário", size=6)
        draw_text(right - 6, y_ben_top - 20, agencia_conta, size=9, align="right")

        x_d1 = left + 95
        x_d2 = left + 190
        x_d3 = left + 250
        x_d4 = left + 295
        x_d5 = left + 345
        for xx in (x_d1, x_d2, x_d3, x_d4, x_d5, x_right):
            draw_line(xx, y_doc_bottom, xx, y_doc_top, width=0.8)

        doc_date = str(boleto_data.get("data_documento") or boleto_data.get("data_documento_display") or vencto or "").strip()
        if not doc_date:
            doc_date = date.today().strftime("%d/%m/%Y")
        proc_date = date.today().strftime("%d/%m/%Y")

        draw_text(left + 6, y_doc_top - 7, "Data do Documento", size=6)
        draw_text(left + 6, y_doc_bottom + 5, doc_date, size=8)
        draw_text(x_d1 + 6, y_doc_top - 7, "Nº do Documento", size=6)
        draw_text(x_d1 + 6, y_doc_bottom + 5, documento, size=8)
        draw_text(x_d2 + 6, y_doc_top - 7, "Espécie Doc.", size=6)
        draw_text(x_d2 + 6, y_doc_bottom + 5, "DM", size=8)
        draw_text(x_d3 + 6, y_doc_top - 7, "Aceite", size=6)
        draw_text(x_d3 + 6, y_doc_bottom + 5, "NÃO", size=8)
        draw_text(x_d4 + 6, y_doc_top - 7, "Data Proc.", size=6)
        draw_text(x_d4 + 6, y_doc_bottom + 5, proc_date, size=8)
        draw_text(x_right + 6, y_doc_top - 7, "Nosso Número", size=6)
        draw_text(right - 6, y_doc_bottom + 5, nosso_numero, size=9, align="right")

        x_u1 = left + 115
        x_u2 = left + 170
        x_u3 = left + 235
        x_u4 = left + 310
        for xx in (x_u1, x_u2, x_u3, x_u4, x_right):
            draw_line(xx, y_uso_bottom, xx, y_uso_top, width=0.8)

        draw_text(left + 6, y_uso_top - 7, "Uso do Banco", size=6)
        draw_text(x_u1 + 6, y_uso_top - 7, "Carteira", size=6)
        draw_text(x_u1 + 6, y_uso_bottom + 5, carteira, size=8)
        draw_text(x_u2 + 6, y_uso_top - 7, "Espécie", size=6)
        draw_text(x_u2 + 6, y_uso_bottom + 5, "REAL", size=8, bold=True)
        draw_text(x_u3 + 6, y_uso_top - 7, "Quantidade", size=6)
        draw_text(x_right + 6, y_uso_top - 7, "(=) Valor do Documento", size=6)
        draw_text(right - 6, y_uso_bottom + 3, valor, size=10, bold=True, align="right")

        draw_text(left + 6, y_instr_top - 10, "Instruções / Observações", size=6)
        show_fixed_instr = (
            bool(mensagem)
            or boleto_data.get("juros_perc_mes") not in (None, "")
            or boleto_data.get("juros_valor_dia") not in (None, "")
            or boleto_data.get("multa_perc") not in (None, "")
            or boleto_data.get("multa_valor") not in (None, "")
            or boleto_data.get("multa_prazo") not in (None, "")
        )
        if show_fixed_instr:
            def _to_float(v):
                try:
                    if v is None:
                        return 0.0
                    if isinstance(v, (int, float)):
                        return float(v)
                    s = str(v).strip().replace(".", "").replace(",", ".")
                    return float(s) if s else 0.0
                except Exception:
                    return 0.0

            def _to_percent(v, default_percent):
                if v in (None, ""):
                    return float(default_percent or 0.0)
                p = _to_float(v)
                if 0 < p <= 1.0:
                    p *= 100.0
                return p

            try:
                amt = float(boleto_data.get("valor") or 0)
            except Exception:
                amt = 0.0
            if not amt:
                s_amt = pix_amount_str(valor) or ""
                try:
                    amt = float(s_amt) if s_amt else 0.0
                except Exception:
                    amt = 0.0

            juros_perc_mes = _to_percent(boleto_data.get("juros_perc_mes"), 1.5)
            multa_perc = _to_percent(boleto_data.get("multa_perc"), 2.0)
            juros_valor_dia = _to_float(boleto_data.get("juros_valor_dia"))
            multa_valor = _to_float(boleto_data.get("multa_valor"))

            daily_interest = juros_valor_dia if juros_valor_dia > 0 else ((amt * (juros_perc_mes / 100.0) / 30.0) if amt and juros_perc_mes else 0.0)
            multa = multa_valor if multa_valor > 0 else ((amt * (multa_perc / 100.0)) if amt and multa_perc else 0.0)

            start_dt_txt = ""
            try:
                v = boleto_data.get("vencto")
                if isinstance(v, datetime):
                    base_dt = v.date()
                elif isinstance(v, date):
                    base_dt = v
                else:
                    base_dt = datetime.strptime(str(vencto or "").strip(), "%d/%m/%Y").date()
                try:
                    multa_prazo = int(boleto_data.get("multa_prazo") or 0)
                except Exception:
                    multa_prazo = 0
                start_dt_txt = (base_dt + timedelta(days=(multa_prazo if multa_prazo > 0 else 1))).strftime("%d/%m/%Y")
            except Exception:
                start_dt_txt = ""

            juros_perc_txt = f"{juros_perc_mes:.2f}".replace(".", ",")
            multa_perc_txt = f"{multa_perc:.2f}".replace(".", ",")

            fixed1 = f"JUROS DE MORA DE {juros_perc_txt}% AO MES (R$ {money_br(daily_interest)} / DIA)".strip() if (juros_perc_mes or daily_interest) else ""
            fixed2 = f"MULTA DE {multa_perc_txt}% (R$ {money_br(multa)}) A PARTIR DE {start_dt_txt}".strip() if (multa_perc or multa) else ""

            wrapped = []
            if fixed1:
                wrapped.append(fixed1)
            if fixed2:
                wrapped.append(fixed2)
            if mensagem:
                wrapped.extend([w for w in wrap_text(mensagem, 70) if str(w or "").strip()])
            start_y = y_instr_top - 35
            step_y = 11.5
            for i, line in enumerate(wrapped[:6]):
                draw_text(left + 6, start_y - (i * step_y), line, size=8, bold=False)

        row_h = (y_instr_top - y_instr_bottom) / 4.0
        for i in range(1, 4):
            draw_line(x_right, y_instr_top - (i * row_h), right, y_instr_top - (i * row_h), width=0.5)

        cell1_top = y_instr_top
        cell1_bottom = y_instr_top - row_h
        cell2_top = cell1_bottom
        cell2_bottom = cell2_top - row_h
        cell3_top = cell2_bottom
        cell3_bottom = cell3_top - row_h
        cell4_top = cell3_bottom
        cell4_bottom = y_instr_bottom

        draw_text(x_right + 6, cell1_top - 12, "(-) Descontos / Abatimentos", size=6)
        draw_text(x_right + 6, cell2_top - 12, "(+) Mora / Multa", size=6)
        draw_text(x_right + 6, cell3_top - 12, "(+) Outros Acréscimos", size=6)
        draw_text(x_right + 6, cell4_top - 12, "(=) Valor Cobrado", size=6)
        draw_text(right - 6, cell4_bottom + 5, valor, size=9, bold=True, align="right")

        x_pag_text = left + 60
        draw_text(left + 6, y_pag_top - 10, "Pagador", size=6)
        pag_lines = wrap_text(f"{sacado_nome} - {sacado_doc}", 70)
        if pag_lines:
            draw_text(x_pag_text, y_pag_top - 10, pag_lines[0], size=8, bold=True)
        draw_text(x_pag_text, y_pag_bottom + 17, sacado_end, size=8)
        draw_text(x_pag_text, y_pag_bottom + 9, cidade_uf, size=8)
        draw_text(left + 6, y_pag_bottom + 2, "Pagador / Avalista", size=6)

        if codigo_barra and codigo_barra.isdigit() and len(codigo_barra) == 44:
            patterns = {
                "0": "00110",
                "1": "10001",
                "2": "01001",
                "3": "11000",
                "4": "00101",
                "5": "10100",
                "6": "01100",
                "7": "00011",
                "8": "10010",
                "9": "01010",
            }
            start_pattern = "0000"
            stop_pattern = "100"
            full_pattern = start_pattern
            for i in range(0, 44, 2):
                p1 = patterns[codigo_barra[i]]
                p2 = patterns[codigo_barra[i + 1]]
                for j in range(5):
                    full_pattern += p1[j] + p2[j]
            full_pattern += stop_pattern

            barcode_w = 280.0
            bx = left
            bh = (y_bar_top - y_bar_bottom) - 6
            by = y_bar_bottom + 6
            units_total = sum(3 if b == "1" else 1 for b in full_pattern)
            bw_fit = (barcode_w / units_total) if units_total else 0.75
            bw_narrow = max(0.55, min(bw_fit, 0.85))
            bw_wide = bw_narrow * 3
            for i, bit in enumerate(full_pattern):
                is_bar = (i % 2 == 0)
                bw = bw_wide if bit == "1" else bw_narrow
                if is_bar:
                    draw_rect(bx, by, bw, bh, fill=True)
                bx += bw
            draw_line(left + barcode_w, y_bar_bottom, left + barcode_w, y_bar_top, width=0.8)
            draw_text(left + barcode_w + 8, y_bar_top - 10, "Autenticação mecânica - Ficha de Compensação", size=5.5, bold=False)

        return build_text_pdf_bytes(ops, images=pdf_images)

    left = 40
    base_y = 350
    width = 515

    draw_line(left, base_y + 450, left + width, base_y + 450, width=1)
    if is_sicredi:
        logo_name = register_png(_resolve_sicredi_logo_path(boleto_data))
        if logo_name:
            logo_meta = next((m for m in pdf_images if m.get("name") == logo_name), None)
            lw = float((logo_meta or {}).get("width") or 0.0)
            lh = float((logo_meta or {}).get("height") or 0.0)
            if lw > 0 and lh > 0:
                scale = min(120.0 / lw, 24.0 / lh)
                draw_image(logo_name, left + 4, base_y + 424, lw * scale, lh * scale)
            draw_text(left + 132, base_y + 435, "RECIBO DO PAGADOR", size=10, bold=True)
        else:
            draw_text(left, base_y + 432, "Sicredi", size=18, bold=True, color=(0.36, 0.76, 0.20))
            draw_text(left + 132, base_y + 435, "RECIBO DO PAGADOR", size=10, bold=True)
    else:
        draw_text(left, base_y + 435, "RECIBO DO PAGADOR", size=10, bold=True)

    curr_y = base_y + 410
    draw_text(left, curr_y, f"Cedente: {cedente_nome}", size=8)
    draw_text(left + 350, curr_y, f"CNPJ/CPF: {cedente_doc}", size=8)

    curr_y -= 12
    draw_text(left, curr_y, f"Sacado: {sacado_nome}", size=8)
    draw_text(left + 350, curr_y, f"Vencimento: {vencto}", size=8)

    curr_y -= 12
    draw_text(left, curr_y, f"Nosso Número: {nosso_numero}", size=8)
    draw_text(left + 350, curr_y, f"Valor: {valor}", size=8)

    draw_line(left, base_y + 370, left + width, base_y + 370, width=0.5)

    y = base_y

    draw_line(left, y + 250, left + width, y + 250, width=1.2)
    draw_line(left, y + 230, left + width, y + 230, width=0.8)
    draw_line(left, y + 205, left + width, y + 205, width=0.8)
    draw_line(left, y + 185, left + width, y + 185, width=0.8)
    draw_line(left, y + 165, left + width, y + 165, width=0.8)
    draw_line(left, y + 145, left + width, y + 145, width=0.8)
    draw_line(left, y + 85, left + width, y + 85, width=0.8)

    draw_line(left + 230, y + 250, left + 230, y + 275, width=1.0)
    draw_line(left + 280, y + 250, left + 280, y + 275, width=1.0)
    draw_line(left + 380, y + 85, left + 380, y + 250, width=0.8)

    draw_line(left + 90, y + 185, left + 90, y + 205, width=0.8)
    draw_line(left + 190, y + 185, left + 190, y + 205, width=0.8)
    draw_line(left + 240, y + 185, left + 240, y + 205, width=0.8)
    draw_line(left + 280, y + 185, left + 280, y + 205, width=0.8)

    draw_line(left + 120, y + 165, left + 120, y + 185, width=0.8)
    draw_line(left + 180, y + 165, left + 180, y + 185, width=0.8)
    draw_line(left + 240, y + 165, left + 240, y + 185, width=0.8)

    for i in range(1, 6):
        draw_line(left + 380, y + 145 - (i * 12), left + width, y + 145 - (i * 12), width=0.5)

    bank_dvs = {"001": "9", "237": "2", "341": "7", "104": "0", "033": "7", "748": "X", "756": "0", "041": "8"}
    bank_dv = bank_dvs.get(bank_code, "9")

    b_size = 9
    if len(bank_name) > 35:
        b_size = 8
    if len(bank_name) > 45:
        b_size = 7
    custom_logo_path = str((boleto_data or {}).get("logo_path") or "").strip()
    custom_logo_name = register_png(custom_logo_path) if (custom_logo_path and os.path.isfile(custom_logo_path)) else ""
    if custom_logo_name and not is_sicredi:
        logo_meta_main = next((m for m in pdf_images if m.get("name") == custom_logo_name), None)
        lw_main = float((logo_meta_main or {}).get("width") or 0.0)
        lh_main = float((logo_meta_main or {}).get("height") or 0.0)
        if lw_main > 0 and lh_main > 0:
            scale_main = min(108.0 / lw_main, 16.0 / lh_main)
            draw_image(custom_logo_name, left + 5, y + 254, lw_main * scale_main, lh_main * scale_main)
        else:
            draw_text(left + 5, y + 260, bank_name, size=b_size, bold=True)
    elif is_sicredi:
        logo_name_main = register_png(_resolve_sicredi_logo_path(boleto_data))
        if logo_name_main:
            logo_meta_main = next((m for m in pdf_images if m.get("name") == logo_name_main), None)
            lw_main = float((logo_meta_main or {}).get("width") or 0.0)
            lh_main = float((logo_meta_main or {}).get("height") or 0.0)
            if lw_main > 0 and lh_main > 0:
                scale_main = min(108.0 / lw_main, 16.0 / lh_main)
                draw_image(logo_name_main, left + 5, y + 254, lw_main * scale_main, lh_main * scale_main)
        else:
            draw_text(left + 5, y + 258, "Sicredi", size=13, bold=True, color=(0.36, 0.76, 0.20))
    else:
        draw_text(left + 5, y + 260, bank_name, size=b_size, bold=True)
    draw_text(left + 235, y + 260, f"{bank_code}-{bank_dv}", size=12, bold=True)
    draw_text(left + 285, y + 260, linha_digitavel, size=8.2, bold=True)

    draw_text(left + 5, y + 242, "Local de Pagamento", size=6)
    if is_sicredi:
        draw_text(left + 5, y + 233, "PAGÁVEL PREFERENCIALMENTE NAS COOPERATIVAS DE CRÉDITO DO SICREDI", size=8, bold=True)
    else:
        draw_text(left + 5, y + 233, "PAGÁVEL EM QUALQUER BANCO ATÉ O VENCIMENTO", size=8)
    draw_text(left + 385, y + 242, "Vencimento", size=6)
    draw_text(left + 385, y + 233, vencto, size=10, bold=True)

    draw_text(left + 5, y + 222, "Beneficiário", size=6)
    beneficiario_lines = wrap_text(f"{cedente_nome} - CNPJ: {cedente_doc}", 75)
    for i, bline in enumerate(beneficiario_lines[:2]):
        draw_text(left + 5, y + 213 - (i * 9), bline, size=8, bold=True)
    draw_text(left + 385, y + 222, "Agência / Código Beneficiário", size=6)
    draw_text(left + 385, y + 213, agencia_conta, size=9)

    draw_text(left + 5, y + 197, "Data do Documento", size=6)
    draw_text(left + 5, y + 188, date.today().strftime("%d/%m/%Y"), size=8)
    draw_text(left + 95, y + 197, "Nº do Documento", size=6)
    draw_text(left + 95, y + 188, documento, size=8)
    draw_text(left + 195, y + 197, "Espécie Doc.", size=6)
    draw_text(left + 195, y + 188, "DM", size=8)
    draw_text(left + 245, y + 197, "Aceite", size=6)
    draw_text(left + 245, y + 188, "NÃO" if is_sicredi else "N", size=8)
    draw_text(left + 285, y + 197, "Data Proc.", size=6)
    draw_text(left + 285, y + 188, date.today().strftime("%d/%m/%Y"), size=8)
    draw_text(left + 385, y + 197, "Nosso Número", size=6)
    draw_text(left + 385, y + 188, nosso_numero, size=9)

    draw_text(left + 5, y + 177, "Uso do Banco", size=6)
    draw_text(left + 125, y + 177, "Carteira", size=6)
    draw_text(left + 125, y + 168, carteira, size=8)
    draw_text(left + 185, y + 177, "Espécie", size=6)
    draw_text(left + 185, y + 168, "REAL" if is_sicredi else "R$", size=8, bold=is_sicredi)
    draw_text(left + 245, y + 177, "Quantidade", size=6)
    draw_text(left + 385, y + 177, "(=) Valor do Documento", size=6)
    draw_text(left + 385, y + 168, valor, size=10, bold=True)

    draw_text(left + 5, y + 157, "Instruções / Observações", size=6)
    max_instr_chars = 45 if pix_code else 75
    if mensagem:
        wrapped_lines = wrap_text(mensagem, max_instr_chars)
        for i, mline in enumerate(wrapped_lines[:6]):
            draw_text(left + 5, y + 130 - (i * 9), mline, size=8)

    draw_text(left + 385, y + 137, "(-) Descontos / Abatimentos", size=6)
    draw_text(left + 385, y + 125, "(+) Mora / Multa", size=6)
    draw_text(left + 385, y + 113, "(+) Outros Acréscimos", size=6)
    draw_text(left + 385, y + 101, "(=) Valor Cobrado", size=6)
    draw_text(left + 385, y + 88, valor, size=9, bold=True)

    if pix_code:
        qr_size = 50
        qr_x = left + 280
        qr_y = y + 95
        matrix = get_qr_matrix(pix_code)
        m_size = len(matrix)
        mod_size = qr_size / m_size
        for row_idx, row_data in enumerate(matrix):
            for col_idx, val in enumerate(row_data):
                if val:
                    draw_rect(
                        qr_x + (col_idx * mod_size),
                        qr_y + qr_size - ((row_idx + 1) * mod_size),
                        mod_size,
                        mod_size,
                        fill=True,
                    )
        draw_rect(qr_x, qr_y, qr_size, qr_size)
        draw_text(qr_x, qr_y - 8, "PAGUE COM PIX", size=6, bold=True)
        draw_text(left + 5, y + 80, "PIX Copia e Cola:", size=5, bold=True)
        draw_text(left + 70, y + 80, pix_code, size=4.5)

    draw_rect(left, y + 15, width, 55)
    draw_text(left + 5, y + 62, "Pagador", size=6)

    sacado_lines = wrap_text(f"{sacado_nome} - {sacado_doc}", 75)
    for i, sline in enumerate(sacado_lines[:2]):
        draw_text(left + 45, y + 62 - (i * 9), sline, size=8, bold=True)
    draw_text(left + 45, y + 43, sacado_end, size=8)
    draw_text(left + 45, y + 34, cidade_uf, size=8)
    draw_text(left + 5, y + 17, "Sacador / Avalista", size=6)

    if codigo_barra and codigo_barra.isdigit() and len(codigo_barra) == 44:
        patterns = {
            "0": "00110",
            "1": "10001",
            "2": "01001",
            "3": "11000",
            "4": "00101",
            "5": "10100",
            "6": "01100",
            "7": "00011",
            "8": "10010",
            "9": "01010",
        }

        start_pattern = "0000"
        stop_pattern = "100"

        full_pattern = start_pattern
        for i in range(0, 44, 2):
            p1 = patterns[codigo_barra[i]]
            p2 = patterns[codigo_barra[i + 1]]
            for j in range(5):
                full_pattern += p1[j] + p2[j]
        full_pattern += stop_pattern

        bx = left
        bh = 50
        footer_text = "Autenticação mecânica - Ficha de Compensação"
        footer_size = 5.5
        footer_step = footer_size + 2.5
        footer_top = (y + 11)
        by = (footer_top - 2) - bh
        units_total = sum(3 if b == "1" else 1 for b in full_pattern)
        bw_fit = (width / units_total) if units_total else 0.75
        bw_narrow = max(0.75, min(bw_fit, 1.15))
        bw_wide = bw_narrow * 3

        for i, bit in enumerate(full_pattern):
            is_bar = (i % 2 == 0)
            bw = bw_wide if bit == "1" else bw_narrow
            if is_bar:
                draw_rect(bx, by, bw, bh, fill=True)
            bx += bw
        draw_text(left, by - footer_step, footer_text, size=footer_size, bold=False)

    return build_text_pdf_bytes(ops, images=pdf_images)

class SimpleDialog(tk.Toplevel):
    def __init__(self, master, title: str, size: str):
        super().__init__(master)
        self.title(title)
        self.geometry(size)
        self.resizable(True, True)
        self.transient(master)
        self.grab_set()
        self.lift()
        self.focus_force()
        self._center()
    def _center(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - self.winfo_width()) // 2
        y = (sh - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 50)}+{max(y, 50)}")
class ConfigWindow(SimpleDialog):
    def __init__(self, master, config_data: Dict[str, Any], on_save):
        self.config_data = deepcopy(config_data)
        self.on_save = on_save
        super().__init__(master, "Configuração local", "820x760")
        self.minsize(820, 760)
        self._build()

    def _build(self):
        outer = ttk.Frame(self, padding=14)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        header = ttk.Frame(outer)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(
            header,
            text="Configuração do sistema",
            font=("Segoe UI", 11, "bold"),
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Label(
            header,
            text="Informe os dados de conexão com o banco e o SMTP para envio de e-mails.",
            wraplength=760,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(0, 12))

        content = ttk.Frame(outer)
        content.grid(row=1, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)

        self.entries = {}

        db_box = ttk.LabelFrame(content, text="Banco de dados", padding=12)
        db_box.grid(row=0, column=0, sticky="ew")
        db_box.columnconfigure(0, weight=1)

        form = ttk.Frame(db_box)
        form.grid(row=0, column=0, sticky="ew")
        form.columnconfigure(1, weight=1)

        conn = self.config_data["connection"]
        fields = [
            ("Host", "host"),
            ("Porta", "port"),
            ("Banco", "dbname"),
            ("Usuário", "user"),
            ("Senha", "password"),
            ("Encoding", "client_encoding"),
        ]
        for i, (label, key) in enumerate(fields):
            ttk.Label(form, text=label).grid(row=i, column=0, sticky="w", padx=(0, 8), pady=5)
            entry = ttk.Entry(form, show="*" if key == "password" else None)
            entry.grid(row=i, column=1, sticky="ew", pady=5)
            entry.insert(0, str(conn.get(key, "")))
            self.entries[key] = entry

        smtp_box = ttk.LabelFrame(content, text="SMTP", padding=12)
        smtp_box.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        smtp_box.columnconfigure(0, weight=1)

        smtp_form = ttk.Frame(smtp_box)
        smtp_form.grid(row=0, column=0, sticky="ew")
        smtp_form.columnconfigure(1, weight=1)

        smtp_cfg = self.config_data.get("smtp", {})
        smtp_fields = [
            ("E-mail remetente", "smtp_email"),
            ("Nome do remetente", "smtp_sender_name"),
            ("Servidor SMTP", "smtp_host"),
            ("Senha", "smtp_password"),
            ("Porta", "smtp_port"),
            ("Espaçamento (seg)", "smtp_delay_seconds"),
        ]
        for i, (label, key) in enumerate(smtp_fields):
            ttk.Label(smtp_form, text=label).grid(row=i, column=0, sticky="w", padx=(0, 8), pady=5)
            entry = ttk.Entry(smtp_form, show="*" if key == "smtp_password" else None)
            entry.grid(row=i, column=1, sticky="ew", pady=5)
            entry.insert(0, str(smtp_cfg.get(key.replace("smtp_", ""), smtp_cfg.get(key, ""))))
            self.entries[key] = entry

        buttons = ttk.Frame(self, padding=(14, 0, 14, 14))
        buttons.pack(side="bottom", fill="x")
        ttk.Button(buttons, text="Testar conexão", command=self._test).pack(side="left")
        ttk.Button(buttons, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(buttons, text="Salvar conexão", command=self._save).pack(side="right", padx=(0, 8))

    def _collect(self):
        cfg = deepcopy(self.config_data)
        try:
            port = int(self.entries["port"].get().strip())
        except Exception:
            raise AppError("A porta deve ser numérica.")
        cfg["connection"]["host"] = self.entries["host"].get().strip()
        cfg["connection"]["port"] = port
        cfg["connection"]["dbname"] = self.entries["dbname"].get().strip()
        cfg["connection"]["user"] = self.entries["user"].get().strip()
        cfg["connection"]["password"] = self.entries["password"].get().strip()
        cfg["connection"]["client_encoding"] = self.entries["client_encoding"].get().strip()

        smtp_port_raw = self.entries["smtp_port"].get().strip()
        if smtp_port_raw:
            try:
                smtp_port = int(smtp_port_raw)
            except Exception:
                raise AppError("A porta SMTP deve ser numérica.")
        else:
            smtp_port = 587

        smtp_delay_raw = self.entries.get("smtp_delay_seconds").get().strip() if self.entries.get("smtp_delay_seconds") else ""
        if smtp_delay_raw:
            try:
                smtp_delay = int(smtp_delay_raw)
            except Exception:
                raise AppError("O espaçamento SMTP deve ser numérico.")
        else:
            smtp_delay = 5
        smtp_delay = max(0, min(300, smtp_delay))

        cfg["smtp"] = {
            "email": self.entries["smtp_email"].get().strip(),
            "sender_name": self.entries.get("smtp_sender_name").get().strip() if self.entries.get("smtp_sender_name") else "",
            "host": self.entries["smtp_host"].get().strip(),
            "password": self.entries["smtp_password"].get().strip(),
            "port": smtp_port,
            "delay_seconds": smtp_delay,
        }

        if not cfg["connection"]["host"]:
            raise AppError("Informe o host.")
        if not cfg["connection"]["dbname"]:
            raise AppError("Informe o nome do banco.")
        if not cfg["connection"]["user"]:
            raise AppError("Informe o usuário.")
        return cfg

    def _test(self):
        try:
            cfg = self._collect()
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Falha ao testar conexão:\n\n{e}", parent=self)
            return

        def _work():
            Database(cfg).test_connection()
            return True

        def _ok(_):
            messagebox.showinfo(APP_TITLE, "Conexão realizada com sucesso.", parent=self)

        def _err(e: Exception):
            messagebox.showerror(APP_TITLE, f"Falha ao testar conexão:\n\n{e}", parent=self)

        run_with_busy(self, "Testando conexão...", _work, _ok, _err)

    def _save(self):
        try:
            cfg = self._collect()
            ConfigManager.save(cfg)
            self.on_save(cfg)
            messagebox.showinfo(APP_TITLE, f"Configuração salva em:\n{CONFIG_PATH}", parent=self)
            self.destroy()
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Erro ao salvar configuração:\n\n{e}", parent=self)

class EmailComposeWindow(SimpleDialog):

    def __init__(self, master, config_data: Dict[str, Any], current_user: str, invoice_row: InvoiceRow, customer_email: str):
        self.config_data = deepcopy(config_data)
        self.current_user = current_user
        self.invoice_row = invoice_row
        self.customer_email = customer_email or ""
        self.boleto_data: Dict[str, Any] = {"exists": False, "email_note": "Observação: consultando documentos."}
        self.nfe_data: Dict[str, Any] = {"exists": False, "attachments": []}
        self.purchase_map: Dict[Any, Dict[str, Any]] = {}
        self.purchase_error: str = ""
        self.attachment_bytes = None
        self.attachment_name = ""
        self.boleto_status_text = "Consultando documentos..."
        self.boleto_status_var = tk.StringVar(value=f"Status do boleto: {self.boleto_status_text}")
        self.send_pix_qrcode_var = tk.BooleanVar(value=False)
        super().__init__(master, "Enviar fatura por e-mail", "760x680")
        self._build()
        self._load_email_data()

    def _load_email_data(self):
        def _work():
            db = Database(self.config_data)
            payload = db.get_boleto_email_payload(self.invoice_row.invoice_id)
            try:
                inv_key = getattr(self.invoice_row, "movto_id", None) if getattr(self.invoice_row, "movto_id", None) not in (None, "", 0, "0") else self.invoice_row.invoice_id
                purchase_map = db.get_purchase_info_bulk([inv_key])
                purchase_error = ""
            except Exception as e:
                purchase_map = {}
                purchase_error = str(e) or "Falha ao consultar itens da venda."
                if len(purchase_error) > 220:
                    purchase_error = purchase_error[:220] + "..."
            try:
                nfe_map = db.get_nfe_attachments_bulk([inv_key], only_invoice_mlid=True, max_nfes_per_invoice=1)
                current = (nfe_map or {}).get(inv_key) or (nfe_map or {}).get(str(inv_key)) or {}
                cur_atts = list((current.get("attachments") or []))
                has_any = bool([a for a in cur_atts if a.get("data") and a.get("filename")])
                if not has_any:
                    nfe_map2 = db.get_nfe_attachments_bulk([inv_key], only_invoice_mlid=False, max_nfes_per_invoice=1)
                    candidate = (nfe_map2 or {}).get(inv_key) or (nfe_map2 or {}).get(str(inv_key)) or {}
                    cand_atts = list((candidate.get("attachments") or []))
                    if bool([a for a in cand_atts if a.get("data") and a.get("filename")]):
                        nfe_map = nfe_map2
            except Exception:
                nfe_map = {}
            return payload, purchase_map, purchase_error, nfe_map, inv_key

        def _ok(result):
            payload, purchase_map, purchase_error, nfe_map, inv_key = result
            self.purchase_map = purchase_map or {}
            self.purchase_error = "" if self.purchase_map else str(purchase_error or "").strip()
            self.nfe_data = (nfe_map or {}).get(inv_key) or (nfe_map or {}).get(str(inv_key)) or {"exists": False, "attachments": []}
            nfe_atts = list((self.nfe_data or {}).get("attachments") or [])
            has_pdf = bool([a for a in nfe_atts if str(a.get("filename") or "").lower().endswith(".pdf") and a.get("data")])
            for a in nfe_atts:
                if has_pdf:
                    break
                name = str(a.get("filename") or "").lower()
                data = a.get("data")
                if not data or not name.endswith(".xml"):
                    continue
                extra_inf = ""
                try:
                    extra_inf = Database(self.config_data).get_placa_km_text_bulk([inv_key]).get(inv_key) or ""
                except Exception:
                    extra_inf = ""
                pdf_bytes, pdf_name = danfe_pdf_from_nfe_xml(
                    data,
                    fallback_suffix=str(inv_key),
                    extra_inf_cpl_text=extra_inf,
                )
                if pdf_bytes and pdf_name:
                    nfe_atts.append({"data": pdf_bytes, "filename": pdf_name, "mime_type": "application/pdf"})
                    has_pdf = True
            self.nfe_data["attachments"] = nfe_atts
            self._apply_boleto_payload(payload)
            self._maybe_refresh_default_body()
            try:
                self.send_btn.state(["!disabled"])
            except Exception:
                pass

        def _err(e: Exception):
            payload = {
                "exists": False,
                "email_note": f"Observação: não foi possível consultar os dados do boleto neste momento. {e}",
            }
            self._apply_boleto_payload(payload)
            self._maybe_refresh_default_body()
            try:
                self.send_btn.state(["!disabled"])
            except Exception:
                pass

        try:
            self.send_btn.state(["disabled"])
        except Exception:
            pass
        run_with_busy(self, "Consultando documentos...", _work, _ok, _err)

    def _apply_boleto_payload(self, payload: Dict[str, Any]):
        self.boleto_data = payload or {}
        if not self.boleto_data.get("exists"):
            self.boleto_status_text = "Boleto ainda não gerado"
            self.attachment_bytes = None
            self.attachment_name = ""
            self.boleto_status_var.set(f"Status do boleto: {self.boleto_status_text}")
            return

        filename = self.boleto_data.get("filename") or f"boleto_{self.invoice_row.invoice_id}.pdf"
        self.attachment_bytes = None
        self.attachment_name = filename
        self.boleto_status_text = "Boleto localizado (será gerado para o envio)"
        self.boleto_status_var.set(f"Status do boleto: {self.boleto_status_text}")

    def _maybe_refresh_default_body(self):
        try:
            current = (self.body_text.get("1.0", "end") or "").strip()
        except Exception:
            return
        new_body = self._default_body()
        if current == (self._default_body_snapshot or "").strip():
            self.body_text.delete("1.0", "end")
            self.body_text.insert("1.0", new_body)
            self._default_body_snapshot = new_body

    def _default_subject(self) -> str:
        due_text = self.invoice_row.due_date_display()
        return f"Fatura a receber - {self.invoice_row.customer_name} - vencimento {due_text}"

    def _default_body(self) -> str:
        company = str(self.invoice_row.company or "").strip()
        account_display = (f"{self.invoice_row.account_code or ''} - {self.invoice_row.account_name or ''}").strip(" -")
        note = ""
        
        try:
            sig = Database(self.config_data).get_sale_signature_pdf(getattr(self.invoice_row, "invoice_id", None))
        except Exception:
            sig = {}
        has_signature = (sig or {}).get("exists") or bool((sig or {}).get("attachments"))
        has_nfe_xml = bool([a for a in ((self.nfe_data or {}).get("attachments") or []) if str(a.get("filename") or "").lower().endswith(".xml") and a.get("data")])
        has_danfe = bool([a for a in ((self.nfe_data or {}).get("attachments") or []) if str(a.get("filename") or "").lower().endswith(".pdf") and a.get("data")])
        has_boleto = bool(self.boleto_data.get("exists"))
        boleto_reason = "" if has_boleto else "ainda não gerado"
        xml_reason = "" if has_nfe_xml else "não encontrado"
        danfe_reason = "não gerada" if (has_nfe_xml and not has_danfe) else ""
        assinatura_reason = "" if has_signature else "não encontrado"
        note = build_attachments_note_text(
            has_boleto=has_boleto,
            has_fatura_pdf=False,
            has_fatura_txt=True,
            has_xml=has_nfe_xml,
            has_danfe=has_danfe,
            has_assinatura=has_signature,
            boleto_reason=boleto_reason,
            xml_reason=xml_reason,
            danfe_reason=danfe_reason,
            assinatura_reason=assinatura_reason,
        )
                
        invoice_id = str(getattr(self.invoice_row, "invoice_id", "") or "").strip()
        doc_str = invoice_id if invoice_id else "N/A"
        
        purchase_text, _ = build_purchase_info_blocks(self.invoice_row, self.purchase_map)
        purchase_err = str(getattr(self, "purchase_error", "") or "").strip()
        if purchase_err and not purchase_text:
            purchase_text = f"\nInformações da compra:\n- Não foi possível consultar itens da venda: {purchase_err}\n\n"
        return (
            f"Prezado(a),\n\n"
            f"Segue a fatura para conferência e programação do pagamento.\n\n"
            f"Cliente: {self.invoice_row.customer_name}\n"
            f"Conta: {account_display}\n"
            f"Total a pagar: {self.invoice_row.open_balance_display()}\n\n"
            f"- Documento / Fatura: {doc_str} | Emissão: {self.invoice_row.issue_date_display()} | Venc.: {self.invoice_row.due_date_display()} | Saldo: {self.invoice_row.open_balance_display()}\n\n"
            f"{purchase_text}"
            f"{note}\n\n"
            f"Em caso de dúvidas, ficamos à disposição.\n\n"
            f"Atenciosamente,\n"
            f"{company}"
        )

    def _default_body_html(self) -> str:
        company = str(self.invoice_row.company or "").strip()
        account_display = (f"{self.invoice_row.account_code or ''} - {self.invoice_row.account_name or ''}").strip(" -")
        note = ""
        
        try:
            sig = Database(self.config_data).get_sale_signature_pdf(getattr(self.invoice_row, "invoice_id", None))
        except Exception:
            sig = {}
        has_signature = (sig or {}).get("exists") or bool((sig or {}).get("attachments"))
        has_nfe_xml = bool([a for a in ((self.nfe_data or {}).get("attachments") or []) if str(a.get("filename") or "").lower().endswith(".xml") and a.get("data")])
        has_danfe = bool([a for a in ((self.nfe_data or {}).get("attachments") or []) if str(a.get("filename") or "").lower().endswith(".pdf") and a.get("data")])
        
        note_bg = "#fff7ed"
        note_border = "#fed7aa"
        note_color = "#9a3412"
        has_boleto = bool(self.boleto_data.get("exists"))
        boleto_reason = "" if has_boleto else "ainda não gerado"
        xml_reason = "" if has_nfe_xml else "não encontrado"
        danfe_reason = "não gerada" if (has_nfe_xml and not has_danfe) else ""
        assinatura_reason = "" if has_signature else "não encontrado"
        html_note = build_attachments_note_html(
            has_boleto=has_boleto,
            has_fatura_pdf=False,
            has_fatura_txt=True,
            has_xml=has_nfe_xml,
            has_danfe=has_danfe,
            has_assinatura=has_signature,
            boleto_reason=boleto_reason,
            xml_reason=xml_reason,
            danfe_reason=danfe_reason,
            assinatura_reason=assinatura_reason,
        )
        _, purchase_html = build_purchase_info_blocks(self.invoice_row, self.purchase_map)
        purchase_err = str(getattr(self, "purchase_error", "") or "").strip()
        if purchase_err and not purchase_html:
            purchase_html = f"<h3>Informações da compra</h3><div class='note'><b>Observação:</b> Não foi possível consultar itens da venda: {html.escape(purchase_err)}</div>"

        invoice_id = str(getattr(self.invoice_row, "invoice_id", "") or "").strip()
        doc_str = invoice_id if invoice_id else "N/A"

        logo_html = f"<div class='logo'><img src='cid:{EMAIL_LOGO_CID}' alt='Logo' width='240' height='60' style='width:240px;height:60px;display:block;border:0;outline:none;text-decoration:none;'></div>"
        return f"""<html>
<head>
<style>
    body {{ font-family: Arial, sans-serif; font-size: 14px; color: #333; }}
    .card {{ max-width: 780px; border: 1px solid #e5e7eb; border-radius: 10px; padding: 16px; background: #ffffff; }}
    .logo {{ margin: 0 0 12px 0; }}
    .logo img {{ max-height: 60px; max-width: 240px; display: block; }}
    .title {{ font-size: 18px; font-weight: 700; color: #2563eb; margin: 0 0 10px 0; }}
    .muted {{ color: #6b7280; margin: 0 0 14px 0; }}
    table {{ border-collapse: collapse; width: 100%; margin-top: 10px; margin-bottom: 12px; }}
    th {{ background-color: #f8fafc; text-align: left; padding: 10px; border: 1px solid #e5e7eb; }}
    td {{ padding: 10px; border: 1px solid #e5e7eb; }}
    .note {{ margin: 12px 0 12px 0; padding: 10px 12px; background: {note_bg}; border: 1px solid {note_border}; border-radius: 8px; color: {note_color}; }}
</style>
</head>
<body>
  <div class="card">
    {logo_html}
    <p class="title">Fatura para programação de pagamento</p>
    <p class="muted">Cliente: <b>{html.escape(self.invoice_row.customer_name)}</b> &nbsp;|&nbsp; Conta: <b>{html.escape(account_display)}</b></p>
    <p><b>Total a pagar:</b> {html.escape(self.invoice_row.open_balance_display())}</p>
    <table>
      <thead>
        <tr>
          <th>Documento / Fatura</th>
          <th>Emissão</th>
          <th>Vencimento</th>
          <th>Valor da fatura</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td>{html.escape(doc_str)}</td>
          <td>{html.escape(self.invoice_row.issue_date_display())}</td>
          <td>{html.escape(self.invoice_row.due_date_display())}</td>
          <td>{html.escape(self.invoice_row.open_balance_display())}</td>
        </tr>
      </tbody>
    </table>
    {purchase_html}
    {html_note}
    <p>Em caso de dúvidas, ficamos à disposição.</p>
    <p>Atenciosamente,<br>{html.escape(company)}</p>
  </div>
</body>
</html>"""

    def _build(self):
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Envio de fatura por e-mail", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 8))
        info = ttk.Frame(frm)
        info.pack(fill="x", pady=(0, 10))
        ttk.Label(info, text=f"Cliente: {self.invoice_row.customer_name}").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Label(info, text=f"Vencimento: {self.invoice_row.due_date_display()}").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Label(info, text=f"Saldo em aberto: {self.invoice_row.open_balance_display()}").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Label(info, textvariable=self.boleto_status_var).grid(row=3, column=0, sticky="w", pady=2)
        ttk.Checkbutton(info, text="Incluir QRCode PIX no boleto (PDF)", variable=self.send_pix_qrcode_var).grid(row=4, column=0, sticky="w", pady=(6, 2))

        form = ttk.Frame(frm)
        form.pack(fill="x", pady=(0, 10))
        ttk.Label(form, text="Para").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.to_var = tk.StringVar(value=self.customer_email)
        ttk.Entry(form, textvariable=self.to_var, width=72).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(form, text="Assunto").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.subject_var = tk.StringVar(value=self._default_subject())
        ttk.Entry(form, textvariable=self.subject_var, width=72).grid(row=1, column=1, sticky="ew", pady=4)
        form.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Mensagem").pack(anchor="w")
        self.body_text = tk.Text(frm, wrap="word", height=18)
        self.body_text.pack(fill="both", expand=True, pady=(4, 0))
        default_body = self._default_body()
        self._default_body_snapshot = default_body
        self.body_text.insert("1.0", default_body)

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(12, 0))
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side="right")
        self.send_btn = ttk.Button(btns, text="Enviar", command=self._send_email)
        self.send_btn.pack(side="right", padx=(0, 8))

    def _send_email(self):
        to_email = self.to_var.get().strip()
        subject = self.subject_var.get().strip()
        body = self.body_text.get("1.0", "end").strip()
        smtp_cfg = self.config_data.get("smtp", {})
        smtp_email = str(smtp_cfg.get("email", "")).strip()
        smtp_host = str(smtp_cfg.get("host", "")).strip()
        smtp_password = str(smtp_cfg.get("password", "")).strip()
        smtp_port = int(smtp_cfg.get("port", 587) or 587)

        if not to_email:
            messagebox.showwarning(APP_TITLE, "Informe o e-mail do destinatário.", parent=self)
            return
        if not subject:
            messagebox.showwarning(APP_TITLE, "Informe o assunto do e-mail.", parent=self)
            return
        if not body:
            messagebox.showwarning(APP_TITLE, "Informe a mensagem do e-mail.", parent=self)
            return
        if not smtp_email or not smtp_host or not smtp_password or not smtp_port:
            messagebox.showwarning(APP_TITLE, "Configure o SMTP na tela de Configuração antes de enviar e-mails.", parent=self)
            return

        msg = EmailMessage()
        msg["From"] = format_smtp_from(smtp_cfg) or smtp_email
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.set_content(body)

        if body.strip() == str(getattr(self, "_default_body_snapshot", "") or "").strip():
            msg.add_alternative(self._default_body_html(), subtype="html")
        else:
            chunks = re.split(r"\n\s*\n", body.strip())
            html_parts = []
            for c in chunks:
                c = c.strip()
                if not c:
                    continue
                html_parts.append(f"<p>{html.escape(c).replace(chr(10), '<br>')}</p>")
            logo_html = f"<div style='margin:0 0 12px 0'><img src='cid:{EMAIL_LOGO_CID}' alt='Logo' width='240' height='60' style='width:240px;height:60px;display:block;border:0;outline:none;text-decoration:none;'></div>"
            html_body = f"""<html>
<head>
<style>
    body {{ font-family: Arial, sans-serif; font-size: 14px; color: #333; }}
</style>
</head>
<body>
{logo_html}
{''.join(html_parts)}
</body>
</html>"""
            msg.add_alternative(html_body, subtype="html")
        try:
            attach_email_logo(msg)
        except Exception:
            pass

        attachment_names: List[str] = []

        include_pix_qrcode = bool(self.send_pix_qrcode_var.get())
        pdf_bytes = None
        if (self.boleto_data or {}).get("exists"):
            if include_pix_qrcode:
                attachment_data = (self.boleto_data or {}).get("attachment_data")
                if attachment_data:
                    try:
                        pdf_bytes = bytes(attachment_data)
                    except Exception:
                        pdf_bytes = None
            if not pdf_bytes:
                try:
                    pdf_bytes = build_boleto_pdf_bytes(self.boleto_data or {}, self.invoice_row, include_pix_qrcode=include_pix_qrcode)
                except Exception:
                    pdf_bytes = None
                    if include_pix_qrcode:
                        attachment_data = (self.boleto_data or {}).get("attachment_data")
                        if attachment_data:
                            try:
                                pdf_bytes = bytes(attachment_data)
                            except Exception:
                                pdf_bytes = None
        else:
            pdf_bytes = self.attachment_bytes

        if pdf_bytes:
            boleto_name = self.attachment_name or f"boleto_{self.invoice_row.invoice_id}.pdf"
            msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=boleto_name)
            attachment_names.append(str(boleto_name))

        try:
            txt_bytes, txt_name = build_faturas_detalhamento_txt_bytes([self.invoice_row], purchase_info_map=getattr(self, "purchase_map", None))
            if txt_bytes and txt_name:
                maintype, subtype = _mime_parts_from_filename(txt_name)
                msg.add_attachment(txt_bytes, maintype=maintype, subtype=subtype, filename=txt_name)
                attachment_names.append(str(txt_name))
        except Exception:
            pass

        try:
            sig = Database(self.config_data).get_sale_signature_pdf(getattr(self.invoice_row, "invoice_id", None))
        except Exception:
            sig = {}
        signature_files: List[Tuple[bytes, str]] = []
        for a in ((sig or {}).get("attachments") or []):
            data = a.get("data")
            name = a.get("filename")
            if data and name:
                signature_files.append((data, name))
        sig_bytes = (sig or {}).get("attachment_data")
        if not signature_files and (sig or {}).get("exists") and sig_bytes:
            name = (sig or {}).get("filename") or f"assinatura_{getattr(self.invoice_row, 'invoice_id', None) or ''}"
            signature_files.append((sig_bytes, name))
        if signature_files:
            inv_part = re.sub(r"[^0-9A-Za-z_-]+", "_", str(getattr(self.invoice_row, "invoice_id", "") or "").strip())[:40]
            zip_name = f"assinaturas_{inv_part}.zip" if inv_part else "assinaturas.zip"
            zip_bytes, _ = zip_named_files(signature_files, zip_filename=zip_name)
            msg.add_attachment(zip_bytes, maintype="application", subtype="zip", filename=zip_name)
            attachment_names.append(str(zip_name))

        nfe = self.nfe_data or {}
        for a in (nfe.get("attachments") or []):
            data = a.get("data")
            name = a.get("filename")
            if not data or not name:
                continue
            maintype, subtype = _mime_parts_from_filename(name)
            msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=name)
            attachment_names.append(str(name))

        try:
            if smtp_port == 465:
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL(smtp_host, smtp_port, context=context, timeout=20) as server:
                    server.login(smtp_email, smtp_password)
                    server.send_message(msg)
            else:
                with smtplib.SMTP(smtp_host, smtp_port, timeout=20) as server:
                    server.ehlo()
                    try:
                        server.starttls(context=ssl.create_default_context())
                        server.ehlo()
                    except Exception:
                        pass
                    server.login(smtp_email, smtp_password)
                    server.send_message(msg)
            AuditLogger.write(
                self.current_user,
                "enviar_email_fatura",
                f"cliente={self.invoice_row.customer_name};para={to_email};invoice={self.invoice_row.invoice_id};anexo_pdf={'sim' if pdf_bytes else 'nao'};pix_incluido_no_boleto={'sim' if include_pix_qrcode else 'nao'}"
            )
            try:
                from app_core.documents_history import DocumentsHistory

                DocumentsHistory().add_event(
                    kind="manual_send",
                    source="faturas_receber",
                    level="info",
                    title="E-mail enviado (manual)",
                    message=f"Cliente: {self.invoice_row.customer_name} | Para: {to_email} | Fatura: {self.invoice_row.invoice_id} | Anexos: {', '.join([str(x) for x in (attachment_names or [])])}",
                )
            except Exception:
                pass
            messagebox.showinfo(APP_TITLE, "E-mail enviado com sucesso.", parent=self)
            self.destroy()
        except Exception as e:
            try:
                from app_core.documents_history import DocumentsHistory

                DocumentsHistory().add_event(
                    kind="manual_send",
                    source="faturas_receber",
                    level="error",
                    title="Falha no envio (manual)",
                    message=f"Cliente: {self.invoice_row.customer_name} | Para: {to_email} | Fatura: {self.invoice_row.invoice_id} | Anexos: {', '.join([str(x) for x in (attachment_names or [])])} | Erro: {e}",
                )
            except Exception:
                pass
            messagebox.showerror(APP_TITLE, f"Falha ao enviar e-mail:\n\n{e}", parent=self)

class CreateUserWindow(SimpleDialog):
    def __init__(self, master, config_data: Dict[str, Any], current_user: str, on_save):
        self.config_data = deepcopy(config_data)
        self.current_user = current_user
        self.on_save = on_save
        super().__init__(master, "Cadastrar usuário", "500x320")
        self._build()
    def _build(self):
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Cadastrar novo usuário", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 6))
        ttk.Label(frm, text="Cadastre um novo usuário local para acesso ao sistema.", wraplength=450, justify="left").pack(anchor="w", pady=(0, 14))
        form = ttk.Frame(frm)
        form.pack(fill="x")
        ttk.Label(form, text="Usuário").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        self.user_entry = ttk.Entry(form, width=34)
        self.user_entry.grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Label(form, text="Senha").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        self.pass_entry = ttk.Entry(form, width=34, show="*")
        self.pass_entry.grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Label(form, text="Confirmar senha").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=5)
        self.confirm_entry = ttk.Entry(form, width=34, show="*")
        self.confirm_entry.grid(row=2, column=1, sticky="ew", pady=5)
        form.columnconfigure(1, weight=1)
        ttk.Frame(frm).pack(fill="both", expand=True)
        btns = ttk.Frame(frm)
        btns.pack(fill="x", side="bottom", pady=(20, 8))
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Cadastrar", command=self._create).pack(side="right", padx=(0, 8))
    def _create(self):
        try:
            username = self.user_entry.get().strip()
            password = self.pass_entry.get()
            confirm = self.confirm_entry.get()
            if not username:
                raise AppError("Informe o usuário.")
            if not password:
                raise AppError("Informe a senha.")
            if password != confirm:
                raise AppError("A confirmação da senha não confere.")
            UserManager.add_user(self.config_data, username, password)
            ConfigManager.save(self.config_data)
            self.on_save(self.config_data)
            AuditLogger.write(self.current_user, "cadastrar_usuario", f"alvo={username}")
            messagebox.showinfo(APP_TITLE, "Usuário cadastrado com sucesso.", parent=self)
            self.destroy()
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e), parent=self)
class ChangeOwnPasswordWindow(SimpleDialog):
    def __init__(self, master, config_data: Dict[str, Any], current_user: str, on_save):
        self.config_data = deepcopy(config_data)
        self.current_user = current_user
        self.on_save = on_save
        super().__init__(master, "Alterar senha", "500x300")
        self._build()
    def _build(self):
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Alterar senha", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 6))
        ttk.Label(frm, text=f"Usuário: {self.current_user}", wraplength=450, justify="left").pack(anchor="w", pady=(0, 14))
        form = ttk.Frame(frm)
        form.pack(fill="x")
        ttk.Label(form, text="Nova senha").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        self.pass_entry = ttk.Entry(form, width=34, show="*")
        self.pass_entry.grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Label(form, text="Confirmar senha").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        self.confirm_entry = ttk.Entry(form, width=34, show="*")
        self.confirm_entry.grid(row=1, column=1, sticky="ew", pady=5)
        form.columnconfigure(1, weight=1)
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(20, 0))
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Salvar alteração", command=self._save).pack(side="right", padx=(0, 8))
    def _save(self):
        try:
            password = self.pass_entry.get()
            confirm = self.confirm_entry.get()
            if not password:
                raise AppError("Informe a nova senha.")
            if password != confirm:
                raise AppError("A confirmação da senha não confere.")
            if not UserManager.find_user(self.config_data, self.current_user):
                raise AppError("Usuário não encontrado.")
            UserManager.update_user_password(self.config_data, self.current_user, password)
            ConfigManager.save(self.config_data)
            self.on_save(self.config_data)
            AuditLogger.write(self.current_user, "alterar_senha", f"alvo={self.current_user}")
            messagebox.showinfo(APP_TITLE, "Senha alterada com sucesso.", parent=self)
            self.destroy()
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e), parent=self)
class MasterManageUsersWindow(SimpleDialog):
    def __init__(self, master, config_data: Dict[str, Any], current_user: str, on_save):
        self.config_data = deepcopy(config_data)
        self.current_user = current_user
        self.on_save = on_save
        self.selected_username = None
        self.users_by_item = {}
        super().__init__(master, "Alterar usuários", "620x430")
        self._build()
        self._load_users()
    def _build(self):
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Alterar usuários", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 6))
        ttk.Label(frm, text="Selecione um usuário para alterar a senha ou excluir. O usuário master não pode excluir a si próprio.", wraplength=580, justify="left").pack(anchor="w", pady=(0, 12))
        mid = ttk.Frame(frm)
        mid.pack(fill="both", expand=True)
        self.tree_users = ttk.Treeview(mid, columns=("username",), show="headings", height=10)
        self.tree_users.heading("username", text="Usuário")
        self.tree_users.column("username", width=260, anchor="w")
        self.tree_users.grid(row=0, column=0, sticky="nsew")
        self.tree_users.bind("<<TreeviewSelect>>", self._on_select)
        yscroll = ttk.Scrollbar(mid, orient="vertical", command=self.tree_users.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.tree_users.configure(yscrollcommand=yscroll.set)
        mid.rowconfigure(0, weight=1)
        mid.columnconfigure(0, weight=1)
        form = ttk.Frame(frm)
        form.pack(fill="x", pady=(14, 0))
        ttk.Label(form, text="Usuário selecionado").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        self.username_var = tk.StringVar(value="")
        ttk.Entry(form, textvariable=self.username_var, width=34, state="readonly").grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Label(form, text="Nova senha").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        self.pass_entry = ttk.Entry(form, width=34, show="*")
        self.pass_entry.grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Label(form, text="Confirmar senha").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=5)
        self.confirm_entry = ttk.Entry(form, width=34, show="*")
        self.confirm_entry.grid(row=2, column=1, sticky="ew", pady=5)
        form.columnconfigure(1, weight=1)
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(18, 0))
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Salvar alteração", command=self._save).pack(side="right", padx=(0, 8))
        ttk.Button(btns, text="Excluir usuário", command=self._remove).pack(side="left")
    def _load_users(self):
        for item in self.tree_users.get_children():
            self.tree_users.delete(item)
        self.users_by_item.clear()
        for user in UserManager.list_users(self.config_data):
            username = user.get("username", "")
            item = self.tree_users.insert("", "end", values=(username,))
            self.users_by_item[item] = username
    def _on_select(self, event=None):
        selected = self.tree_users.selection()
        if not selected:
            self.selected_username = None
            self.username_var.set("")
            return
        item = selected[0]
        self.selected_username = self.users_by_item.get(item)
        self.username_var.set(self.selected_username or "")
    def _require_selected(self):
        username = (self.selected_username or "").strip()
        if not username:
            raise AppError("Selecione um usuário.")
        return username
    def _save(self):
        try:
            username = self._require_selected()
            password = self.pass_entry.get()
            confirm = self.confirm_entry.get()
            if not password:
                raise AppError("Informe a nova senha.")
            if password != confirm:
                raise AppError("A confirmação da senha não confere.")
            if not UserManager.find_user(self.config_data, username):
                raise AppError("Usuário não encontrado.")
            UserManager.update_user_password(self.config_data, username, password)
            ConfigManager.save(self.config_data)
            self.on_save(self.config_data)
            AuditLogger.write(self.current_user, "editar_usuario", f"alvo={username}")
            self.pass_entry.delete(0, "end")
            self.confirm_entry.delete(0, "end")
            messagebox.showinfo(APP_TITLE, "Senha alterada com sucesso.", parent=self)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e), parent=self)
    def _remove(self):
        try:
            username = self._require_selected()
            if username.lower() == MASTER_USERNAME.lower():
                raise AppError("O usuário master não pode ser excluído.")
            if not UserManager.find_user(self.config_data, username):
                raise AppError("Usuário não encontrado.")
            if messagebox.askyesno(APP_TITLE, f"Deseja excluir o usuário '{username}'?", parent=self):
                UserManager.remove_user(self.config_data, username)
                ConfigManager.save(self.config_data)
                self.on_save(self.config_data)
                AuditLogger.write(self.current_user, "excluir_usuario", f"alvo={username}")
                self.selected_username = None
                self.username_var.set("")
                self.pass_entry.delete(0, "end")
                self.confirm_entry.delete(0, "end")
                self._load_users()
                messagebox.showinfo(APP_TITLE, "Usuário excluído com sucesso.", parent=self)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e), parent=self)
class InactiveCustomersWindow(tk.Toplevel):
    FILTER_OPTIONS = {
        "Todos": None,
        "Ativos": "Ativo",
        "Inativos": "Inativo",
        "Deletados": "Deletado",
        "Com limite": "__HAS_CREDIT__",
        "Possui conta": "__HAS_ACCOUNT__",
    }
    def __init__(self, master, config_data: Dict[str, Any], current_user: str, on_config_saved):
        super().__init__(master)
        self.master_app = master
        self.config_data = deepcopy(config_data)
        self.current_user = current_user
        self.on_config_saved = on_config_saved
        self.rows: List[CustomerRow] = []
        self.filtered_rows: List[CustomerRow] = []
        self.tree_items: Dict[str, CustomerRow] = {}
        self.filter_var = tk.StringVar(value="Todos")
        self.inactive_amount_var = tk.StringVar(value="2")
        self.inactive_unit_var = tk.StringVar(value="Anos")
        self.sort_column: Optional[str] = None
        self.sort_reverse = False
        self.status_var = tk.StringVar(value="Pronto.")
        self.title(f"{APP_TITLE} - Clientes inativos")
        self.geometry("1500x780")
        self.minsize(1320, 700)
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._setup_style()
        self._build_ui()
        self._center_window()
        self.load_data()
    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.configure("Treeview", rowheight=24, font=("Segoe UI", 10), background="#ffffff", fieldbackground="#ffffff")
        self.configure(background="#f5f7fb")
    def _center_window(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - self.winfo_width()) // 2
        y = (sh - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 20)}+{max(y, 20)}")
    def _build_ui(self):
        top = ttk.Frame(self, padding=(12, 12, 12, 8))
        top.pack(fill="x")
        row1 = ttk.Frame(top)
        row1.pack(fill="x")
        left_actions = ttk.Frame(row1)
        left_actions.pack(side="left")
        ttk.Button(left_actions, text="Atualizar lista", command=self.load_data).pack(side="left")
        ttk.Button(left_actions, text="Marcar todos", command=self.mark_all).pack(side="left", padx=(8, 0))
        ttk.Button(left_actions, text="Desmarcar todos", command=self.unmark_all).pack(side="left", padx=(8, 0))
        
        ttk.Button(row1, text="Voltar ao início", command=self._close).pack(side="right")
        
        actions = ttk.Frame(row1)
        actions.pack(side="right", padx=(0, 16))
        ttk.Button(actions, text="Inativar cliente", command=lambda: self.run_action("inactivate_customer_sql", "Inativar cliente", "Inativo")).pack(side="left")
        ttk.Button(actions, text="Excluir cliente", command=lambda: self.run_action("delete_customer_sql", "Excluir cliente", "Deletado")).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Zerar limite", command=lambda: self.run_action("disable_credit_sql", "Zerar limite", None)).pack(side="left", padx=(8, 0))
        row2 = ttk.Frame(top)
        row2.pack(fill="x", pady=(8, 0))
        filter_box = ttk.Frame(row2)
        filter_box.pack(side="left")
        ttk.Label(filter_box, text="Mostrar:").pack(side="left", padx=(0, 6))
        filtro = ttk.Combobox(filter_box, textvariable=self.filter_var, values=list(self.FILTER_OPTIONS.keys()), state="readonly", width=12)
        filtro.pack(side="left")
        filtro.bind("<<ComboboxSelected>>", lambda e: self.apply_filter())

        ttk.Separator(filter_box, orient="vertical").pack(side="left", padx=12, fill="y")
        ttk.Label(filter_box, text="Tempo inativo:").pack(side="left", padx=(0, 6))
        tempo_spin = ttk.Spinbox(filter_box, from_=1, to=240, textvariable=self.inactive_amount_var, width=6)
        tempo_spin.pack(side="left")
        tempo_spin.bind("<Return>", lambda e: self.load_data())
        tempo_spin.bind("<FocusOut>", lambda e: self._sync_inactive_amount())
        tempo_unidade = ttk.Combobox(filter_box, textvariable=self.inactive_unit_var, values=["Meses", "Anos"], state="readonly", width=7)
        tempo_unidade.pack(side="left", padx=(6, 0))
        tempo_unidade.bind("<<ComboboxSelected>>", lambda e: self.load_data())
        middle = ttk.Frame(self, padding=(12, 0, 12, 0))
        middle.pack(fill="both", expand=True)
        columns = ("company", "code", "name", "account", "credit_limit", "last_date", "status")
        self.tree = ttk.Treeview(middle, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("company", text="Empresa da última compra", command=lambda: self.sort_by("company"))
        self.tree.heading("code", text="Código", command=lambda: self.sort_by("code"))
        self.tree.heading("name", text="Cliente", command=lambda: self.sort_by("name"))
        self.tree.heading("account", text="Conta", command=lambda: self.sort_by("account"))
        self.tree.heading("credit_limit", text="Lim. crédito", command=lambda: self.sort_by("credit_limit"))
        self.tree.heading("last_date", text="Última compra", command=lambda: self.sort_by("last_date"))
        self.tree.heading("status", text="Status", command=lambda: self.sort_by("status"))
        self.tree.column("company", width=200, minwidth=180, anchor="w", stretch=False)
        self.tree.column("code", width=80, minwidth=76, anchor="center", stretch=False)
        self.tree.column("name", width=280, minwidth=240, anchor="w", stretch=True)
        self.tree.column("account", width=170, minwidth=150, anchor="w", stretch=False)
        self.tree.column("credit_limit", width=110, minwidth=100, anchor="e", stretch=False)
        self.tree.column("last_date", width=110, minwidth=100, anchor="center", stretch=False)
        self.tree.column("status", width=95, minwidth=90, anchor="center", stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<Control-a>", self._select_all_rows)
        self.tree.bind("<Control-A>", self._select_all_rows)
        yscroll = ttk.Scrollbar(middle, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll = ttk.Scrollbar(middle, orient="horizontal", command=self.tree.xview)
        xscroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        middle.rowconfigure(0, weight=1)
        middle.columnconfigure(0, weight=1)
        bottom = ttk.Frame(self, padding=(12, 8, 12, 10))
        bottom.pack(fill="x")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left")
        ttk.Label(bottom, text=f"Usuário: {self.current_user}").pack(side="right")
    def _sync_inactive_amount(self):
        try:
            value = int(self.inactive_amount_var.get() or 0)
        except Exception:
            value = 0
        if value < 1:
            self.inactive_amount_var.set(1)
    def _inactive_months(self) -> int:
        self._sync_inactive_amount()
        try:
            value = int(self.inactive_amount_var.get() or 1)
        except Exception:
            value = 1
        unit = (self.inactive_unit_var.get() or "Meses").strip()
        if unit.lower().startswith("ano"):
            return max(1, value) * 12
        return max(1, value)
    def _close(self):
        self.destroy()
        if hasattr(self.master_app, "inactive_window"):
            self.master_app.inactive_window = None
        if hasattr(self.master_app, "show_home"):
            self.master_app.show_home()
    def open_config(self):
        ConfigWindow(self, self.config_data, self._apply_new_config)
    def _apply_new_config(self, cfg: Dict[str, Any]):
        self.config_data = cfg
        self.on_config_saved(cfg)
        self.load_data()
    def set_status(self, text: str):
        self.status_var.set(text)
        self.update_idletasks()
    def _normalize_status(self, value: Any) -> str:
        txt = str(value or "").strip()
        if txt in ("A", "Ativo"):
            return "Ativo"
        if txt in ("I", "Inativo"):
            return "Inativo"
        if txt in ("D", "Deletado"):
            return "Deletado"
        return txt
    def _sort_value(self, row: CustomerRow, column: str):
        if column == "company":
            return (row.last_purchase_company or "").lower()
        if column == "code":
            try:
                return (0, int(str(row.customer_code)))
            except Exception:
                return (1, str(row.customer_code or "").lower())
        if column == "name":
            return (row.customer_name or "").lower()
        if column == "account":
            return (row.account_name or "").lower()
        if column == "credit_limit":
            try:
                return float(row.credit_limit or 0)
            except Exception:
                return 0.0
        if column == "last_date":
            value = row.last_purchase_date
            if value is None:
                return datetime.min
            if isinstance(value, datetime):
                return value
            if isinstance(value, date):
                return datetime.combine(value, time.min)
            return str(value)
        if column == "status":
            order = {"Ativo": 0, "Inativo": 1, "Deletado": 2}
            return (order.get(row.customer_status, 99), row.customer_status)
        return ""
    def sort_by(self, column: str):
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False
        self.filtered_rows.sort(key=lambda r: self._sort_value(r, column), reverse=self.sort_reverse)
        self._refresh_tree()
        self._update_heading_titles()
    def _update_heading_titles(self):
        labels = {
            "company": "Empresa da última compra",
            "code": "Código",
            "name": "Cliente",
            "account": "Conta",
            "credit_limit": "Lim. crédito",
            "last_date": "Última compra",
            "status": "Status",
        }
        for col, label in labels.items():
            suffix = ""
            if self.sort_column == col:
                suffix = " ▼" if self.sort_reverse else " ▲"
            self.tree.heading(col, text=label + suffix, command=lambda c=col: self.sort_by(c))
    def load_data(self):
        self.set_status("Conectando ao banco e carregando clientes.")

        def _work():
            return Database(self.config_data).list_inactive_customers(inactive_months=self._inactive_months())

        def _ok(data):
            self.rows = [
                CustomerRow(
                    customer_id=row.get("customer_id"),
                    customer_code=row.get("customer_code"),
                    customer_name=row.get("customer_name") or "",
                    last_purchase_date=row.get("last_purchase_date"),
                    last_purchase_company=row.get("last_purchase_company") or "",
                    account_name=row.get("account_name") or "Sem conta",
                    customer_status=self._normalize_status(row.get("customer_status")),
                    has_account=bool(row.get("has_account")),
                    credit_limit=row.get("credit_limit"),
                    selected=False,
                )
                for row in (data or [])
            ]
            self.apply_filter()
            self.set_status(f"{len(self.filtered_rows)} cliente(s) encontrado(s).")
            AuditLogger.write(self.current_user, "carregar_lista", f"tipo=clientes_inativos;quantidade={len(self.filtered_rows)}")

        def _err(e: Exception):
            self.set_status("Falha ao carregar dados.")
            messagebox.showerror(APP_TITLE, f"Erro ao carregar clientes:\n\n{e}", parent=self)
            AuditLogger.write(self.current_user, "erro_carregar_lista", f"tipo=clientes_inativos;erro={e}")

        run_with_busy(self, "Carregando clientes...", _work, _ok, _err)
    def apply_filter(self):
        selected_filter = self.FILTER_OPTIONS.get(self.filter_var.get())
        if selected_filter is None:
            self.filtered_rows = list(self.rows)
        elif selected_filter == "__HAS_CREDIT__":
            self.filtered_rows = [r for r in self.rows if float(r.credit_limit or 0) > 0]
        elif selected_filter == "__HAS_ACCOUNT__":
            self.filtered_rows = [r for r in self.rows if bool(r.has_account)]
        else:
            self.filtered_rows = [r for r in self.rows if r.customer_status == selected_filter]
        if self.sort_column:
            self.filtered_rows.sort(key=lambda r: self._sort_value(r, self.sort_column), reverse=self.sort_reverse)
        self._refresh_tree()
        self._update_heading_titles()
        self.set_status(f"{len(self.filtered_rows)} cliente(s) encontrado(s).")
    def _row_values(self, row: CustomerRow):
        return (
            row.last_purchase_company,
            row.customer_code,
            row.customer_name,
            row.account_name,
            row.credit_limit_display(),
            row.last_purchase_date_display(),
            row.customer_status,
        )
    def _refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree_items.clear()
        for row in self.filtered_rows:
            item_id = self.tree.insert("", "end", values=self._row_values(row))
            self.tree_items[item_id] = row


    def mark_all(self):
        items = self.tree.get_children()
        if items:
            self.tree.selection_set(items)

    def unmark_all(self):
        selected_items = self.tree.selection()
        if selected_items:
            self.tree.selection_remove(selected_items)

    def _select_all_rows(self, event=None):
        self.mark_all()
        return "break"
    def selected_rows(self) -> List[CustomerRow]:
        selected_items = self.tree.selection()
        if not selected_items:
            return []
        return [self.tree_items[item_id] for item_id in selected_items if item_id in self.tree_items]
    def run_action(self, query_key: str, action_name: str, new_status: Optional[str]):
        selected = self.selected_rows()
        if not selected:
            messagebox.showwarning(APP_TITLE, "Selecione pelo menos um cliente.", parent=self)
            return
        sql_text = (self.config_data.get("queries", {}).get(query_key) or "").strip()
        if not sql_text:
            messagebox.showwarning(APP_TITLE, "A SQL da ação não está configurada.", parent=self)
            return
        if not messagebox.askyesno(APP_TITLE, f"Deseja executar a ação '{action_name}' para {len(selected)} cliente(s)?", parent=self):
            return
        customer_ids = [r.customer_id for r in selected]

        def _work():
            return Database(self.config_data).execute_action(sql_text, customer_ids)

        def _ok(affected):
            if new_status:
                for row in selected:
                    row.customer_status = new_status
            if query_key == "disable_credit_sql":
                for row in selected:
                    row.credit_limit = 0
            AuditLogger.write(self.current_user, "acao_operacional", f"acao={action_name};selecionados={len(selected)};afetados={affected}")
            self.load_data()
            messagebox.showinfo(APP_TITLE, f"Ação '{action_name}' executada com sucesso.", parent=self)

        def _err(e: Exception):
            AuditLogger.write(self.current_user, "erro_acao_operacional", f"acao={action_name};erro={e}")
            messagebox.showerror(APP_TITLE, f"Erro ao executar a ação:\n\n{e}", parent=self)

        run_with_busy(self, f"Executando '{action_name}'...", _work, _ok, _err)
class OpenInvoicesWindow(tk.Toplevel):
    GROUP_OPTIONS = {
        "Não agrupar": "none",
        "Cliente": "customer",
        "Vencimento": "due_date",
        "Conta": "account_group",
    }
    def __init__(self, master, config_data: Dict[str, Any], current_user: str, on_config_saved):
        super().__init__(master)
        self.master_app = master
        self.config_data = deepcopy(config_data)
        self.current_user = current_user
        self.on_config_saved = on_config_saved
        self.raw_rows: List[InvoiceRow] = []
        self.rows: List[InvoiceRow] = []
        self.tree_items: Dict[str, InvoiceRow] = {}
        self.sort_column: Optional[str] = None
        self.sort_reverse = False
        self.status_var = tk.StringVar(value="Pronto.")
        today = date.today()
        self.period_start_var = tk.StringVar(value=today.strftime("%d/%m/%Y"))
        self.period_end_var = tk.StringVar(value=today.strftime("%d/%m/%Y"))
        self.group_by_var = tk.StringVar(value="Não agrupar")
        self.customer_var = tk.StringVar(value="Todos")
        self.account_var = tk.StringVar(value="Todas")
        self.customer_options_map: Dict[str, Any] = {"Todos": None}
        self.all_customer_options_map: Dict[str, Any] = {"Todos": None}
        self.account_options_map: Dict[str, Any] = {"Todas": None}
        self.all_account_options_map: Dict[str, Any] = {"Todas": None}
        self._auto_filter_job = None
        self._loading = False
        self._pending_reload = False
        self.title(f"{APP_TITLE} - Faturas a receber (apenas com boletos vinculados)")
        self.geometry("1680x760")
        self.minsize(1500, 700)
        self.resizable(True, True)
        self.transient(master)
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._setup_style()
        self._build_ui()
        self._center_window()
        self._initial_load()

    def _initial_load(self):
        self.set_status("Conectando ao banco e carregando dados.")

        def _work():
            date_from = self._parse_date(self.period_start_var.get())
            date_to = self._parse_date(self.period_end_var.get())
            if date_from and date_to and date_from > date_to:
                raise AppError("O período inicial não pode ser maior que o período final.")
            db = Database(self.config_data)
            customers = db.list_open_invoice_customers()
            accounts = db.list_open_invoice_accounts()
            invoices = db.list_open_invoices(due_date_from=date_from, due_date_to=date_to, customer_id=None, account_code=None)
            invoice_ids = [r.get("movto_id") for r in (invoices or [])]
            try:
                sig_map = db.get_sale_signatures_pdf_bulk(invoice_ids)
            except Exception:
                sig_map = {}
            try:
                boleto_map = db.check_boleto_exists_bulk(invoice_ids)
            except Exception:
                boleto_map = {}
            try:
                nf_map = db.check_nota_fiscal_exists_bulk(invoice_ids)
            except Exception:
                nf_map = {}
            for r in (invoices or []):
                inv_id = r.get("movto_id")
                sig = sig_map.get(inv_id) or sig_map.get(str(inv_id)) or {}
                r["has_signed_doc"] = bool(sig.get("exists") or (sig.get("attachments") or []))
                r["has_boleto"] = bool(boleto_map.get(inv_id) or boleto_map.get(str(inv_id)))
                r["has_nota_fiscal"] = bool(nf_map.get(inv_id) or nf_map.get(str(inv_id)))
            return customers, accounts, invoices

        def _ok(res):
            customers, accounts, invoices = res

            options_map = {"Todos": None}
            for row in (customers or []):
                label = f"{row.get('codigo_cliente')} - {row.get('cliente')}"
                options_map[label] = row.get("customer_id")
            self.all_customer_options_map = options_map
            self.customer_options_map = dict(options_map)
            self.customer_var.set("Todos")
            self._hide_customer_suggestions()

            acc_map = {"Todas": None}
            for row in (accounts or []):
                label = f"{row.get('conta')} - {row.get('conta_nome')}"
                acc_map[label] = row.get("conta")
            self.all_account_options_map = acc_map
            self.account_options_map = dict(acc_map)
            self.account_var.set("Todas")
            self._hide_account_suggestions()

            self._apply_invoices_data(invoices)

        def _err(e: Exception):
            self.set_status("Falha ao carregar dados.")
            messagebox.showerror(APP_TITLE, f"Erro ao carregar faturas a receber:\n\n{e}", parent=self)

        run_with_busy(self, "Carregando faturas e filtros...", _work, _ok, _err)

    def _apply_invoices_data(self, data):
        self.raw_rows = [
            InvoiceRow(
                invoice_id=row.get("movto_id"),
                company=row.get("empresa") or "",
                customer_id=row.get("customer_id"),
                customer_code=row.get("codigo_cliente"),
                customer_name=row.get("cliente") or "",
                has_signed_doc=bool(row.get("has_signed_doc")),
                has_boleto=bool(row.get("has_boleto")),
                has_nota_fiscal=bool(row.get("has_nota_fiscal")),
                motive_code=row.get("motivo"),
                motive_name=f"Motivo {row.get('motivo')}" if row.get("motivo") not in (None, "") else "",
                account_code=row.get("conta") or "",
                account_name=row.get("conta_nome") or "",
                issue_date=row.get("data"),
                due_date=row.get("vencto"),
                amount=row.get("valor"),
                discount_amount=row.get("valor_desconto"),
                paid_amount=row.get("valor_baixado"),
                open_balance=row.get("saldo_em_aberto"),
                customer_email=row.get("customer_email") or "",
            )
            for row in (data or [])
        ]
        self.rows = self._group_rows(self.raw_rows)
        if self.sort_column:
            self.rows.sort(key=lambda r: self._sort_value(r, self.sort_column), reverse=self.sort_reverse)
        self._refresh_tree()
        self._update_heading_titles()
        total_open = sum(float(r.open_balance or 0) for r in self.rows)
        self.set_status(f"{len(self.rows)} registro(s). Total em aberto: {total_open:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        AuditLogger.write(self.current_user, "carregar_lista", f"tipo=faturas_receber;quantidade={len(self.rows)};agrupar={self.GROUP_OPTIONS.get(self.group_by_var.get(), 'none')}")
    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.configure("Treeview", rowheight=24, font=("Segoe UI", 10), background="#ffffff", fieldbackground="#ffffff")
        self.configure(background="#f5f7fb")
    def _center_window(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - self.winfo_width()) // 2
        y = (sh - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 20)}+{max(y, 20)}")
    def _build_ui(self):
        top = ttk.Frame(self, padding=(12, 12, 12, 8))
        top.pack(fill="x")
        actions = ttk.Frame(top)
        actions.pack(side="left")
        ttk.Button(actions, text="Limpar filtros", command=self.clear_filters).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Selecionar tudo", command=self._select_all_tree_rows).pack(side="left", padx=(12, 0))
        ttk.Button(actions, text="Limpar seleção", command=self._clear_tree_selection).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Enviar fatura por e-mail", command=self.open_email_window).pack(side="left", padx=(12, 0))
        
        ttk.Button(top, text="Voltar ao início", command=self._close).pack(side="right")
        
        filters = ttk.Frame(self, padding=(12, 0, 12, 8))
        filters.pack(fill="x")
        ttk.Label(filters, text="Vencimento inicial").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        self.period_start_entry = ttk.Entry(filters, textvariable=self.period_start_var, width=14)
        self.period_start_entry.grid(row=0, column=1, sticky="w", pady=4)
        bind_date_entry_shortcuts(self.period_start_entry)
        self.period_start_entry.bind("<FocusOut>", self._schedule_auto_filter, add="+")
        ttk.Label(filters, text="Vencimento final").grid(row=0, column=2, sticky="w", padx=(16, 6), pady=4)
        self.period_end_entry = ttk.Entry(filters, textvariable=self.period_end_var, width=14)
        self.period_end_entry.grid(row=0, column=3, sticky="w", pady=4)
        bind_date_entry_shortcuts(self.period_end_entry)
        self.period_end_entry.bind("<FocusOut>", self._schedule_auto_filter, add="+")
        ttk.Label(filters, text="Cliente").grid(row=0, column=4, sticky="w", padx=(16, 6), pady=4)
        self.customer_entry = ttk.Entry(filters, textvariable=self.customer_var, width=28)
        self.customer_entry.grid(row=0, column=5, sticky="ew", pady=4)
        self.customer_entry.bind("<KeyRelease>", self._on_customer_keyrelease)
        self.customer_entry.bind("<FocusIn>", self._on_customer_focus_in)
        self.customer_entry.bind("<FocusOut>", self._on_customer_focus_out)
        self.customer_entry.bind("<Down>", self._on_customer_arrow_down)
        self.customer_entry.bind("<Up>", self._on_customer_arrow_up)
        self.customer_entry.bind("<Return>", self._on_customer_entry_confirm)
        self.customer_suggestions_frame = ttk.Frame(filters)
        self.customer_suggestions_frame.grid(row=1, column=4, columnspan=2, sticky="nsew", padx=(16, 0), pady=(0, 2))
        self.customer_suggestions_frame.grid_remove()
        self.customer_listbox = tk.Listbox(self.customer_suggestions_frame, height=6, exportselection=False)
        self.customer_listbox.pack(fill="x", expand=True)
        self.customer_listbox.bind("<<ListboxSelect>>", self._on_customer_listbox_select)
        self.customer_listbox.bind("<ButtonRelease-1>", self._on_customer_listbox_confirm)
        self.customer_listbox.bind("<Double-Button-1>", self._on_customer_listbox_confirm)
        self.customer_listbox.bind("<Return>", self._on_customer_listbox_confirm)

        ttk.Label(filters, text="Conta").grid(row=0, column=6, sticky="w", padx=(16, 6), pady=4)
        self.account_entry = ttk.Entry(filters, textvariable=self.account_var, width=28)
        self.account_entry.grid(row=0, column=7, sticky="ew", pady=4)
        self.account_entry.bind("<KeyRelease>", self._on_account_keyrelease)
        self.account_entry.bind("<FocusIn>", self._on_account_focus_in)
        self.account_entry.bind("<FocusOut>", self._on_account_focus_out)
        self.account_entry.bind("<Down>", self._on_account_arrow_down)
        self.account_entry.bind("<Up>", self._on_account_arrow_up)
        self.account_entry.bind("<Return>", self._on_account_entry_confirm)
        self.account_suggestions_frame = ttk.Frame(filters)
        self.account_suggestions_frame.grid(row=1, column=6, columnspan=2, sticky="nsew", padx=(16, 0), pady=(0, 2))
        self.account_suggestions_frame.grid_remove()
        self.account_listbox = tk.Listbox(self.account_suggestions_frame, height=6, exportselection=False)
        self.account_listbox.pack(fill="x", expand=True)
        self.account_listbox.bind("<<ListboxSelect>>", self._on_account_listbox_select)
        self.account_listbox.bind("<ButtonRelease-1>", self._on_account_listbox_confirm)
        self.account_listbox.bind("<Double-Button-1>", self._on_account_listbox_confirm)
        self.account_listbox.bind("<Return>", self._on_account_listbox_confirm)

        ttk.Label(filters, text="Agrupar por").grid(row=0, column=8, sticky="w", padx=(16, 6), pady=4)
        self.group_combo = ttk.Combobox(filters, textvariable=self.group_by_var, state="readonly", width=18, values=list(self.GROUP_OPTIONS.keys()))
        self.group_combo.grid(row=0, column=9, sticky="w", pady=4)
        self.group_combo.bind("<<ComboboxSelected>>", self._schedule_auto_filter)
        ttk.Label(filters, text="Formato: DD/MM/AAAA").grid(row=2, column=0, columnspan=4, sticky="w", pady=(2, 0))
        filters.columnconfigure(5, weight=1)
        filters.columnconfigure(7, weight=1)
        middle = ttk.Frame(self, padding=(12, 0, 12, 0))
        middle.pack(fill="both", expand=True)
        columns = ("company", "account_display", "invoice", "code", "name", "issue_date", "due_date", "open_balance", "signed_doc", "boleto", "nota_fiscal", "fatura_anexo")
        self.tree = ttk.Treeview(middle, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("company", text="Empresa", command=lambda: self.sort_by("company"))
        self.tree.heading("account_display", text="Conta", command=lambda: self.sort_by("account_display"))
        self.tree.heading("invoice", text="Nº Fatura", command=lambda: self.sort_by("invoice"))
        self.tree.heading("code", text="Código", command=lambda: self.sort_by("code"))
        self.tree.heading("name", text="Cliente", command=lambda: self.sort_by("name"))
        self.tree.heading("issue_date", text="Data", command=lambda: self.sort_by("issue_date"))
        self.tree.heading("due_date", text="Vencimento", command=lambda: self.sort_by("due_date"))
        self.tree.heading("open_balance", text="Valor da fatura", command=lambda: self.sort_by("open_balance"))
        self.tree.heading("signed_doc", text="Assinado", command=lambda: self.sort_by("signed_doc"))
        self.tree.heading("boleto", text="Boleto", command=lambda: self.sort_by("boleto"))
        self.tree.heading("nota_fiscal", text="NF", command=lambda: self.sort_by("nota_fiscal"))
        self.tree.heading("fatura_anexo", text="Fatura", command=lambda: self.sort_by("fatura_anexo"))
        self.tree.column("company", width=180, minwidth=160, anchor="w", stretch=False)
        self.tree.column("account_display", width=260, minwidth=240, anchor="w", stretch=False)
        self.tree.column("invoice", width=90, minwidth=80, anchor="center", stretch=False)
        self.tree.column("code", width=80, minwidth=70, anchor="center", stretch=False)
        self.tree.column("name", width=250, minwidth=220, anchor="w", stretch=True)
        self.tree.column("issue_date", width=100, minwidth=90, anchor="center", stretch=False)
        self.tree.column("due_date", width=100, minwidth=90, anchor="center", stretch=False)
        self.tree.column("open_balance", width=130, minwidth=120, anchor="e", stretch=False)
        self.tree.column("signed_doc", width=90, minwidth=80, anchor="center", stretch=False)
        self.tree.column("boleto", width=80, minwidth=70, anchor="center", stretch=False)
        self.tree.column("nota_fiscal", width=70, minwidth=60, anchor="center", stretch=False)
        self.tree.column("fatura_anexo", width=70, minwidth=60, anchor="center", stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(middle, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll = ttk.Scrollbar(middle, orient="horizontal", command=self.tree.xview)
        xscroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        try:
            self.tree.bind("<Control-a>", lambda e: (self._select_all_tree_rows() or "break"))
            self.tree.bind("<Control-A>", lambda e: (self._select_all_tree_rows() or "break"))
        except Exception:
            pass
        middle.rowconfigure(0, weight=1)
        middle.columnconfigure(0, weight=1)
        bottom = ttk.Frame(self, padding=(12, 8, 12, 10))
        bottom.pack(fill="x")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left")
        ttk.Label(bottom, text=f"Usuário: {self.current_user}").pack(side="right")
    def _close(self):
        self.destroy()
        if hasattr(self.master_app, "invoices_window"):
            self.master_app.invoices_window = None
        if hasattr(self.master_app, "show_home"):
            self.master_app.show_home()
    def open_config(self):
        ConfigWindow(self, self.config_data, self._apply_new_config)
    def _apply_new_config(self, cfg: Dict[str, Any]):
        self.config_data = cfg
        self.on_config_saved(cfg)
        current_customer = (self.customer_var.get() or "").strip() or "Todos"
        current_account = (self.account_var.get() or "").strip() or "Todas"

        def _work():
            db = Database(self.config_data)
            customers = db.list_open_invoice_customers()
            accounts = db.list_open_invoice_accounts()
            return customers, accounts

        def _ok(res):
            customers, accounts = res
            options_map = {"Todos": None}
            for row in (customers or []):
                label = f"{row.get('codigo_cliente')} - {row.get('cliente')}"
                options_map[label] = row.get("customer_id")
            self.all_customer_options_map = options_map
            self.customer_options_map = dict(options_map)
            self.customer_var.set(current_customer if current_customer in options_map else "Todos")
            self._hide_customer_suggestions()

            acc_map = {"Todas": None}
            for row in (accounts or []):
                label = f"{row.get('conta')} - {row.get('conta_nome')}"
                acc_map[label] = row.get("conta")
            self.all_account_options_map = acc_map
            self.account_options_map = dict(acc_map)
            self.account_var.set(current_account if current_account in acc_map else "Todas")
            self._hide_account_suggestions()

            self.load_data()

        def _err(e: Exception):
            messagebox.showerror(APP_TITLE, f"Erro ao recarregar filtros:\n\n{e}", parent=self)
            self.load_data()

        run_with_busy(self, "Recarregando filtros...", _work, _ok, _err)
    def _schedule_auto_filter(self, event=None):
        if self._auto_filter_job is not None:
            try:
                self.after_cancel(self._auto_filter_job)
            except Exception:
                pass
        self._auto_filter_job = self.after(350, self._run_auto_filter)
        return None

    def _run_auto_filter(self):
        self._auto_filter_job = None
        self.load_data()

    def clear_filters(self):
        today = date.today()
        self.period_start_var.set(today.strftime("%d/%m/%Y"))
        self.period_end_var.set(today.strftime("%d/%m/%Y"))
        self.customer_var.set("Todos")
        self.account_var.set("Todas")
        self.group_by_var.set("Não agrupar")
        self._hide_customer_suggestions()
        self._hide_account_suggestions()
        self.load_data()
    def set_status(self, text: str):
        self.status_var.set(text)
        self.update_idletasks()
    def _parse_date(self, value: str):
        value = (value or "").strip()
        if not value:
            return None
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                pass
        raise AppError("Data inválida. Use DD/MM/AAAA.")

    def _selected_customer_id(self):
        typed = (self.customer_var.get() or "").strip()
        if not typed or typed == "Todos":
            return None

        exact = self.all_customer_options_map.get(typed)
        if exact is not None:
            return exact

        typed_lower = typed.lower()
        matches = [
            customer_id
            for label, customer_id in self.all_customer_options_map.items()
            if label != "Todos" and typed_lower in label.lower()
        ]
        if len(matches) == 1:
            return matches[0]
        return None

    def _load_customer_options(self):
        current = (self.customer_var.get() or "").strip() or "Todos"
        options_map = {"Todos": None}
        try:
            rows = Database(self.config_data).list_open_invoice_customers()
            for row in rows:
                label = f"{row.get('codigo_cliente')} - {row.get('cliente')}"
                options_map[label] = row.get("customer_id")
        except Exception:
            pass

        self.all_customer_options_map = options_map
        self.customer_options_map = dict(options_map)
        self.customer_var.set(current if current in options_map else "Todos")
        self._hide_customer_suggestions()

    def _matching_customer_labels(self, typed: str):
        typed = (typed or "").strip().lower()
        if not typed or typed == "todos":
            return []
        parts = [p for p in re.split(r"\s+", typed) if p]
        labels = []
        for label in self.all_customer_options_map.keys():
            if label == "Todos":
                continue
            low = label.lower()
            if all(p in low for p in parts):
                labels.append(label)
        return labels[:200]

    def _show_customer_suggestions(self, labels):
        if not labels:
            self._hide_customer_suggestions()
            return
        self.customer_listbox.delete(0, "end")
        for label in labels:
            self.customer_listbox.insert("end", label)
        self.customer_suggestions_frame.grid()
        self.customer_listbox.selection_clear(0, "end")
        self.customer_listbox.selection_set(0)
        self.customer_listbox.activate(0)
        self.customer_listbox.see(0)

    def _hide_customer_suggestions(self):
        if hasattr(self, "customer_suggestions_frame"):
            self.customer_suggestions_frame.grid_remove()

    def _apply_customer_search(self, typed: str):
        typed = (typed or "").strip()
        if not typed or typed == "Todos":
            self.customer_options_map = dict(self.all_customer_options_map)
            self._hide_customer_suggestions()
            return
        labels = self._matching_customer_labels(typed)
        self.customer_options_map = {label: self.all_customer_options_map.get(label) for label in labels}
        self._show_customer_suggestions(labels)

    def _on_customer_keyrelease(self, event=None):
        keysym = getattr(event, "keysym", "")
        if keysym in ("Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return None
        if not (self.customer_var.get() or "").strip():
            self.customer_var.set("Todos")
            self._hide_customer_suggestions()
            self._schedule_auto_filter()
            return None
        self._apply_customer_search(self.customer_var.get())
        return None

    def _on_customer_focus_in(self, event=None):
        typed = (self.customer_var.get() or "").strip()
        if typed and typed != "Todos":
            self._apply_customer_search(typed)
        return None

    def _on_customer_focus_out(self, event=None):
        self.after(150, self._hide_customer_suggestions)
        if not (self.customer_var.get() or "").strip():
            self.customer_var.set("Todos")
            self._schedule_auto_filter()
        return None

    def _on_customer_arrow_down(self, event=None):
        if not self.customer_suggestions_frame.winfo_ismapped() or self.customer_listbox.size() == 0:
            self._apply_customer_search(self.customer_var.get())
        if not self.customer_suggestions_frame.winfo_ismapped() or self.customer_listbox.size() == 0:
            return "break"
        selection = self.customer_listbox.curselection()
        index = selection[0] + 1 if selection else 0
        if index >= self.customer_listbox.size():
            index = self.customer_listbox.size() - 1
        self.customer_listbox.selection_clear(0, "end")
        self.customer_listbox.selection_set(index)
        self.customer_listbox.activate(index)
        self.customer_listbox.see(index)
        return "break"

    def _on_customer_arrow_up(self, event=None):
        if not self.customer_suggestions_frame.winfo_ismapped() or self.customer_listbox.size() == 0:
            self._apply_customer_search(self.customer_var.get())
        if not self.customer_suggestions_frame.winfo_ismapped() or self.customer_listbox.size() == 0:
            return "break"
        selection = self.customer_listbox.curselection()
        index = selection[0] - 1 if selection else 0
        if index < 0:
            index = 0
        self.customer_listbox.selection_clear(0, "end")
        self.customer_listbox.selection_set(index)
        self.customer_listbox.activate(index)
        self.customer_listbox.see(index)
        return "break"

    def _focus_customer_listbox(self, event=None):
        if self.customer_suggestions_frame.winfo_ismapped() and self.customer_listbox.size() > 0:
            selection = self.customer_listbox.curselection()
            index = selection[0] if selection else 0
            self.customer_listbox.selection_clear(0, "end")
            self.customer_listbox.selection_set(index)
            self.customer_listbox.activate(index)
            self.customer_listbox.see(index)
            return "break"
        return None

    def _on_customer_entry_confirm(self, event=None):
        typed = (self.customer_var.get() or "").strip()
        if self.customer_suggestions_frame.winfo_ismapped() and self.customer_listbox.size() > 0:
            selection = self.customer_listbox.curselection()
            if selection:
                label = self.customer_listbox.get(selection[0])
            else:
                label = self.customer_listbox.get(0)
            self.customer_var.set(label)
        elif typed:
            matches = self._matching_customer_labels(typed)
            if len(matches) == 1:
                self.customer_var.set(matches[0])
        self._hide_customer_suggestions()
        self._schedule_auto_filter()
        return "break"

    def _on_customer_listbox_select(self, event=None):
        # Navegação com seta para cima/baixo não deve aplicar filtro
        # nem alterar o conteúdo do campo. A confirmação ocorre apenas
        # com Enter ou clique do mouse.
        return None

    def _on_customer_listbox_confirm(self, event=None):
        selection = self.customer_listbox.curselection()
        if selection:
            label = self.customer_listbox.get(selection[0])
            self.customer_var.set(label)
            self._hide_customer_suggestions()
            self.customer_entry.focus_set()
            self.customer_entry.icursor("end")
            self._schedule_auto_filter()
        return "break"


    def _selected_account_code(self):
        typed = (self.account_var.get() or "").strip()
        if not typed or typed == "Todas":
            return None

        exact = self.all_account_options_map.get(typed)
        if exact is not None:
            return exact

        typed_lower = typed.lower()
        matches = [
            account_code
            for label, account_code in self.all_account_options_map.items()
            if label != "Todas" and typed_lower in label.lower()
        ]
        if len(matches) == 1:
            return matches[0]
        return None

    def _load_account_options(self):
        current = (self.account_var.get() or "").strip() or "Todas"
        options_map = {"Todas": None}
        try:
            rows = Database(self.config_data).list_open_invoice_accounts()
            for row in rows:
                label = f"{row.get('conta')} - {row.get('conta_nome')}"
                options_map[label] = row.get("conta")
        except Exception:
            pass

        self.all_account_options_map = options_map
        self.account_options_map = dict(options_map)
        self.account_var.set(current if current in options_map else "Todas")
        self._hide_account_suggestions()

    def _matching_account_labels(self, typed: str):
        typed = (typed or "").strip().lower()
        if not typed or typed == "todas":
            return []
        parts = [p for p in re.split(r"\s+", typed) if p]
        labels = []
        for label in self.all_account_options_map.keys():
            if label == "Todas":
                continue
            low = label.lower()
            if all(p in low for p in parts):
                labels.append(label)
        return labels[:200]

    def _show_account_suggestions(self, labels):
        if not labels:
            self._hide_account_suggestions()
            return
        self.account_listbox.delete(0, "end")
        for label in labels:
            self.account_listbox.insert("end", label)
        self.account_suggestions_frame.grid()
        self.account_listbox.selection_clear(0, "end")
        self.account_listbox.selection_set(0)
        self.account_listbox.activate(0)
        self.account_listbox.see(0)

    def _hide_account_suggestions(self):
        if hasattr(self, "account_suggestions_frame"):
            self.account_suggestions_frame.grid_remove()

    def _apply_account_search(self, typed: str):
        typed = (typed or "").strip()
        if not typed or typed == "Todas":
            self.account_options_map = dict(self.all_account_options_map)
            self._hide_account_suggestions()
            return
        labels = self._matching_account_labels(typed)
        self.account_options_map = {label: self.all_account_options_map.get(label) for label in labels}
        self._show_account_suggestions(labels)

    def _on_account_keyrelease(self, event=None):
        keysym = getattr(event, "keysym", "")
        if keysym in ("Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return None
        if not (self.account_var.get() or "").strip():
            self.account_var.set("Todas")
            self._hide_account_suggestions()
            self._schedule_auto_filter()
            return None
        self._apply_account_search(self.account_var.get())
        return None

    def _on_account_focus_in(self, event=None):
        typed = (self.account_var.get() or "").strip()
        if typed and typed != "Todas":
            self._apply_account_search(typed)
        return None

    def _on_account_focus_out(self, event=None):
        self.after(150, self._hide_account_suggestions)
        if not (self.account_var.get() or "").strip():
            self.account_var.set("Todas")
            self._schedule_auto_filter()
        return None

    def _on_account_arrow_down(self, event=None):
        if not self.account_suggestions_frame.winfo_ismapped() or self.account_listbox.size() == 0:
            self._apply_account_search(self.account_var.get())
        if not self.account_suggestions_frame.winfo_ismapped() or self.account_listbox.size() == 0:
            return "break"
        selection = self.account_listbox.curselection()
        index = selection[0] + 1 if selection else 0
        if index >= self.account_listbox.size():
            index = self.account_listbox.size() - 1
        self.account_listbox.selection_clear(0, "end")
        self.account_listbox.selection_set(index)
        self.account_listbox.activate(index)
        self.account_listbox.see(index)
        return "break"

    def _on_account_arrow_up(self, event=None):
        if not self.account_suggestions_frame.winfo_ismapped() or self.account_listbox.size() == 0:
            self._apply_account_search(self.account_var.get())
        if not self.account_suggestions_frame.winfo_ismapped() or self.account_listbox.size() == 0:
            return "break"
        selection = self.account_listbox.curselection()
        index = selection[0] - 1 if selection else 0
        if index < 0:
            index = 0
        self.account_listbox.selection_clear(0, "end")
        self.account_listbox.selection_set(index)
        self.account_listbox.activate(index)
        self.account_listbox.see(index)
        return "break"

    def _on_account_entry_confirm(self, event=None):
        typed = (self.account_var.get() or "").strip()
        if self.account_suggestions_frame.winfo_ismapped() and self.account_listbox.size() > 0:
            selection = self.account_listbox.curselection()
            if selection:
                label = self.account_listbox.get(selection[0])
            else:
                label = self.account_listbox.get(0)
            self.account_var.set(label)
        elif typed:
            matches = self._matching_account_labels(typed)
            if len(matches) == 1:
                self.account_var.set(matches[0])
        self._hide_account_suggestions()
        self._schedule_auto_filter()
        return "break"

    def _on_account_listbox_select(self, event=None):
        return None

    def _on_account_listbox_confirm(self, event=None):
        selection = self.account_listbox.curselection()
        if selection:
            label = self.account_listbox.get(selection[0])
            self.account_var.set(label)
            self._hide_account_suggestions()
            self.account_entry.focus_set()
            self.account_entry.icursor("end")
            self._schedule_auto_filter()
        return "break"

    def _sort_value(self, row: InvoiceRow, column: str):
        if column == "company":
            return (row.company or "").lower()
        if column == "account_display":
            return (f"{row.account_code or ''} - {row.account_name or ''}").lower().strip(" -")
        if column == "invoice":
            inv_id = getattr(row, "invoice_id", None)
            inv_str = "" if inv_id in (None, "") else str(inv_id)
            if inv_str.startswith("grp_"):
                return (1, 0)
            try:
                return (0, int(inv_str))
            except Exception:
                return (0, inv_str.lower())
        if column == "code":
            try:
                return (0, int(str(row.customer_code)))
            except Exception:
                return (1, str(row.customer_code or "").lower())
        if column == "name":
            return (row.customer_name or "").lower()
        if column == "issue_date":
            return row.issue_date or datetime.min
        if column == "due_date":
            return row.due_date or datetime.min
        if column == "open_balance":
            return float(row.open_balance or 0)
        if column == "signed_doc":
            return 1 if getattr(row, "has_signed_doc", False) else 0
        if column == "boleto":
            return 1 if getattr(row, "has_boleto", False) else 0
        if column == "nota_fiscal":
            return 1 if getattr(row, "has_nota_fiscal", False) else 0
        if column == "fatura_anexo":
            return 1 if self._has_fatura_attachment(row) else 0
        return ""
    def sort_by(self, column: str):
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False
        self.rows.sort(key=lambda r: self._sort_value(r, column), reverse=self.sort_reverse)
        self._refresh_tree()
        self._update_heading_titles()
    def _update_heading_titles(self):
        labels = {
            "company": "Empresa",
            "account_display": "Conta",
            "invoice": "Nº Fatura",
            "code": "Código",
            "name": "Cliente",
            "issue_date": "Data",
            "due_date": "Vencimento",
            "open_balance": "Valor da fatura",
            "signed_doc": "Assinado",
            "boleto": "Boleto",
            "nota_fiscal": "NF",
            "fatura_anexo": "Fatura",
        }
        for col, label in labels.items():
            suffix = ""
            if self.sort_column == col:
                suffix = " ▼" if self.sort_reverse else " ▲"
            self.tree.heading(col, text=label + suffix, command=lambda c=col: self.sort_by(c))
    def _group_rows(self, rows: List[InvoiceRow]) -> List[InvoiceRow]:
        mode = self.GROUP_OPTIONS.get(self.group_by_var.get(), "none")
        if mode == "none":
            return list(rows)
        grouped = {}
        for row in rows:
            if mode == "customer":
                key = (row.customer_code, row.customer_name)
            elif mode == "due_date":
                key = (row.company, row.due_date)
            elif mode == "account_group":
                key = (row.company, row.account_code, row.account_name)
            else:
                key = (row.invoice_id,)
            if key not in grouped:
                if mode == "customer":
                    grouped[key] = InvoiceRow(
                        invoice_id=f"grp_customer_{row.customer_code}",
                        company=row.company,
                        customer_id=row.customer_id,
                        customer_code=row.customer_code,
                        customer_name=row.customer_name,
                        motive_code="",
                        motive_name="",
                        account_code="",
                        account_name="Várias Contas",
                        issue_date=None,
                        due_date=None,
                        amount=0,
                        discount_amount=0,
                        paid_amount=0,
                        open_balance=0,
                        customer_email=row.customer_email,
                    )
                    grouped[key].original_rows = []
                elif mode == "due_date":
                    grouped[key] = InvoiceRow(
                        invoice_id=f"grp_due_{row.company}_{row.due_date}",
                        company=row.company,
                        customer_id="",
                        customer_code="",
                        customer_name="Agrupado por vencimento",
                        motive_code="",
                        motive_name="",
                        account_code=row.account_code,
                        account_name=row.account_name,
                        issue_date=None,
                        due_date=row.due_date,
                        amount=0,
                        discount_amount=0,
                        paid_amount=0,
                        open_balance=0,
                    )
                elif mode == "account_group":
                    grouped[key] = InvoiceRow(
                        invoice_id=f"grp_account_{row.company}_{row.account_code}",
                        company=row.company,
                        customer_id="",
                        customer_code="",
                        customer_name="Agrupado por conta",
                        motive_code="",
                        motive_name="",
                        account_code=row.account_code,
                        account_name=row.account_name,
                        issue_date=None,
                        due_date=None,
                        amount=0,
                        discount_amount=0,
                        paid_amount=0,
                        open_balance=0,
                    )
            grouped[key].amount = float(grouped[key].amount or 0) + float(row.amount or 0)
            grouped[key].discount_amount = float(grouped[key].discount_amount or 0) + float(row.discount_amount or 0)
            grouped[key].paid_amount = float(grouped[key].paid_amount or 0) + float(row.paid_amount or 0)
            grouped[key].open_balance = float(grouped[key].open_balance or 0) + float(row.open_balance or 0)
            grouped[key].has_signed_doc = bool(getattr(grouped[key], "has_signed_doc", False) or getattr(row, "has_signed_doc", False))
            grouped[key].has_boleto = bool(getattr(grouped[key], "has_boleto", False) or getattr(row, "has_boleto", False))
            grouped[key].has_nota_fiscal = bool(getattr(grouped[key], "has_nota_fiscal", False) or getattr(row, "has_nota_fiscal", False))
            if hasattr(grouped[key], "original_rows"):
                grouped[key].original_rows.append(row)
        return list(grouped.values())
    def load_data(self):
        if self._loading:
            self._pending_reload = True
            return
        try:
            date_from = self._parse_date(self.period_start_var.get())
            date_to = self._parse_date(self.period_end_var.get())
            if date_from and date_to and date_from > date_to:
                raise AppError("O período inicial não pode ser maior que o período final.")
        except Exception as e:
            self.set_status("Falha ao carregar dados.")
            messagebox.showerror(APP_TITLE, f"Erro ao carregar faturas a receber:\n\n{e}", parent=self)
            return

        customer_id = self._selected_customer_id()
        account_code = self._selected_account_code()
        self.set_status("Conectando ao banco e carregando faturas a receber.")
        self._loading = True

        def _work():
            db = Database(self.config_data)
            invoices = db.list_open_invoices(
                due_date_from=date_from,
                due_date_to=date_to,
                customer_id=customer_id,
                account_code=account_code,
            )
            invoice_ids = [r.get("movto_id") for r in (invoices or [])]
            try:
                sig_map = db.get_sale_signatures_pdf_bulk(invoice_ids)
            except Exception:
                sig_map = {}
            try:
                boleto_map = db.check_boleto_exists_bulk(invoice_ids)
            except Exception:
                boleto_map = {}
            try:
                nf_map = db.check_nota_fiscal_exists_bulk(invoice_ids)
            except Exception:
                nf_map = {}
            for r in (invoices or []):
                inv_id = r.get("movto_id")
                sig = sig_map.get(inv_id) or sig_map.get(str(inv_id)) or {}
                r["has_signed_doc"] = bool(sig.get("exists") or (sig.get("attachments") or []))
                r["has_boleto"] = bool(boleto_map.get(inv_id) or boleto_map.get(str(inv_id)))
                r["has_nota_fiscal"] = bool(nf_map.get(inv_id) or nf_map.get(str(inv_id)))
            return invoices

        def _ok(data):
            self._apply_invoices_data(data)
            self._loading = False
            if self._pending_reload:
                self._pending_reload = False
                self.load_data()

        def _err(e: Exception):
            self.set_status("Falha ao carregar dados.")
            messagebox.showerror(APP_TITLE, f"Erro ao carregar faturas a receber:\n\n{e}", parent=self)
            self._loading = False
            if self._pending_reload:
                self._pending_reload = False

        run_with_busy(self, "Carregando faturas...", _work, _ok, _err)
    def _selected_invoice_row(self) -> Optional[InvoiceRow]:
        selected = self.tree.selection()
        if not selected:
            return None
        return self.tree_items.get(selected[0])

    def _selected_invoice_rows(self) -> List[InvoiceRow]:
        selected = self.tree.selection() or []
        rows: List[InvoiceRow] = []
        for item_id in selected:
            r = self.tree_items.get(item_id)
            if r is not None:
                rows.append(r)
        return rows

    def _select_all_tree_rows(self):
        try:
            items = self.tree.get_children()
            if items:
                self.tree.selection_set(items)
        except Exception:
            pass

    def _clear_tree_selection(self):
        try:
            self.tree.selection_remove(self.tree.selection())
        except Exception:
            pass

    def _send_emails_queue(self, invoices: List[InvoiceRow]):
        invoices = [i for i in (invoices or []) if getattr(i, "invoice_id", None) not in (None, "", 0, "0") and not str(getattr(i, "invoice_id", "")).startswith("grp_")]
        if not invoices:
            messagebox.showwarning(APP_TITLE, "Nenhuma fatura válida selecionada para envio.", parent=self)
            return

        smtp_cfg = self.config_data.get("smtp", {})
        smtp_email = str(smtp_cfg.get("email", "")).strip()
        smtp_host = str(smtp_cfg.get("host", "")).strip()
        smtp_password = str(smtp_cfg.get("password", "")).strip()
        smtp_port = int(smtp_cfg.get("port", 587) or 587)
        if not smtp_email or not smtp_host or not smtp_password or not smtp_port:
            messagebox.showwarning(APP_TITLE, "Configure o SMTP na tela de Configuração antes de enviar e-mails.", parent=self)
            return

        if not messagebox.askyesno(APP_TITLE, f"Deseja enviar e-mails em fila para {len(invoices)} fatura(s) selecionada(s)?", parent=self):
            return

        def _work():
            def _send_smtp_message(msg: EmailMessage):
                if smtp_port == 465:
                    context = ssl.create_default_context()
                    with smtplib.SMTP_SSL(smtp_host, smtp_port, context=context, timeout=20) as server:
                        server.login(smtp_email, smtp_password)
                        server.send_message(msg)
                else:
                    with smtplib.SMTP(smtp_host, smtp_port, timeout=20) as server:
                        server.ehlo()
                        try:
                            server.starttls(context=ssl.create_default_context())
                            server.ehlo()
                        except Exception:
                            pass
                        server.login(smtp_email, smtp_password)
                        server.send_message(msg)

            def _set_progress(
                done: int,
                total: int,
                customer_name: str = "",
                to_email: str = "",
                detail: str = "",
                sent: int = 0,
                failed_count: int = 0,
                skipped: int = 0,
            ):
                try:
                    done = int(done or 0)
                except Exception:
                    done = 0
                try:
                    total = int(total or 0)
                except Exception:
                    total = 0
                done = max(0, done)
                if total > 0:
                    done = min(done, total)
                try:
                    sent = int(sent or 0)
                except Exception:
                    sent = 0
                try:
                    failed_count = int(failed_count or 0)
                except Exception:
                    failed_count = 0
                try:
                    skipped = int(skipped or 0)
                except Exception:
                    skipped = 0

                def _do(
                    done=done,
                    total=total,
                    customer_name=customer_name,
                    to_email=to_email,
                    detail=detail,
                    sent=sent,
                    failed_count=failed_count,
                    skipped=skipped,
                ):
                    overlay = _ensure_busy_overlay(self)
                    if total > 0:
                        current_label = total if done >= total else min(total, done + 1)
                        msg = f"Enviando {current_label} de {total}"
                    else:
                        msg = "Enviando..."
                    overlay.set_message(msg)
                    extra = []
                    if customer_name or to_email:
                        extra.append(f"{customer_name} <{to_email}>".strip())
                    if detail:
                        extra.append(detail)
                    extra.append(f"Enviados: {sent} | Falhas: {failed_count} | Sem e-mail: {skipped}")
                    overlay.set_detail(" | ".join([e for e in extra if e]))
                    overlay.set_progress(done, total if total > 0 else None)

                try:
                    self.after(0, _do)
                except Exception:
                    pass

            emails_sent = 0
            skipped_no_email = 0
            failed = 0
            missing_total = 0
            try:
                delay_seconds = int(smtp_cfg.get("delay_seconds", 5) or 0)
            except Exception:
                delay_seconds = 5
            delay_seconds = max(0, min(300, delay_seconds))
            first_email = True
            db = Database(self.config_data)

            invoice_ids = [i.invoice_id for i in invoices]
            boleto_map: Dict[Any, Dict[str, Any]] = {}
            try:
                boleto_map = db.get_boletos_email_payload_bulk(invoice_ids)
            except Exception:
                boleto_map = {}
            signature_map: Dict[Any, Dict[str, Any]] = {}
            try:
                signature_map = db.get_sale_signatures_pdf_bulk(invoice_ids)
            except Exception:
                signature_map = {}
            nfe_map: Dict[Any, Dict[str, Any]] = {}
            try:
                nfe_map = db.get_nfe_attachments_bulk(invoice_ids, only_invoice_mlid=True, max_nfes_per_invoice=1)
                missing_nfe_ids = []
                for inv_id in (invoice_ids or []):
                    current = (nfe_map or {}).get(inv_id) or (nfe_map or {}).get(str(inv_id)) or {}
                    cur_atts = list((current.get("attachments") or []))
                    if not bool([a for a in cur_atts if a.get("data") and a.get("filename")]):
                        missing_nfe_ids.append(inv_id)
                if missing_nfe_ids:
                    nfe_map2 = db.get_nfe_attachments_bulk(missing_nfe_ids, only_invoice_mlid=False, max_nfes_per_invoice=1)
                    for inv_id in missing_nfe_ids:
                        candidate = (nfe_map2 or {}).get(inv_id) or (nfe_map2 or {}).get(str(inv_id)) or {}
                        cand_atts = list((candidate.get("attachments") or []))
                        if bool([a for a in cand_atts if a.get("data") and a.get("filename")]):
                            nfe_map[inv_id] = candidate
                            nfe_map[str(inv_id)] = candidate
            except Exception:
                nfe_map = {}
            purchase_map: Dict[Any, Dict[str, Any]] = {}
            try:
                purchase_map = db.get_purchase_info_bulk(invoice_ids)
            except Exception:
                purchase_map = {}

            total = len(invoices)
            _set_progress(0, total, detail="Iniciando...", sent=0, failed_count=0, skipped=0)
            for idx, inv in enumerate(invoices, start=1):
                done_before = int(emails_sent or 0) + int(failed or 0) + int(skipped_no_email or 0)
                to_email = str(getattr(inv, "customer_email", "") or "").strip()
                if not to_email and getattr(inv, "customer_id", None) not in (None, "", 0, "0"):
                    try:
                        to_email = str(db.get_customer_email(getattr(inv, "customer_id", None)) or "").strip()
                    except Exception:
                        to_email = ""
                if not to_email:
                    skipped_no_email += 1
                    done_after = int(emails_sent or 0) + int(failed or 0) + int(skipped_no_email or 0)
                    _set_progress(
                        done_after,
                        total,
                        inv.customer_name,
                        "",
                        f"Sem e-mail (fatura {inv.invoice_id})",
                        sent=emails_sent,
                        failed_count=failed,
                        skipped=skipped_no_email,
                    )
                    continue

                _set_progress(
                    done_before,
                    total,
                    inv.customer_name,
                    to_email,
                    f"Preparando anexos (fatura {inv.invoice_id})",
                    sent=emails_sent,
                    failed_count=failed,
                    skipped=skipped_no_email,
                )

                attachments: List[Tuple[bytes, str]] = []
                missing = 0
                boleto_data = boleto_map.get(inv.invoice_id) or boleto_map.get(str(inv.invoice_id)) or {}
                try:
                    if boleto_data.get("exists"):
                        filename = boleto_data.get("filename") or f"boleto_{inv.invoice_id}.pdf"
                        try:
                            pdf_bytes = build_boleto_pdf_bytes(boleto_data, inv, include_pix_qrcode=False)
                        except Exception:
                            pdf_bytes = None
                        if pdf_bytes:
                            attachments.append((pdf_bytes, filename))
                        else:
                            missing += 1
                    else:
                        missing += 1
                except Exception:
                    missing += 1
                sig = signature_map.get(inv.invoice_id) or signature_map.get(str(inv.invoice_id)) or {}
                signature_files: List[Tuple[bytes, str]] = []
                for a in (sig.get("attachments") or []):
                    sdata = a.get("data")
                    sname = a.get("filename")
                    if sdata and sname:
                        signature_files.append((sdata, sname))
                sig_bytes = sig.get("attachment_data")
                if not signature_files and sig.get("exists") and sig_bytes:
                    signature_files.append((sig_bytes, sig.get("filename") or f"assinatura_{inv.invoice_id}"))
                if signature_files:
                    inv_part = re.sub(r"[^0-9A-Za-z_-]+", "_", str(getattr(inv, "invoice_id", "") or "").strip())[:40]
                    zip_name = f"assinaturas_{inv_part}.zip" if inv_part else "assinaturas.zip"
                    zip_bytes, _ = zip_named_files(signature_files, zip_filename=zip_name)
                    attachments.append((zip_bytes, zip_name))

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
                    extra_inf = ""
                    try:
                        extra_inf = Database(self.config_data).get_placa_km_text_bulk([inv.invoice_id]).get(inv.invoice_id) or ""
                    except Exception:
                        extra_inf = ""
                    pdf_bytes, pdf_name = danfe_pdf_from_nfe_xml(
                        ndata,
                        fallback_suffix=str(inv.invoice_id),
                        extra_inf_cpl_text=extra_inf,
                    )
                    if pdf_bytes and pdf_name:
                        nfe_atts.append({"data": pdf_bytes, "filename": pdf_name, "mime_type": "application/pdf"})
                        has_pdf = True
                for a in nfe_atts:
                    ndata = a.get("data")
                    nname = a.get("filename")
                    if ndata and nname:
                        attachments.append((ndata, nname))

                fatura_txt: Optional[Tuple[bytes, str]] = None
                try:
                    txt_bytes, txt_name = build_faturas_detalhamento_txt_bytes([inv], purchase_info_map=purchase_map)
                    if txt_bytes and txt_name:
                        fatura_txt = (txt_bytes, txt_name)
                except Exception:
                    fatura_txt = None

                missing_total += missing

                subject = f"Fatura a receber - {inv.customer_name} - vencimento {inv.due_date_display()}"
                flag_names = [name for data, name in attachments if data]
                if fatura_txt:
                    flag_names.append(fatura_txt[1])
                flags = _detect_email_attachment_flags(flag_names)
                text_body, html_body = build_agenda_email_body(
                    inv.customer_name,
                    inv.due_date_display(),
                    [inv],
                    missing,
                    "",
                    context_label="Vencimento",
                    purchase_info_map=purchase_map,
                    attachment_flags=flags,
                )

                msg = EmailMessage()
                msg["From"] = format_smtp_from(smtp_cfg) or smtp_email
                msg["To"] = to_email
                msg["Subject"] = subject
                msg.set_content(text_body)
                msg.add_alternative(html_body, subtype="html")
                try:
                    attach_email_logo(msg)
                except Exception:
                    pass
                if fatura_txt:
                    maintype, subtype = _mime_parts_from_filename(fatura_txt[1])
                    msg.add_attachment(fatura_txt[0], maintype=maintype, subtype=subtype, filename=fatura_txt[1])
                for data, name in attachments:
                    if not data or not name:
                        continue
                    maintype, subtype = _mime_parts_from_filename(name)
                    msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=name)

                try:
                    if not first_email and delay_seconds > 0:
                        time_module.sleep(delay_seconds)
                    first_email = False
                    _send_smtp_message(msg)
                    emails_sent += 1
                    done_after = int(emails_sent or 0) + int(failed or 0) + int(skipped_no_email or 0)
                    _set_progress(
                        done_after,
                        total,
                        inv.customer_name,
                        to_email,
                        f"Enviado (fatura {inv.invoice_id})",
                        sent=emails_sent,
                        failed_count=failed,
                        skipped=skipped_no_email,
                    )
                    AuditLogger.write(
                        self.current_user,
                        "enviar_email_fatura_fila",
                        f"cliente={inv.customer_name};para={to_email};invoice={inv.invoice_id};anexos={len(attachments) + (1 if fatura_txt else 0)}"
                    )
                except Exception as e:
                    failed += 1
                    done_after = int(emails_sent or 0) + int(failed or 0) + int(skipped_no_email or 0)
                    _set_progress(
                        done_after,
                        total,
                        inv.customer_name,
                        to_email,
                        f"Falha (fatura {inv.invoice_id})",
                        sent=emails_sent,
                        failed_count=failed,
                        skipped=skipped_no_email,
                    )
                    AuditLogger.write(self.current_user, "enviar_email_fatura_fila_erro", f"cliente={inv.customer_name};para={to_email};invoice={inv.invoice_id};erro={e}")

            try:
                from app_core.documents_history import DocumentsHistory

                level = "info"
                if int(failed or 0) > 0:
                    level = "error"
                elif int(skipped_no_email or 0) > 0:
                    level = "warn"
                DocumentsHistory().add_event(
                    kind="manual_send_queue",
                    source="faturas_receber",
                    level=level,
                    title="Envio manual em fila (Faturas a receber)",
                    message=f"Selecionadas: {len(invoices)} | Enviados: {int(emails_sent or 0)} | Sem e-mail: {int(skipped_no_email or 0)} | Falhas: {int(failed or 0)}",
                )
            except Exception:
                pass
            return {"emails_sent": emails_sent, "skipped_no_email": skipped_no_email, "failed": failed, "missing_total": missing_total}

        def _ok(result):
            result = result or {}
            sent = int(result.get("emails_sent") or 0)
            skipped = int(result.get("skipped_no_email") or 0)
            failed = int(result.get("failed") or 0)
            messagebox.showinfo(APP_TITLE, f"Envio concluído.\n\nEnviados: {sent}\nSem e-mail: {skipped}\nFalhas: {failed}", parent=self)

        def _err(e: Exception):
            messagebox.showerror(APP_TITLE, f"Erro no envio:\n\n{e}", parent=self)

        run_with_busy(self, "Enviando e-mails...", _work, _ok, _err)

    def open_email_window(self):
        selected_rows = self._selected_invoice_rows()
        if not selected_rows:
            messagebox.showwarning(APP_TITLE, "Selecione uma ou mais faturas na grade.", parent=self)
            return

        if len(selected_rows) > 1:
            expanded: List[InvoiceRow] = []
            for r in selected_rows:
                is_grouped_row = self.GROUP_OPTIONS.get(self.group_by_var.get()) != "none" or str(r.invoice_id).startswith("grp_")
                if is_grouped_row and getattr(r, "original_rows", []):
                    expanded.extend(list(r.original_rows))
                else:
                    expanded.append(r)
            unique: List[InvoiceRow] = []
            seen = set()
            for r in expanded:
                key = getattr(r, "invoice_id", None)
                if key in seen:
                    continue
                seen.add(key)
                unique.append(r)
            self._send_emails_queue(unique)
            return

        row = selected_rows[0]

        is_grouped = self.GROUP_OPTIONS.get(self.group_by_var.get()) != "none" or str(row.invoice_id).startswith("grp_")
        
        if is_grouped:
            if not getattr(row, "original_rows", []):
                messagebox.showwarning(APP_TITLE, "Não há faturas detalhadas para este agrupamento.", parent=self)
                return
            target_rows = row.original_rows
        else:
            target_rows = [row]

        def _open_for_row(idx: int):
            if idx >= len(target_rows):
                return
            r = target_rows[idx]
            if not r.customer_id:
                messagebox.showwarning(APP_TITLE, f"Não foi possível identificar o cliente da fatura {r.invoice_id}.", parent=self)
                _open_for_row(idx + 1)
                return
            
            email = str(r.customer_email or "").strip()
            if email:
                win = EmailComposeWindow(self, self.config_data, self.current_user, r, email)
                self.wait_window(win)
                _open_for_row(idx + 1)
                return

            def _work():
                return Database(self.config_data).get_customer_email(r.customer_id)

            def _ok(found_email):
                win = EmailComposeWindow(self, self.config_data, self.current_user, r, str(found_email or "").strip())
                self.wait_window(win)
                _open_for_row(idx + 1)

            def _err(e: Exception):
                messagebox.showerror(APP_TITLE, f"Erro ao buscar o e-mail do cliente para fatura {r.invoice_id}:\n\n{e}", parent=self)
                _open_for_row(idx + 1)

            run_with_busy(self, f"Buscando e-mail... ({idx+1}/{len(target_rows)})", _work, _ok, _err)

        if is_grouped:
            if not messagebox.askyesno(APP_TITLE, f"Esta linha agrupada contém {len(target_rows)} fatura(s).\nDeseja enviar e-mails separadamente para cada uma?", parent=self):
                return
                
        _open_for_row(0)

    def _has_fatura_attachment(self, row: InvoiceRow) -> bool:
        inv_id = getattr(row, "invoice_id", None)
        if inv_id in (None, "", 0, "0"):
            return False
        inv_str = str(inv_id)
        if inv_str.startswith("grp_"):
            return bool(getattr(row, "original_rows", []))
        return True

    def _row_values(self, row: InvoiceRow):
        invoice_id = getattr(row, "invoice_id", None)
        invoice_display = "" if invoice_id in (None, "") else str(invoice_id)
        if invoice_display.startswith("grp_"):
            invoice_display = ""
        return (
            row.company,
            (f"{row.account_code or ''} - {row.account_name or ''}").strip(" -"),
            invoice_display,
            row.customer_code,
            row.customer_name,
            row.issue_date_display(),
            row.due_date_display(),
            row.open_balance_display(),
            "Sim" if getattr(row, "has_signed_doc", False) else "Não",
            "Sim" if getattr(row, "has_boleto", False) else "Não",
            "Sim" if getattr(row, "has_nota_fiscal", False) else "Não",
            "Sim" if self._has_fatura_attachment(row) else "Não",
        )
    def _refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree_items.clear()
        for row in self.rows:
            item_id = self.tree.insert("", "end", values=self._row_values(row))
            self.tree_items[item_id] = row


class FinanceiroAlertasWindow(tk.Toplevel):
    def __init__(self, master, config_data: Dict[str, Any], current_user: str, on_config_saved):
        super().__init__(master)
        self.master_app = master
        self.config_data = deepcopy(config_data)
        self.current_user = current_user
        self.on_config_saved = on_config_saved
        self.status_var = tk.StringVar(value="Pronto.")
        self.selected_id = None
        self.title(f"{APP_TITLE} - Alertas de vencimento")
        self.geometry("1400x680")
        self.minsize(1000, 620)
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._build_ui()
        self._center_window()
        self.reload()

    def _center_window(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - self.winfo_width()) // 2
        y = (sh - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 20)}+{max(y, 20)}")

    def _build_ui(self):
        header = ttk.Frame(self, padding=(16, 12, 16, 0))
        header.pack(fill="x")
        ttk.Label(header, text="Alertas de Vencimento", font=("Segoe UI", 14, "bold"), foreground="#2563eb").pack(side="left")
        ttk.Label(header, text="Envio automático por vencimento do boleto (apenas com boletos vinculados).", font=("Segoe UI", 10), foreground="#6b7280").pack(side="left", padx=(12, 0), pady=(4, 0))

        top = ttk.Frame(self, padding=(12, 10, 12, 10))
        top.pack(fill="x")
        ttk.Button(top, text="+ Novo Alerta", command=self._new).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Editar", command=self._edit).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Detalhes do último envio", command=self._details).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Excluir", command=self._delete).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Atualizar", command=self.reload).pack(side="left", padx=(12, 0))
        ttk.Button(top, text="Voltar ao início", command=self._close).pack(side="right")

        mid = ttk.Frame(self, padding=(12, 0, 12, 10))
        mid.pack(fill="both", expand=True)
        cols = (
            "nome",
            "ativo",
            "hora",
            "antes",
            "depois",
            "grupo",
            "portador",
            "ultimo_envio",
            "vencimentos",
            "planejados",
            "enviados",
            "sem_email",
            "falhas",
            "atraso",
            "fora_horario",
        )
        self.tree = ttk.Treeview(mid, columns=cols, show="headings")
        self.tree.heading("nome", text="Nome")
        self.tree.heading("ativo", text="Ativo")
        self.tree.heading("hora", text="Hora")
        self.tree.heading("antes", text="Antes (dias)")
        self.tree.heading("depois", text="Depois (dias)")
        self.tree.heading("grupo", text="Grupo")
        self.tree.heading("portador", text="Portador")
        self.tree.heading("ultimo_envio", text="Último envio")
        self.tree.heading("vencimentos", text="Vencimentos (base)")
        self.tree.heading("planejados", text="Planejados")
        self.tree.heading("enviados", text="Enviados")
        self.tree.heading("sem_email", text="Sem e-mail")
        self.tree.heading("falhas", text="Falhas")
        self.tree.heading("atraso", text="Atraso (min)")
        self.tree.heading("fora_horario", text="Fora do horário")
        self.tree.column("nome", width=260, anchor="w")
        self.tree.column("ativo", width=70, anchor="center")
        self.tree.column("hora", width=70, anchor="center")
        self.tree.column("antes", width=90, anchor="center")
        self.tree.column("depois", width=95, anchor="center")
        self.tree.column("grupo", width=200, anchor="w")
        self.tree.column("portador", width=200, anchor="w")
        self.tree.column("ultimo_envio", width=140, anchor="center")
        self.tree.column("vencimentos", width=170, anchor="center")
        self.tree.column("planejados", width=90, anchor="center")
        self.tree.column("enviados", width=80, anchor="center")
        self.tree.column("sem_email", width=80, anchor="center")
        self.tree.column("falhas", width=70, anchor="center")
        self.tree.column("atraso", width=90, anchor="center")
        self.tree.column("fora_horario", width=110, anchor="center")
        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(mid, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        mid.columnconfigure(0, weight=1)
        mid.rowconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", lambda e: self._edit())

        bottom = ttk.Frame(self, padding=(12, 0, 12, 10))
        bottom.pack(fill="x")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left")
        ttk.Label(bottom, text=f"Usuário: {self.current_user}").pack(side="right")

    def _agenda_rows(self):
        agendas = self.config_data.get("financeiro_agendas", []) or []
        return [a for a in agendas if isinstance(a, dict)]

    def reload(self):
        self.status_var.set("Carregando alertas...")
        agendas = self._agenda_rows()

        def _work():
            groups_map: Dict[str, str] = {}
            port_map: Dict[str, str] = {}
            try:
                db = Database(self.config_data)
                for g in db.list_grupos_pessoa():
                    groups_map[str(g.get("grid"))] = str(g.get("nome") or "").strip()
                for p in db.list_portadores():
                    port_map[str(p.get("grid"))] = str(p.get("nome") or "").strip()
            except Exception:
                groups_map = {}
                port_map = {}
            return groups_map, port_map

        def _ok(res):
            groups_map, port_map = res
            try:
                for it in self.tree.get_children():
                    self.tree.delete(it)

                agendas.sort(key=lambda a: (str(a.get("name") or "").lower(), str(a.get("id") or "")))
                for a in agendas:
                    iid = str(a.get("id") or "")
                    gid = a.get("group_id")
                    pid = a.get("portador_id")
                    group_name = groups_map.get(str(gid), "") if gid not in (None, "", 0, "0") else ""
                    port_name = port_map.get(str(pid), "") if pid not in (None, "", 0, "0") else ""
                    before_days = int(a.get("days_before_due") or 0)
                    after_days = int(a.get("days_after_due") or 0)
                    last_run_at = str(a.get("last_run_at") or "").strip()
                    last_run_txt = ""
                    if last_run_at:
                        try:
                            last_run_txt = datetime.fromisoformat(last_run_at).strftime("%d/%m/%Y %H:%M")
                        except Exception:
                            last_run_txt = last_run_at
                    last_due_date = str(a.get("last_due_date") or "").strip()
                    try:
                        late_minutes = int(a.get("last_late_minutes") or 0)
                    except Exception:
                        late_minutes = 0
                    out_of_time = bool(a.get("last_out_of_time"))
                    last_result = a.get("last_result") or {}
                    enviados = int(last_result.get("emails_sent") or 0)
                    sem_email = int(last_result.get("skipped_no_email") or 0)
                    falhas = int(last_result.get("failed") or 0)
                    planejados = int(last_result.get("emails_planned") or 0)
                    self.tree.insert(
                        "",
                        "end",
                        iid=iid,
                        values=(
                            str(a.get("name") or "").strip(),
                            "Sim" if a.get("enabled") else "Não",
                            str(a.get("send_time") or "06:00").strip(),
                            before_days,
                            after_days,
                            group_name,
                            port_name,
                            last_run_txt,
                            last_due_date,
                            planejados,
                            enviados,
                            sem_email,
                            falhas,
                            late_minutes,
                            "Sim" if out_of_time else "Não",
                        ),
                    )
                self.status_var.set(f"{len(agendas)} alerta(s).")
            except Exception as e:
                self.status_var.set("Falha ao carregar alertas.")
                messagebox.showerror(APP_TITLE, f"Erro ao carregar alertas:\n\n{e}", parent=self)

        def _err(e: Exception):
            self.status_var.set("Falha ao carregar alertas.")
            messagebox.showerror(APP_TITLE, f"Erro ao carregar alertas:\n\n{e}", parent=self)

        run_with_busy(self, "Carregando alertas...", _work, _ok, _err)

    def _on_select(self, event=None):
        sel = self.tree.selection()
        self.selected_id = sel[0] if sel else None

    def _new(self):
        FinanceiroAlertaWindow(self, self.config_data, self.current_user, self._after_saved, agenda_id=None)

    def _edit(self):
        if not self.selected_id:
            messagebox.showwarning(APP_TITLE, "Selecione um alerta.", parent=self)
            return
        FinanceiroAlertaWindow(self, self.config_data, self.current_user, self._after_saved, agenda_id=self.selected_id)

    def _details(self):
        if not self.selected_id:
            messagebox.showwarning(APP_TITLE, "Selecione um alerta.", parent=self)
            return
        agenda = None
        for a in self._agenda_rows():
            if str(a.get("id") or "") == str(self.selected_id):
                agenda = a
                break
        if not agenda:
            messagebox.showwarning(APP_TITLE, "Não foi possível carregar o alerta selecionado.", parent=self)
            return
        FinanceiroAlertaLastRunDetailsWindow(self, agenda=agenda)

    def _delete(self):
        if not self.selected_id:
            messagebox.showwarning(APP_TITLE, "Selecione um alerta.", parent=self)
            return
        if not messagebox.askyesno(APP_TITLE, "Deseja excluir o alerta selecionado?", parent=self):
            return
        agendas = self._agenda_rows()
        agendas = [a for a in agendas if str(a.get("id") or "") != str(self.selected_id)]
        self.config_data["financeiro_agendas"] = agendas
        ConfigManager.save(self.config_data)
        self.config_data = ConfigManager.load()
        self.on_config_saved(self.config_data)
        self.selected_id = None
        self.reload()

    def _after_saved(self, cfg: Dict[str, Any]):
        self.config_data = cfg
        self.on_config_saved(cfg)
        self.reload()

    def _close(self):
        self.destroy()
        if hasattr(self.master_app, "alerts_window"):
            self.master_app.alerts_window = None
        if hasattr(self.master_app, "show_home"):
            self.master_app.show_home()


class FinanceiroAlertaLastRunDetailsWindow(tk.Toplevel):
    def __init__(self, parent: tk.Toplevel, *, agenda: Dict[str, Any]):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()
        self.agenda = dict(agenda or {})
        self.last_result = self.agenda.get("last_result") if isinstance(self.agenda.get("last_result"), dict) else {}
        self.rows_cache: List[Dict[str, Any]] = []
        self.item_map: Dict[str, Dict[str, Any]] = {}

        self.title(f"{APP_TITLE} - Detalhes do último envio")
        self.geometry("1200x720")
        self.minsize(980, 560)
        self._build_ui()
        self._center_window()
        self._load_rows()

    def _center_window(self):
        try:
            self.update_idletasks()
        except Exception:
            pass
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - self.winfo_width()) // 2
        y = (sh - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 20)}+{max(y, 20)}")

    def _format_dt(self, value: str) -> str:
        raw = str(value or "").strip()
        if not raw:
            return ""
        try:
            return datetime.fromisoformat(raw).strftime("%d/%m/%Y %H:%M")
        except Exception:
            return raw

    def _money_br(self, value: Any) -> str:
        if value in (None, ""):
            return "0,00"
        try:
            num = float(value)
            return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return str(value)

    def _flags_label(self, flags: Dict[str, Any]) -> str:
        if not isinstance(flags, dict):
            flags = {}
        order = [
            ("boleto", "boleto"),
            ("danfe", "danfe"),
            ("xml", "xml"),
            ("assinatura", "assinaturas"),
            ("fatura_txt", "fatura"),
            ("fatura_pdf", "fatura (PDF)"),
        ]
        labels = []
        for k, label in order:
            try:
                if bool(flags.get(k)):
                    labels.append(label)
            except Exception:
                pass
        return ", ".join(labels)

    def _invoice_ids_label(self, invoices: List[Dict[str, Any]]) -> str:
        ids = []
        for inv in invoices or []:
            iid = inv.get("invoice_id")
            if iid in (None, "", 0, "0"):
                continue
            ids.append(str(iid))
        if not ids:
            return ""
        extra = ""
        if len(ids) > 8:
            extra = f" (+{len(ids) - 8})"
            ids = ids[:8]
        return ", ".join(ids) + extra

    def _build_ui(self):
        header = ttk.Frame(self, padding=(16, 12, 16, 0))
        header.pack(fill="x")
        ttk.Label(header, text="Detalhes do último envio", font=("Segoe UI", 14, "bold"), foreground="#111827").pack(side="left")
        ttk.Button(header, text="Copiar", command=self._copy).pack(side="right")
        ttk.Button(header, text="Fechar", command=self.destroy).pack(side="right", padx=(0, 8))

        info = ttk.Frame(self, padding=(16, 10, 16, 8))
        info.pack(fill="x")
        info.columnconfigure(1, weight=1)
        info.columnconfigure(3, weight=1)

        name = str(self.agenda.get("name") or "").strip() or str(self.agenda.get("id") or "").strip() or "Alerta"
        last_run_at = self._format_dt(str(self.agenda.get("last_run_at") or "").strip())
        due_dates = str(self.agenda.get("last_due_date") or "").strip()
        try:
            late_minutes = int(self.agenda.get("last_late_minutes") or 0)
        except Exception:
            late_minutes = 0
        out_of_time = bool(self.agenda.get("last_out_of_time"))

        planned = int(self.last_result.get("emails_planned") or 0)
        sent = int(self.last_result.get("emails_sent") or 0)
        skipped = int(self.last_result.get("skipped_no_email") or 0)
        failed = int(self.last_result.get("failed") or 0)
        atts_total = int(self.last_result.get("attachments_total") or 0)
        missing_total = int(self.last_result.get("missing_total") or 0)

        ttk.Label(info, text="Alerta").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=2)
        ttk.Label(info, text=name).grid(row=0, column=1, sticky="w", pady=2)
        ttk.Label(info, text="Último envio").grid(row=0, column=2, sticky="w", padx=(20, 10), pady=2)
        ttk.Label(info, text=last_run_at or "-").grid(row=0, column=3, sticky="w", pady=2)

        ttk.Label(info, text="Vencimentos (base)").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=2)
        ttk.Label(info, text=due_dates or "-").grid(row=1, column=1, sticky="w", pady=2)
        ttk.Label(info, text="Atraso / Fora do horário").grid(row=1, column=2, sticky="w", padx=(20, 10), pady=2)
        ttk.Label(info, text=f"{late_minutes} min / {'Sim' if out_of_time else 'Não'}").grid(row=1, column=3, sticky="w", pady=2)

        ttk.Label(info, text="Resumo").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=(6, 2))
        ttk.Label(
            info,
            text=f"Planejados: {planned} | Enviados: {sent} | Sem e-mail: {skipped} | Falhas: {failed} | Anexos: {atts_total} | Boletos faltando: {missing_total}",
        ).grid(row=2, column=1, columnspan=3, sticky="w", pady=(6, 2))

        body = ttk.Frame(self, padding=(12, 0, 12, 12))
        body.pack(fill="both", expand=True)
        body.rowconfigure(0, weight=2)
        body.rowconfigure(1, weight=1)
        body.columnconfigure(0, weight=1)

        cols = ("status", "cliente", "para", "faturas", "docs", "msg")
        self.tree = ttk.Treeview(body, columns=cols, show="headings")
        self.tree.heading("status", text="Status")
        self.tree.heading("cliente", text="Cliente")
        self.tree.heading("para", text="Para")
        self.tree.heading("faturas", text="Faturas")
        self.tree.heading("docs", text="Documentos")
        self.tree.heading("msg", text="Mensagem")
        self.tree.column("status", width=110, anchor="center")
        self.tree.column("cliente", width=260, anchor="w")
        self.tree.column("para", width=240, anchor="w")
        self.tree.column("faturas", width=200, anchor="w")
        self.tree.column("docs", width=180, anchor="w")
        self.tree.column("msg", width=250, anchor="w")
        vsb = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self.detail_text = tk.Text(body, wrap="word", height=10)
        self.detail_text.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(10, 0))
        try:
            self.detail_text.configure(state="disabled")
        except Exception:
            pass

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

    def _load_rows(self):
        details = self.last_result.get("details") if isinstance(self.last_result.get("details"), list) else []
        truncated = bool(self.last_result.get("details_truncated"))

        rows: List[Dict[str, Any]] = []
        for d in details or []:
            if not isinstance(d, dict):
                continue
            emails = d.get("emails") if isinstance(d.get("emails"), list) else []
            if emails:
                for e in emails:
                    if not isinstance(e, dict):
                        continue
                    rows.append({"target": d, "email": e})
            else:
                rows.append({"target": d, "email": {}})
            if len(rows) >= 500:
                truncated = True
                break

        if not rows and not details:
            rows = [{"target": {"status": "no_details", "customer_name": "", "to_email": "", "invoices": []}, "email": {}}]

        self.rows_cache = rows

        for it in self.tree.get_children():
            self.tree.delete(it)
        self.item_map.clear()

        def _status_label(value: str) -> str:
            v = str(value or "").strip().lower()
            if v == "sent":
                return "Enviado"
            if v == "failed":
                return "Falha"
            if v == "skipped_no_email":
                return "Sem e-mail"
            if v == "planned":
                return "Planejado"
            if v == "no_details":
                return "-"
            return v or "-"

        for row in self.rows_cache:
            t = row.get("target") or {}
            e = row.get("email") or {}
            status = str(t.get("status") or "").strip()
            customer = str(t.get("customer_name") or "").strip()
            to_email = str(t.get("to_email") or "").strip()
            invoices = t.get("invoices") if isinstance(t.get("invoices"), list) else []
            msg = str(e.get("error") or "").strip()
            if not msg and status == "no_details":
                msg = "Este alerta ainda não registrou detalhes. Rode o agendamento novamente para capturar."
            flags = e.get("flags") if isinstance(e.get("flags"), dict) else {}
            docs = self._flags_label(flags)
            faturas = self._invoice_ids_label(invoices)
            item_id = self.tree.insert(
                "",
                "end",
                values=(
                    _status_label(status),
                    customer,
                    to_email,
                    faturas,
                    docs,
                    msg,
                ),
            )
            self.item_map[item_id] = row

        if truncated:
            self._set_detail_text("Detalhes truncados (muitos destinatários/batches). Use os logs para auditoria completa.")
        else:
            self._set_detail_text("")

    def _set_detail_text(self, text: str):
        try:
            self.detail_text.configure(state="normal")
        except Exception:
            pass
        try:
            self.detail_text.delete("1.0", "end")
            if text:
                self.detail_text.insert("1.0", text)
        except Exception:
            pass
        try:
            self.detail_text.configure(state="disabled")
        except Exception:
            pass

    def _on_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        row = self.item_map.get(sel[0]) or {}
        t = row.get("target") or {}
        e = row.get("email") or {}

        customer = str(t.get("customer_name") or "").strip()
        to_email = str(t.get("to_email") or "").strip()
        status = str(t.get("status") or "").strip()
        batch = e.get("batch")
        batches_total = e.get("batches_total")
        subject = str(e.get("subject") or "").strip()
        error = str(e.get("error") or "").strip()
        missing = int(t.get("missing_count") or 0)

        lines: List[str] = []
        if customer:
            lines.append(f"Cliente: {customer}")
        if to_email:
            lines.append(f"Para: {to_email}")
        if status:
            lines.append(f"Status: {status}")
        if batch and batches_total:
            lines.append(f"E-mail: {batch}/{batches_total}")
        if subject:
            lines.append(f"Assunto: {subject}")
        if missing:
            lines.append(f"Boletos sem anexo: {missing}")
        if error:
            lines.append(f"Erro: {error}")

        attachments = e.get("attachments") if isinstance(e.get("attachments"), list) else []
        if attachments:
            lines.append("")
            lines.append("Anexos:")
            for name in attachments[:200]:
                n = str(name or "").strip()
                if n:
                    lines.append(f"- {n}")
            if bool(e.get("attachments_truncated")):
                lines.append("- (lista truncada)")

        invoices = t.get("invoices") if isinstance(t.get("invoices"), list) else []
        if invoices:
            lines.append("")
            lines.append("Faturas / documentos vinculados:")
            for inv in invoices[:200]:
                if not isinstance(inv, dict):
                    continue
                iid = inv.get("invoice_id")
                issue = str(inv.get("issue_date") or "").strip()
                due = str(inv.get("due_date") or "").strip()
                bal = self._money_br(inv.get("open_balance"))
                docs = inv.get("docs") if isinstance(inv.get("docs"), dict) else {}
                parts = [f"- {iid}"]
                if issue:
                    parts.append(f"Emissão: {issue}")
                if due:
                    parts.append(f"Venc.: {due}")
                parts.append(f"Aberto: {bal}")
                inv_docs = []
                if docs:
                    if bool(docs.get("boleto")):
                        inv_docs.append("boleto")
                    if bool(docs.get("danfe")):
                        inv_docs.append("danfe")
                    if bool(docs.get("xml")):
                        inv_docs.append("xml")
                    if bool(docs.get("assinatura")):
                        inv_docs.append("assinatura")
                if inv_docs:
                    parts.append(f"Docs: {', '.join(inv_docs)}")
                lines.append(" | ".join(parts))
                boleto_fn = str(docs.get("boleto_filename") or "").strip()
                boleto_doc = str(docs.get("boleto_documento") or "").strip()
                boleto_vencto = str(docs.get("boleto_vencto_display") or "").strip()
                if bool(docs.get("boleto")) or boleto_fn or boleto_doc:
                    parts = []
                    if boleto_doc:
                        parts.append(f"boleto {boleto_doc}")
                    else:
                        parts.append("boleto")
                    if boleto_vencto:
                        parts.append(f"venc.: {boleto_vencto}")
                    elif due:
                        parts.append(f"venc.: {due}")
                    if boleto_fn:
                        parts.append(f"arquivo: {boleto_fn}")
                    lines.append(f"  - {' | '.join(parts)}")

                purchase_docs = docs.get("purchase_documents") if isinstance(docs.get("purchase_documents"), list) else []
                if purchase_docs:
                    shown = [str(x or "").strip() for x in purchase_docs if str(x or "").strip()]
                    if shown:
                        extra = " (lista truncada)" if bool(docs.get("purchase_documents_truncated")) else ""
                        lines.append(f"  - compras/documentos: {', '.join(shown)}{extra}")

                if bool(docs.get("assinatura")):
                    if purchase_docs:
                        lines.append("  - assinatura: sim (referente às compras/documentos acima)")
                    else:
                        lines.append("  - assinatura: sim")

                nfe_names = docs.get("nfe_filenames") if isinstance(docs.get("nfe_filenames"), list) else []
                if nfe_names:
                    count = len([n for n in nfe_names if str(n or '').strip()])
                    if count:
                        lines.append(f"  - NF-e: {count} arquivo(s)")
                    for n in nfe_names[:20]:
                        nn = str(n or "").strip()
                        if nn:
                            lines.append(f"    - {nn}")
            if bool(t.get("invoices_truncated")):
                lines.append("- (lista truncada)")

        self._set_detail_text("\n".join(lines).strip())

    def _copy(self):
        try:
            text = str(self.detail_text.get("1.0", "end") or "").strip()
        except Exception:
            text = ""
        if not text:
            name = str(self.agenda.get("name") or "").strip()
            last_run_at = self._format_dt(str(self.agenda.get("last_run_at") or "").strip())
            text = f"Alerta: {name}\nÚltimo envio: {last_run_at}"
        try:
            self.clipboard_clear()
            self.clipboard_append(text)
        except Exception:
            pass


class FinanceiroAlertaWindow(tk.Toplevel):
    def __init__(self, master, config_data: Dict[str, Any], current_user: str, on_config_saved, agenda_id: Optional[str] = None):
        super().__init__(master)
        self.master_app = master
        self.config_data = deepcopy(config_data)
        self.current_user = current_user
        self.on_config_saved = on_config_saved
        self.agenda_id = str(agenda_id) if agenda_id is not None else ""
        self.name_var = tk.StringVar(value="")
        self.enabled_var = tk.BooleanVar(value=True)
        self.send_time_var = tk.StringVar(value="06:00")
        self.send_pix_qrcode_var = tk.BooleanVar(value=False)
        self.days_before_var = tk.IntVar(value=5)
        self.days_after_var = tk.IntVar(value=0)
        self._days_trace_lock = False
        self._days_last_changed = "before"
        self.base_date_var = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        self.group_var = tk.StringVar(value="Todos")
        self.portador_var = tk.StringVar(value="Todos")
        self.customer_var = tk.StringVar(value="Todos")
        self.status_var = tk.StringVar(value="Pronto.")
        self.group_options_map: Dict[str, Any] = {"Todos": None}
        self.portador_options_map: Dict[str, Any] = {"Todos": None}
        self.all_group_options_map: Dict[str, Any] = {"Todos": None}
        self.all_portador_options_map: Dict[str, Any] = {"Todos": None}
        self.customer_options_map: Dict[str, Any] = {"Todos": None}
        self.all_customer_options_map: Dict[str, Any] = {"Todos": None}
        self._customer_filter_job = None
        self._group_filter_job = None
        self._portador_filter_job = None
        self._send_thread = None
        self._extra_body_value = ""

        self.title(f"{APP_TITLE} - Alerta de vencimento de fatura")
        self.geometry("1280x820")
        self.minsize(1120, 720)
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self._close)

        agenda_cfg = None
        if self.agenda_id:
            for a in (self.config_data.get("financeiro_agendas", []) or []):
                if isinstance(a, dict) and str(a.get("id") or "") == str(self.agenda_id):
                    agenda_cfg = a
                    break
        if agenda_cfg:
            self.name_var.set(str(agenda_cfg.get("name") or "").strip())
            self.enabled_var.set(bool(agenda_cfg.get("enabled", False)))
            self.send_time_var.set(str(agenda_cfg.get("send_time") or "06:00").strip() or "06:00")
            self.send_pix_qrcode_var.set(bool(agenda_cfg.get("send_pix_qrcode", False)))
            try:
                self.days_before_var.set(int(agenda_cfg.get("days_before_due", 5) or 5))
            except Exception:
                self.days_before_var.set(5)
            try:
                self.days_after_var.set(int(agenda_cfg.get("days_after_due", 0) or 0))
            except Exception:
                self.days_after_var.set(0)
            try:
                if int(self.days_after_var.get() or 0) > 0:
                    self._days_last_changed = "after"
                elif int(self.days_before_var.get() or 0) > 0:
                    self._days_last_changed = "before"
            except Exception:
                self._days_last_changed = "before"
            self._extra_body_value = str(agenda_cfg.get("extra_body") or "")
            saved_group = agenda_cfg.get("group_id")
            saved_portador = agenda_cfg.get("portador_id")
            saved_customer = agenda_cfg.get("customer_id")
        else:
            saved_group = None
            saved_portador = None
            saved_customer = None

        self._build_ui()
        self._setup_days_exclusivity()
        self._center_window()
        self._preview_loading = False
        self._preview_pending = False
        self._initial_load(saved_group, saved_portador, saved_customer)

    def _initial_load(self, saved_group=None, saved_portador=None, saved_customer=None):
        self._update_hint_text()

        def _work():
            db = Database(self.config_data)
            groups = db.list_grupos_pessoa()
            portadores = db.list_portadores()
            customers = db.list_customer_options_tipo_c()
            return groups, portadores, customers

        def _ok(res):
            groups, portadores, customers = res
            self._load_group_options(saved_group, grupos=groups)
            self._load_portador_options(saved_portador, portadores=portadores)
            self._load_customer_options(saved_customer, customers=customers)
            self.refresh_preview()

        def _err(e: Exception):
            self.status_var.set("Falha ao carregar dados.")
            messagebox.showerror(APP_TITLE, f"Erro ao carregar dados:\n\n{e}", parent=self)

        run_with_busy(self, "Carregando opções...", _work, _ok, _err)

    def _center_window(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - self.winfo_width()) // 2
        y = (sh - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 20)}+{max(y, 20)}")

    def _build_ui(self):
        header = ttk.Frame(self, padding=(16, 12, 16, 0))
        header.pack(fill="x")
        ttk.Label(header, text="Alerta de Vencimento de Fatura", font=("Segoe UI", 14, "bold"), foreground="#2563eb").pack(side="left")
        ttk.Label(header, text="Somente faturas com boleto vinculado serão listadas/enviadas.", font=("Segoe UI", 10), foreground="#d9534f").pack(side="left", padx=(12, 0), pady=(4, 0))
        hint = ttk.Frame(self, padding=(16, 2, 16, 8))
        hint.pack(fill="x")
        self.mode_hint_label = ttk.Label(hint, text="", font=("Segoe UI", 10), foreground="#6b7280")
        self.mode_hint_label.pack(side="left")

        settings_container = ttk.Frame(self, padding=(12, 8, 12, 8))
        settings_container.pack(fill="x")

        general_frame = ttk.LabelFrame(settings_container, text=" Configuração ", padding=(12, 10, 12, 10))
        general_frame.pack(side="left", fill="both", expand=True, padx=(0, 6))
        ttk.Label(general_frame, text="Nome:").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        ttk.Entry(general_frame, textvariable=self.name_var, width=42).grid(row=0, column=1, sticky="w", pady=4)
        ttk.Checkbutton(general_frame, text="Ativo", variable=self.enabled_var).grid(row=1, column=1, sticky="w", pady=4)
        ttk.Label(general_frame, text="Hora do disparo:").grid(row=2, column=0, sticky="w", padx=(0, 6), pady=4)
        ttk.Entry(general_frame, textvariable=self.send_time_var, width=10).grid(row=2, column=1, sticky="w", pady=4)
        ttk.Checkbutton(general_frame, text="Incluir QRCode PIX no boleto (PDF)", variable=self.send_pix_qrcode_var).grid(row=3, column=1, sticky="w", pady=(6, 4))

        rules_frame = ttk.LabelFrame(settings_container, text=" Regras ", padding=(12, 10, 12, 10))
        rules_frame.pack(side="left", fill="both", expand=True, padx=(6, 0))
        ttk.Label(rules_frame, text="Dias antes do venc.:").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        self.days_spin = ttk.Spinbox(rules_frame, from_=0, to=60, textvariable=self.days_before_var, width=10, command=self._schedule_refresh)
        self.days_spin.grid(row=0, column=1, sticky="w", pady=4)
        self.days_spin.bind("<KeyRelease>", lambda e: self._schedule_refresh())
        ttk.Label(rules_frame, text="Dias após o venc.:").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=4)
        self.days_after_spin = ttk.Spinbox(rules_frame, from_=0, to=60, textvariable=self.days_after_var, width=10, command=self._schedule_refresh)
        self.days_after_spin.grid(row=1, column=1, sticky="w", pady=4)
        self.days_after_spin.bind("<KeyRelease>", lambda e: self._schedule_refresh())
        ttk.Label(rules_frame, text="Grupo de cliente:").grid(row=2, column=0, sticky="w", padx=(0, 6), pady=4)
        self.group_entry = ttk.Entry(rules_frame, textvariable=self.group_var, width=35)
        self.group_entry.grid(row=2, column=1, sticky="w", pady=4)
        self.group_entry.bind("<KeyRelease>", self._on_group_keyrelease)
        self.group_entry.bind("<FocusIn>", self._on_group_focus_in)
        self.group_entry.bind("<FocusOut>", self._on_group_focus_out)
        self.group_entry.bind("<Down>", self._on_group_arrow_down)
        self.group_entry.bind("<Up>", self._on_group_arrow_up)
        self.group_entry.bind("<Return>", self._on_group_entry_confirm)
        self.group_suggestions_frame = ttk.Frame(rules_frame)
        self.group_suggestions_frame.grid(row=3, column=1, sticky="nsew", pady=(0, 2))
        self.group_suggestions_frame.grid_remove()
        self.group_listbox = tk.Listbox(self.group_suggestions_frame, height=6, exportselection=False)
        self.group_listbox.pack(fill="x", expand=True)
        self.group_listbox.bind("<ButtonRelease-1>", self._on_group_listbox_confirm)
        self.group_listbox.bind("<Double-Button-1>", self._on_group_listbox_confirm)
        self.group_listbox.bind("<Return>", self._on_group_listbox_confirm)

        ttk.Label(rules_frame, text="Portador:").grid(row=4, column=0, sticky="w", padx=(0, 6), pady=4)
        self.portador_entry = ttk.Entry(rules_frame, textvariable=self.portador_var, width=35)
        self.portador_entry.grid(row=4, column=1, sticky="w", pady=4)
        self.portador_entry.bind("<KeyRelease>", self._on_portador_keyrelease)
        self.portador_entry.bind("<FocusIn>", self._on_portador_focus_in)
        self.portador_entry.bind("<FocusOut>", self._on_portador_focus_out)
        self.portador_entry.bind("<Down>", self._on_portador_arrow_down)
        self.portador_entry.bind("<Up>", self._on_portador_arrow_up)
        self.portador_entry.bind("<Return>", self._on_portador_entry_confirm)
        self.portador_suggestions_frame = ttk.Frame(rules_frame)
        self.portador_suggestions_frame.grid(row=5, column=1, sticky="nsew", pady=(0, 2))
        self.portador_suggestions_frame.grid_remove()
        self.portador_listbox = tk.Listbox(self.portador_suggestions_frame, height=6, exportselection=False)
        self.portador_listbox.pack(fill="x", expand=True)
        self.portador_listbox.bind("<ButtonRelease-1>", self._on_portador_listbox_confirm)
        self.portador_listbox.bind("<Double-Button-1>", self._on_portador_listbox_confirm)
        self.portador_listbox.bind("<Return>", self._on_portador_listbox_confirm)

        ttk.Label(rules_frame, text="Cliente:").grid(row=6, column=0, sticky="w", padx=(0, 6), pady=4)
        self.customer_entry = ttk.Entry(rules_frame, textvariable=self.customer_var, width=35)
        self.customer_entry.grid(row=6, column=1, sticky="w", pady=4)
        self.customer_entry.bind("<KeyRelease>", self._on_customer_keyrelease)
        self.customer_entry.bind("<FocusIn>", self._on_customer_focus_in)
        self.customer_entry.bind("<FocusOut>", self._on_customer_focus_out)
        self.customer_entry.bind("<Down>", self._on_customer_arrow_down)
        self.customer_entry.bind("<Up>", self._on_customer_arrow_up)
        self.customer_entry.bind("<Return>", self._on_customer_entry_confirm)
        self.customer_suggestions_frame = ttk.Frame(rules_frame)
        self.customer_suggestions_frame.grid(row=7, column=1, sticky="nsew", pady=(0, 2))
        self.customer_suggestions_frame.grid_remove()
        self.customer_listbox = tk.Listbox(self.customer_suggestions_frame, height=6, exportselection=False)
        self.customer_listbox.pack(fill="x", expand=True)
        self.customer_listbox.bind("<ButtonRelease-1>", self._on_customer_listbox_confirm)
        self.customer_listbox.bind("<Double-Button-1>", self._on_customer_listbox_confirm)
        self.customer_listbox.bind("<Return>", self._on_customer_listbox_confirm)

        preview_header = ttk.Frame(self, padding=(12, 0, 12, 6))
        preview_header.pack(fill="x")
        ttk.Label(preview_header, text="Prévia dos Clientes Abrangidos", font=("Segoe UI", 11, "bold")).pack(side="left")
        ttk.Label(preview_header, text="Data base (simulação):").pack(side="left", padx=(30, 6))
        self.base_date_entry = ttk.Entry(preview_header, textvariable=self.base_date_var, width=14)
        self.base_date_entry.pack(side="left")
        bind_date_entry_shortcuts(self.base_date_entry)
        self.base_date_entry.bind("<FocusOut>", lambda e: self._schedule_refresh(), add="+")

        preview = ttk.Frame(self, padding=(12, 0, 12, 10))
        preview.pack(fill="both", expand=True)
        cols = ("cliente", "grupo", "email", "situacao", "titulos", "total", "portador")
        self.tree = ttk.Treeview(preview, columns=cols, show="headings")
        self.tree.heading("cliente", text="Cliente")
        self.tree.heading("grupo", text="Grupo")
        self.tree.heading("email", text="E-mail")
        self.tree.heading("situacao", text="Situação")
        self.tree.heading("titulos", text="Títulos")
        self.tree.heading("total", text="Total")
        self.tree.heading("portador", text="Portador")
        self.tree.column("cliente", width=260, anchor="w")
        self.tree.column("grupo", width=160, anchor="w")
        self.tree.column("email", width=220, anchor="w")
        self.tree.column("situacao", width=220, anchor="w")
        self.tree.column("titulos", width=70, anchor="center")
        self.tree.column("total", width=100, anchor="e")
        self.tree.column("portador", width=160, anchor="w")
        vsb = ttk.Scrollbar(preview, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(preview, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        preview.columnconfigure(0, weight=1)
        preview.rowconfigure(0, weight=1)

        bottom = ttk.Frame(self, padding=(12, 8, 12, 18))
        bottom.pack(fill="x")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left")
        btns = ttk.Frame(bottom)
        btns.pack(side="right")
        ttk.Button(btns, text="Simular envio", command=self.simulate_now).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Enviar agora", command=self.send_now).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Salvar Alerta", command=self.save_settings).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Fechar", command=self._close).pack(side="left")

    def _load_group_options(self, selected_id=None, grupos=None):
        if grupos is None:
            try:
                grupos = Database(self.config_data).list_grupos_pessoa()
            except Exception:
                grupos = []
        options = {"Todos": None}
        for g in grupos:
            gid = g.get("grid")
            name = str(g.get("nome") or "").strip()
            if gid in (None, "", 0, "0") or not name:
                continue
            options[name] = gid
        self.all_group_options_map = options
        self.group_options_map = options
        if selected_id not in (None, "", 0, "0"):
            for label, gid in options.items():
                if str(gid) == str(selected_id):
                    self.group_var.set(label)
                    break
        if not (self.group_var.get() or "").strip():
            self.group_var.set("Todos")

    def _load_portador_options(self, selected_id=None, portadores=None):
        if portadores is None:
            try:
                portadores = Database(self.config_data).list_portadores()
            except Exception:
                portadores = []
        options = {"Todos": None}
        for p in portadores:
            pid = p.get("grid")
            name = str(p.get("nome") or "").strip()
            if pid in (None, "", 0, "0") or not name:
                continue
            options[name] = pid
        self.all_portador_options_map = options
        self.portador_options_map = options
        if selected_id not in (None, "", 0, "0"):
            for label, pid in options.items():
                if str(pid) == str(selected_id):
                    self.portador_var.set(label)
                    break
        if not (self.portador_var.get() or "").strip():
            self.portador_var.set("Todos")

    def _load_customer_options(self, selected_customer_id=None, customers=None):
        if customers is None:
            try:
                customers = Database(self.config_data).list_customer_options_tipo_c()
            except Exception:
                customers = []
        options = {"Todos": None}
        for c in customers:
            cid = c.get("customer_id")
            code = str(c.get("codigo_cliente") or "").strip()
            name = str(c.get("cliente") or "").strip()
            if cid in (None, "", 0, "0") or not name:
                continue
            label = f"{code} - {name}".strip(" -")
            options[label] = cid
        self.all_customer_options_map = options
        self.customer_options_map = dict(options)
        if selected_customer_id not in (None, "", 0, "0"):
            for label, cid in options.items():
                if str(cid) == str(selected_customer_id):
                    self.customer_var.set(label)
                    break
        if self.customer_var.get() not in options:
            self.customer_var.set("Todos")

    def _selected_group_id(self):
        typed = (self.group_var.get() or "").strip()
        if not typed or typed == "Todos":
            return None
        exact = self.all_group_options_map.get(typed)
        if exact is not None:
            return exact
        typed_lower = typed.lower()
        matches = [gid for label, gid in self.all_group_options_map.items() if label != "Todos" and typed_lower in label.lower()]
        return matches[0] if len(matches) == 1 else None

    def _selected_portador_id(self):
        typed = (self.portador_var.get() or "").strip()
        if not typed or typed == "Todos":
            return None
        exact = self.all_portador_options_map.get(typed)
        if exact is not None:
            return exact
        typed_lower = typed.lower()
        matches = [pid for label, pid in self.all_portador_options_map.items() if label != "Todos" and typed_lower in label.lower()]
        return matches[0] if len(matches) == 1 else None

    def _selected_customer_id(self):
        typed = (self.customer_var.get() or "").strip()
        if not typed or typed == "Todos":
            return None
        exact = self.all_customer_options_map.get(typed)
        if exact is not None:
            return exact
        typed_lower = typed.lower()
        matches = [cid for label, cid in self.all_customer_options_map.items() if label != "Todos" and typed_lower in label.lower()]
        return matches[0] if len(matches) == 1 else None

    def _matching_group_labels(self, typed: str):
        typed = (typed or "").strip().lower()
        if not typed or typed == "todos":
            return []
        parts = [p for p in re.split(r"\s+", typed) if p]
        labels = []
        for label in self.all_group_options_map.keys():
            if label == "Todos":
                continue
            low = label.lower()
            if all(p in low for p in parts):
                labels.append(label)
        return labels[:200]

    def _show_group_suggestions(self, labels):
        if not labels:
            self._hide_group_suggestions()
            return
        self.group_listbox.delete(0, "end")
        for label in labels:
            self.group_listbox.insert("end", label)
        self.group_suggestions_frame.grid()
        self.group_listbox.selection_clear(0, "end")
        self.group_listbox.selection_set(0)
        self.group_listbox.activate(0)
        self.group_listbox.see(0)

    def _hide_group_suggestions(self):
        self.group_suggestions_frame.grid_remove()

    def _on_group_keyrelease(self, event=None):
        keysym = getattr(event, "keysym", "")
        if keysym in ("Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return None
        if not (self.group_var.get() or "").strip():
            self.group_var.set("Todos")
            self._hide_group_suggestions()
            self._schedule_refresh()
            return None
        if self._group_filter_job is not None:
            try:
                self.after_cancel(self._group_filter_job)
            except Exception:
                pass
        self._group_filter_job = self.after(150, self._apply_group_search)
        return None

    def _apply_group_search(self):
        self._group_filter_job = None
        typed = (self.group_var.get() or "").strip()
        if not typed or typed == "Todos":
            self._hide_group_suggestions()
            return
        self._show_group_suggestions(self._matching_group_labels(typed))

    def _on_group_focus_in(self, event=None):
        typed = (self.group_var.get() or "").strip()
        if typed and typed != "Todos":
            self._show_group_suggestions(self._matching_group_labels(typed))
        return None

    def _on_group_focus_out(self, event=None):
        self.after(120, self._hide_group_suggestions)
        if not (self.group_var.get() or "").strip():
            self.group_var.set("Todos")
        self._schedule_refresh()

    def _on_group_arrow_down(self, event=None):
        if not self.group_suggestions_frame.winfo_ismapped() or self.group_listbox.size() == 0:
            self._show_group_suggestions(self._matching_group_labels(self.group_var.get()))
        if not self.group_suggestions_frame.winfo_ismapped() or self.group_listbox.size() == 0:
            return "break"
        selection = self.group_listbox.curselection()
        index = selection[0] + 1 if selection else 0
        if index >= self.group_listbox.size():
            index = self.group_listbox.size() - 1
        self.group_listbox.selection_clear(0, "end")
        self.group_listbox.selection_set(index)
        self.group_listbox.activate(index)
        self.group_listbox.see(index)
        return "break"

    def _on_group_arrow_up(self, event=None):
        if not self.group_suggestions_frame.winfo_ismapped() or self.group_listbox.size() == 0:
            return "break"
        selection = self.group_listbox.curselection()
        index = selection[0] - 1 if selection else 0
        if index < 0:
            index = 0
        self.group_listbox.selection_clear(0, "end")
        self.group_listbox.selection_set(index)
        self.group_listbox.activate(index)
        self.group_listbox.see(index)
        return "break"

    def _on_group_entry_confirm(self, event=None):
        typed = (self.group_var.get() or "").strip()
        if not typed:
            self.group_var.set("Todos")
        if self.group_suggestions_frame.winfo_ismapped() and self.group_listbox.size() > 0:
            sel = self.group_listbox.curselection()
            label = self.group_listbox.get(sel[0]) if sel else self.group_listbox.get(0)
            self.group_var.set(label)
        elif typed and typed != "Todos":
            matches = self._matching_group_labels(typed)
            if len(matches) == 1:
                self.group_var.set(matches[0])
        self._hide_group_suggestions()
        self._schedule_refresh()
        return "break"

    def _on_group_listbox_confirm(self, event=None):
        sel = self.group_listbox.curselection()
        if sel:
            label = self.group_listbox.get(sel[0])
        elif self.group_listbox.size() > 0:
            label = self.group_listbox.get(0)
        else:
            return "break"
        self.group_var.set(label)
        self._hide_group_suggestions()
        try:
            self.group_entry.focus_set()
            self.group_entry.icursor("end")
        except Exception:
            pass
        self._schedule_refresh()
        return "break"

    def _matching_portador_labels(self, typed: str):
        typed = (typed or "").strip().lower()
        if not typed or typed == "todos":
            return []
        parts = [p for p in re.split(r"\s+", typed) if p]
        labels = []
        for label in self.all_portador_options_map.keys():
            if label == "Todos":
                continue
            low = label.lower()
            if all(p in low for p in parts):
                labels.append(label)
        return labels[:200]

    def _show_portador_suggestions(self, labels):
        if not labels:
            self._hide_portador_suggestions()
            return
        self.portador_listbox.delete(0, "end")
        for label in labels:
            self.portador_listbox.insert("end", label)
        self.portador_suggestions_frame.grid()
        self.portador_listbox.selection_clear(0, "end")
        self.portador_listbox.selection_set(0)
        self.portador_listbox.activate(0)
        self.portador_listbox.see(0)

    def _hide_portador_suggestions(self):
        self.portador_suggestions_frame.grid_remove()

    def _on_portador_keyrelease(self, event=None):
        keysym = getattr(event, "keysym", "")
        if keysym in ("Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return None
        if not (self.portador_var.get() or "").strip():
            self.portador_var.set("Todos")
            self._hide_portador_suggestions()
            self._schedule_refresh()
            return None
        if self._portador_filter_job is not None:
            try:
                self.after_cancel(self._portador_filter_job)
            except Exception:
                pass
        self._portador_filter_job = self.after(150, self._apply_portador_search)
        return None

    def _apply_portador_search(self):
        self._portador_filter_job = None
        typed = (self.portador_var.get() or "").strip()
        if not typed or typed == "Todos":
            self._hide_portador_suggestions()
            return
        self._show_portador_suggestions(self._matching_portador_labels(typed))

    def _on_portador_focus_in(self, event=None):
        typed = (self.portador_var.get() or "").strip()
        if typed and typed != "Todos":
            self._show_portador_suggestions(self._matching_portador_labels(typed))
        return None

    def _on_portador_focus_out(self, event=None):
        self.after(120, self._hide_portador_suggestions)
        if not (self.portador_var.get() or "").strip():
            self.portador_var.set("Todos")
        self._schedule_refresh()

    def _on_portador_arrow_down(self, event=None):
        if not self.portador_suggestions_frame.winfo_ismapped() or self.portador_listbox.size() == 0:
            self._show_portador_suggestions(self._matching_portador_labels(self.portador_var.get()))
        if not self.portador_suggestions_frame.winfo_ismapped() or self.portador_listbox.size() == 0:
            return "break"
        selection = self.portador_listbox.curselection()
        index = selection[0] + 1 if selection else 0
        if index >= self.portador_listbox.size():
            index = self.portador_listbox.size() - 1
        self.portador_listbox.selection_clear(0, "end")
        self.portador_listbox.selection_set(index)
        self.portador_listbox.activate(index)
        self.portador_listbox.see(index)
        return "break"

    def _on_portador_arrow_up(self, event=None):
        if not self.portador_suggestions_frame.winfo_ismapped() or self.portador_listbox.size() == 0:
            return "break"
        selection = self.portador_listbox.curselection()
        index = selection[0] - 1 if selection else 0
        if index < 0:
            index = 0
        self.portador_listbox.selection_clear(0, "end")
        self.portador_listbox.selection_set(index)
        self.portador_listbox.activate(index)
        self.portador_listbox.see(index)
        return "break"

    def _on_portador_entry_confirm(self, event=None):
        typed = (self.portador_var.get() or "").strip()
        if not typed:
            self.portador_var.set("Todos")
        if self.portador_suggestions_frame.winfo_ismapped() and self.portador_listbox.size() > 0:
            sel = self.portador_listbox.curselection()
            label = self.portador_listbox.get(sel[0]) if sel else self.portador_listbox.get(0)
            self.portador_var.set(label)
        elif typed and typed != "Todos":
            matches = self._matching_portador_labels(typed)
            if len(matches) == 1:
                self.portador_var.set(matches[0])
        self._hide_portador_suggestions()
        self._schedule_refresh()
        return "break"

    def _on_portador_listbox_confirm(self, event=None):
        sel = self.portador_listbox.curselection()
        if sel:
            label = self.portador_listbox.get(sel[0])
        elif self.portador_listbox.size() > 0:
            label = self.portador_listbox.get(0)
        else:
            return "break"
        self.portador_var.set(label)
        self._hide_portador_suggestions()
        try:
            self.portador_entry.focus_set()
            self.portador_entry.icursor("end")
        except Exception:
            pass
        self._schedule_refresh()
        return "break"

    def _matching_customer_labels(self, typed: str):
        typed = (typed or "").strip().lower()
        if not typed or typed == "todos":
            return []
        parts = [p for p in re.split(r"\s+", typed) if p]
        labels = []
        for label in self.all_customer_options_map.keys():
            if label == "Todos":
                continue
            low = label.lower()
            if all(p in low for p in parts):
                labels.append(label)
        return labels[:200]

    def _show_customer_suggestions(self, labels):
        if not labels:
            self._hide_customer_suggestions()
            return
        self.customer_listbox.delete(0, "end")
        for label in labels:
            self.customer_listbox.insert("end", label)
        self.customer_suggestions_frame.grid()
        self.customer_listbox.selection_clear(0, "end")
        self.customer_listbox.selection_set(0)
        self.customer_listbox.activate(0)
        self.customer_listbox.see(0)

    def _hide_customer_suggestions(self):
        self.customer_suggestions_frame.grid_remove()

    def _on_customer_keyrelease(self, event=None):
        keysym = getattr(event, "keysym", "")
        if keysym in ("Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return None
        if not (self.customer_var.get() or "").strip():
            self.customer_var.set("Todos")
            self._hide_customer_suggestions()
            self._schedule_refresh()
            return None
        if self._customer_filter_job is not None:
            try:
                self.after_cancel(self._customer_filter_job)
            except Exception:
                pass
        self._customer_filter_job = self.after(150, self._apply_customer_search)
        return None

    def _apply_customer_search(self):
        self._customer_filter_job = None
        typed = (self.customer_var.get() or "").strip()
        if not typed or typed == "Todos":
            self._hide_customer_suggestions()
            return
        self._show_customer_suggestions(self._matching_customer_labels(typed))

    def _on_customer_focus_in(self, event=None):
        typed = (self.customer_var.get() or "").strip()
        if typed and typed != "Todos":
            self._show_customer_suggestions(self._matching_customer_labels(typed))
        return None

    def _on_customer_focus_out(self, event=None):
        self.after(120, self._hide_customer_suggestions)
        if not (self.customer_var.get() or "").strip():
            self.customer_var.set("Todos")
        self._schedule_refresh()
        return None

    def _on_customer_arrow_down(self, event=None):
        if not self.customer_suggestions_frame.winfo_ismapped() or self.customer_listbox.size() == 0:
            self._apply_customer_search()
        if not self.customer_suggestions_frame.winfo_ismapped() or self.customer_listbox.size() == 0:
            return "break"
        selection = self.customer_listbox.curselection()
        index = selection[0] + 1 if selection else 0
        if index >= self.customer_listbox.size():
            index = self.customer_listbox.size() - 1
        self.customer_listbox.selection_clear(0, "end")
        self.customer_listbox.selection_set(index)
        self.customer_listbox.activate(index)
        self.customer_listbox.see(index)
        return "break"

    def _on_customer_arrow_up(self, event=None):
        if not self.customer_suggestions_frame.winfo_ismapped() or self.customer_listbox.size() == 0:
            self._apply_customer_search()
        if not self.customer_suggestions_frame.winfo_ismapped() or self.customer_listbox.size() == 0:
            return "break"
        selection = self.customer_listbox.curselection()
        index = selection[0] - 1 if selection else 0
        if index < 0:
            index = 0
        self.customer_listbox.selection_clear(0, "end")
        self.customer_listbox.selection_set(index)
        self.customer_listbox.activate(index)
        self.customer_listbox.see(index)
        return "break"

    def _on_customer_entry_confirm(self, event=None):
        typed = (self.customer_var.get() or "").strip()
        if not typed:
            self.customer_var.set("Todos")
        if self.customer_suggestions_frame.winfo_ismapped() and self.customer_listbox.size() > 0:
            selection = self.customer_listbox.curselection()
            label = self.customer_listbox.get(selection[0]) if selection else self.customer_listbox.get(0)
            self.customer_var.set(label)
        elif typed and typed != "Todos":
            matches = self._matching_customer_labels(typed)
            if len(matches) == 1:
                self.customer_var.set(matches[0])
        self._hide_customer_suggestions()
        self._schedule_refresh()
        return "break"

    def _on_customer_listbox_confirm(self, event=None):
        selection = self.customer_listbox.curselection()
        if selection:
            label = self.customer_listbox.get(selection[0])
        elif self.customer_listbox.size() > 0:
            label = self.customer_listbox.get(0)
        else:
            return "break"
        self.customer_var.set(label)
        self._hide_customer_suggestions()
        try:
            self.customer_entry.focus_set()
            self.customer_entry.icursor("end")
        except Exception:
            pass
        self._schedule_refresh()
        return "break"

    def _parse_base_date(self):
        return OpenInvoicesWindow._parse_date(self, self.base_date_var.get()) or date.today()

    def _setup_days_exclusivity(self):
        def _on_before(*_):
            self._enforce_days_exclusive(changed="before")

        def _on_after(*_):
            self._enforce_days_exclusive(changed="after")

        try:
            if hasattr(self.days_before_var, "trace_add"):
                self.days_before_var.trace_add("write", _on_before)
                self.days_after_var.trace_add("write", _on_after)
            else:
                self.days_before_var.trace("w", _on_before)
                self.days_after_var.trace("w", _on_after)
        except Exception:
            pass
        self._enforce_days_exclusive(changed=None)

    def _get_days_values(self, max_days: int, changed: Optional[str] = None) -> Tuple[int, int]:
        try:
            before_days = int(self.days_before_var.get() or 0)
        except Exception:
            before_days = 0
        try:
            after_days = int(self.days_after_var.get() or 0)
        except Exception:
            after_days = 0
        before_days = max(0, min(int(max_days or 0), before_days))
        after_days = max(0, min(int(max_days or 0), after_days))
        if before_days > 0 and after_days > 0:
            if changed == "before":
                after_days = 0
            elif changed == "after":
                before_days = 0
            else:
                if str(getattr(self, "_days_last_changed", "before") or "before") == "after":
                    before_days = 0
                else:
                    after_days = 0
        return before_days, after_days

    def _enforce_days_exclusive(self, changed: Optional[str]):
        if self._days_trace_lock:
            return
        before_days, after_days = self._get_days_values(60, changed=changed)
        self._days_trace_lock = True
        try:
            if changed == "before" and before_days > 0:
                self._days_last_changed = "before"
                self.days_before_var.set(before_days)
                self.days_after_var.set(after_days)
            elif changed == "after" and after_days > 0:
                self._days_last_changed = "after"
                self.days_before_var.set(before_days)
                self.days_after_var.set(after_days)
            else:
                self.days_before_var.set(before_days)
                self.days_after_var.set(after_days)
        finally:
            self._days_trace_lock = False
        try:
            self._schedule_refresh()
        except Exception:
            pass

    def _update_hint_text(self):
        before_days, after_days = self._get_days_values(60)
        parts = []
        if before_days > 0:
            parts.append(f"{before_days} dia(s) antes do vencimento do boleto")
        if after_days > 0:
            parts.append(f"{after_days} dia(s) após o vencimento do boleto")
        self.mode_hint_label.configure(text=("Alerta: envia quando o boleto estiver em " + " ou ".join(parts) + ".") if parts else "Defina os dias antes/depois do vencimento para montar o alerta.")

    def _schedule_refresh(self):
        self._update_hint_text()
        try:
            if hasattr(self, "_auto_refresh_job") and self._auto_refresh_job:
                self.after_cancel(self._auto_refresh_job)
        except Exception:
            pass
        self._auto_refresh_job = self.after(250, self.refresh_preview)

    def refresh_preview(self):
        if self._preview_loading:
            self._preview_pending = True
            return
        try:
            base_date = self._parse_base_date()
            group_id = self._selected_group_id()
            portador_id = self._selected_portador_id()
            customer_id = self._selected_customer_id()
            before_days, after_days = self._get_days_values(60)
        except Exception as e:
            self.status_var.set("Falha ao carregar dados.")
            messagebox.showerror(APP_TITLE, f"Erro ao atualizar alerta:\n\n{e}", parent=self)
            return
        due_dates = []
        if before_days > 0:
            due_dates.append(base_date + timedelta(days=before_days))
        if after_days > 0:
            due_dates.append(base_date - timedelta(days=after_days))
        due_dates = sorted({d for d in due_dates})

        if not due_dates:
            self.status_var.set("Defina os dias antes/depois do vencimento para visualizar a prévia.")
            for it in self.tree.get_children():
                self.tree.delete(it)
            self._preview_rows = []
            self._preview_base_date = base_date
            return

        self._preview_loading = True

        def _work():
            rows: List[Dict[str, Any]] = []
            db = Database(self.config_data)
            for d in due_dates:
                rows.extend(db.list_agenda_invoices(d, group_id=group_id, portador_id=portador_id, customer_id=customer_id))

            def _status_label(vencto: date) -> str:
                diff = (vencto - base_date).days
                if diff == 0:
                    return "Vence hoje"
                if diff > 0:
                    return f"Vence em {diff} dia(s)"
                return f"Vencido há {abs(diff)} dia(s)"

            grouped: Dict[Any, Dict[str, Any]] = {}
            for r in rows:
                cid = r.get("customer_id")
                key = cid if cid not in (None, "", 0, "0") else f"sem_{r.get('cliente')}"
                item = grouped.get(key)
                if not item:
                    item = {"cliente": str(r.get("cliente") or "").strip(), "grupo": str(r.get("customer_group_name") or "").strip(), "email": str(r.get("customer_email") or "").strip(), "situacoes": set(), "titulos": 0, "total": 0.0, "portador": set()}
                    grouped[key] = item
                item["titulos"] += 1
                try:
                    item["total"] += float(r.get("saldo_em_aberto") or 0)
                except Exception:
                    pass
                vd = r.get("vencto")
                if isinstance(vd, date):
                    item["situacoes"].add(f"{_status_label(vd)} ({vd.strftime('%d/%m/%Y')})")
                pname = str(r.get("portador_nome") or "").strip()
                if pname:
                    item["portador"].add(pname)

            items = list(grouped.values())
            items.sort(key=lambda x: (x["cliente"] or "").lower())
            total_sum = sum(float(it.get("total") or 0) for it in items)
            dd_txt = " / ".join([d.strftime("%d/%m/%Y") for d in due_dates])
            return {"rows": rows, "items": items, "total_sum": total_sum, "dd_txt": dd_txt}

        def _ok(payload):
            self._preview_loading = False
            rows = payload.get("rows") or []
            items = payload.get("items") or []
            for it in self.tree.get_children():
                self.tree.delete(it)
            for it in items:
                portador_txt = ", ".join(sorted(it.get("portador") or [])) if it.get("portador") else ""
                sset = sorted(list(it.get("situacoes") or []))
                situacao_txt = sset[0] if len(sset) == 1 else ("Múltiplos vencimentos" if sset else "")
                self.tree.insert("", "end", values=(it.get("cliente") or "", it.get("grupo") or "", it.get("email") or "", situacao_txt, it.get("titulos") or 0, money_br(it.get("total") or 0), portador_txt))
            self.status_var.set(f"{len(items)} cliente(s). Total previsto: {money_br(payload.get('total_sum') or 0)}. Datas-alvo: {payload.get('dd_txt')}")
            self._preview_rows = rows
            self._preview_base_date = base_date
            if self._preview_pending:
                self._preview_pending = False
                self.refresh_preview()

        def _err(e: Exception):
            self._preview_loading = False
            self.status_var.set("Falha ao carregar dados.")
            messagebox.showerror(APP_TITLE, f"Erro ao atualizar alerta:\n\n{e}", parent=self)
            if self._preview_pending:
                self._preview_pending = False

        run_with_busy(self, "Carregando prévia...", _work, _ok, _err)

    def _get_extra_body(self) -> str:
        return str(getattr(self, "_extra_body_value", "") or "")

    def _alert_settings_payload(self) -> Dict[str, Any]:
        group_id = self._selected_group_id()
        portador_id = self._selected_portador_id()
        customer_id = self._selected_customer_id()
        before_days, after_days = self._get_days_values(365)
        return {
            "id": (self.agenda_id or ""),
            "name": str(self.name_var.get() or "").strip() or "Alerta de vencimento",
            "enabled": bool(self.enabled_var.get()),
            "send_time": str(self.send_time_var.get() or "06:00").strip() or "06:00",
            "send_pix_qrcode": bool(self.send_pix_qrcode_var.get()),
            "days_before_due": max(0, min(365, before_days)),
            "days_after_due": max(0, min(365, after_days)),
            "group_id": group_id,
            "portador_id": portador_id,
            "customer_id": customer_id,
            "extra_body": self._get_extra_body(),
        }

    def save_settings(self):
        payload = self._alert_settings_payload()
        agendas = self.config_data.get("financeiro_agendas", []) or []
        if not isinstance(agendas, list):
            agendas = []
        updated = False
        if payload["id"]:
            for i, a in enumerate(agendas):
                if isinstance(a, dict) and str(a.get("id") or "") == str(payload["id"]):
                    last_run_date = str(a.get("last_run_date") or "")
                    last_run_at = str(a.get("last_run_at") or "")
                    last_due_date = str(a.get("last_due_date") or "")
                    last_result = a.get("last_result") if isinstance(a.get("last_result"), dict) else {}
                    merged = dict(a)
                    merged.update(payload)
                    merged["last_run_date"] = last_run_date
                    merged["last_run_at"] = last_run_at
                    merged["last_due_date"] = last_due_date
                    merged["last_result"] = last_result
                    agendas[i] = merged
                    updated = True
                    break
        if not updated:
            payload["id"] = payload["id"] or str(len(agendas) + 1)
            agendas.append(payload)
        self.config_data["financeiro_agendas"] = agendas
        ConfigManager.save(self.config_data)
        self.config_data = ConfigManager.load()
        messagebox.showinfo(APP_TITLE, "Alerta salvo com sucesso.", parent=self)
        self.on_config_saved(self.config_data)
        self._close()

    def simulate_now(self):
        try:
            base_date = self._parse_base_date()
            group_id = self._selected_group_id()
            portador_id = self._selected_portador_id()
            customer_id = self._selected_customer_id()
            before_days, after_days = self._get_days_values(60)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Erro ao preparar simulação:\n\n{e}", parent=self)
            return
        due_dates = []
        if before_days > 0:
            due_dates.append(base_date + timedelta(days=before_days))
        if after_days > 0:
            due_dates.append(base_date - timedelta(days=after_days))
        due_dates = sorted({d for d in due_dates})
        if not due_dates:
            messagebox.showinfo(APP_TITLE, "Defina os dias antes/depois do vencimento para simular.", parent=self)
            return

        def _work():
            rows: List[Dict[str, Any]] = []
            db = Database(self.config_data)
            for d in due_dates:
                rows.extend(db.list_agenda_invoices(d, group_id=group_id, portador_id=portador_id, customer_id=customer_id))
            return rows

        def _ok(rows):
            rows = rows or []
            if not rows:
                messagebox.showinfo(APP_TITLE, "Nenhum título encontrado para as regras atuais.", parent=self)
                return
            messagebox.showinfo(APP_TITLE, f"{len(rows)} título(s) encontrado(s) para envio.", parent=self)

        def _err(e: Exception):
            messagebox.showerror(APP_TITLE, f"Erro ao simular alerta:\n\n{e}", parent=self)

        run_with_busy(self, "Simulando alerta...", _work, _ok, _err)

    def _send_for_rows_grouped(self, rows: List[Dict[str, Any]], dry_run: bool = False, base_date: Optional[date] = None):
        base_date = base_date or date.today()
        smtp_cfg = self.config_data.get("smtp", {})
        smtp_email = str(smtp_cfg.get("email", "")).strip()
        smtp_host = str(smtp_cfg.get("host", "")).strip()
        smtp_password = str(smtp_cfg.get("password", "")).strip()
        smtp_port = int(smtp_cfg.get("port", 587) or 587)
        if not smtp_email or not smtp_host or not smtp_password or not smtp_port:
            raise AppError("SMTP não configurado.")

        def _send_smtp_message(msg: EmailMessage):
            if smtp_port == 465:
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL(smtp_host, smtp_port, context=context, timeout=20) as server:
                    server.login(smtp_email, smtp_password)
                    server.send_message(msg)
            else:
                with smtplib.SMTP(smtp_host, smtp_port, timeout=20) as server:
                    server.ehlo()
                    try:
                        server.starttls(context=ssl.create_default_context())
                        server.ehlo()
                    except Exception:
                        pass
                    server.login(smtp_email, smtp_password)
                    server.send_message(msg)

        invoices: List[InvoiceRow] = []
        for r in rows:
            invoices.append(
                InvoiceRow(
                    invoice_id=r.get("movto_id"),
                    company=str(r.get("empresa") or "").strip(),
                    customer_id=r.get("customer_id"),
                    customer_code=r.get("codigo_cliente"),
                    customer_name=str(r.get("cliente") or "").strip(),
                    motive_code="",
                    motive_name="",
                    account_code=str(r.get("conta") or "").strip(),
                    account_name=str(r.get("conta_nome") or "").strip(),
                    issue_date=r.get("data"),
                    due_date=r.get("vencto"),
                    amount=r.get("valor"),
                    discount_amount=r.get("valor_desconto"),
                    paid_amount=r.get("valor_baixado"),
                    open_balance=r.get("saldo_em_aberto"),
                    customer_email=str(r.get("customer_email") or "").strip(),
                )
            )

        invoice_ids = [i.invoice_id for i in invoices if i.invoice_id not in (None, "", 0, "0")]
        boleto_map: Dict[Any, Dict[str, Any]] = {}
        try:
            boleto_map = Database(self.config_data).get_boletos_email_payload_bulk(invoice_ids)
        except Exception:
            boleto_map = {}

        signature_map: Dict[Any, Dict[str, Any]] = {}
        try:
            signature_map = Database(self.config_data).get_sale_signatures_pdf_bulk(invoice_ids)
        except Exception:
            signature_map = {}

        nfe_map: Dict[Any, Dict[str, Any]] = {}
        try:
            nfe_map = Database(self.config_data).get_nfe_attachments_bulk(invoice_ids, only_invoice_mlid=True, max_nfes_per_invoice=1)
        except Exception:
            nfe_map = {}

        grouped: Dict[str, Dict[str, Any]] = {}
        for inv in invoices:
            cid = inv.customer_id
            key = str(cid) if cid not in (None, "", 0, "0") else (inv.customer_email or inv.customer_name or str(inv.customer_code))
            item = grouped.get(key)
            if not item:
                item = {"customer_id": cid, "customer_name": inv.customer_name, "customer_email": inv.customer_email, "invoices": []}
                grouped[key] = item
            item["invoices"].append(inv)

        emails_sent = 0
        skipped_no_email = 0
        failed = 0
        attachments_total = 0
        missing_total = 0
        details: List[Dict[str, Any]] = []
        details_truncated = False
        include_pix_qrcode = bool(self.send_pix_qrcode_var.get())
        try:
            delay_seconds = int(smtp_cfg.get("delay_seconds", 5) or 0)
        except Exception:
            delay_seconds = 5
        delay_seconds = max(0, min(300, delay_seconds))
        first_email = True
        db = Database(self.config_data)
        targets: List[Dict[str, Any]] = []

        def _invoice_items(invs: List[InvoiceRow], purchase_info_map: Optional[Dict[Any, Dict[str, Any]]] = None) -> List[Dict[str, Any]]:
            items: List[Dict[str, Any]] = []
            for inv in invs or []:
                boleto_data = boleto_map.get(inv.invoice_id) or {}
                nfe = nfe_map.get(inv.invoice_id) or nfe_map.get(str(inv.invoice_id)) or {}
                nfe_atts = nfe.get("attachments") if isinstance(nfe.get("attachments"), list) else []
                sig = signature_map.get(inv.invoice_id) or {}
                pinfo = {}
                if isinstance(purchase_info_map, dict):
                    pinfo = purchase_info_map.get(inv.invoice_id) or purchase_info_map.get(str(inv.invoice_id)) or {}
                pdocs = []
                try:
                    for d in (pinfo.get("documents") or []):
                        docno = str((d or {}).get("documento") or "").strip()
                        if docno:
                            pdocs.append(docno)
                except Exception:
                    pdocs = []
                seen = set()
                uniq_pdocs = []
                for docno in pdocs:
                    if docno in seen:
                        continue
                    seen.add(docno)
                    uniq_pdocs.append(docno)
                nfe_files = []
                for a in nfe_atts:
                    fn = str(a.get("filename") or "").strip()
                    if fn:
                        nfe_files.append(fn)
                items.append(
                    {
                        "invoice_id": inv.invoice_id,
                        "issue_date": str(getattr(inv, "issue_date", "") or ""),
                        "due_date": str(getattr(inv, "due_date", "") or ""),
                        "open_balance": float(inv.open_balance or 0) if inv.open_balance not in (None, "") else 0.0,
                        "docs": {
                            "boleto": bool(boleto_data.get("exists")),
                            "boleto_filename": str(boleto_data.get("filename") or ""),
                            "boleto_documento": str(boleto_data.get("documento") or "").strip(),
                            "boleto_vencto_display": str(boleto_data.get("vencto_display") or "").strip(),
                            "xml": bool([a for a in nfe_atts if str(a.get("filename") or "").lower().endswith(".xml") and a.get("data")]),
                            "danfe": bool([a for a in nfe_atts if str(a.get("filename") or "").lower().endswith(".pdf") and a.get("data")]),
                            "nfe_filenames": nfe_files[:10],
                            "assinatura": bool(sig.get("exists") or sig.get("attachments")),
                            "purchase_documents": uniq_pdocs[:20],
                            "purchase_documents_truncated": bool(len(uniq_pdocs) > 20),
                        },
                    }
                )
            return items

        for g in grouped.values():
            invs: List[InvoiceRow] = g.get("invoices") or []
            if not invs:
                continue
            to_email = (g.get("customer_email") or "").strip()
            if not to_email and g.get("customer_id") not in (None, "", 0, "0"):
                try:
                    to_email = db.get_customer_email(g.get("customer_id"))
                except Exception:
                    to_email = ""
            if not to_email:
                skipped_no_email += 1
                if len(details) < 200:
                    inv_items = []
                    try:
                        inv_items = _invoice_items(invs, purchase_info_map={})
                    except Exception:
                        inv_items = []
                    details.append(
                        {
                            "status": "skipped_no_email",
                            "customer_id": g.get("customer_id"),
                            "customer_name": str(g.get("customer_name") or "").strip(),
                            "to_email": "",
                            "invoice_count": len(invs),
                            "invoices": inv_items[:80],
                            "invoices_truncated": bool(len(inv_items) > 80),
                            "emails": [],
                            "missing_count": 0,
                        }
                    )
                else:
                    details_truncated = True
                continue
            g["_resolved_email"] = to_email
            targets.append(g)

        total_targets = len(targets)
        emails_planned = total_targets

        def _set_progress(current: int, total: int, customer_name: str = "", to_email: str = "", detail: str = ""):
            def _do():
                overlay = _ensure_busy_overlay(self)
                msg = f"Enviando {current} de {total}" if total > 0 else "Enviando..."
                overlay.set_message(msg)
                extra = []
                if customer_name or to_email:
                    extra.append(f"{customer_name} <{to_email}>".strip())
                if detail:
                    extra.append(detail)
                extra.append(f"Enviados: {emails_sent} | Falhas: {failed} | Sem e-mail: {skipped_no_email}")
                overlay.set_detail(" | ".join([e for e in extra if e]))
                overlay.set_progress(current, total if total > 0 else None)

            try:
                self.after(0, _do)
            except Exception:
                pass

        for i, g in enumerate(targets, start=1):
            invs: List[InvoiceRow] = g.get("invoices") or []
            if not invs:
                continue
            to_email = (g.get("_resolved_email") or "").strip()
            if not to_email:
                skipped_no_email += 1
                continue
            _set_progress(i, total_targets, invs[0].customer_name, to_email, "Preparando anexos")

            purchase_map = {}
            try:
                invoice_ids = [inv.invoice_id for inv in invs if inv.invoice_id not in (None, "", 0, "0")]
                purchase_map = Database(self.config_data).get_purchase_info_bulk(invoice_ids)
            except Exception:
                purchase_map = {}

            attachments: List[Tuple[bytes, str]] = []
            missing = 0
            signature_files: List[Tuple[bytes, str]] = []
            for inv in invs:
                boleto_data = boleto_map.get(inv.invoice_id) or {}
                try:
                    if boleto_data.get("exists"):
                        attachment_data = boleto_data.get("attachment_data")
                        filename = boleto_data.get("filename") or f"boleto_{inv.invoice_id}.pdf"
                        data = None
                        if include_pix_qrcode and attachment_data:
                            try:
                                data = bytes(attachment_data)
                            except Exception:
                                data = None
                        if not data:
                            try:
                                data = build_boleto_pdf_bytes(boleto_data, inv, include_pix_qrcode=include_pix_qrcode)
                            except Exception:
                                data = None
                                if include_pix_qrcode and attachment_data:
                                    try:
                                        data = bytes(attachment_data)
                                    except Exception:
                                        data = None
                        if data:
                            attachments.append((data, filename))
                        else:
                            missing += 1
                    else:
                        missing += 1
                except Exception:
                    missing += 1
                sig = signature_map.get(inv.invoice_id) or {}
                added_for_inv = False
                for a in (sig.get("attachments") or []):
                    sdata = a.get("data")
                    sname = a.get("filename")
                    if sdata and sname:
                        signature_files.append((sdata, sname))
                        added_for_inv = True
                sig_bytes = sig.get("attachment_data")
                if not added_for_inv and sig.get("exists") and sig_bytes:
                    signature_files.append((sig_bytes, sig.get("filename") or f"assinatura_{inv.invoice_id}"))
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
                    extra_inf = ""
                    try:
                        extra_inf = Database(self.config_data).get_placa_km_text_bulk([inv.invoice_id]).get(inv.invoice_id) or ""
                    except Exception:
                        extra_inf = ""
                    pdf_bytes, pdf_name = danfe_pdf_from_nfe_xml(
                        ndata,
                        fallback_suffix=str(inv.invoice_id),
                        extra_inf_cpl_text=extra_inf,
                    )
                    if pdf_bytes and pdf_name:
                        nfe_atts.append({"data": pdf_bytes, "filename": pdf_name, "mime_type": "application/pdf"})
                        has_pdf = True
                for a in (nfe.get("attachments") or []):
                    ndata = a.get("data")
                    nname = a.get("filename")
                    if ndata and nname:
                        attachments.append((ndata, nname))
                for a in nfe_atts:
                    ndata = a.get("data")
                    nname = a.get("filename")
                    if ndata and nname and (ndata, nname) not in attachments:
                        attachments.append((ndata, nname))
            if signature_files:
                cust_part = re.sub(r"[^0-9A-Za-z_-]+", "_", str(invs[0].customer_name or "").strip())[:40]
                zip_name = f"assinaturas_{cust_part}.zip" if cust_part else "assinaturas.zip"
                zip_bytes, _ = zip_named_files(signature_files, zip_filename=zip_name)
                attachments.append((zip_bytes, zip_name))
            attachments_total += len(attachments)
            missing_total += missing

            if dry_run:
                continue
            subject = f"Alerta de vencimento de boleto - {invs[0].customer_name}"

            fatura_txt: Optional[Tuple[bytes, str]] = None
            try:
                txt_bytes, txt_name = build_faturas_detalhamento_txt_bytes(invs, purchase_info_map=purchase_map)
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
            batches: List[List[Tuple[bytes, str]]] = []
            current: List[Tuple[bytes, str]] = []
            current_bytes = 0
            for data, name in attachments:
                size = len(data) if data else 0
                if current and (len(current) >= usable_max_attachments or (current_bytes + size) > usable_max_bytes):
                    batches.append(current)
                    current = []
                    current_bytes = 0
                current.append((data, name))
                current_bytes += size
            if current or not attachments:
                batches.append(current)

            for idx, batch in enumerate(batches, start=1):
                if len(batches) > 1:
                    _set_progress(i, total_targets, invs[0].customer_name, to_email, f"E-mail {idx}/{len(batches)}")
                else:
                    _set_progress(i, total_targets, invs[0].customer_name, to_email, "")
                if not first_email and delay_seconds > 0:
                    time_module.sleep(delay_seconds)
                first_email = False
                flag_names = [name for data, name in batch if data]
                if fatura_txt:
                    flag_names.append(fatura_txt[1])
                flags = _detect_email_attachment_flags(flag_names)
                text_body, html_body = build_due_alert_email_body(
                    invs[0].customer_name,
                    base_date,
                    invs,
                    missing,
                    self._get_extra_body(),
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
                if fatura_txt:
                    maintype, subtype = _mime_parts_from_filename(fatura_txt[1])
                    msg.add_attachment(fatura_txt[0], maintype=maintype, subtype=subtype, filename=fatura_txt[1])
                for data, name in batch:
                    if not data:
                        continue
                    maintype, subtype = _mime_parts_from_filename(name)
                    msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=name)
                try:
                    _send_smtp_message(msg)
                    emails_sent += 1
                    _set_progress(i, total_targets, invs[0].customer_name, to_email, f"Enviado ({idx}/{len(batches)})" if len(batches) > 1 else "Enviado")
                    AuditLogger.write(self.current_user, "alerta_envio_email", f"alerta_id={self.agenda_id};cliente={invs[0].customer_name};para={to_email};titulos={len(invs)};anexos={len(batch) + (1 if fatura_txt else 0)};pix_incluido_no_boleto={'sim' if include_pix_qrcode else 'nao'}")
                    if len(details) < 200:
                        inv_items = []
                        try:
                            inv_items = _invoice_items(invs, purchase_info_map=purchase_map)
                        except Exception:
                            inv_items = []
                        att_names = [name for data, name in batch if data]
                        if fatura_txt:
                            att_names.append(fatura_txt[1])
                        details.append(
                            {
                                "status": "sent",
                                "customer_id": g.get("customer_id"),
                                "customer_name": str(g.get("customer_name") or "").strip(),
                                "to_email": to_email,
                                "invoice_count": len(invs),
                                "invoices": inv_items[:80],
                                "invoices_truncated": bool(len(inv_items) > 80),
                                "emails": [
                                    {
                                        "batch": int(idx),
                                        "batches_total": int(len(batches)),
                                        "subject": str(msg.get("Subject") or ""),
                                        "attachments": att_names[:60],
                                        "attachments_truncated": bool(len(att_names) > 60),
                                        "flags": dict(flags or {}),
                                    }
                                ],
                                "missing_count": int(missing or 0),
                            }
                        )
                    else:
                        details_truncated = True
                except Exception as e:
                    failed += 1
                    _set_progress(i, total_targets, invs[0].customer_name, to_email, f"Falha ({idx}/{len(batches)})" if len(batches) > 1 else "Falha")
                    AuditLogger.write(self.current_user, "alerta_envio_email_erro", f"alerta_id={self.agenda_id};cliente={invs[0].customer_name};para={to_email};erro={e}")
                    if len(details) < 200:
                        inv_items = []
                        try:
                            inv_items = _invoice_items(invs, purchase_info_map=purchase_map)
                        except Exception:
                            inv_items = []
                        att_names = [name for data, name in batch if data]
                        if fatura_txt:
                            att_names.append(fatura_txt[1])
                        details.append(
                            {
                                "status": "failed",
                                "customer_id": g.get("customer_id"),
                                "customer_name": str(g.get("customer_name") or "").strip(),
                                "to_email": to_email,
                                "invoice_count": len(invs),
                                "invoices": inv_items[:80],
                                "invoices_truncated": bool(len(inv_items) > 80),
                                "emails": [
                                    {
                                        "batch": int(idx),
                                        "batches_total": int(len(batches)),
                                        "subject": str(msg.get("Subject") or ""),
                                        "attachments": att_names[:60],
                                        "attachments_truncated": bool(len(att_names) > 60),
                                        "flags": dict(flags or {}),
                                        "error": str(e),
                                    }
                                ],
                                "missing_count": int(missing or 0),
                            }
                        )
                    else:
                        details_truncated = True

        return {
            "emails_sent": emails_sent,
            "skipped_no_email": skipped_no_email,
            "failed": failed,
            "attachments_total": attachments_total,
            "missing_total": missing_total,
            "emails_planned": int(emails_planned or 0),
            "details_version": 2,
            "details_truncated": bool(details_truncated),
            "details": details,
        }

    def send_now(self):
        if self._send_thread and getattr(self._send_thread, "is_alive", lambda: False)():
            messagebox.showwarning(APP_TITLE, "O envio já está em andamento.", parent=self)
            return
        try:
            base_date = self._parse_base_date()
            group_id = self._selected_group_id()
            portador_id = self._selected_portador_id()
            customer_id = self._selected_customer_id()
            before_days, after_days = self._get_days_values(60)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Erro ao preparar envio:\n\n{e}", parent=self)
            return
        due_dates = []
        if before_days > 0:
            due_dates.append(base_date + timedelta(days=before_days))
        if after_days > 0:
            due_dates.append(base_date - timedelta(days=after_days))
        due_dates = sorted({d for d in due_dates})
        if not due_dates:
            messagebox.showinfo(APP_TITLE, "Defina os dias antes/depois do vencimento para enviar.", parent=self)
            return

        def _work():
            run_dt = datetime.now()
            run_at = run_dt.isoformat(timespec="seconds")
            run_date = run_dt.date().isoformat()
            due_dates_str = ",".join([d.isoformat() for d in due_dates])
            rows: List[Dict[str, Any]] = []
            db = Database(self.config_data)
            for d in due_dates:
                rows.extend(db.list_agenda_invoices(d, group_id=group_id, portador_id=portador_id, customer_id=customer_id))
            if not rows:
                return {"empty": True, "run_at": run_at, "run_date": run_date, "due_dates": due_dates_str}
            res = self._send_for_rows_grouped(rows, dry_run=False, base_date=base_date)
            return {"empty": False, "result": res, "run_at": run_at, "run_date": run_date, "due_dates": due_dates_str}

        def _ok(payload):
            if payload.get("empty"):
                messagebox.showinfo(APP_TITLE, "Nenhum título encontrado para as regras atuais.", parent=self)
                return
            result = payload.get("result") or {}
            sent = int(result.get("emails_sent") or 0)
            skipped = int(result.get("skipped_no_email") or 0)
            failed = int(result.get("failed") or 0)
            if self.agenda_id:
                try:
                    agendas = self.config_data.get("financeiro_agendas", []) or []
                    if not isinstance(agendas, list):
                        agendas = []
                    for i, a in enumerate(agendas):
                        if not isinstance(a, dict):
                            continue
                        if str(a.get("id") or "") != str(self.agenda_id):
                            continue
                        merged = dict(a)
                        merged["last_run_date"] = str(payload.get("run_date") or "").strip()
                        merged["last_run_at"] = str(payload.get("run_at") or "").strip()
                        merged["last_due_date"] = str(payload.get("due_dates") or "").strip()
                        merged["last_late_minutes"] = 0
                        merged["last_out_of_time"] = False
                        merged["last_result"] = {
                            "emails_sent": sent,
                            "skipped_no_email": skipped,
                            "failed": failed,
                            "attachments_total": int(result.get("attachments_total") or 0),
                            "missing_total": int(result.get("missing_total") or 0),
                            "emails_planned": int(result.get("emails_planned") or 0),
                            "details_version": int(result.get("details_version") or 2),
                            "details_truncated": bool(result.get("details_truncated")),
                            "details": list(result.get("details") or []),
                        }
                        agendas[i] = merged
                        self.config_data["financeiro_agendas"] = agendas
                        ConfigManager.save(self.config_data)
                        self.config_data = ConfigManager.load()
                        self.on_config_saved(self.config_data)
                        break
                except Exception:
                    pass
            try:
                from app_core.documents_history import DocumentsHistory

                level = "info"
                if failed > 0:
                    level = "error"
                elif skipped > 0:
                    level = "warn"
                DocumentsHistory().add_event(
                    kind="manual_due_alert",
                    source="alertas_vencimento",
                    level=level,
                    title="Alerta de vencimento (execução manual)",
                    message=f"Enviados: {sent} | Sem e-mail: {skipped} | Falhas: {failed}",
                )
            except Exception:
                pass
            messagebox.showinfo(APP_TITLE, f"Envio concluído.\n\nEnviados: {sent}\nSem e-mail: {skipped}\nFalhas: {failed}", parent=self)

        def _err(e: Exception):
            try:
                from app_core.documents_history import DocumentsHistory

                DocumentsHistory().add_event(
                    kind="manual_due_alert",
                    source="alertas_vencimento",
                    level="error",
                    title="Falha no alerta de vencimento (manual)",
                    message=str(e),
                )
            except Exception:
                pass
            messagebox.showerror(APP_TITLE, f"Erro no envio:\n\n{e}", parent=self)

        run_with_busy(self, "Enviando alertas...", _work, _ok, _err)

    def _close(self):
        self.destroy()


class CentralAlertDetailsWindow(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Toplevel,
        *,
        created_at: str,
        level: str,
        source: str,
        title: str,
        message: str,
        documents: str = "",
    ):
        super().__init__(parent)
        self.title(f"{APP_TITLE} - Detalhes do alerta")
        self.geometry("860x520")
        self.minsize(720, 420)
        self.transient(parent)
        self.grab_set()

        self.created_at = str(created_at or "").strip()
        self.level = str(level or "").strip()
        self.source = str(source or "").strip()
        self.alert_title = str(title or "").strip()
        self.message = str(message or "")
        self.documents = str(documents or "")

        self._build_ui()
        self._center_window()

    def _build_ui(self):
        created_at_txt = format_iso_datetime_br(self.created_at, with_seconds=True) or "-"
        header = ttk.Frame(self, padding=(16, 12, 16, 0))
        header.pack(fill="x")
        ttk.Label(header, text="Detalhes do alerta", font=("Segoe UI", 14, "bold"), foreground="#111827").pack(side="left")
        ttk.Button(header, text="Copiar", command=self._copy).pack(side="right")
        ttk.Button(header, text="Fechar", command=self.destroy).pack(side="right", padx=(0, 8))

        body = ttk.Frame(self, padding=(16, 12, 16, 14))
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(4, weight=1)

        ttk.Label(body, text="Data/Hora").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Label(body, text=created_at_txt).grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(body, text="Nível").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Label(body, text=self.level or "-").grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(body, text="Origem").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Label(body, text=self.source or "-").grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(body, text="Título").grid(row=3, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Label(body, text=self.alert_title or "-").grid(row=3, column=1, sticky="w", pady=4)

        ttk.Label(body, text="Mensagem").grid(row=4, column=0, sticky="nw", padx=(0, 10), pady=(8, 4))
        self.msg_text = tk.Text(body, wrap="word", height=14)
        self.msg_text.grid(row=4, column=1, sticky="nsew", pady=(8, 4))
        self.msg_text.insert("1.0", self.message or "")
        try:
            self.msg_text.configure(state="disabled")
        except Exception:
            pass

        docs = str(self.documents or "").strip()
        if docs:
            body.rowconfigure(5, weight=1)
            ttk.Label(body, text="Documentos enviados").grid(row=5, column=0, sticky="nw", padx=(0, 10), pady=(8, 4))
            self.docs_text = tk.Text(body, wrap="word", height=8)
            self.docs_text.grid(row=5, column=1, sticky="nsew", pady=(8, 4))
            self.docs_text.insert("1.0", docs)
            try:
                self.docs_text.configure(state="disabled")
            except Exception:
                pass

    def _copy(self):
        created_at_txt = format_iso_datetime_br(self.created_at, with_seconds=True) or "-"
        parts = [
            f"Data/Hora: {created_at_txt}",
            f"Nível: {self.level or '-'}",
            f"Origem: {self.source or '-'}",
            f"Título: {self.alert_title or '-'}",
            "",
            str(self.message or ""),
        ]
        docs = str(self.documents or "").strip()
        if docs:
            parts.extend(["", "Documentos enviados:", docs])
        text = "\n".join(
            [
                str(x)
                for x in parts
            ]
        ).strip()
        try:
            self.clipboard_clear()
            self.clipboard_append(text)
        except Exception:
            pass

    def _center_window(self, min_x: int = 20, min_y: int = 20):
        try:
            self.update_idletasks()
        except Exception:
            pass
        width = max(self.winfo_width(), self.winfo_reqwidth())
        height = max(self.winfo_height(), self.winfo_reqheight())
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = int((screen_w - width) / 2)
        y = int((screen_h - height) / 2)
        self.geometry(f"+{max(x, min_x)}+{max(y, min_y)}")


class FinanceiroProblemasDocumentosWindow(tk.Toplevel):
    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.title(f"{APP_TITLE} - Central de alertas")
        self.geometry("1250x560")
        self.minsize(980, 400)
        self.transient(parent)
        self.grab_set()
        self.master_app = parent
        self.details_window = None
        self.show_info_var = tk.BooleanVar(value=True)
        self.show_warn_var = tk.BooleanVar(value=True)
        self.show_error_var = tk.BooleanVar(value=True)
        self._alerts_cache = []
        self._alert_by_iid = {}
        self._build_ui()
        self._load_data()
        self._center_window()

    def _build_ui(self):
        header = ttk.Frame(self, padding=(16, 12, 16, 0))
        header.pack(fill="x")
        ttk.Label(header, text="Central de alertas", font=("Segoe UI", 14, "bold"), foreground="#111827").pack(side="left")
        ttk.Button(header, text="Limpar histórico", command=self._clear_history).pack(side="right")
        ttk.Button(header, text="Atualizar", command=self._load_data).pack(side="right")
        ttk.Button(header, text="Fechar", command=self.destroy).pack(side="right", padx=(0, 8))

        filters = ttk.Frame(self, padding=(16, 8, 16, 0))
        filters.pack(fill="x")
        ttk.Label(filters, text="Tipo:").pack(side="left", padx=(0, 8))
        ttk.Checkbutton(filters, text="Info", variable=self.show_info_var, command=self._apply_filters).pack(side="left")
        ttk.Checkbutton(filters, text="Alerta", variable=self.show_warn_var, command=self._apply_filters).pack(side="left", padx=(10, 0))
        ttk.Checkbutton(filters, text="Erro", variable=self.show_error_var, command=self._apply_filters).pack(side="left", padx=(10, 0))

        body = ttk.Frame(self, padding=(12, 10, 12, 10))
        body.pack(fill="both", expand=True)

        columns = ("created_at", "level", "source", "title", "message")
        self.tree = ttk.Treeview(body, columns=columns, show="headings")
        self.tree.heading("created_at", text="Data/Hora")
        self.tree.column("created_at", width=165, anchor="center")
        self.tree.heading("level", text="Nível")
        self.tree.column("level", width=90, anchor="center")
        self.tree.heading("source", text="Origem")
        self.tree.column("source", width=150, anchor="center")
        self.tree.heading("title", text="Título")
        self.tree.column("title", width=260, anchor="w")
        self.tree.heading("message", text="Mensagem")
        self.tree.column("message", width=520, anchor="w")
        self.tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.bind("<Double-1>", self._on_double_click)

    def _load_data(self):
        from app_core.central_alerts import list_central_alerts

        cfg = {}
        try:
            cfg = getattr(self.master_app, "config_data", {}) if self.master_app is not None else {}
        except Exception:
            cfg = {}
        alerts = list_central_alerts(cfg, limit=1000)
        level_map = {"info": "Info", "warn": "Alerta", "error": "Erro"}
        source_map = {
            "docs": "Docs",
            "envio_auto_docs": "Envio auto",
            "alertas_vencimento": "Vencimento",
            "conferencia_fatura": "Conferência",
            "faturas": "Faturas a receber",
            "faturas_receber": "Faturas a receber",
        }
        cache = []
        for a in alerts:
            created_at = str(getattr(a, "created_at", "") or "").strip()
            created_at_txt = format_iso_datetime_br(created_at, with_seconds=True) or created_at
            level_raw = str(getattr(a, "level", "") or "").strip().lower()
            source_raw = str(getattr(a, "source", "") or "").strip()
            title = str(getattr(a, "title", "") or "").strip()
            message = str(getattr(a, "message", "") or "").strip()
            cache.append(
                {
                    "created_at": created_at,
                    "created_at_txt": created_at_txt,
                    "level_raw": level_raw,
                    "level_label": level_map.get(level_raw, level_raw or "-"),
                    "source_raw": source_raw,
                    "source_label": source_map.get(source_raw, source_raw or "-"),
                    "title": title,
                    "message": message,
                }
            )
        self._alerts_cache = cache
        self._apply_filters()

    def _apply_filters(self):
        levels = set()
        if bool(self.show_info_var.get()):
            levels.add("info")
        if bool(self.show_warn_var.get()):
            levels.add("warn")
        if bool(self.show_error_var.get()):
            levels.add("error")

        for item in self.tree.get_children():
            self.tree.delete(item)
        self._alert_by_iid = {}

        for a in self._alerts_cache or []:
            level_raw = str(a.get("level_raw") or "").strip().lower()
            if level_raw in ("info", "warn", "error"):
                if level_raw not in levels:
                    continue
            else:
                if "info" not in levels:
                    continue

            iid = self.tree.insert(
                "",
                "end",
                values=(
                    str(a.get("created_at_txt") or a.get("created_at") or ""),
                    str(a.get("level_label") or "-"),
                    str(a.get("source_label") or "-"),
                    str(a.get("title") or ""),
                    str(a.get("message") or ""),
                ),
            )
            self._alert_by_iid[str(iid)] = dict(a)

    @staticmethod
    def _extract_first_email(text: str) -> str:
        try:
            m = re.search(r"([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})", str(text or ""))
            return (m.group(1) or "").strip() if m else ""
        except Exception:
            return ""

    @staticmethod
    def _extract_docs_list(text: str) -> List[str]:
        raw = str(text or "")
        m = re.search(r"\bDocs:\s*([0-9,\s]+)", raw, flags=re.IGNORECASE)
        if not m:
            return []
        part = str(m.group(1) or "")
        out: List[str] = []
        for tok in re.split(r"[,\s]+", part):
            tok = str(tok or "").strip()
            if tok.isdigit():
                out.append(tok)
        return out

    def _build_documents_text(self, alert_row: Dict[str, Any]) -> str:
        created_at = str(alert_row.get("created_at") or "").strip()
        title = str(alert_row.get("title") or "").strip().lower()
        message = str(alert_row.get("message") or "")
        if "enviado" not in title and "enviad" not in message.lower():
            return ""

        docs = self._extract_docs_list(message)
        to_email = self._extract_first_email(message)

        try:
            from app_core.documents_history import DocumentsHistory
        except Exception:
            DocumentsHistory = None

        lines: List[str] = []
        if docs and DocumentsHistory is not None:
            rec_by_grid = DocumentsHistory().list_sent_by_grids(docs)
            for bg in docs:
                rec = rec_by_grid.get(str(bg))
                if rec:
                    sent_at = format_iso_datetime_br(getattr(rec, "sent_at", "") or "", with_seconds=True) or str(getattr(rec, "sent_at", "") or "").strip()
                    docno = str(getattr(rec, "documento", "") or "").strip()
                    if docno:
                        lines.append(f"- ID Boleto: {bg} | Documento: {docno} | Enviado em: {sent_at}")
                    else:
                        lines.append(f"- ID Boleto: {bg} | Enviado em: {sent_at}")
                else:
                    lines.append(f"- ID Boleto: {bg}")
            return "\n".join(lines).strip()

        if to_email and created_at and DocumentsHistory is not None:
            try:
                rows = DocumentsHistory().list_sent_for_email_around(to_email, created_at, window_minutes=10, limit=300)
            except Exception:
                rows = []
            if rows:
                for r in rows:
                    bg = str(getattr(r, "boleto_grid", "") or "").strip()
                    docno = str(getattr(r, "documento", "") or "").strip()
                    sent_at = format_iso_datetime_br(getattr(r, "sent_at", "") or "", with_seconds=True) or str(getattr(r, "sent_at", "") or "").strip()
                    if bg and docno:
                        lines.append(f"- ID Boleto: {bg} | Documento: {docno} | Enviado em: {sent_at}")
                    elif bg:
                        lines.append(f"- ID Boleto: {bg} | Enviado em: {sent_at}")
                if lines:
                    return "\n".join(lines).strip()

        m = re.search(r"\bAnexos:\s*(.+)$", message, flags=re.IGNORECASE)
        if m:
            an = str(m.group(1) or "").strip()
            if an:
                return an

        return ""

    def _on_double_click(self, event):
        item = self.tree.selection()
        if not item:
            return
        row = self._alert_by_iid.get(str(item[0])) or {}
        if row:
            created_at = str(row.get("created_at") or "")
            created_at_txt = str(row.get("created_at_txt") or created_at)
            level = str(row.get("level_label") or row.get("level_raw") or "")
            source = str(row.get("source_label") or row.get("source_raw") or "")
            title = str(row.get("title") or "Detalhes do alerta")
            message = str(row.get("message") or "")
            documents = self._build_documents_text(row)
        else:
            values = self.tree.item(item[0], "values")
            if not (values and len(values) >= 5):
                return
            created_at_txt = str(values[0] or "")
            created_at = created_at_txt
            level = str(values[1] or "")
            source = str(values[2] or "")
            title = str(values[3] or "Detalhes do alerta")
            message = str(values[4] or "")
            documents = ""
            try:
                if self.details_window is not None and self.details_window.winfo_exists():
                    try:
                        self.details_window.destroy()
                    except Exception:
                        pass
            except Exception:
                pass
        self.details_window = CentralAlertDetailsWindow(self, created_at=created_at, level=level, source=source, title=title, message=message, documents=documents)

    def _clear_history(self):
        if not messagebox.askyesno(
            APP_TITLE,
            "Deseja limpar o histórico local?\n\n"
            "Isso remove registros da Central de alertas e o histórico local de envios.\n"
            "Pode permitir reenvio de documentos já enviados (pois o deduplicador usa esse histórico).",
            parent=self,
        ):
            return
        try:
            from app_core.documents_history import DocumentsHistory

            DocumentsHistory().clear_all()
            try:
                DocumentsHistory().vacuum()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Falha ao limpar histórico:\n\n{e}", parent=self)
            return
        try:
            if self.master_app is not None and hasattr(self.master_app, "_mark_central_alerts_seen"):
                self.master_app._mark_central_alerts_seen()
            if self.master_app is not None and hasattr(self.master_app, "_check_problems"):
                self.master_app._check_problems()
        except Exception:
            pass
        self._load_data()
        messagebox.showinfo(APP_TITLE, "Histórico limpo.", parent=self)

    def _center_window(self, min_x: int = 20, min_y: int = 20):
        try:
            self.update_idletasks()
        except Exception:
            pass
        width = max(self.winfo_width(), self.winfo_reqwidth())
        height = max(self.winfo_height(), self.winfo_reqheight())
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = int((screen_w - width) / 2)
        y = int((screen_h - height) / 2)
        self.geometry(f"+{max(x, min_x)}+{max(y, min_y)}")

class FinanceiroEnviadosDocumentosWindow(tk.Toplevel):
    def __init__(self, parent: tk.Tk, config_data: Optional[Dict[str, Any]] = None):
        super().__init__(parent)
        self.config_data = config_data if isinstance(config_data, dict) else {}
        self.title(f"{APP_TITLE} - Documentos enviados")
        self.geometry("960x520")
        self.minsize(860, 420)
        self.transient(parent)
        self.grab_set()
        self._build_ui()
        self._load_data()
        self._center_window()

    def _build_ui(self):
        header = ttk.Frame(self, padding=(16, 12, 16, 0))
        header.pack(fill="x")
        ttk.Label(header, text="Documentos enviados", font=("Segoe UI", 14, "bold"), foreground="#16a34a").pack(side="left")
        ttk.Button(header, text="Atualizar", command=self._load_data).pack(side="right")
        ttk.Button(header, text="Limpar seleção p/ reenviar", command=self._clear_selected).pack(side="right", padx=(0, 8))

        body = ttk.Frame(self, padding=(12, 10, 12, 10))
        body.pack(fill="both", expand=True)

        columns = ("sent_at", "boleto_grid", "documento", "customer_name", "customer_email")
        self.tree = ttk.Treeview(body, columns=columns, show="headings")
        self.tree.heading("sent_at", text="Enviado em")
        self.tree.column("sent_at", width=160, anchor="center")
        self.tree.heading("boleto_grid", text="ID Boleto")
        self.tree.column("boleto_grid", width=90, anchor="center")
        self.tree.heading("documento", text="Documento")
        self.tree.column("documento", width=140, anchor="center")
        self.tree.heading("customer_name", text="Cliente")
        self.tree.column("customer_name", width=260, anchor="w")
        self.tree.heading("customer_email", text="E-mail")
        self.tree.column("customer_email", width=300, anchor="w")
        self.tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=vsb.set)

    def _center_window(self, min_x: int = 20, min_y: int = 20):
        try:
            self.update_idletasks()
        except Exception:
            pass
        width = max(self.winfo_width(), self.winfo_reqwidth())
        height = max(self.winfo_height(), self.winfo_reqheight())
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = int((screen_w - width) / 2)
        y = int((screen_h - height) / 2)
        self.geometry(f"+{max(x, min_x)}+{max(y, min_y)}")

    def _load_data(self):
        def _work():
            from app_core.documents_history import DocumentsHistory

            rows = DocumentsHistory().list_sent(limit=1000)
            name_by_grid: Dict[str, str] = {}

            try:
                missing_grids = [r.boleto_grid for r in rows if not str(r.customer_name or "").strip()]
                if missing_grids and isinstance(self.config_data, dict) and self.config_data:
                    from app_core.database import Database

                    db_rows = Database(self.config_data).list_boletos_by_grids(missing_grids, include_closed=True)
                    for d in db_rows or []:
                        bg = str(d.get("boleto_grid") or "").strip()
                        cliente = str(d.get("cliente") or "").strip()
                        if bg and cliente and bg not in name_by_grid:
                            name_by_grid[bg] = cliente
            except Exception:
                pass

            return {"rows": rows, "name_by_grid": name_by_grid}

        def _ok(out):
            rows = out.get("rows") if isinstance(out, dict) else []
            name_by_grid = out.get("name_by_grid") if isinstance(out, dict) else {}
            if not isinstance(rows, list):
                rows = []
            if not isinstance(name_by_grid, dict):
                name_by_grid = {}

            for item in self.tree.get_children():
                self.tree.delete(item)
            for r in rows:
                customer_name = str(getattr(r, "customer_name", "") or "").strip()
                if not customer_name:
                    bg = str(getattr(r, "boleto_grid", "") or "").strip()
                    customer_name = str(name_by_grid.get(bg) or "")
                self.tree.insert(
                    "",
                    "end",
                    values=(
                        getattr(r, "sent_at", ""),
                        getattr(r, "boleto_grid", ""),
                        getattr(r, "documento", ""),
                        customer_name,
                        getattr(r, "customer_email", ""),
                    ),
                )

        def _err(e: Exception):
            messagebox.showerror(APP_TITLE, f"Falha ao carregar histórico:\n\n{e}", parent=self)

        run_with_busy(self, "Carregando...", _work, _ok, _err)

    def _clear_selected(self):
        selected = list(self.tree.selection() or [])
        if not selected:
            messagebox.showinfo(APP_TITLE, "Selecione um ou mais itens para limpar do histórico e reenviar.", parent=self)
            return
        grids: List[str] = []
        for iid in selected:
            values = self.tree.item(iid, "values")
            if values and len(values) >= 2:
                bg = str(values[1] or "").strip()
                if bg:
                    grids.append(bg)
        grids = sorted(set(grids))
        if not grids:
            return
        if not messagebox.askyesno(
            APP_TITLE,
            f"Limpar {len(grids)} item(ns) do histórico de enviados para permitir reenviar?\n\n"
            f"Isso remove o status 'enviado' e volta para pendente.",
            parent=self,
        ):
            return
        try:
            from app_core.documents_history import DocumentsHistory

            DocumentsHistory().reset_to_pending(grids)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Falha ao limpar histórico:\n\n{e}", parent=self)
            return
        self._load_data()
        messagebox.showinfo(APP_TITLE, "Histórico limpo. Esses documentos poderão ser reenviados.", parent=self)

class FinanceiroEnvioAutomaticoDocumentosAgendamentoEditorWindow(tk.Toplevel):
    def __init__(
        self,
        master,
        *,
        schedule_id: str,
        initial_name: str,
        config_data: Dict[str, Any],
        current_user: str,
        on_saved,
    ):
        super().__init__(master)
        self.master_app = master
        self.schedule_id = str(schedule_id or "").strip()
        self.initial_name = str(initial_name or "").strip()
        self.config_data = deepcopy(config_data)
        self.current_user = current_user
        self.on_saved = on_saved

        self.name_var = tk.StringVar(value="")
        self.enabled_var = tk.BooleanVar(value=False)
        self.interval_var = tk.StringVar(value="4")
        self.first_run_time_var = tk.StringVar(value="")
        self.send_pix_qrcode_var = tk.BooleanVar(value=False)
        self.batch_size_var = tk.StringVar(value="2000")
        self._first_run_time_saved_raw = ""
        self._schedule_anchor_saved_raw = ""

        self.title(f"{APP_TITLE} - Agendamento")
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = min(920, sw - 60)
        h = min(680, sh - 80)
        w = max(740, w)
        h = max(520, h)
        self.geometry(f"{w}x{h}")
        self.minsize(min(740, w), min(520, h))
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self._close)

        self._build_ui()
        self._center_window()
        self._load_from_disk()

    def _center_window(self):
        try:
            self.update_idletasks()
        except Exception:
            pass
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - self.winfo_width()) // 2
        y = (sh - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 20)}+{max(y, 20)}")

    def _schedules_key(self) -> str:
        return "financeiro_envio_auto_documentos_agendamentos"

    def _active_id_key(self) -> str:
        return "financeiro_envio_auto_documentos_active_id"

    def _parse_first_run_time_value(self, raw: str):
        from datetime import datetime

        raw = str(raw or "").strip().lower().replace(" ", "")
        if not raw:
            return None
        raw = raw.replace("h", ":")
        if raw.endswith(":"):
            raw = raw + "00"
        if ":" not in raw:
            raw = raw + ":00"
        parts = raw.split(":")
        if len(parts) >= 2:
            raw = f"{parts[0]}:{parts[1]}"
        try:
            return datetime.strptime(raw, "%H:%M").time()
        except Exception:
            return None

    def _on_first_run_time_focus_out(self, event=None):
        raw = str(self.first_run_time_var.get() or "").strip()
        t = self._parse_first_run_time_value(raw)
        if t is not None:
            self.first_run_time_var.set(t.strftime("%H:%M"))

    def _find_schedule(self, cfg: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        schedules = cfg.get(self._schedules_key())
        if not isinstance(schedules, list):
            return None
        for s in schedules:
            if isinstance(s, dict) and str(s.get("id") or "").strip() == self.schedule_id:
                return s
        return None

    def _ensure_schedule_exists(self, cfg: Dict[str, Any]) -> Dict[str, Any]:
        schedules = cfg.get(self._schedules_key())
        if not isinstance(schedules, list):
            schedules = []
            cfg[self._schedules_key()] = schedules
        s = self._find_schedule(cfg)
        if isinstance(s, dict):
            return s
        name = self.initial_name or "Agendamento"
        s = {
            "id": self.schedule_id,
            "name": name,
            "enabled": False,
            "interval_hours": 4,
            "first_run_time": "",
            "schedule_anchor_at": "",
            "pending_batch_size": 2000,
            "send_pix_qrcode": False,
            "extra_body": "",
            "last_scan_end": "",
            "last_run_at": "",
            "last_result": {},
        }
        schedules.append(s)
        cfg[self._schedules_key()] = schedules
        return s

    def _sync_effective_cfg_from_active(self, cfg: Dict[str, Any]):
        schedules = cfg.get(self._schedules_key())
        active_id = str(cfg.get(self._active_id_key()) or "").strip()
        enabled = [s for s in schedules if isinstance(s, dict) and bool(s.get("enabled"))] if isinstance(schedules, list) else []
        active = None
        for s in enabled:
            if str(s.get("id") or "").strip() == active_id:
                active = s
                break
        if active is None and enabled:
            active = enabled[0]
        cfg["financeiro_envio_auto_documentos"] = active if isinstance(active, dict) else {"enabled": False}

    def _load_from_disk(self):
        try:
            cfg = ConfigManager.load()
        except Exception:
            cfg = deepcopy(self.config_data)
        s = self._find_schedule(cfg)
        if isinstance(s, dict):
            self.name_var.set(str(s.get("name") or "").strip() or self.initial_name or "Agendamento")
            self.enabled_var.set(bool(s.get("enabled", False)))
            self.interval_var.set(str(s.get("interval_hours") or 4).strip() or "4")
            self.first_run_time_var.set(str(s.get("first_run_time") or "").strip())
            self.send_pix_qrcode_var.set(bool(s.get("send_pix_qrcode", False)))
            self.batch_size_var.set(str(s.get("pending_batch_size") or 2000).strip() or "2000")
            self._first_run_time_saved_raw = str(s.get("first_run_time") or "").strip()
            self._schedule_anchor_saved_raw = str(s.get("schedule_anchor_at") or "").strip()
            self.extra_text.delete("1.0", "end")
            self.extra_text.insert("1.0", str(s.get("extra_body") or ""))
        else:
            self.name_var.set(self.initial_name or "Agendamento")
            self.enabled_var.set(False)
            self.interval_var.set("4")
            self.first_run_time_var.set("")
            self.send_pix_qrcode_var.set(False)
            self.batch_size_var.set("2000")
            self._first_run_time_saved_raw = ""
            self._schedule_anchor_saved_raw = ""
            self.extra_text.delete("1.0", "end")
            self.extra_text.insert("1.0", "")
        self.config_data = cfg

    def _persist(self):
        from datetime import datetime, date

        cfg = ConfigManager.load()
        s = self._ensure_schedule_exists(cfg)

        s["name"] = str(self.name_var.get() or "").strip() or str(s.get("name") or "").strip() or self.initial_name or "Agendamento"
        s["enabled"] = bool(self.enabled_var.get())
        try:
            s["interval_hours"] = max(1, min(72, int(str(self.interval_var.get() or "").strip() or 4)))
        except Exception:
            s["interval_hours"] = 4
        first_raw = str(self.first_run_time_var.get() or "").strip()
        parsed_time = self._parse_first_run_time_value(first_raw)
        prev_time = str(s.get("first_run_time") or "").strip()
        prev_anchor = str(s.get("schedule_anchor_at") or "").strip()
        s["first_run_time"] = first_raw
        if not first_raw:
            s["schedule_anchor_at"] = ""
        elif parsed_time is not None and (not prev_anchor or prev_time != first_raw):
            anchor_dt = datetime.combine(date.today(), parsed_time)
            s["schedule_anchor_at"] = anchor_dt.isoformat(timespec="seconds")
        s["send_pix_qrcode"] = bool(self.send_pix_qrcode_var.get())
        try:
            s["pending_batch_size"] = max(50, min(5000, int(str(self.batch_size_var.get() or "").strip() or 2000)))
        except Exception:
            s["pending_batch_size"] = 2000
        s["extra_body"] = self.extra_text.get("1.0", "end").strip()

        schedules = cfg.get(self._schedules_key())
        if not isinstance(schedules, list):
            schedules = [s]
        if bool(s.get("enabled")):
            for other in schedules:
                if not isinstance(other, dict):
                    continue
                if str(other.get("id") or "").strip() != self.schedule_id:
                    other["enabled"] = False
            cfg[self._active_id_key()] = self.schedule_id

        cfg[self._schedules_key()] = schedules
        self._sync_effective_cfg_from_active(cfg)
        ConfigManager.save(cfg)
        self.config_data = ConfigManager.load()

    def _save(self):
        try:
            self._persist()
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Erro ao salvar:\n\n{e}", parent=self)
            return
        try:
            if callable(self.on_saved):
                self.on_saved()
        except Exception:
            pass
        self._close()

    def _build_ui(self):
        header = ttk.Frame(self, padding=(16, 12, 16, 0))
        header.pack(fill="x")
        ttk.Label(header, text="Editar agendamento", font=("Segoe UI", 14, "bold"), foreground="#2563eb").pack(side="left")

        top = ttk.Frame(self, padding=(12, 10, 12, 10))
        top.pack(fill="x")
        ttk.Button(top, text="Salvar", command=self._save).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Cancelar", command=self._close).pack(side="left")
        ttk.Button(top, text="Excluir", command=self._delete).pack(side="right")

        body = ttk.Frame(self, padding=(12, 0, 12, 10))
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(6, weight=1)

        ttk.Label(body, text="Nome do agendamento").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Entry(body, textvariable=self.name_var, width=44).grid(row=0, column=1, sticky="w", pady=5)

        ttk.Checkbutton(body, text="Ativar este agendamento (somente 1 ativo por vez)", variable=self.enabled_var).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(0, 10)
        )

        ttk.Label(body, text="Intervalo (horas)").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Entry(body, textvariable=self.interval_var, width=12).grid(row=2, column=1, sticky="w", pady=5)

        ttk.Label(body, text="Primeira execução (HH:MM)").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=5)
        self.first_run_time_entry = ttk.Entry(body, textvariable=self.first_run_time_var, width=12)
        self.first_run_time_entry.grid(row=3, column=1, sticky="w", pady=5)
        bind_time_entry_shortcuts(self.first_run_time_entry)
        self.first_run_time_entry.bind("<FocusOut>", self._on_first_run_time_focus_out, add="+")

        ttk.Checkbutton(body, text="Incluir QRCode PIX no boleto (PDF)", variable=self.send_pix_qrcode_var).grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(2, 8)
        )

        adv = ttk.Frame(body)
        adv.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Label(adv, text="Lote pendências").pack(side="left")
        ttk.Entry(adv, textvariable=self.batch_size_var, width=10).pack(side="left", padx=(8, 0))

        extra_box = ttk.LabelFrame(body, text="Texto extra do e-mail", padding=10)
        extra_box.grid(row=6, column=0, columnspan=2, sticky="nsew", pady=(12, 0))
        extra_box.rowconfigure(0, weight=1)
        extra_box.columnconfigure(0, weight=1)
        self.extra_text = tk.Text(extra_box, wrap="word", height=6)
        self.extra_text.grid(row=0, column=0, sticky="nsew")
        vsb = ttk.Scrollbar(extra_box, orient="vertical", command=self.extra_text.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.extra_text.configure(yscrollcommand=vsb.set)

        bottom = ttk.Frame(self, padding=(12, 0, 12, 10))
        bottom.pack(fill="x")
        ttk.Label(bottom, text=f"Usuário: {self.current_user}").pack(side="right")

    def _close(self):
        self.destroy()

    def _delete(self):
        if not self.schedule_id:
            return
        if not messagebox.askyesno(
            APP_TITLE,
            "Deseja excluir este agendamento?\n\nIsso remove o agendamento da lista.",
            parent=self,
        ):
            return
        try:
            cfg = ConfigManager.load()
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Falha ao carregar configurações:\n\n{e}", parent=self)
            return
        schedules = cfg.get(self._schedules_key())
        if not isinstance(schedules, list):
            schedules = []
        new_schedules = []
        removed = False
        for s in schedules:
            if not isinstance(s, dict):
                continue
            sid = str(s.get("id") or "").strip()
            if sid and sid == self.schedule_id:
                removed = True
                continue
            new_schedules.append(s)
        if not removed:
            self._close()
            return
        cfg[self._schedules_key()] = new_schedules
        if str(cfg.get(self._active_id_key()) or "").strip() == self.schedule_id:
            cfg[self._active_id_key()] = ""
        if not str(cfg.get(self._active_id_key()) or "").strip() and new_schedules:
            cfg[self._active_id_key()] = str(new_schedules[0].get("id") or "").strip()
        self._sync_effective_cfg_from_active(cfg)
        try:
            ConfigManager.save(cfg)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Falha ao excluir:\n\n{e}", parent=self)
            return
        try:
            if callable(self.on_saved):
                self.on_saved()
        except Exception:
            pass
        self._close()

class FinanceiroEnvioAutomaticoDocumentosWindow(tk.Toplevel):
    def __init__(self, master, config_data: Dict[str, Any], current_user: str, on_config_saved):
        super().__init__(master)
        self.master_app = master
        self.config_data = deepcopy(config_data)
        self.current_user = current_user
        self.on_config_saved = on_config_saved
        self._run_thread = None
        self._info_refresh_after_id = None
        self.status_var = tk.StringVar(value="Pronto.")
        self.progress_message_var = tk.StringVar(value="")
        self.progress_detail_var = tk.StringVar(value="")
        self._progress_mode = "determinate"
        self._progress_running = False
        self.active_schedule_label_var = tk.StringVar(value="")
        self.new_docs_hint_var = tk.StringVar(value="")
        self.last_scan_var = tk.StringVar(value="")
        self.last_run_var = tk.StringVar(value="")
        self.next_run_var = tk.StringVar(value="")
        self.last_result_var = tk.StringVar(value="")
        self._selected_schedule_id = ""
        self.title(f"{APP_TITLE} - Envio automático de documentos")
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = min(1280, sw - 60)
        h = min(900, sh - 80)
        w = max(980, w)
        h = max(720, h)
        w = min(w, sw - 20)
        h = min(h, sh - 60)
        self.geometry(f"{w}x{h}")
        self.minsize(min(980, w), min(680, h))
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._build_ui()
        self._center_window()
        self._ensure_schedules()
        self._load_from_config()
        self._start_info_refresh_loop()
        try:
            self.after(50, self._set_initial_focus)
        except Exception:
            pass
        try:
            self.after(250, self._auto_list_boletos_on_open)
        except Exception:
            pass

    def _center_window(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - self.winfo_width()) // 2
        y = (sh - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 20)}+{max(y, 20)}")

    def _auto_list_boletos_on_open(self):
        try:
            self._list_boletos(silent=True)
        except Exception:
            pass

    def _auto_cfg(self) -> Dict[str, Any]:
        cfg = self.config_data.get("financeiro_envio_auto_documentos")
        if not isinstance(cfg, dict):
            cfg = {}
            self.config_data["financeiro_envio_auto_documentos"] = cfg
        return cfg

    def _schedules_key(self) -> str:
        return "financeiro_envio_auto_documentos_agendamentos"

    def _active_id_key(self) -> str:
        return "financeiro_envio_auto_documentos_active_id"

    def _ensure_schedules(self):
        schedules = self.config_data.get(self._schedules_key())
        active_id = str(self.config_data.get(self._active_id_key()) or "").strip()
        if isinstance(schedules, list):
            if not schedules:
                self._sync_effective_cfg_from_active()
                return
            cleaned = []
            for s in schedules:
                if not isinstance(s, dict):
                    continue
                sid = str(s.get("id") or "").strip()
                if not sid:
                    continue
                cleaned.append(s)
            if cleaned != schedules:
                self.config_data[self._schedules_key()] = cleaned
                schedules = cleaned
            if not active_id:
                active_id = str(cleaned[0].get("id") or "").strip()
                self.config_data[self._active_id_key()] = active_id
            enabled_ids = [str(s.get("id") or "").strip() for s in schedules if isinstance(s, dict) and s.get("enabled")]
            if len(enabled_ids) > 1:
                keep_id = active_id if active_id in enabled_ids else enabled_ids[0]
                for s in schedules:
                    if not isinstance(s, dict):
                        continue
                    sid = str(s.get("id") or "").strip()
                    if sid and sid != keep_id:
                        s["enabled"] = False
                self.config_data[self._active_id_key()] = keep_id
            self._sync_effective_cfg_from_active()
            return

        legacy = self._auto_cfg()
        from uuid import uuid4

        sid = uuid4().hex
        schedule = {"id": sid, "name": "Padrão"}
        for k in (
            "enabled",
            "interval_hours",
            "first_run_time",
            "schedule_anchor_at",
            "pending_batch_size",
            "send_pix_qrcode",
            "extra_body",
            "last_scan_end",
            "last_run_at",
            "last_result",
        ):
            if k in legacy:
                schedule[k] = legacy.get(k)
        self.config_data[self._schedules_key()] = [schedule]
        self.config_data[self._active_id_key()] = sid
        self._sync_effective_cfg_from_active()
        try:
            ConfigManager.save(self.config_data)
            self.config_data = ConfigManager.load()
        except Exception:
            pass

    def _schedules(self) -> List[Dict[str, Any]]:
        schedules = self.config_data.get(self._schedules_key())
        if not isinstance(schedules, list):
            schedules = []
            self.config_data[self._schedules_key()] = schedules
        out: List[Dict[str, Any]] = []
        for s in schedules:
            if isinstance(s, dict) and str(s.get("id") or "").strip():
                out.append(s)
        self.config_data[self._schedules_key()] = out
        return out

    def _active_schedule_id(self) -> str:
        return str(self.config_data.get(self._active_id_key()) or "").strip()

    def _set_active_schedule_id(self, schedule_id: str):
        self.config_data[self._active_id_key()] = str(schedule_id or "").strip()

    def _get_schedule(self, schedule_id: str) -> Optional[Dict[str, Any]]:
        schedule_id = str(schedule_id or "").strip()
        if not schedule_id:
            return None
        for s in self._schedules():
            if str(s.get("id") or "").strip() == schedule_id:
                return s
        return None

    def _get_enabled_schedule(self) -> Optional[Dict[str, Any]]:
        schedules = self._schedules()
        enabled = [s for s in schedules if isinstance(s, dict) and bool(s.get("enabled"))]
        if not enabled:
            return None
        active_id = self._active_schedule_id()
        for s in enabled:
            if str(s.get("id") or "").strip() == active_id:
                return s
        return enabled[0]

    def _sync_effective_cfg_from_active(self):
        active = self._get_enabled_schedule()
        if active is None:
            self.config_data["financeiro_envio_auto_documentos"] = {"enabled": False}
            return
        self.config_data["financeiro_envio_auto_documentos"] = active

    def _load_from_config(self):
        schedules = self._schedules()
        if not self._selected_schedule_id:
            sid = str(self._active_schedule_id() or "").strip()
            if not sid and schedules:
                sid = str(schedules[0].get("id") or "").strip()
            self._selected_schedule_id = sid
        active = self._get_enabled_schedule()
        if active is not None:
            self.active_schedule_label_var.set(str(active.get("name") or ""))
        else:
            self.active_schedule_label_var.set("Nenhum")
        cfg = active if isinstance(active, dict) else {}
        self.last_scan_var.set(str(cfg.get("last_scan_end") or "").strip())
        self.last_run_var.set(str(cfg.get("last_run_at") or "").strip())
        base_raw = str(cfg.get("last_run_at") or cfg.get("last_scan_end") or "").strip()
        self.new_docs_hint_var.set(f"Novos desde: {base_raw}" if base_raw else "Novos desde: (primeira execução)")
        lr = cfg.get("last_result") if isinstance(cfg.get("last_result"), dict) else {}
        if lr:
            self.last_result_var.set(
                f"docs_encontrados={lr.get('discovered')} pendentes_antes={lr.get('pending_before')} emails_planejados={lr.get('emails_planned')} emails_enviados={lr.get('emails_sent')} falhas_email={lr.get('failed_emails')} docs_enviados={lr.get('docs_sent')} docs_falha={lr.get('docs_failed')} sem_email={lr.get('docs_no_email')}"
            )
        else:
            self.last_result_var.set("")
        self._update_next_run_var()
        try:
            self._refresh_schedules_tree()
        except Exception:
            pass

    def _parse_first_run_time_value(self, raw: str):
        from datetime import datetime

        raw = str(raw or "").strip().lower().replace(" ", "")
        if not raw:
            return None
        raw = raw.replace("h", ":")
        if raw.endswith(":"):
            raw = raw + "00"
        if ":" not in raw:
            raw = raw + ":00"
        parts = raw.split(":")
        if len(parts) >= 2:
            raw = f"{parts[0]}:{parts[1]}"
        try:
            return datetime.strptime(raw, "%H:%M").time()
        except Exception:
            return None

    def _update_next_run_var(self):
        active = self._get_enabled_schedule()
        if not isinstance(active, dict) or not bool(active.get("enabled")):
            self.next_run_var.set("Desativado")
            return
        self.next_run_var.set(self._compute_next_run_text_for_schedule(active))

    def _refresh_last_info_from_disk(self):
        try:
            cfg = ConfigManager.load()
            if isinstance(cfg, dict):
                self.config_data = cfg
            schedules = cfg.get(self._schedules_key()) if isinstance(cfg, dict) else None
            active_id = str(cfg.get(self._active_id_key()) or "").strip()
            enabled = [s for s in schedules if isinstance(s, dict) and bool(s.get("enabled"))] if isinstance(schedules, list) else []
            schedule = None
            for s in enabled:
                if str(s.get("id") or "").strip() == active_id:
                    schedule = s
                    break
            if schedule is None and enabled:
                schedule = enabled[0]
            if schedule is None or not isinstance(schedule, dict):
                schedule = {}
            self.active_schedule_label_var.set(str(schedule.get("name") or "Nenhum") if bool(schedule.get("enabled")) else "Nenhum")
            self.last_scan_var.set(str(schedule.get("last_scan_end") or "").strip())
            self.last_run_var.set(str(schedule.get("last_run_at") or "").strip())
            base_raw = str(schedule.get("last_run_at") or schedule.get("last_scan_end") or "").strip()
            self.new_docs_hint_var.set(f"Novos desde: {base_raw}" if base_raw else "Novos desde: (primeira execução)")
            lr = schedule.get("last_result") if isinstance(schedule.get("last_result"), dict) else {}
            if lr:
                self.last_result_var.set(
                    f"docs_encontrados={lr.get('discovered')} pendentes_antes={lr.get('pending_before')} emails_planejados={lr.get('emails_planned')} emails_enviados={lr.get('emails_sent')} falhas_email={lr.get('failed_emails')} docs_enviados={lr.get('docs_sent')} docs_falha={lr.get('docs_failed')} sem_email={lr.get('docs_no_email')}"
                )
            else:
                self.last_result_var.set("")
        except Exception:
            return

    def _start_info_refresh_loop(self):
        if self._info_refresh_after_id:
            try:
                self.after_cancel(self._info_refresh_after_id)
            except Exception:
                pass
            self._info_refresh_after_id = None
        self._info_refresh_after_id = self.after(30 * 1000, self._info_refresh_tick)

    def _info_refresh_tick(self):
        try:
            self._refresh_last_info_from_disk()
            self._update_next_run_var()
        except Exception:
            pass
        try:
            self._info_refresh_after_id = self.after(30 * 1000, self._info_refresh_tick)
        except Exception:
            self._info_refresh_after_id = None

    def _open_app_folder(self):
        try:
            from app_core.constants import log_dir
            os.startfile(str(log_dir()))
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Falha ao abrir pasta:\n\n{e}", parent=self)

    def _open_sent(self):
        try:
            cfg = ConfigManager.load()
        except Exception:
            cfg = self.config_data
        FinanceiroEnviadosDocumentosWindow(self, cfg)

    def _build_runtime_payload(self, schedule_cfg: Dict[str, Any]) -> Dict[str, Any]:
        payload = deepcopy(self.config_data)
        payload["financeiro_envio_auto_documentos"] = dict(schedule_cfg or {})
        return payload

    def _set_controls_running(self, running: bool):
        state = "disabled" if running else "normal"
        for b in (
            getattr(self, "btn_simulate", None),
            getattr(self, "btn_run_now", None),
            getattr(self, "btn_open_folder", None),
            getattr(self, "btn_open_sent", None),
            getattr(self, "btn_list_boletos", None),
            getattr(self, "btn_schedule_edit", None),
            getattr(self, "btn_schedule_activate", None),
            getattr(self, "btn_schedule_deactivate", None),
            getattr(self, "btn_schedule_new", None),
        ):
            if b is None:
                continue
            try:
                b.configure(state=state)
            except Exception:
                pass

    def _progress_set_mode(self, mode: str):
        mode = str(mode or "").strip().lower()
        if mode not in ("determinate", "indeterminate"):
            mode = "determinate"
        if self._progress_mode == mode:
            return
        self._progress_mode = mode
        try:
            self.progress_bar.stop()
        except Exception:
            pass
        self._progress_running = False
        try:
            self.progress_bar.configure(mode=mode)
        except Exception:
            pass

    def _progress_update(self, *, message: str = "", detail: str = "", current: Optional[int] = None, total: Optional[int] = None):
        self.progress_message_var.set(str(message or "").strip())
        self.progress_detail_var.set(str(detail or "").strip())
        if total is not None:
            try:
                total = int(total)
            except Exception:
                total = None
        if current is not None:
            try:
                current = int(current)
            except Exception:
                current = None

        determinate = bool(total and total > 0 and current is not None)
        if determinate:
            self._progress_set_mode("determinate")
            try:
                self.progress_bar.configure(maximum=total)
            except Exception:
                pass
            try:
                self.progress_bar.configure(value=max(0, min(current, total)))
            except Exception:
                pass
            try:
                self.progress_bar.update_idletasks()
            except Exception:
                pass
            return

        self._progress_set_mode("indeterminate")
        try:
            self.progress_bar.configure(value=0)
        except Exception:
            pass
        if not self._progress_running:
            try:
                self.progress_bar.start(10)
                self._progress_running = True
            except Exception:
                self._progress_running = False
        try:
            self.progress_bar.update_idletasks()
        except Exception:
            pass

    def _progress_clear(self):
        self.progress_message_var.set("")
        self.progress_detail_var.set("")
        self._progress_set_mode("determinate")
        try:
            self.progress_bar.configure(maximum=1, value=0)
        except Exception:
            pass

    def _apply_progress_event(self, ev2: Dict[str, Any], *, auto: bool = False):
        if not self.winfo_exists():
            return
        ev2 = ev2 if isinstance(ev2, dict) else {}
        stage = str(ev2.get("stage") or "").strip()
        message = str(ev2.get("message") or "").strip()
        customer_name = str(ev2.get("customer_name") or "").strip()
        to_email = str(ev2.get("to_email") or "").strip()
        current = ev2.get("current")
        total = ev2.get("total")
        batch_idx = ev2.get("batch_idx")
        batch_total = ev2.get("batch_total")
        emails_sent = ev2.get("emails_sent")
        failed_emails = ev2.get("failed_emails")
        docs_no_email = ev2.get("docs_no_email")

        if auto:
            if stage == "start":
                self._set_controls_running(True)
                self.status_var.set("Executando automaticamente...")
            elif stage == "done":
                self._set_controls_running(False)
                self.status_var.set("Pronto.")
                current = 1
                total = 1
                if not message:
                    message = "Concluído."

        header = message or "Processando..."
        if total not in (None, "", 0, "0") and current is not None and stage in ("target_start", "target_ready", "batch_sending", "batch_sent", "batch_failed"):
            try:
                header = f"Enviando {int(current)} de {int(total)}"
            except Exception:
                header = header
        elif stage == "done":
            header = "Concluído."

        parts = []
        if customer_name or to_email:
            parts.append(f"{customer_name} <{to_email}>".strip())
        if batch_total not in (None, "", 0, "0") and batch_idx not in (None, "", 0, "0"):
            try:
                bt = int(batch_total)
                bi = int(batch_idx)
                if bt > 1:
                    parts.append(f"E-mail {bi}/{bt}")
            except Exception:
                pass
        if message and header != message and stage not in ("start", "done"):
            parts.append(message)

        try:
            es = int(emails_sent or 0)
        except Exception:
            es = 0
        try:
            fe = int(failed_emails or 0)
        except Exception:
            fe = 0
        try:
            ne = int(docs_no_email or 0)
        except Exception:
            ne = 0
        parts.append(f"Enviados: {es} | Falhas: {fe} | Sem e-mail: {ne}")

        self._progress_update(message=header, detail=" | ".join([p for p in parts if p]), current=current, total=total)

    def _run(self, dry_run: bool, force: bool = False, allow_resend: bool = False):
        if self._run_thread is not None and getattr(self._run_thread, "is_alive", lambda: False)():
            return
        cancelled = {"value": False}

        def _on_progress(ev: Dict[str, Any]):
            ev_snapshot = ev if isinstance(ev, dict) else {}
            try:
                self.after(0, lambda ev2=ev_snapshot: (None if cancelled["value"] else self._apply_progress_event(ev2, auto=False)))
            except Exception:
                pass

        def _work():
            from app_core.auto_documents import run_auto_documents
            try:
                self.config_data = ConfigManager.load()
            except Exception:
                pass
            active = self._get_enabled_schedule()
            if not isinstance(active, dict) or not bool(active.get("enabled")):
                return {"result": {"skipped": True, "reason": "disabled"}, "payload": None, "schedule_id": ""}
            schedule_id = str(active.get("id") or "").strip()
            payload = self._build_runtime_payload(active)
            res = run_auto_documents(
                payload,
                dry_run=dry_run,
                user_label=self.current_user,
                force=bool(force),
                allow_resend=bool(allow_resend),
                progress_cb=_on_progress,
            )
            return {"result": res, "payload": payload, "schedule_id": schedule_id}

        def _ok(out):
            self._set_controls_running(False)
            res = out.get("result") if isinstance(out, dict) else out
            payload = out.get("payload") if isinstance(out, dict) else None
            schedule_id = out.get("schedule_id") if isinstance(out, dict) else ""
            schedule_id = str(schedule_id or "").strip()

            if not dry_run:
                try:
                    if isinstance(payload, dict):
                        src = payload.get("financeiro_envio_auto_documentos")
                        if isinstance(src, dict):
                            schedules = self._schedules()
                            target = None
                            for s in schedules:
                                if isinstance(s, dict) and str(s.get("id") or "").strip() == schedule_id:
                                    target = s
                                    break
                            if isinstance(target, dict):
                                target["last_scan_end"] = src.get("last_scan_end")
                                target["last_run_at"] = src.get("last_run_at")
                                target["last_result"] = src.get("last_result")
                            self.config_data[self._schedules_key()] = schedules
                            self._sync_effective_cfg_from_active()
                    try:
                        ConfigManager.save(self.config_data)
                        self.config_data = ConfigManager.load()
                        self.on_config_saved(self.config_data)
                    except Exception:
                        pass
                except Exception:
                    pass
            self._load_from_config()
            self._update_next_run_var()
            self.status_var.set("Concluído.")
            self._progress_update(message="Concluído.", detail="", current=1, total=1)
            title = "Simulação concluída" if dry_run else "Execução concluída"
            if isinstance(res, dict) and res.get("skipped"):
                reason = str(res.get("reason") or "").strip()
                if reason == "disabled":
                    messagebox.showinfo(
                        APP_TITLE,
                        "Envio não executado porque não há agendamento ativo.\n\nUse Ativar na lista de agendamentos (ou marque como ativo ao editar o agendamento) para habilitar a execução automática. Para executar uma vez manualmente, use \"Executar agora\".",
                        parent=self,
                    )
                    return
                if reason == "already_running":
                    messagebox.showinfo(APP_TITLE, "Envio não executado porque já existe uma execução em andamento.", parent=self)
                    return
            if isinstance(res, dict):
                sent_to = res.get("sent_to") if isinstance(res.get("sent_to"), list) else []
                sent_to_text = ", ".join([str(x) for x in sent_to if str(x).strip()]) if sent_to else ""
                planned_to = res.get("planned_to") if isinstance(res.get("planned_to"), list) else []
                planned_to_text = ", ".join([str(x) for x in planned_to if str(x).strip()]) if planned_to else ""
                skipped_dups = int(res.get("skipped_duplicates") or 0)
                already_sent = int(res.get("already_sent") or 0)
                lines = [
                    title + ".",
                    "",
                    f"Docs encontrados: {res.get('discovered')}",
                    f"Pendentes antes: {res.get('pending_before')}",
                    f"E-mails planejados: {res.get('emails_planned')}",
                    f"E-mails enviados: {res.get('emails_sent')} (falhas: {res.get('failed_emails')})",
                    f"Docs enviados: {res.get('docs_sent')} (falhas: {res.get('docs_failed')})",
                    f"Sem e-mail: {res.get('docs_no_email')}",
                ]
                if (skipped_dups + already_sent) > 0:
                    lines.append(f"Duplicados ignorados: {skipped_dups + already_sent}")
                if planned_to_text:
                    lines.append(f"Planejado para: {planned_to_text}")
                if sent_to_text:
                    lines.append(f"Enviado para: {sent_to_text}")
                if not dry_run and int(res.get("emails_sent") or 0) > 0 and int(res.get("failed_emails") or 0) == 0:
                    lines.append("")
                    lines.append("Se não chegou, verifique Spam/Lixo eletrônico. Para ver os logs, use o botão Abrir pasta de logs (system.log).")
                try:
                    from app_core.documents_history import DocumentsHistory

                    if not dry_run:
                        level = "info"
                        if int(res.get("failed_emails") or 0) > 0 or int(res.get("docs_failed") or 0) > 0:
                            level = "error"
                        elif int(res.get("docs_no_email") or 0) > 0:
                            level = "warn"
                        DocumentsHistory().add_event(
                            kind="manual_auto_docs_run",
                            source="envio_auto_docs",
                            level=level,
                            title="Envio automático de documentos (execução manual)",
                            message=f"E-mails enviados: {int(res.get('emails_sent') or 0)} | Falhas: {int(res.get('failed_emails') or 0)} | Docs enviados: {int(res.get('docs_sent') or 0)} | Docs falha: {int(res.get('docs_failed') or 0)} | Sem e-mail: {int(res.get('docs_no_email') or 0)}",
                        )
                except Exception:
                    pass
                messagebox.showinfo(APP_TITLE, "\n".join(lines), parent=self)
                return
            messagebox.showinfo(APP_TITLE, f"{title}.\n\n{res}", parent=self)

        def _err(e: Exception):
            self._set_controls_running(False)
            self.status_var.set("Falha.")
            self._progress_update(message="Falha.", detail=str(e), current=1, total=1)
            try:
                from app_core.documents_history import DocumentsHistory

                DocumentsHistory().add_event(
                    kind="manual_auto_docs_run",
                    source="envio_auto_docs",
                    level="error",
                    title="Falha no envio automático (manual)",
                    message=str(e),
                )
            except Exception:
                pass
            messagebox.showerror(APP_TITLE, f"Erro:\n\n{e}", parent=self)

        self.status_var.set("Processando...")
        self._set_controls_running(True)
        self._progress_update(message="Processando...", detail="", current=None, total=None)

        def _run_bg():
            try:
                result = _work()
            except Exception as e:
                try:
                    self.after(0, lambda: _err(e))
                except Exception:
                    cancelled["value"] = True
                return
            try:
                self.after(0, lambda: _ok(result))
            except Exception:
                cancelled["value"] = True

        t = threading.Thread(target=_run_bg, daemon=True)
        self._run_thread = t
        t.start()

    def _simulate(self):
        self._run(True)

    def _run_now(self):
        if not messagebox.askyesno(APP_TITLE, "Deseja executar o envio agora?", parent=self):
            return
        try:
            self.config_data = ConfigManager.load()
        except Exception:
            pass
        active = self._get_enabled_schedule()
        if not isinstance(active, dict) or not bool(active.get("enabled")):
            messagebox.showinfo(APP_TITLE, "Nenhum agendamento ativo para executar.", parent=self)
            return
        allow_resend = False
        try:
            from datetime import datetime, timedelta
            from app_core.database import Database
            from app_core.documents_history import DocumentsHistory

            payload = self._build_runtime_payload(active)
            auto_cfg = payload.get("financeiro_envio_auto_documentos") if isinstance(payload.get("financeiro_envio_auto_documentos"), dict) else {}
            try:
                interval_hours = int(auto_cfg.get("interval_hours") or 4)
            except Exception:
                interval_hours = 4
            now = datetime.now()
            window_start = now - timedelta(hours=interval_hours)
            rows = Database(payload).list_generated_boletos(window_start, now)
            grids = [str(r.get("boleto_grid") or "").strip() for r in rows if str(r.get("boleto_grid") or "").strip()]
            movto_ids = [str(r.get("movto_id") or "").strip() for r in rows if str(r.get("movto_id") or "").strip()]
            history = DocumentsHistory()
            sent_by_grid = history.list_sent_by_grids(grids)
            sent_by_movto = history.list_sent_by_movto_ids(movto_ids)
            duplicates = []
            for r in rows:
                bg = str(r.get("boleto_grid") or "").strip()
                mid = str(r.get("movto_id") or "").strip()
                rec = sent_by_grid.get(bg) or (sent_by_movto.get(mid) if mid else None)
                if not rec:
                    continue
                duplicates.append(
                    {
                        "cliente": str(r.get("cliente") or "").strip(),
                        "documento": str(r.get("documento") or "").strip(),
                        "email": str(r.get("customer_email") or "").strip(),
                        "sent_at": str(rec.sent_at or "").strip(),
                    }
                )
            if duplicates:
                preview = []
                for d in duplicates[:5]:
                    preview.append(f"- {d.get('cliente') or '-'} | {d.get('documento') or '-'} | {d.get('email') or '-'} | enviado em {d.get('sent_at') or '-'}")
                msg = (
                    f"Foram encontrados {len(duplicates)} documentos que já constam como enviados no histórico.\n\n"
                    f"Deseja reenviar mesmo assim?\n\n"
                    + "\n".join(preview)
                    + ("\n\n(Exibindo até 5 exemplos.)" if len(duplicates) > 5 else "")
                )
                allow_resend = bool(messagebox.askyesno(APP_TITLE, msg, parent=self))
        except Exception:
            allow_resend = False
        self._run(False, force=True, allow_resend=allow_resend)

    def _build_ui(self):
        header = ttk.Frame(self, padding=(16, 12, 16, 0))
        header.pack(fill="x")
        ttk.Label(header, text="Envio automático de documentos", font=("Segoe UI", 14, "bold"), foreground="#2563eb").pack(side="left")

        top = ttk.Frame(self, padding=(12, 10, 12, 10))
        top.pack(fill="x")
        self.btn_simulate = ttk.Button(top, text="Simular", command=self._simulate)
        self.btn_simulate.pack(side="left", padx=(0, 8))
        self.btn_run_now = ttk.Button(top, text="Executar agora", command=self._run_now)
        self.btn_run_now.pack(side="left", padx=(0, 8))
        self.btn_open_folder = ttk.Button(top, text="Abrir pasta de logs", command=self._open_app_folder)
        self.btn_open_folder.pack(side="left", padx=(12, 0))
        self.btn_open_sent = ttk.Button(top, text="Ver enviados", command=self._open_sent)
        self.btn_open_sent.pack(side="left", padx=(8, 0))
        ttk.Button(top, text="Voltar ao início", command=self._close).pack(side="right")

        body = ttk.Frame(self, padding=(12, 0, 12, 10))
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(1, weight=1)

        schedules_box = ttk.LabelFrame(body, text="Agendamentos", padding=10)
        schedules_box.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        schedules_box.columnconfigure(0, weight=1)
        schedules_box.rowconfigure(0, weight=1)

        self.schedules_tree = ttk.Treeview(
            schedules_box,
            columns=("ativo", "nome", "intervalo", "primeira", "proxima", "ultima"),
            show="headings",
            height=5,
            selectmode="extended",
        )
        self.schedules_tree.heading("ativo", text="Ativo")
        self.schedules_tree.column("ativo", width=60, anchor="center")
        self.schedules_tree.heading("nome", text="Agendamento")
        self.schedules_tree.column("nome", width=260, anchor="w")
        self.schedules_tree.heading("intervalo", text="Intervalo")
        self.schedules_tree.column("intervalo", width=80, anchor="center")
        self.schedules_tree.heading("primeira", text="1ª exec.")
        self.schedules_tree.column("primeira", width=70, anchor="center")
        self.schedules_tree.heading("proxima", text="Próxima")
        self.schedules_tree.column("proxima", width=160, anchor="center")
        self.schedules_tree.heading("ultima", text="Última")
        self.schedules_tree.column("ultima", width=160, anchor="center")
        self.schedules_tree.grid(row=0, column=0, sticky="nsew")
        self.schedules_tree.bind("<Control-a>", self._select_all_schedules_rows)
        self.schedules_tree.bind("<Control-A>", self._select_all_schedules_rows)

        schedules_vsb = ttk.Scrollbar(schedules_box, orient="vertical", command=self.schedules_tree.yview)
        schedules_vsb.grid(row=0, column=1, sticky="ns")
        self.schedules_tree.configure(yscrollcommand=schedules_vsb.set)
        self.schedules_tree.tag_configure("active", background="#dbeafe")
        self.schedules_tree.bind("<<TreeviewSelect>>", self._on_schedule_selected, add="+")
        self.schedules_tree.bind("<Double-1>", self._on_schedule_double_click, add="+")

        schedules_actions = ttk.Frame(schedules_box)
        schedules_actions.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Label(schedules_actions, text="Ativo:").pack(side="left")
        ttk.Label(schedules_actions, textvariable=self.active_schedule_label_var, font=("Segoe UI", 9, "bold")).pack(side="left", padx=(6, 0))
        self.btn_schedule_edit = ttk.Button(schedules_actions, text="Editar", command=self._edit_selected_schedule)
        self.btn_schedule_edit.pack(side="left", padx=(12, 0))
        self.btn_schedule_new = ttk.Button(schedules_actions, text="Novo", command=self._new_schedule)
        self.btn_schedule_new.pack(side="right")
        self.btn_schedule_deactivate = ttk.Button(schedules_actions, text="Desativar", command=self._deactivate_selected_schedule)
        self.btn_schedule_deactivate.pack(side="right", padx=(8, 0))
        self.btn_schedule_activate = ttk.Button(schedules_actions, text="Ativar", command=self._activate_selected_schedule)
        self.btn_schedule_activate.pack(side="right", padx=(8, 0))

        list_box = ttk.LabelFrame(body, text="Boletos novos após o último envio (agendamento ativo)", padding=10)
        list_box.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(12, 0))
        list_box.rowconfigure(1, weight=1)
        list_box.columnconfigure(0, weight=1)

        list_header = ttk.Frame(list_box)
        list_header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        list_header.columnconfigure(0, weight=1)
        ttk.Label(list_header, textvariable=self.new_docs_hint_var).grid(row=0, column=0, sticky="w")
        self.btn_list_boletos = ttk.Button(list_header, text="Listar boletos novos", command=self._list_boletos)
        self.btn_list_boletos.grid(row=0, column=1, sticky="e")

        columns = ("boleto_grid", "documento", "cliente", "customer_email", "generated_at", "valor")
        self.boletos_tree = ttk.Treeview(list_box, columns=columns, show="headings", height=10, selectmode="extended")
        self.boletos_tree.heading("boleto_grid", text="ID")
        self.boletos_tree.column("boleto_grid", width=80, anchor="center")
        self.boletos_tree.heading("documento", text="Documento")
        self.boletos_tree.column("documento", width=120, anchor="center")
        self.boletos_tree.heading("cliente", text="Cliente")
        self.boletos_tree.column("cliente", width=220, anchor="w")
        self.boletos_tree.heading("customer_email", text="E-mail")
        self.boletos_tree.column("customer_email", width=250, anchor="w")
        self.boletos_tree.heading("generated_at", text="Gerado em")
        self.boletos_tree.column("generated_at", width=150, anchor="center")
        self.boletos_tree.heading("valor", text="Valor")
        self.boletos_tree.column("valor", width=110, anchor="e")
        self.boletos_tree.grid(row=1, column=0, sticky="nsew")
        self.boletos_tree.bind("<Control-a>", self._select_all_boletos_rows)
        self.boletos_tree.bind("<Control-A>", self._select_all_boletos_rows)
        
        tree_vsb = ttk.Scrollbar(list_box, orient="vertical", command=self.boletos_tree.yview)
        tree_vsb.grid(row=1, column=1, sticky="ns")
        self.boletos_tree.configure(yscrollcommand=tree_vsb.set)

        info = ttk.LabelFrame(body, text="Última execução", padding=10)
        info.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        info.columnconfigure(1, weight=1)
        ttk.Label(info, text="Último scan até").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Label(info, textvariable=self.last_scan_var).grid(row=0, column=1, sticky="w", pady=4)
        ttk.Label(info, text="Última execução").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Label(info, textvariable=self.last_run_var).grid(row=1, column=1, sticky="w", pady=4)
        ttk.Label(info, text="Próxima execução prevista").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Label(info, textvariable=self.next_run_var).grid(row=2, column=1, sticky="w", pady=4)
        ttk.Label(info, text="Resultado").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Label(info, textvariable=self.last_result_var, wraplength=1100, justify="left").grid(row=3, column=1, sticky="w", pady=4)

        bottom = ttk.Frame(self, padding=(12, 0, 12, 10))
        bottom.pack(fill="x")
        left = ttk.Frame(bottom)
        left.pack(side="left", fill="x", expand=True)
        ttk.Label(left, textvariable=self.status_var).pack(side="top", anchor="w")
        self.progress_bar = ttk.Progressbar(left, mode="determinate", length=420)
        self.progress_bar.pack(side="top", fill="x", pady=(6, 0))
        ttk.Label(left, textvariable=self.progress_message_var, font=("Segoe UI", 9, "bold")).pack(side="top", anchor="w", pady=(6, 0))
        ttk.Label(left, textvariable=self.progress_detail_var, wraplength=1100, justify="left").pack(side="top", anchor="w")
        self._progress_clear()
        self.bind("<Control-a>", self._on_ctrl_a, add="+")
        self.bind("<Control-A>", self._on_ctrl_a, add="+")

    def _set_initial_focus(self):
        try:
            self.schedules_tree.focus_set()
        except Exception:
            pass

    def _on_ctrl_a(self, event=None):
        w = None
        try:
            w = self.focus_get()
        except Exception:
            w = None
        if w is self.boletos_tree:
            return self._select_all_boletos_rows(event)
        try:
            self.schedules_tree.focus_set()
        except Exception:
            pass
        return self._select_all_schedules_rows(event)

    def _select_all_schedules_rows(self, event=None):
        try:
            items = self.schedules_tree.get_children()
            if items:
                self.schedules_tree.selection_set(items)
        except Exception:
            pass
        return "break"

    def _select_all_boletos_rows(self, event=None):
        try:
            items = self.boletos_tree.get_children()
            if items:
                self.boletos_tree.selection_set(items)
        except Exception:
            pass
        return "break"

    def _compute_next_run_text_for_schedule(self, s: Dict[str, Any]) -> str:
        from datetime import datetime, timedelta, date

        if not bool(s.get("enabled")):
            return "Desativado"
        try:
            interval_hours = int(s.get("interval_hours") or 4)
        except Exception:
            interval_hours = 4
        interval_hours = max(1, min(72, interval_hours))
        now = datetime.now()
        first_raw = str(s.get("first_run_time") or "").strip()
        anchor_raw = str(s.get("schedule_anchor_at") or "").strip()
        anchor_dt = None
        if first_raw:
            if anchor_raw:
                try:
                    anchor_dt = datetime.fromisoformat(anchor_raw)
                except Exception:
                    anchor_dt = None
            if anchor_dt is None:
                t = self._parse_first_run_time_value(first_raw)
                if t is not None:
                    anchor_dt = datetime.combine(date.today(), t)
        last_run_raw = str(s.get("last_run_at") or "").strip()
        last_run_dt = None
        if last_run_raw:
            try:
                last_run_dt = datetime.fromisoformat(last_run_raw)
            except Exception:
                last_run_dt = None
        if anchor_dt is not None:
            if now < anchor_dt:
                return anchor_dt.strftime("%Y-%m-%d %H:%M:%S")
            interval_td = timedelta(hours=interval_hours)
            elapsed = now - anchor_dt
            slots = int(elapsed.total_seconds() // interval_td.total_seconds()) if interval_td.total_seconds() > 0 else 0
            due_slot_dt = anchor_dt + (interval_td * slots)
            if last_run_dt is None or last_run_dt < due_slot_dt:
                return due_slot_dt.strftime("%Y-%m-%d %H:%M:%S") + " (atrasado)"
            next_dt = due_slot_dt + interval_td
            return next_dt.strftime("%Y-%m-%d %H:%M:%S")
        if last_run_dt is None:
            due_dt = now
        else:
            due_dt = last_run_dt + timedelta(hours=interval_hours)
        due_fmt = due_dt.strftime("%Y-%m-%d %H:%M:%S")
        return f"{due_fmt} (já pode executar)" if due_dt <= now else due_fmt

    def _refresh_schedules_tree(self):
        if not hasattr(self, "schedules_tree"):
            return
        for item in self.schedules_tree.get_children():
            self.schedules_tree.delete(item)
        schedules = self._schedules()
        active = self._get_enabled_schedule()
        active_id = str(active.get("id") or "").strip() if isinstance(active, dict) else ""
        for s in schedules:
            sid = str(s.get("id") or "").strip()
            is_active = bool(s.get("enabled"))
            name = str(s.get("name") or "").strip() or "Agendamento"
            intervalo = str(s.get("interval_hours") or "").strip()
            primeira = str(s.get("first_run_time") or "").strip()
            proxima = self._compute_next_run_text_for_schedule(s)
            ultima = str(s.get("last_run_at") or "").strip()
            tags = ("active",) if (sid and sid == active_id and is_active) else ()
            self.schedules_tree.insert("", "end", iid=sid, values=("SIM" if is_active else "", name, intervalo, primeira, proxima, ultima), tags=tags)
        if self._selected_schedule_id and self.schedules_tree.exists(self._selected_schedule_id):
            self.schedules_tree.selection_set(self._selected_schedule_id)
            self.schedules_tree.see(self._selected_schedule_id)

    def _on_schedule_selected(self, event=None):
        try:
            selected = list(self.schedules_tree.selection() or [])
        except Exception:
            selected = []
        if not selected:
            return
        sid = str(selected[0] or "").strip()
        if not sid or sid == self._selected_schedule_id:
            return
        self._selected_schedule_id = sid
        self._load_from_config()

    def _on_schedule_double_click(self, event=None):
        try:
            selected = list(self.schedules_tree.selection() or [])
        except Exception:
            selected = []
        if not selected:
            return
        sid = str(selected[0] or "").strip()
        if not sid:
            return
        self._selected_schedule_id = sid
        self._open_schedule_editor(sid)

    def _new_schedule(self):
        from uuid import uuid4

        sid = uuid4().hex
        base_name = "Agendamento"
        schedules = self._schedules()
        existing = {str(s.get("name") or "").strip() for s in schedules if isinstance(s, dict)}
        name = base_name
        if name in existing:
            i = 2
            while f"{base_name} {i}" in existing:
                i += 1
            name = f"{base_name} {i}"
        self._open_schedule_editor(sid, initial_name=name)

    def _activate_selected_schedule(self):
        schedule_id = str(self._selected_schedule_id or "").strip()
        if not schedule_id:
            return
        try:
            self.config_data = ConfigManager.load()
        except Exception:
            pass
        s = self._get_schedule(schedule_id)
        if not isinstance(s, dict):
            return
        s["enabled"] = True
        for other in self._schedules():
            if not isinstance(other, dict):
                continue
            if str(other.get("id") or "").strip() != str(s.get("id") or "").strip():
                other["enabled"] = False
        self._set_active_schedule_id(str(s.get("id") or "").strip())
        self._sync_effective_cfg_from_active()
        try:
            ConfigManager.save(self.config_data)
            self.config_data = ConfigManager.load()
            self.on_config_saved(self.config_data)
        except Exception:
            pass
        self._load_from_config()
        self.status_var.set("Agendamento ativado.")

    def _deactivate_selected_schedule(self):
        schedule_id = str(self._selected_schedule_id or "").strip()
        if not schedule_id:
            return
        try:
            self.config_data = ConfigManager.load()
        except Exception:
            pass
        s = self._get_schedule(schedule_id)
        if not isinstance(s, dict):
            return
        s["enabled"] = False
        self._sync_effective_cfg_from_active()
        try:
            ConfigManager.save(self.config_data)
            self.config_data = ConfigManager.load()
            self.on_config_saved(self.config_data)
        except Exception:
            pass
        self._load_from_config()
        self.status_var.set("Agendamento desativado.")

    def _edit_selected_schedule(self):
        sid = str(self._selected_schedule_id or "").strip()
        if not sid:
            return
        self._open_schedule_editor(sid)

    def _open_schedule_editor(self, schedule_id: str, initial_name: str = ""):
        schedule_id = str(schedule_id or "").strip()
        if not schedule_id:
            return

        def _saved():
            try:
                self.config_data = ConfigManager.load()
            except Exception:
                pass
            self._ensure_schedules()
            self._load_from_config()
            self.status_var.set("Agendamento salvo.")

        FinanceiroEnvioAutomaticoDocumentosAgendamentoEditorWindow(
            self,
            schedule_id=schedule_id,
            initial_name=str(initial_name or "").strip(),
            config_data=self.config_data,
            current_user=self.current_user,
            on_saved=_saved,
        )

    def _list_boletos(self, silent: bool = False):
        try:
            self.config_data = ConfigManager.load()
        except Exception:
            pass
        active = self._get_enabled_schedule()
        if not isinstance(active, dict) or not bool(active.get("enabled")):
            for item in self.boletos_tree.get_children():
                self.boletos_tree.delete(item)
            self.status_var.set("Nenhum agendamento ativo.")
            if not silent:
                messagebox.showinfo(APP_TITLE, "Nenhum agendamento ativo para listar boletos novos.", parent=self)
            return
        try:
            interval_hours = int(active.get("interval_hours") or 4)
        except Exception:
            interval_hours = 4

        def _work():
            from app_core.database import Database
            from datetime import datetime, timedelta
            
            now = datetime.now()
            base_raw = str(active.get("last_run_at") or "").strip() or str(active.get("last_scan_end") or "").strip()
            window_start = None
            if base_raw:
                try:
                    window_start = datetime.fromisoformat(base_raw)
                except Exception:
                    window_start = None
            if window_start is None:
                window_start = now - timedelta(hours=interval_hours)
            if window_start > now:
                window_start = now
            
            db = Database(self.config_data)
            rows = db.list_generated_boletos(window_start, now)
            return {"rows": rows, "window_start": window_start, "window_end": now}

        def _ok(out):
            rows = out.get("rows") if isinstance(out, dict) else out
            window_start = out.get("window_start") if isinstance(out, dict) else None
            window_end = out.get("window_end") if isinstance(out, dict) else None
            if not isinstance(rows, list):
                rows = []

            try:
                from app_core.documents_history import DocumentsHistory

                grids = [str(r.get("boleto_grid") or "").strip() for r in rows if str(r.get("boleto_grid") or "").strip()]
                movto_ids = [str(r.get("movto_id") or "").strip() for r in rows if str(r.get("movto_id") or "").strip()]
                history = DocumentsHistory()
                sent_by_grid = history.list_sent_by_grids(grids)
                sent_by_movto = history.list_sent_by_movto_ids(movto_ids)
                filtered = []
                for r in rows:
                    bg = str(r.get("boleto_grid") or "").strip()
                    mid = str(r.get("movto_id") or "").strip()
                    rec = sent_by_grid.get(bg) or (sent_by_movto.get(mid) if mid else None)
                    if rec:
                        continue
                    filtered.append(r)
                rows = filtered
            except Exception:
                pass

            for item in self.boletos_tree.get_children():
                self.boletos_tree.delete(item)
            
            for r in rows:
                self.boletos_tree.insert(
                    "", "end",
                    values=(
                        r.get("boleto_grid", ""),
                        r.get("documento", ""),
                        r.get("cliente", ""),
                        r.get("customer_email", ""),
                        r.get("generated_at", ""),
                        money_br(r.get("valor"))
                    )
                )
            if window_start is not None and window_end is not None:
                try:
                    self.new_docs_hint_var.set(f"Novos desde: {window_start.strftime('%Y-%m-%d %H:%M:%S')}")
                except Exception:
                    pass
            self.status_var.set(f"{len(rows)} boleto(s) novo(s) desde o último envio.")
            try:
                self.boletos_tree.focus_set()
            except Exception:
                pass

        def _err(e: Exception):
            self.status_var.set("Falha ao buscar documentos.")
            msg = str(e) if e else ""
            if not msg and e is not None:
                msg = getattr(e, "pgerror", "") or repr(e)
            if not msg:
                msg = "Erro desconhecido."
            messagebox.showerror(APP_TITLE, f"Erro:\n\n{msg}", parent=self)

        self.status_var.set("Buscando documentos...")
        run_with_busy(self, "Buscando...", _work, _ok, _err)

    def _close(self):
        if self._info_refresh_after_id:
            try:
                self.after_cancel(self._info_refresh_after_id)
            except Exception:
                pass
            self._info_refresh_after_id = None
        self.destroy()

class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.title(APP_TITLE)
        self.geometry("460x300")
        self.minsize(460, 300)
        self.config_data = ConfigManager.load()
        self.first_time = not CONFIG_PATH.exists()
        self.current_user = ""
        self.current_user_is_master = False
        self.license_data = None
        self.user_var = tk.StringVar(value="")
        self.home_message_var = tk.StringVar(value="")
        self.login_message_var = tk.StringVar(value="")
        self.inactive_window: Optional[InactiveCustomersWindow] = None
        self.invoices_window: Optional[OpenInvoicesWindow] = None
        self.alerts_window: Optional[FinanceiroAlertasWindow] = None
        self.envio_docs_window: Optional[FinanceiroEnvioAutomaticoDocumentosWindow] = None
        self._financeiro_alert_after_id = None
        self._financeiro_alert_running: set[str] = set()
        self._financeiro_alert_attempted: Dict[str, str] = {}
        self._auto_docs_after_id = None
        self._auto_docs_running = False
        self._setup_style()
        self._build_menu()
        self._build_frames()
        self._start_application_flow()
        self._center_window()

    def _center_window(self, min_x: int = 20, min_y: int = 20):
        try:
            self.update_idletasks()
        except Exception:
            pass
        width = max(self.winfo_width(), self.winfo_reqwidth())
        height = max(self.winfo_height(), self.winfo_reqheight())
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = int((screen_w - width) / 2)
        y = int((screen_h - height) / 2)
        self.geometry(f"+{max(x, min_x)}+{max(y, min_y)}")
    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        base_font = ("Segoe UI", 10)
        self.option_add("*Font", base_font)
        style.configure("TFrame", background="#f5f7fb")
        style.configure("TLabel", background="#f5f7fb", foreground="#1f2937", font=base_font)
        style.configure("TButton", font=("Segoe UI", 9, "bold"), padding=(10, 6))
        style.configure("Bell.TButton", font=("Segoe UI", 12, "bold"), padding=(6, 2), foreground="#374151", background="#f5f7fb")
        style.configure("BellAlert.TButton", font=("Segoe UI", 12, "bold"), padding=(6, 2), foreground="#b91c1c", background="#fee2e2")
        style.configure("BellWarn.TButton", font=("Segoe UI", 12, "bold"), padding=(6, 2), foreground="#92400e", background="#ffedd5")
        style.configure("Status.TFrame", background="#eef2f7")
        style.configure("Status.TLabel", background="#eef2f7", foreground="#4b5563", font=("Segoe UI", 9))
        self.configure(background="#f5f7fb")
    def _build_menu(self):
        menubar = tk.Menu(self)
        self.login_menu = tk.Menu(menubar, tearoff=0)
        self.login_menu.add_command(label="Cadastrar usuário", command=self.open_create_user)
        self.login_menu.add_command(label="Alterar senha", command=self.open_change_own_password)
        self.login_menu.add_command(label="Alterar usuários", command=self.open_master_manage_users)
        menubar.add_cascade(label="Login", menu=self.login_menu)
        self.cadastro_menu = tk.Menu(menubar, tearoff=0)
        self.cadastro_menu.add_command(label="Clientes inativos", command=self.open_inactive_customers_screen)
        menubar.add_cascade(label="Cadastro", menu=self.cadastro_menu)
        self.financeiro_menu = tk.Menu(menubar, tearoff=0)
        self.financeiro_menu.add_command(label="Faturas a receber", command=self.open_invoices_screen)
        self.financeiro_menu.add_command(label="Envio automático de documentos", command=self.open_financeiro_envio_docs_screen)
        self.financeiro_menu.add_command(label="Alertas de vencimento", command=self.open_financeiro_alertas_screen)
        menubar.add_cascade(label="Financeiro", menu=self.financeiro_menu)
        self.config_menu = tk.Menu(menubar, tearoff=0)
        self.config_menu.add_command(label="Configuração local", command=self.open_config)
        menubar.add_cascade(label="Configurações", menu=self.config_menu)
        self.config(menu=menubar)
        self._update_menu_visibility()
    def _update_menu_visibility(self):
        logged_state = "normal" if self.current_user else "disabled"
        master_state = "normal" if self.current_user_is_master else "disabled"
        self.login_menu.entryconfig("Cadastrar usuário", state=logged_state)
        self.login_menu.entryconfig("Alterar senha", state=logged_state)
        self.login_menu.entryconfig("Alterar usuários", state=master_state)
        self.cadastro_menu.entryconfig("Clientes inativos", state=logged_state)
        self.config_menu.entryconfig("Configuração local", state=logged_state)
        self.financeiro_menu.entryconfig("Faturas a receber", state=logged_state)
        self.financeiro_menu.entryconfig("Envio automático de documentos", state=logged_state)
        self.financeiro_menu.entryconfig("Alertas de vencimento", state=logged_state)
    def _build_frames(self):
        self.login_frame = ttk.Frame(self, padding=24)
        self.setup_user_frame = ttk.Frame(self, padding=24)
        self.home_frame = ttk.Frame(self, padding=0)
        self.license_block_frame = ttk.Frame(self, padding=24)
        self._build_login_frame()
        self._build_setup_user_frame()
        self._build_home_frame()
        self._build_license_block_frame()
    def _clear_frames(self):
        for frame in (self.login_frame, self.setup_user_frame, self.home_frame, self.license_block_frame):
            frame.pack_forget()
    def _start_application_flow(self):
        try:
            self.license_data = LicenseManager.validate_file()
        except Exception as e:
            self._show_license_block(str(e))
            return
        self._show_login_frame()
    def _build_license_block_frame(self):
        card = ttk.Frame(self.license_block_frame, padding=20)
        card.place(relx=0.5, rely=0.5, anchor="center")
        ttk.Label(card, text="Licença inválida", font=("Segoe UI", 12, "bold")).pack(anchor="center", pady=(0, 10))
        self.license_error_var = tk.StringVar(value="")
        ttk.Label(card, textvariable=self.license_error_var, wraplength=380, justify="center").pack(anchor="center", pady=(0, 14))
        ttk.Label(card, text=f"Arquivo esperado: {LICENSE_FILENAME}", justify="center").pack(anchor="center", pady=(0, 10))
        ttk.Button(card, text="Fechar", command=self.destroy).pack()
    def _show_license_block(self, error_text: str):
        self.geometry("460x250")
        self.minsize(460, 250)
        self._center_window()
        self.license_error_var.set(error_text)
        self._clear_frames()
        self.license_block_frame.pack(fill="both", expand=True)
    def _show_login_frame(self):
        self._clear_frames()
        self.login_frame.pack(fill="both", expand=True)
        self.after(50, lambda: self.login_user_entry.focus_set())
    def _show_setup_user_frame(self):
        self._clear_frames()
        self.setup_user_frame.pack(fill="both", expand=True)
    def _build_login_frame(self):
        card = ttk.Frame(self.login_frame, padding=20)
        card.place(relx=0.5, rely=0.5, anchor="center")
        ttk.Label(card, text="Acesso ao sistema", font=("Segoe UI", 12, "bold")).pack(anchor="center", pady=(2, 6))
        ttk.Label(card, textvariable=self.login_message_var, justify="center", wraplength=320).pack(anchor="center", pady=(0, 14))
        form = ttk.Frame(card)
        form.pack(fill="x")
        ttk.Label(form, text="Usuário").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=6)
        self.login_user_entry = ttk.Entry(form, width=24)
        self.login_user_entry.grid(row=0, column=1, sticky="ew", pady=6)
        ttk.Label(form, text="Senha").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=6)
        self.login_pass_entry = ttk.Entry(form, width=24, show="*")
        self.login_pass_entry.grid(row=1, column=1, sticky="ew", pady=6)
        form.columnconfigure(1, weight=1)
        btns = ttk.Frame(card)
        btns.pack(fill="x", pady=(18, 0))
        ttk.Button(btns, text="Sair", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Entrar", command=self._handle_login).pack(side="right", padx=(0, 8))
        self.login_pass_entry.bind("<Return>", lambda e: self._handle_login())
    def _build_setup_user_frame(self):
        card = ttk.Frame(self.setup_user_frame, padding=28)
        card.place(relx=0.5, rely=0.5, anchor="center")
        ttk.Label(card, text="Criar usuário local", font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 8))
        ttk.Label(card, text="Você pode criar agora um usuário local para acessar o sistema. Se não quiser, clique em Continuar.", wraplength=500, justify="left").pack(anchor="w", pady=(0, 16))
        form = ttk.Frame(card)
        form.pack(fill="x")
        ttk.Label(form, text="Usuário").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        self.first_user_entry = ttk.Entry(form, width=34)
        self.first_user_entry.grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Label(form, text="Senha").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        self.first_pass_entry = ttk.Entry(form, width=34, show="*")
        self.first_pass_entry.grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Label(form, text="Confirmar senha").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=5)
        self.first_confirm_entry = ttk.Entry(form, width=34, show="*")
        self.first_confirm_entry.grid(row=2, column=1, sticky="ew", pady=5)
        form.columnconfigure(1, weight=1)
        btns = ttk.Frame(card)
        btns.pack(fill="x", pady=(20, 0))
        ttk.Button(btns, text="Continuar", command=self._continue_first_access).pack(side="right")
        ttk.Button(btns, text="Criar usuário", command=self._create_first_user).pack(side="right", padx=(0, 8))
    def _build_home_frame(self):
        body = ttk.Frame(self.home_frame, padding=24)
        body.pack(fill="both", expand=True)
        
        self.bell_btn = ttk.Button(self.home_frame, text="🔔", style="Bell.TButton", command=self._open_problems_window)
        self.bell_btn.place(relx=0.98, rely=0.02, anchor="ne")

        center = ttk.Frame(body)
        center.place(relx=0.5, rely=0.42, anchor="center")
        ttk.Label(center, text="DataHub", font=("Segoe UI", 18, "bold")).pack(anchor="center", pady=(0, 10))
        ttk.Label(center, text="Selecione uma opção no menu para abrir uma funcionalidade.", justify="center", wraplength=520).pack(anchor="center", pady=(0, 6))

        status_bar = ttk.Frame(self.home_frame, style="Status.TFrame")
        status_bar.pack(side="bottom", fill="x")
        ttk.Separator(self.home_frame, orient="horizontal").pack(side="bottom", fill="x")
        ttk.Label(status_bar, textvariable=self.home_message_var, style="Status.TLabel").pack(side="left", padx=(10, 10), pady=6)
        ttk.Label(status_bar, textvariable=self.user_var, style="Status.TLabel").pack(side="right", padx=(10, 10), pady=6)

    def _open_problems_window(self):
        if hasattr(self, 'problems_window') and self.problems_window and self.problems_window.winfo_exists():
            self.problems_window.lift()
            return
        try:
            self._mark_central_alerts_seen()
        except Exception:
            pass
        self.problems_window = FinanceiroProblemasDocumentosWindow(self)
        try:
            self._check_problems()
        except Exception:
            pass

    def _mark_central_alerts_seen(self):
        try:
            from datetime import datetime
            from app_core.config_manager import ConfigManager

            self.config_data = ConfigManager.load()
            self.config_data["central_alerts_last_seen_at"] = datetime.now().isoformat(timespec="seconds")
            ConfigManager.save(self.config_data)
            self.config_data = ConfigManager.load()
        except Exception:
            return

    def _superscript_number(self, n: int) -> str:
        try:
            n = int(n or 0)
        except Exception:
            n = 0
        if n <= 0:
            return ""
        sup = {"0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴", "5": "⁵", "6": "⁶", "7": "⁷", "8": "⁸", "9": "⁹"}
        return "".join([sup.get(ch, "") for ch in str(n) if ch.isdigit()])

    def _set_bell_state(self, *, unseen: int, has_problems: bool):
        badge = self._superscript_number(unseen)
        self.bell_btn.config(text=("🔔" + badge if badge else "🔔"))
        if unseen > 0:
            self.bell_btn.config(style="BellAlert.TButton")
            return
        if has_problems:
            self.bell_btn.config(style="BellWarn.TButton")
            return
        self.bell_btn.config(style="Bell.TButton")

    def _check_problems(self):
        if not self.current_user:
            self.after(30000, self._check_problems)
            return
            
        try:
            from app_core.config_manager import ConfigManager
            from app_core.central_alerts import count_problem_central_alerts, count_unseen_central_alerts

            try:
                self.config_data = ConfigManager.load()
            except Exception:
                pass
            unseen = int(count_unseen_central_alerts(self.config_data, limit=500) or 0)
            has_problems = int(count_problem_central_alerts(self.config_data) or 0) > 0
            self._set_bell_state(unseen=unseen, has_problems=has_problems)
        except Exception:
            pass
        self.after(30000, self._check_problems)

    def show_home(self):
        self._clear_frames()
        self.home_frame.pack(fill="both", expand=True)
        self._update_home_message()
        self._check_problems()

    def _update_home_message(self):
        name = ""
        if isinstance(self.license_data, dict):
            name = str(self.license_data.get("customer_name", "")).strip()
        self.home_message_var.set(f"Licença ativa para: {name}" if name else "")
    def _handle_login(self):
        username = self.login_user_entry.get().strip()
        password = self.login_pass_entry.get()
        if not username:
            messagebox.showwarning(APP_TITLE, "Informe o usuário.", parent=self)
            return
        if not password:
            messagebox.showwarning(APP_TITLE, "Informe a senha.", parent=self)
            return
        if username == MASTER_USERNAME and password == MASTER_PASSWORD:
            AuditLogger.write(username, "login_sucesso", "tipo=master")
            self._after_login_success(username, True)
            return
        if self.first_time:
            AuditLogger.write(username or "-", "login_falha", "primeiro_acesso_sem_master")
            messagebox.showerror(APP_TITLE, "No primeiro acesso, apenas o login master é permitido.", parent=self)
            return
        if UserManager.validate_login(self.config_data, username, password):
            AuditLogger.write(username, "login_sucesso", "tipo=local")
            self._after_login_success(username, False)
            return
        AuditLogger.write(username or "-", "login_falha", "credenciais_invalidas")
        messagebox.showerror(APP_TITLE, "Usuário ou senha inválidos.", parent=self)
    def _after_login_success(self, username: str, is_master: bool):
        self.current_user = username
        self.current_user_is_master = is_master
        self.user_var.set(f"Usuário: {username}")
        self._update_menu_visibility()
        self.login_user_entry.delete(0, "end")
        self.login_pass_entry.delete(0, "end")
        if self.first_time:
            self._show_setup_user_frame()
            return
        try:
            LicenseManager.validate_against_database(self.config_data, self.license_data)
        except Exception as e:
            self._show_license_block(str(e))
            return
        self.geometry("900x520")
        self.minsize(700, 420)
        self._center_window()
        self.show_home()
        self._start_financeiro_alert_loop()
        self._start_auto_docs_loop()
    def _create_first_user(self):
        try:
            username = self.first_user_entry.get().strip()
            password = self.first_pass_entry.get()
            confirm = self.first_confirm_entry.get()
            if not username:
                raise AppError("Informe o usuário.")
            if not password:
                raise AppError("Informe a senha.")
            if password != confirm:
                raise AppError("A confirmação da senha não confere.")
            UserManager.add_user(self.config_data, username, password)
            ConfigManager.save(self.config_data)
            self.config_data = ConfigManager.load()
            AuditLogger.write(MASTER_USERNAME, "cadastrar_usuario", f"alvo={username};primeiro_acesso=sim")
            self.first_user_entry.delete(0, "end")
            self.first_pass_entry.delete(0, "end")
            self.first_confirm_entry.delete(0, "end")
            messagebox.showinfo(APP_TITLE, "Usuário local criado com sucesso.", parent=self)
            self._continue_first_access()
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e), parent=self)
    def _continue_first_access(self):
        ConfigWindow(self, self.config_data, self._after_first_connection_saved)
    def _after_first_connection_saved(self, cfg: Dict[str, Any]):
        try:
            LicenseManager.validate_against_database(cfg, self.license_data)
        except Exception as e:
            self._show_license_block(str(e))
            return
        self.config_data = cfg
        self.first_time = False
        self._update_menu_visibility()
        self.geometry("900x520")
        self.minsize(700, 420)
        self._center_window()
        self.show_home()
        self._start_financeiro_alert_loop()
        self._start_auto_docs_loop()
    def open_create_user(self):
        if not self.current_user:
            return
        CreateUserWindow(self, self.config_data, self.current_user, self._apply_security_config)
    def open_change_own_password(self):
        if not self.current_user:
            return
        if self.current_user_is_master:
            messagebox.showwarning(APP_TITLE, "A senha do usuário master é fixa no sistema e não pode ser alterada.", parent=self)
            return
        ChangeOwnPasswordWindow(self, self.config_data, self.current_user, self._apply_security_config)
    def open_master_manage_users(self):
        if not self.current_user_is_master:
            messagebox.showwarning(APP_TITLE, "Somente o usuário master pode acessar esta tela.", parent=self)
            return
        MasterManageUsersWindow(self, self.config_data, self.current_user, self._apply_security_config)

    def open_config(self):
        if not self.current_user:
            return
        ConfigWindow(self, self.config_data, self._apply_new_config)

    def open_inactive_customers_screen(self):
        if not self.current_user:
            return
        if self.inactive_window is not None and self.inactive_window.winfo_exists():
            self.inactive_window.deiconify()
            self.inactive_window.lift()
            self.inactive_window.focus_force()
            return
        self.inactive_window = InactiveCustomersWindow(self, self.config_data, self.current_user, self._apply_new_config)
    def open_invoices_screen(self):
        if not self.current_user:
            return
        if self.invoices_window is not None and self.invoices_window.winfo_exists():
            self.invoices_window.deiconify()
            self.invoices_window.lift()
            self.invoices_window.focus_force()
            return
        self.invoices_window = OpenInvoicesWindow(self, self.config_data, self.current_user, self._apply_new_config)

    def open_financeiro_alertas_screen(self):
        if not self.current_user:
            return
        if self.alerts_window is not None and self.alerts_window.winfo_exists():
            self.alerts_window.deiconify()
            self.alerts_window.lift()
            self.alerts_window.focus_force()
            return
        self.alerts_window = FinanceiroAlertasWindow(self, self.config_data, self.current_user, self._apply_new_config)

    def open_financeiro_envio_docs_screen(self):
        if not self.current_user:
            return
        if self.envio_docs_window is not None and self.envio_docs_window.winfo_exists():
            self.envio_docs_window.deiconify()
            self.envio_docs_window.lift()
            self.envio_docs_window.focus_force()
            return
        self.envio_docs_window = FinanceiroEnvioAutomaticoDocumentosWindow(self, self.config_data, self.current_user, self._apply_new_config)

    def _start_financeiro_alert_loop(self):
        if self._financeiro_alert_after_id:
            try:
                self.after_cancel(self._financeiro_alert_after_id)
            except Exception:
                pass
            self._financeiro_alert_after_id = None
        self._financeiro_alert_after_id = self.after(2500, self._financeiro_alert_tick)

    def _financeiro_alert_tick(self):
        self._financeiro_alert_after_id = self.after(60 * 1000, self._financeiro_alert_tick)
        if not self.current_user:
            return
        now = datetime.now()
        today_key = date.today().isoformat()
        agendas = self.config_data.get("financeiro_agendas", []) or []
        if not isinstance(agendas, list) or not agendas:
            return

        def parse_time(value: str):
            value = str(value or "").strip()
            m = re.match(r"^(\d{1,2}):(\d{2})$", value)
            if not m:
                return time(6, 0)
            hh = int(m.group(1))
            mm = int(m.group(2))
            if hh < 0 or hh > 23 or mm < 0 or mm > 59:
                return time(6, 0)
            return time(hh, mm)

        def send_for_alert(alert_id: str, alert_cfg: Dict[str, Any]):
            started_at = datetime.now()
            target_time = parse_time(alert_cfg.get("send_time"))
            late_minutes = 0
            try:
                late_minutes = int((started_at - datetime.combine(started_at.date(), target_time)).total_seconds() // 60)
                if late_minutes < 0:
                    late_minutes = 0
            except Exception:
                late_minutes = 0
            out_of_time = late_minutes > 15
            include_pix_qrcode = bool(alert_cfg.get("send_pix_qrcode", False))
            base_date = date.today()
            try:
                days_before = int(alert_cfg.get("days_before_due") or 0)
            except Exception:
                days_before = 0
            try:
                days_after = int(alert_cfg.get("days_after_due") or 0)
            except Exception:
                days_after = 0
            days_before = max(0, min(365, days_before))
            days_after = max(0, min(365, days_after))
            if days_after > 0:
                days_before = 0
            elif days_before > 0:
                days_after = 0
            due_dates = []
            if days_before > 0:
                due_dates.append(base_date + timedelta(days=days_before))
            if days_after > 0:
                due_dates.append(base_date - timedelta(days=days_after))
            due_dates = sorted({d for d in due_dates})
            due_dates_str = ",".join([d.isoformat() for d in due_dates])
            try:
                rows: List[Dict[str, Any]] = []
                db = Database(self.config_data)
                for d in due_dates:
                    rows.extend(
                        db.list_agenda_invoices(
                            d,
                            group_id=alert_cfg.get("group_id"),
                            portador_id=alert_cfg.get("portador_id"),
                            customer_id=alert_cfg.get("customer_id"),
                        )
                    )
                extra_body = str(alert_cfg.get("extra_body") or "").strip()

                smtp_cfg = self.config_data.get("smtp", {})
                smtp_email = str(smtp_cfg.get("email", "")).strip()
                smtp_host = str(smtp_cfg.get("host", "")).strip()
                smtp_password = str(smtp_cfg.get("password", "")).strip()
                smtp_port = int(smtp_cfg.get("port", 587) or 587)
                if not smtp_email or not smtp_host or not smtp_password or not smtp_port:
                    raise AppError("SMTP não configurado.")

                def _send_smtp_message(msg: EmailMessage):
                    if smtp_port == 465:
                        context = ssl.create_default_context()
                        with smtplib.SMTP_SSL(smtp_host, smtp_port, context=context, timeout=20) as server:
                            server.login(smtp_email, smtp_password)
                            server.send_message(msg)
                    else:
                        with smtplib.SMTP(smtp_host, smtp_port, timeout=20) as server:
                            server.ehlo()
                            try:
                                server.starttls(context=ssl.create_default_context())
                                server.ehlo()
                            except Exception:
                                pass
                            server.login(smtp_email, smtp_password)
                            server.send_message(msg)

                invoices: List[InvoiceRow] = []
                for r in rows:
                    invoices.append(
                        InvoiceRow(
                            invoice_id=r.get("movto_id"),
                            company=str(r.get("empresa") or "").strip(),
                            customer_id=r.get("customer_id"),
                            customer_code=r.get("codigo_cliente"),
                            customer_name=str(r.get("cliente") or "").strip(),
                            motive_code="",
                            motive_name="",
                            account_code=str(r.get("conta") or "").strip(),
                            account_name=str(r.get("conta_nome") or "").strip(),
                            issue_date=r.get("data"),
                            due_date=r.get("vencto"),
                            amount=r.get("valor"),
                            discount_amount=r.get("valor_desconto"),
                            paid_amount=r.get("valor_baixado"),
                            open_balance=r.get("saldo_em_aberto"),
                            customer_email=str(r.get("customer_email") or "").strip(),
                        )
                    )

                invoice_ids = [i.invoice_id for i in invoices if i.invoice_id not in (None, "", 0, "0")]
                boleto_map: Dict[Any, Dict[str, Any]] = {}
                try:
                    boleto_map = Database(self.config_data).get_boletos_email_payload_bulk(invoice_ids)
                except Exception:
                    boleto_map = {}

                signature_map: Dict[Any, Dict[str, Any]] = {}
                try:
                    signature_map = Database(self.config_data).get_sale_signatures_pdf_bulk(invoice_ids)
                except Exception:
                    signature_map = {}

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
                try:
                    delay_seconds = int(smtp_cfg.get("delay_seconds", 5) or 0)
                except Exception:
                    delay_seconds = 5
                delay_seconds = max(0, min(300, delay_seconds))
                first_email = True
                for g in grouped.values():
                    invs: List[InvoiceRow] = g.get("invoices") or []
                    if not invs:
                        continue
                    to_email = (g.get("customer_email") or "").strip()
                    if not to_email and g.get("customer_id") not in (None, "", 0, "0"):
                        try:
                            to_email = Database(self.config_data).get_customer_email(g.get("customer_id"))
                        except Exception:
                            to_email = ""
                    if not to_email:
                        skipped_no_email += 1
                        continue

                    attachments: List[Tuple[bytes, str]] = []
                    missing = 0
                    signature_files: List[Tuple[bytes, str]] = []
                    for inv in invs:
                        boleto_data = boleto_map.get(inv.invoice_id) or {}
                        try:
                            if boleto_data.get("exists"):
                                attachment_data = boleto_data.get("attachment_data")
                                filename = boleto_data.get("filename") or f"boleto_{inv.invoice_id}.pdf"
                                data = None
                                if include_pix_qrcode and attachment_data:
                                    try:
                                        data = bytes(attachment_data)
                                    except Exception:
                                        data = None
                                if not data:
                                    try:
                                        data = build_boleto_pdf_bytes(boleto_data, inv, include_pix_qrcode=include_pix_qrcode)
                                    except Exception:
                                        data = None
                                        if include_pix_qrcode and attachment_data:
                                            try:
                                                data = bytes(attachment_data)
                                            except Exception:
                                                data = None
                                if data:
                                    attachments.append((data, filename))
                                else:
                                    missing += 1
                            else:
                                missing += 1
                        except Exception:
                            missing += 1
                        sig = signature_map.get(inv.invoice_id) or {}
                        added_for_inv = False
                        for a in (sig.get("attachments") or []):
                            sdata = a.get("data")
                            sname = a.get("filename")
                            if sdata and sname:
                                signature_files.append((sdata, sname))
                                added_for_inv = True
                        sig_bytes = sig.get("attachment_data")
                        if not added_for_inv and sig.get("exists") and sig_bytes:
                            signature_files.append((sig_bytes, sig.get("filename") or f"assinatura_{inv.invoice_id}"))
                    if signature_files:
                        cust_part = re.sub(r"[^0-9A-Za-z_-]+", "_", str(invs[0].customer_name or "").strip())[:40]
                        zip_name = f"assinaturas_{cust_part}.zip" if cust_part else "assinaturas.zip"
                        zip_bytes, _ = zip_named_files(signature_files, zip_filename=zip_name)
                        attachments.append((zip_bytes, zip_name))

                    emails_planned += 1
                    attachments_total += len(attachments)
                    missing_total += missing
                    purchase_map = {}
                    try:
                        invoice_ids = [inv.invoice_id for inv in invs if inv.invoice_id not in (None, "", 0, "0")]
                        purchase_map = Database(self.config_data).get_purchase_info_bulk(invoice_ids)
                    except Exception:
                        purchase_map = {}
                    subject = f"Alerta de vencimento de boleto - {invs[0].customer_name}"

                    max_attachments = 18
                    max_bytes = 15 * 1024 * 1024
                    fatura_txt: Optional[Tuple[bytes, str]] = None
                    try:
                        txt_bytes, txt_name = build_faturas_detalhamento_txt_bytes(invs, purchase_info_map=purchase_map)
                        if txt_bytes and txt_name:
                            fatura_txt = (txt_bytes, txt_name)
                    except Exception:
                        fatura_txt = None
                    if fatura_txt:
                        attachments_total += 1
                    reserved_count = 1 if fatura_txt else 0
                    reserved_bytes = len(fatura_txt[0]) if fatura_txt else 0
                    usable_max_attachments = max(1, max_attachments - reserved_count)
                    usable_max_bytes = max(1024, max_bytes - reserved_bytes)
                    batches: List[List[Tuple[bytes, str]]] = []
                    current: List[Tuple[bytes, str]] = []
                    current_bytes = 0
                    for data, name in attachments:
                        size = len(data) if data else 0
                        if current and (len(current) >= usable_max_attachments or (current_bytes + size) > usable_max_bytes):
                            batches.append(current)
                            current = []
                            current_bytes = 0
                        current.append((data, name))
                        current_bytes += size
                    if current or not attachments:
                        batches.append(current)

                    for idx, batch in enumerate(batches, start=1):
                        if not first_email and delay_seconds > 0:
                            time_module.sleep(delay_seconds)
                        first_email = False
                        flag_names = [name for data, name in batch if data]
                        if fatura_txt:
                            flag_names.append(fatura_txt[1])
                        flags = _detect_email_attachment_flags(flag_names)
                        text_body, html_body = build_due_alert_email_body(
                            invs[0].customer_name,
                            base_date,
                            invs,
                            missing,
                            extra_body,
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
                        if fatura_txt:
                            maintype, subtype = _mime_parts_from_filename(fatura_txt[1])
                            msg.add_attachment(fatura_txt[0], maintype=maintype, subtype=subtype, filename=fatura_txt[1])
                        for data, name in batch:
                            if not data:
                                continue
                            maintype, subtype = _mime_parts_from_filename(name)
                            msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=name)
                        try:
                            _send_smtp_message(msg)
                            emails_sent += 1
                            AuditLogger.write(self.current_user, "alerta_envio_email_auto", f"alerta_id={alert_id};cliente={invs[0].customer_name};para={to_email};titulos={len(invs)};anexos={len(batch) + (1 if fatura_txt else 0)};pix_incluido_no_boleto={'sim' if include_pix_qrcode else 'nao'}")
                        except Exception as e:
                            failed += 1
                            AuditLogger.write(self.current_user, "alerta_envio_email_auto_erro", f"alerta_id={alert_id};cliente={invs[0].customer_name};para={to_email};erro={e}")

                updated = False
                new_agendas = []
                for a in (self.config_data.get("financeiro_agendas", []) or []):
                    if isinstance(a, dict) and str(a.get("id") or "") == str(alert_id):
                        merged = dict(a)
                        merged["last_run_date"] = today_key
                        merged["last_run_at"] = started_at.isoformat(timespec="seconds")
                        merged["last_due_date"] = due_dates_str
                        merged["last_late_minutes"] = late_minutes
                        merged["last_out_of_time"] = bool(out_of_time)
                        merged["last_result"] = {"emails_sent": emails_sent, "skipped_no_email": skipped_no_email, "failed": failed, "attachments_total": attachments_total, "missing_total": missing_total, "emails_planned": emails_planned}
                        new_agendas.append(merged)
                        updated = True
                    else:
                        new_agendas.append(a)
                if updated:
                    self.config_data["financeiro_agendas"] = new_agendas
                    ConfigManager.save(self.config_data)
                    self.config_data = ConfigManager.load()
                AuditLogger.write(self.current_user, "alerta_execucao", f"alerta_id={alert_id};due_dates={due_dates_str};enviados={emails_sent};sem_email={skipped_no_email};falhas={failed};atraso_min={late_minutes}")
            except Exception as e:
                AuditLogger.write(self.current_user, "alerta_execucao_erro", f"alerta_id={alert_id};erro={e}")
            finally:
                try:
                    self._financeiro_alert_running.discard(str(alert_id))
                except Exception:
                    pass

        for a in agendas:
            if not isinstance(a, dict):
                continue
            if not a.get("enabled"):
                continue
            alert_id = str(a.get("id") or "")
            if not alert_id:
                continue
            if alert_id in self._financeiro_alert_running:
                continue
            if str(self._financeiro_alert_attempted.get(alert_id) or "") == today_key:
                continue
            if str(a.get("last_run_date") or "").strip() == today_key:
                continue
            target_time = parse_time(a.get("send_time"))
            if now.time() < target_time:
                continue
            self._financeiro_alert_running.add(alert_id)
            self._financeiro_alert_attempted[alert_id] = today_key
            threading.Thread(target=send_for_alert, args=(alert_id, dict(a)), daemon=True).start()

    def _start_auto_docs_loop(self):
        if self._auto_docs_after_id:
            try:
                self.after_cancel(self._auto_docs_after_id)
            except Exception:
                pass
            self._auto_docs_after_id = None
        self._auto_docs_after_id = self.after(5000, self._auto_docs_tick)

    def _auto_docs_tick(self):
        from datetime import date

        self._auto_docs_after_id = self.after(60 * 1000, self._auto_docs_tick)
        if not self.current_user:
            return
        if self._auto_docs_running:
            return
        try:
            self.config_data = ConfigManager.load()
        except Exception:
            pass
        schedules = self.config_data.get("financeiro_envio_auto_documentos_agendamentos")
        active_id = str(self.config_data.get("financeiro_envio_auto_documentos_active_id") or "").strip()
        auto_cfg = None
        if isinstance(schedules, list) and schedules:
            enabled = []
            for s in schedules:
                if not isinstance(s, dict):
                    continue
                sid = str(s.get("id") or "").strip()
                if not sid:
                    continue
                if s.get("enabled"):
                    enabled.append(s)
            if enabled:
                auto_cfg = None
                for s in enabled:
                    if str(s.get("id") or "").strip() == active_id:
                        auto_cfg = s
                        break
                if auto_cfg is None:
                    auto_cfg = enabled[0]
                keep_id = str(auto_cfg.get("id") or "").strip()
                if len(enabled) > 1:
                    for s in schedules:
                        if not isinstance(s, dict):
                            continue
                        sid = str(s.get("id") or "").strip()
                        if sid and sid != keep_id:
                            s["enabled"] = False
                    self.config_data["financeiro_envio_auto_documentos_agendamentos"] = schedules
                    self.config_data["financeiro_envio_auto_documentos_active_id"] = keep_id
                    self.config_data["financeiro_envio_auto_documentos"] = auto_cfg
                    try:
                        ConfigManager.save(self.config_data)
                        self.config_data = ConfigManager.load()
                    except Exception:
                        pass
        if auto_cfg is None:
            auto_cfg = self.config_data.get("financeiro_envio_auto_documentos")
        if not isinstance(auto_cfg, dict) or not auto_cfg.get("enabled"):
            return
        try:
            interval_hours = int(auto_cfg.get("interval_hours") or 4)
        except Exception:
            interval_hours = 4
        interval_hours = max(1, min(72, interval_hours))
        first_run_time_raw = str(auto_cfg.get("first_run_time") or "").strip()
        schedule_anchor_raw = str(auto_cfg.get("schedule_anchor_at") or "").strip()
        if first_run_time_raw:
            anchor_dt = None
            if schedule_anchor_raw:
                try:
                    anchor_dt = datetime.fromisoformat(schedule_anchor_raw)
                except Exception:
                    anchor_dt = None
            if anchor_dt is None:
                raw = first_run_time_raw.lower().replace(" ", "").replace("h", ":")
                if raw.endswith(":"):
                    raw = raw + "00"
                if ":" not in raw:
                    raw = raw + ":00"
                parts = raw.split(":")
                if len(parts) >= 2:
                    raw = f"{parts[0]}:{parts[1]}"
                try:
                    t = datetime.strptime(raw, "%H:%M").time()
                    anchor_dt = datetime.combine(date.today(), t)
                    auto_cfg["schedule_anchor_at"] = anchor_dt.isoformat(timespec="seconds")
                    self.config_data["financeiro_envio_auto_documentos"] = auto_cfg
                    try:
                        ConfigManager.save(self.config_data)
                    except Exception:
                        pass
                except Exception:
                    anchor_dt = None
            if anchor_dt is not None:
                now = datetime.now()
                if now < anchor_dt:
                    return
                interval_td = timedelta(hours=interval_hours)
                elapsed = now - anchor_dt
                slots = int(elapsed.total_seconds() // interval_td.total_seconds()) if interval_td.total_seconds() > 0 else 0
                due_slot_dt = anchor_dt + (interval_td * slots)
                last_run_at_raw = str(auto_cfg.get("last_run_at") or "").strip()
                if last_run_at_raw:
                    try:
                        last_run_at_dt = datetime.fromisoformat(last_run_at_raw)
                        if last_run_at_dt >= due_slot_dt:
                            return
                    except Exception:
                        pass
            else:
                first_run_time_raw = ""
        if not first_run_time_raw:
            last_run_at_raw = str(auto_cfg.get("last_run_at") or "").strip()
            if last_run_at_raw:
                try:
                    last_run_at_dt = datetime.fromisoformat(last_run_at_raw)
                    if (datetime.now() - last_run_at_dt) < timedelta(hours=interval_hours):
                        return
                except Exception:
                    pass

        def _run():
            self._auto_docs_running = True
            win = None
            try:
                w = getattr(self, "envio_docs_window", None)
                if w is not None and w.winfo_exists():
                    win = w
            except Exception:
                win = None

            def _progress_cb(ev: Dict[str, Any]):
                if win is None:
                    return
                ev_snapshot = ev if isinstance(ev, dict) else {}
                try:
                    win.after(0, lambda ev2=ev_snapshot: win._apply_progress_event(ev2, auto=True))
                except Exception:
                    return

            try:
                from app_core.auto_documents import run_auto_documents
                res = run_auto_documents(self.config_data, dry_run=False, user_label=f"auto:{self.current_user}", progress_cb=_progress_cb)
                ConfigManager.save(self.config_data)
                self.config_data = ConfigManager.load()
                AuditLogger.write(
                    self.current_user,
                    "auto_docs_execucao",
                    f"emails_enviados={res.get('emails_sent')};falhas_email={res.get('failed_emails')};docs_enviados={res.get('docs_sent')};docs_falha={res.get('docs_failed')};sem_email={res.get('docs_no_email')}",
                )
            except Exception as e:
                try:
                    AuditLogger.write(self.current_user, "auto_docs_execucao_erro", f"erro={e}")
                except Exception:
                    pass
            finally:
                self._auto_docs_running = False
                if win is not None:
                    try:
                        win.after(0, lambda: win._set_controls_running(False))
                    except Exception:
                        pass

        threading.Thread(target=_run, daemon=True).start()
    def _apply_security_config(self, cfg: Dict[str, Any]):
        self.config_data = ConfigManager.load()
    def _apply_new_config(self, cfg: Dict[str, Any]):
        try:
            LicenseManager.validate_against_database(cfg, self.license_data)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e), parent=self)
            return
        self.config_data = cfg
        self._start_financeiro_alert_loop()
        self._start_auto_docs_loop()
def main():
    from app_core.logging_setup import init_logging
    init_logging()
    app = MainApp()
    app.mainloop()
if __name__ == "__main__":
    main()
