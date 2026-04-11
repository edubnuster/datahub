from __future__ import annotations

from datetime import datetime
import re
from typing import Any, Dict, List, Optional, Tuple
import xml.etree.ElementTree as ET


def _pdf_escape(text: str) -> str:
    if not text:
        return ""
    text = str(text)
    rep = {
        "á": r"\341",
        "à": r"\340",
        "â": r"\342",
        "ã": r"\343",
        "é": r"\351",
        "ê": r"\352",
        "í": r"\355",
        "ó": r"\363",
        "ô": r"\364",
        "õ": r"\365",
        "ú": r"\372",
        "ç": r"\347",
        "Á": r"\301",
        "À": r"\300",
        "Â": r"\302",
        "Ã": r"\303",
        "É": r"\311",
        "Ê": r"\312",
        "Í": r"\315",
        "Ó": r"\323",
        "Ô": r"\324",
        "Õ": r"\325",
        "Ú": r"\332",
        "Ç": r"\307",
        "º": r"\272",
        "ª": r"\252",
    }
    text = text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    for char, escape in rep.items():
        text = text.replace(char, escape)
    return text


def _build_pdf_bytes(ops: List[str]) -> bytes:
    page_width = 595
    page_height = 842
    stream = ("\n".join(ops)).encode("latin-1", errors="replace")

    objects = []
    objects.append(b"1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n")
    objects.append(b"2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n")
    objects.append(
        f"3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {page_width} {page_height}] /Resources << /Font << /F1 5 0 R /F2 6 0 R >> >> /Contents 4 0 R >>\nendobj\n".encode(
            "latin-1"
        )
    )
    objects.append(b"4 0 obj\n<< /Length " + str(len(stream)).encode("ascii") + b" >>\nstream\n" + stream + b"\nendstream\nendobj\n")
    objects.append(b"5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj\n")
    objects.append(b"6 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>\nendobj\n")

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


def _sanitize_xml_text(value: str) -> str:
    s = str(value or "").strip()
    if not s:
        return ""
    s = re.sub(r"`(https?://[^`]+)`", r"\1", s)
    s = re.sub(r'xmlns="([^"]*?)\s+(https?://[^"]+?)\s*"', r'xmlns="\2"', s)
    s = s.replace('xmlns=" http://', 'xmlns="http://')
    s = s.replace('xmlns="https://', 'xmlns="https://')
    s = s.replace('xmlns="http://', 'xmlns="http://')
    return s


def _xml_ns(tag: str) -> str:
    if tag.startswith("{") and "}" in tag:
        return tag[1 : tag.index("}")]
    return ""


def _t(el: Optional[ET.Element]) -> str:
    if el is None:
        return ""
    return str(el.text or "").strip()


def _dt_display(value: str) -> str:
    v = str(value or "").strip()
    if not v:
        return ""
    try:
        if v.endswith("Z"):
            v = v[:-1] + "+00:00"
        dt = datetime.fromisoformat(v)
        return dt.strftime("%d/%m/%Y %H:%M")
    except Exception:
        pass
    try:
        dt = datetime.strptime(v[:19], "%Y-%m-%dT%H:%M:%S")
        return dt.strftime("%d/%m/%Y %H:%M")
    except Exception:
        return v


def _money_br(value: str) -> str:
    s = str(value or "").strip()
    if not s:
        return ""
    try:
        num = float(s)
        return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return s


def _wrap(text: str, max_chars: int) -> List[str]:
    text = str(text or "").strip()
    if not text:
        return []
    lines: List[str] = []
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


def _extract_nfe_fields(xml_text: str) -> Dict[str, Any]:
    root = ET.fromstring(xml_text)
    ns = _xml_ns(root.tag)
    q = (lambda t: f"{{{ns}}}{t}") if ns else (lambda t: t)

    inf = root.find(q("infNFe")) if root.tag.endswith("NFe") else root.find(".//" + q("infNFe"))
    if inf is None:
        raise ValueError("XML inválido: infNFe não encontrado.")

    ide = inf.find(q("ide"))
    emit = inf.find(q("emit"))
    dest = inf.find(q("dest"))
    total = inf.find(q("total"))
    icms_tot = total.find(q("ICMSTot")) if total is not None else None
    cobr = inf.find(q("cobr"))
    fat = cobr.find(q("fat")) if cobr is not None else None
    inf_adic = inf.find(q("infAdic"))

    key = str(inf.get("Id") or "").strip()
    if key.startswith("NFe"):
        key = key[3:]

    items = []
    for det in (inf.findall(q("det")) or []):
        prod = det.find(q("prod"))
        items.append(
            {
                "nItem": str(det.get("nItem") or "").strip(),
                "cProd": _t(prod.find(q("cProd")) if prod is not None else None),
                "xProd": _t(prod.find(q("xProd")) if prod is not None else None),
                "uCom": _t(prod.find(q("uCom")) if prod is not None else None),
                "qCom": _t(prod.find(q("qCom")) if prod is not None else None),
                "vUnCom": _t(prod.find(q("vUnCom")) if prod is not None else None),
                "vProd": _t(prod.find(q("vProd")) if prod is not None else None),
                "CFOP": _t(prod.find(q("CFOP")) if prod is not None else None),
            }
        )

    dups = []
    if cobr is not None:
        for dup in (cobr.findall(q("dup")) or []):
            dups.append(
                {
                    "nDup": _t(dup.find(q("nDup"))),
                    "dVenc": _t(dup.find(q("dVenc"))),
                    "vDup": _t(dup.find(q("vDup"))),
                }
            )

    emit_end = emit.find(q("enderEmit")) if emit is not None else None
    dest_end = dest.find(q("enderDest")) if dest is not None else None
    transp = inf.find(q("transp"))
    transporta = transp.find(q("transporta")) if transp is not None else None
    vol = transp.find(q("vol")) if transp is not None else None
    ide_tpNF = _t(ide.find(q("tpNF")) if ide is not None else None)

    out = {
        "key": key,
        "nNF": _t(ide.find(q("nNF")) if ide is not None else None),
        "serie": _t(ide.find(q("serie")) if ide is not None else None),
        "dhEmi": _t(ide.find(q("dhEmi")) if ide is not None else None),
        "dhSaiEnt": _t(ide.find(q("dhSaiEnt")) if ide is not None else None),
        "natOp": _t(ide.find(q("natOp")) if ide is not None else None),
        "tpNF": ide_tpNF,  # 0=Entrada, 1=Saída
        "emit_xNome": _t(emit.find(q("xNome")) if emit is not None else None),
        "emit_xFant": _t(emit.find(q("xFant")) if emit is not None else None),
        "emit_CNPJ": _t(emit.find(q("CNPJ")) if emit is not None else None),
        "emit_CPF": _t(emit.find(q("CPF")) if emit is not None else None),
        "emit_IE": _t(emit.find(q("IE")) if emit is not None else None),
        "emit_IEST": _t(emit.find(q("IEST")) if emit is not None else None),
        "emit_xMun": _t(emit_end.find(q("xMun")) if emit_end is not None else None),
        "emit_UF": _t(emit_end.find(q("UF")) if emit_end is not None else None),
        "emit_xLgr": _t(emit_end.find(q("xLgr")) if emit_end is not None else None),
        "emit_nro": _t(emit_end.find(q("nro")) if emit_end is not None else None),
        "emit_xCpl": _t(emit_end.find(q("xCpl")) if emit_end is not None else None),
        "emit_xBairro": _t(emit_end.find(q("xBairro")) if emit_end is not None else None),
        "emit_CEP": _t(emit_end.find(q("CEP")) if emit_end is not None else None),
        "emit_fone": _t(emit_end.find(q("fone")) if emit_end is not None else None),
        "dest_xNome": _t(dest.find(q("xNome")) if dest is not None else None),
        "dest_CNPJ": _t(dest.find(q("CNPJ")) if dest is not None else None),
        "dest_CPF": _t(dest.find(q("CPF")) if dest is not None else None),
        "dest_IE": _t(dest.find(q("IE")) if dest is not None else None),
        "dest_xMun": _t(dest_end.find(q("xMun")) if dest_end is not None else None),
        "dest_UF": _t(dest_end.find(q("UF")) if dest_end is not None else None),
        "dest_xLgr": _t(dest_end.find(q("xLgr")) if dest_end is not None else None),
        "dest_nro": _t(dest_end.find(q("nro")) if dest_end is not None else None),
        "dest_xCpl": _t(dest_end.find(q("xCpl")) if dest_end is not None else None),
        "dest_xBairro": _t(dest_end.find(q("xBairro")) if dest_end is not None else None),
        "dest_CEP": _t(dest_end.find(q("CEP")) if dest_end is not None else None),
        "dest_fone": _t(dest_end.find(q("fone")) if dest_end is not None else None),
        "vBC": _t(icms_tot.find(q("vBC")) if icms_tot is not None else None),
        "vICMS": _t(icms_tot.find(q("vICMS")) if icms_tot is not None else None),
        "vBCST": _t(icms_tot.find(q("vBCST")) if icms_tot is not None else None),
        "vST": _t(icms_tot.find(q("vST")) if icms_tot is not None else None),
        "vProd": _t(icms_tot.find(q("vProd")) if icms_tot is not None else None),
        "vFrete": _t(icms_tot.find(q("vFrete")) if icms_tot is not None else None),
        "vSeg": _t(icms_tot.find(q("vSeg")) if icms_tot is not None else None),
        "vDesc": _t(icms_tot.find(q("vDesc")) if icms_tot is not None else None),
        "vIPI": _t(icms_tot.find(q("vIPI")) if icms_tot is not None else None),
        "vOutro": _t(icms_tot.find(q("vOutro")) if icms_tot is not None else None),
        "vNF": _t(icms_tot.find(q("vNF")) if icms_tot is not None else None),
        "vTotTrib": _t(icms_tot.find(q("vTotTrib")) if icms_tot is not None else None),
        "transp_modFrete": _t(transp.find(q("modFrete")) if transp is not None else None),
        "transp_xNome": _t(transporta.find(q("xNome")) if transporta is not None else None),
        "transp_CNPJ": _t(transporta.find(q("CNPJ")) if transporta is not None else None),
        "transp_CPF": _t(transporta.find(q("CPF")) if transporta is not None else None),
        "transp_IE": _t(transporta.find(q("IE")) if transporta is not None else None),
        "transp_xEnder": _t(transporta.find(q("xEnder")) if transporta is not None else None),
        "transp_xMun": _t(transporta.find(q("xMun")) if transporta is not None else None),
        "transp_UF": _t(transporta.find(q("UF")) if transporta is not None else None),
        "vol_qVol": _t(vol.find(q("qVol")) if vol is not None else None),
        "vol_esp": _t(vol.find(q("esp")) if vol is not None else None),
        "vol_marca": _t(vol.find(q("marca")) if vol is not None else None),
        "vol_nVol": _t(vol.find(q("nVol")) if vol is not None else None),
        "vol_pesoB": _t(vol.find(q("pesoB")) if vol is not None else None),
        "vol_pesoL": _t(vol.find(q("pesoL")) if vol is not None else None),
        "fat_nFat": _t(fat.find(q("nFat")) if fat is not None else None),
        "fat_vOrig": _t(fat.find(q("vOrig")) if fat is not None else None),
        "fat_vLiq": _t(fat.find(q("vLiq")) if fat is not None else None),
        "infCpl": _t(inf_adic.find(q("infCpl")) if inf_adic is not None else None),
        "infAdicFisco": _t(inf_adic.find(q("infAdicFisco")) if inf_adic is not None else None),
        "items": items,
        "dups": dups,
    }
    return out


def danfe_pdf_from_nfe_xml(xml_data: Any, *, fallback_suffix: str = "") -> Tuple[Optional[bytes], str]:
    if xml_data is None:
        return None, ""
    if isinstance(xml_data, (bytes, bytearray)):
        xml_text = xml_data.decode("utf-8", errors="replace")
    else:
        xml_text = str(xml_data)
    xml_text = _sanitize_xml_text(xml_text)
    if not xml_text.strip():
        return None, ""

    fields = _extract_nfe_fields(xml_text)
    key = str(fields.get("key") or "").strip()
    nNF = str(fields.get("nNF") or "").strip()
    serie = str(fields.get("serie") or "").strip()
    suffix = ""
    if nNF:
        suffix = nNF
        if serie:
            suffix = f"{suffix}_serie{serie}"
    elif key:
        suffix = key[-12:]
    else:
        suffix = str(fallback_suffix or "").strip() or "nfe"

    filename = f"danfe_{suffix}.pdf"

    # A4 dimensions in points: 595 x 842
    page_width = 595
    left = 20
    top = 822
    width = 555
    right = left + width

    ops: List[str] = []

    def draw_text(x, y, text, size=9, bold=False, max_len=None, align="left"):
        text = str(text or "").strip()
        if max_len and len(text) > max_len:
            text = text[: max_len - 3] + "..."
        if not text:
            return
        
        # very simple approximation of text width
        approx_width = len(text) * (size * 0.5)
        if align == "center":
            x = x - (approx_width / 2)
        elif align == "right":
            x = x - approx_width

        text = _pdf_escape(text)
        font = "/F2" if bold else "/F1"
        ops.append("BT")
        ops.append(f"{font} {size} Tf")
        ops.append(f"{x} {y} Td")
        ops.append(f"({text}) Tj")
        ops.append("ET")

    def draw_line(x1, y1, x2, y2, lw=0.6):
        ops.append(f"{lw} w")
        ops.append(f"{x1} {y1} m")
        ops.append(f"{x2} {y2} l")
        ops.append("S")

    def draw_rect(x, y, w, h, fill=False, lw=0.6):
        ops.append(f"{lw} w")
        ops.append(f"{x} {y} {w} {h} re")
        ops.append("f" if fill else "S")

    def format_doc(cnpj, cpf):
        d = cnpj or cpf or ""
        if len(d) == 14:
            return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
        if len(d) == 11:
            return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"
        return d
    
    def format_cep(cep):
        c = str(cep or "")
        if len(c) == 8:
            return f"{c[:5]}-{c[5:]}"
        return c

    def draw_itf_barcode(x, y, w, h, data):
        digits = "".join([c for c in str(data or "") if c.isdigit()])
        if not digits:
            return
        if len(digits) % 2 == 1:
            digits = "0" + digits

        dpat = {
            "0": "nnwwn",
            "1": "wnnnw",
            "2": "nwnnw",
            "3": "wwnnn",
            "4": "nnwnw",
            "5": "wnwnn",
            "6": "nwwnn",
            "7": "nnnww",
            "8": "wnnwn",
            "9": "nwnwn",
        }

        seq: List[Tuple[bool, int]] = []
        seq.extend([(True, 1), (False, 1), (True, 1), (False, 1)])
        for i in range(0, len(digits), 2):
            a = dpat.get(digits[i], "nnnnn")
            b = dpat.get(digits[i + 1], "nnnnn")
            for j in range(5):
                seq.append((True, 3 if a[j] == "w" else 1))
                seq.append((False, 3 if b[j] == "w" else 1))
        seq.extend([(True, 3), (False, 1), (True, 1)])

        quiet = 10
        total_units = (quiet * 2) + sum([u for _, u in seq])
        if total_units <= 0:
            return
        unit_w = float(w) / float(total_units)
        cx = x + (quiet * unit_w)

        ops.append("0 0 0 rg")
        for is_bar, units in seq:
            bw = units * unit_w
            if is_bar and bw > 0:
                ops.append(f"{cx} {y} {bw} {h} re")
                ops.append("f")
            cx += bw

    # --- Canhoto ---
    canhoto_h = 40
    y = top - canhoto_h
    draw_rect(left, y, width - 100, canhoto_h)
    draw_rect(right - 100, y, 100, canhoto_h)
    draw_text(left + 2, y + canhoto_h - 6, "RECEBEMOS DE", size=5)
    draw_text(left + 50, y + canhoto_h - 6, f"{fields.get('emit_xNome')} OS PRODUTOS E/OU SERVIÇOS CONSTANTES DA NOTA FISCAL INDICADA ABAIXO", size=5)
    
    draw_line(left, y + canhoto_h - 12, left + width - 100, y + canhoto_h - 12)
    draw_line(left + 100, y, left + 100, y + canhoto_h - 12)
    draw_text(left + 2, y + canhoto_h - 18, "DATA DE RECEBIMENTO", size=5)
    draw_text(left + 102, y + canhoto_h - 18, "IDENTIFICAÇÃO E ASSINATURA DO RECEBEDOR", size=5)
    
    draw_text(right - 50, y + canhoto_h - 12, "NF-e", size=12, bold=True, align="center")
    draw_text(right - 50, y + canhoto_h - 24, f"Nº {nNF}", size=8, bold=True, align="center")
    draw_text(right - 50, y + canhoto_h - 34, f"Série {serie}", size=8, bold=True, align="center")

    # dashed line
    y -= 5
    ops.append("[2 2] 0 d")
    draw_line(left, y, right, y)
    ops.append("[] 0 d") # reset dash

    # --- Emitente ---
    y -= 110
    emit_h = 105
    draw_rect(left, y, width, emit_h)
    
    # Emitente left box
    w_emit_left = 230
    draw_line(left + w_emit_left, y, left + w_emit_left, y + emit_h)
    draw_text(left + w_emit_left/2, y + emit_h - 8, "IDENTIFICAÇÃO DO EMITENTE", size=6, align="center")
    draw_text(left + w_emit_left/2, y + emit_h - 25, fields.get('emit_xNome'), size=9, bold=True, align="center", max_len=45)
    
    end_emit = f"{fields.get('emit_xLgr')}, {fields.get('emit_nro')}"
    if fields.get('emit_xCpl'):
        end_emit += f" - {fields.get('emit_xCpl')}"
    draw_text(left + w_emit_left/2, y + emit_h - 40, end_emit, size=7, align="center", max_len=55)
    draw_text(left + w_emit_left/2, y + emit_h - 50, f"{fields.get('emit_xBairro')} - CEP {format_cep(fields.get('emit_CEP'))}", size=7, align="center")
    draw_text(left + w_emit_left/2, y + emit_h - 60, f"{fields.get('emit_xMun')} - {fields.get('emit_UF')} Fone: {fields.get('emit_fone') or ''}", size=7, align="center")

    # Middle box (DANFE info)
    w_emit_mid = 85
    draw_line(left + w_emit_left + w_emit_mid, y, left + w_emit_left + w_emit_mid, y + emit_h)
    draw_text(left + w_emit_left + w_emit_mid/2, y + emit_h - 15, "DANFE", size=12, bold=True, align="center")
    draw_text(left + w_emit_left + w_emit_mid/2, y + emit_h - 25, "Documento Auxiliar da Nota", size=6, align="center")
    draw_text(left + w_emit_left + w_emit_mid/2, y + emit_h - 31, "Fiscal Eletrônica", size=6, align="center")
    
    tpNF = fields.get('tpNF') or '1'
    draw_text(left + w_emit_left + 5, y + emit_h - 45, "0 - ENTRADA", size=6)
    draw_text(left + w_emit_left + 5, y + emit_h - 53, "1 - SAÍDA", size=6)
    draw_rect(left + w_emit_left + 65, y + emit_h - 55, 12, 14)
    draw_text(left + w_emit_left + 71, y + emit_h - 51, tpNF, size=9, bold=True, align="center")

    draw_text(left + w_emit_left + w_emit_mid/2, y + emit_h - 70, f"Nº {nNF}", size=9, bold=True, align="center")
    draw_text(left + w_emit_left + w_emit_mid/2, y + emit_h - 80, f"SÉRIE {serie}", size=9, bold=True, align="center")
    draw_text(left + w_emit_left + w_emit_mid/2, y + emit_h - 90, "Folha 1/1", size=7, align="center")

    # Right box (Barcode & Key)
    bc_x = left + w_emit_left + w_emit_mid + 10
    bc_y = y + emit_h - 30
    bc_w = width - w_emit_left - w_emit_mid - 20
    draw_rect(bc_x, bc_y, bc_w, 25, fill=False)
    ops.append("1 1 1 rg")
    draw_rect(bc_x + 1, bc_y + 1, bc_w - 2, 23, fill=True, lw=0)
    ops.append("0 0 0 rg")
    draw_itf_barcode(bc_x + 3, bc_y + 3, bc_w - 6, 19, key)

    draw_line(left + w_emit_left + w_emit_mid, y + emit_h - 35, right, y + emit_h - 35)
    draw_text(left + w_emit_left + w_emit_mid + 2, y + emit_h - 41, "CHAVE DE ACESSO", size=5)
    formatted_key = f"{key[:4]} {key[4:8]} {key[8:12]} {key[12:16]} {key[16:20]} {key[20:24]} {key[24:28]} {key[28:32]} {key[32:36]} {key[36:40]} {key[40:44]}" if len(key)==44 else key
    draw_text(left + w_emit_left + w_emit_mid + 120, y + emit_h - 52, formatted_key, size=8, bold=True, align="center")
    
    draw_line(left + w_emit_left + w_emit_mid, y + emit_h - 58, right, y + emit_h - 58)
    draw_text(left + w_emit_left + w_emit_mid + 120, y + emit_h - 66, "Consulta de autenticidade no portal nacional da NF-e", size=7, align="center")
    draw_text(left + w_emit_left + w_emit_mid + 120, y + emit_h - 74, "www.nfe.fazenda.gov.br/portal ou no site da Sefaz Autorizadora", size=7, align="center")

    # Natureza / Protocolo
    y -= 25
    draw_rect(left, y, width, 25)
    draw_line(left + 330, y, left + 330, y + 25)
    draw_text(left + 2, y + 19, "NATUREZA DA OPERAÇÃO", size=5)
    draw_text(left + 2, y + 5, fields.get('natOp'), size=9, bold=True, max_len=60)
    draw_text(left + 332, y + 19, "PROTOCOLO DE AUTORIZAÇÃO DE USO", size=5)

    # Inscrição Estadual etc
    y -= 20
    draw_rect(left, y, width, 20)
    w_ie = 185
    draw_line(left + w_ie, y, left + w_ie, y + 20)
    draw_line(left + w_ie * 2, y, left + w_ie * 2, y + 20)
    draw_text(left + 2, y + 14, "INSCRIÇÃO ESTADUAL", size=5)
    draw_text(left + 2, y + 4, fields.get('emit_IE'), size=8, bold=True)
    draw_text(left + w_ie + 2, y + 14, "INSCRIÇÃO ESTADUAL DO SUBST. TRIBUT.", size=5)
    draw_text(left + w_ie + 2, y + 4, fields.get('emit_IEST'), size=8, bold=True)
    draw_text(left + w_ie * 2 + 2, y + 14, "CNPJ", size=5)
    draw_text(left + w_ie * 2 + 2, y + 4, format_doc(fields.get('emit_CNPJ'), fields.get('emit_CPF')), size=8, bold=True)

    # --- Destinatário / Remetente ---
    y -= 12
    draw_text(left, y, "DESTINATÁRIO / REMETENTE", size=7, bold=True)
    y -= 20
    draw_rect(left, y, width, 20)
    draw_line(right - 180, y, right - 180, y + 20)
    draw_line(right - 70, y, right - 70, y + 20)
    draw_text(left + 2, y + 14, "NOME / RAZÃO SOCIAL", size=5)
    draw_text(left + 2, y + 4, fields.get('dest_xNome'), size=8, bold=True, max_len=60)
    draw_text(right - 178, y + 14, "CNPJ / CPF", size=5)
    draw_text(right - 178, y + 4, format_doc(fields.get('dest_CNPJ'), fields.get('dest_CPF')), size=8, bold=True)
    draw_text(right - 68, y + 14, "DATA DA EMISSÃO", size=5)
    draw_text(right - 68, y + 4, _dt_display(fields.get('dhEmi'))[:10], size=8, bold=True)

    y -= 20
    draw_rect(left, y, width, 20)
    draw_line(right - 260, y, right - 260, y + 20)
    draw_line(right - 130, y, right - 130, y + 20)
    draw_line(right - 70, y, right - 70, y + 20)
    draw_text(left + 2, y + 14, "ENDEREÇO", size=5)
    end_dest = f"{fields.get('dest_xLgr')}, {fields.get('dest_nro')}"
    if fields.get('dest_xCpl'):
         end_dest += f" - {fields.get('dest_xCpl')}"
    draw_text(left + 2, y + 4, end_dest, size=8, bold=True, max_len=50)
    draw_text(right - 258, y + 14, "BAIRRO / DISTRITO", size=5)
    draw_text(right - 258, y + 4, fields.get('dest_xBairro'), size=8, bold=True, max_len=25)
    draw_text(right - 128, y + 14, "CEP", size=5)
    draw_text(right - 128, y + 4, format_cep(fields.get('dest_CEP')), size=8, bold=True)
    draw_text(right - 68, y + 14, "DATA DA SAÍDA/ENTRADA", size=5)
    draw_text(right - 68, y + 4, _dt_display(fields.get('dhSaiEnt'))[:10] if fields.get('dhSaiEnt') else "", size=8, bold=True)

    y -= 20
    draw_rect(left, y, width, 20)
    draw_line(right - 280, y, right - 280, y + 20)
    draw_line(right - 260, y, right - 260, y + 20)
    draw_line(right - 150, y, right - 150, y + 20)
    draw_line(right - 70, y, right - 70, y + 20)
    draw_text(left + 2, y + 14, "MUNICÍPIO", size=5)
    draw_text(left + 2, y + 4, fields.get('dest_xMun'), size=8, bold=True, max_len=50)
    draw_text(right - 278, y + 14, "UF", size=5)
    draw_text(right - 278, y + 4, fields.get('dest_UF'), size=8, bold=True)
    draw_text(right - 258, y + 14, "FONE / FAX", size=5)
    draw_text(right - 258, y + 4, fields.get('dest_fone'), size=8, bold=True)
    draw_text(right - 148, y + 14, "INSCRIÇÃO ESTADUAL", size=5)
    draw_text(right - 148, y + 4, fields.get('dest_IE'), size=8, bold=True)
    draw_text(right - 68, y + 14, "HORA DA SAÍDA", size=5)
    hora_saida = ""
    if fields.get('dhSaiEnt') and len(fields.get('dhSaiEnt')) > 15:
        hora_saida = _dt_display(fields.get('dhSaiEnt'))[11:16]
    draw_text(right - 68, y + 4, hora_saida, size=8, bold=True)

    # --- Fatura ---
    dups = fields.get("dups") or []
    if dups or fields.get("fat_nFat"):
        y -= 12
        draw_text(left, y, "FATURA / DUPLICATAS", size=7, bold=True)
        y -= 25
        draw_rect(left, y, width, 25)
        
        dx = left + 2
        for i, dup in enumerate(dups[:6]):
            draw_text(dx, y + 17, f"Num: {dup.get('nDup')}", size=5)
            draw_text(dx, y + 10, f"Venc: {_dt_display(dup.get('dVenc'))[:10]}", size=5)
            draw_text(dx, y + 3, f"Valor: R$ {_money_br(dup.get('vDup'))}", size=5)
            dx += 90
            if i < len(dups[:6]) - 1:
                draw_line(dx - 5, y, dx - 5, y + 25)

    # --- Cálculo do Imposto ---
    y -= 12
    draw_text(left, y, "CÁLCULO DO IMPOSTO", size=7, bold=True)
    
    y -= 20
    draw_rect(left, y, width, 20)
    w_imp = width / 7
    for i in range(1, 7):
        draw_line(left + w_imp * i, y, left + w_imp * i, y + 20)
    
    draw_text(left + 2, y + 14, "BASE DE CÁLCULO DO ICMS", size=4)
    draw_text(left + w_imp - 2, y + 4, _money_br(fields.get('vBC')), size=8, bold=True, align="right")
    draw_text(left + w_imp + 2, y + 14, "VALOR DO ICMS", size=4)
    draw_text(left + w_imp*2 - 2, y + 4, _money_br(fields.get('vICMS')), size=8, bold=True, align="right")
    draw_text(left + w_imp*2 + 2, y + 14, "BASE DE CÁLCULO DO ICMS ST", size=4)
    draw_text(left + w_imp*3 - 2, y + 4, _money_br(fields.get('vBCST')), size=8, bold=True, align="right")
    draw_text(left + w_imp*3 + 2, y + 14, "VALOR DO ICMS SUBSTITUIÇÃO", size=4)
    draw_text(left + w_imp*4 - 2, y + 4, _money_br(fields.get('vST')), size=8, bold=True, align="right")
    draw_text(left + w_imp*4 + 2, y + 14, "VALOR TOTAL TRIBUTOS", size=4)
    draw_text(left + w_imp*5 - 2, y + 4, _money_br(fields.get('vTotTrib')), size=8, bold=True, align="right")
    draw_text(left + w_imp*5 + 2, y + 14, "VALOR DO PIS", size=4)
    # Not extracted by default, keep 0,00 if missing or extract if needed.
    draw_text(left + w_imp*6 - 2, y + 4, "0,00", size=8, bold=True, align="right")
    draw_text(left + w_imp*6 + 2, y + 14, "VALOR TOTAL DOS PRODUTOS", size=4)
    draw_text(right - 2, y + 4, _money_br(fields.get('vProd')), size=8, bold=True, align="right")

    y -= 20
    draw_rect(left, y, width, 20)
    for i in range(1, 7):
        draw_line(left + w_imp * i, y, left + w_imp * i, y + 20)

    draw_text(left + 2, y + 14, "VALOR DO FRETE", size=4)
    draw_text(left + w_imp - 2, y + 4, _money_br(fields.get('vFrete')), size=8, bold=True, align="right")
    draw_text(left + w_imp + 2, y + 14, "VALOR DO SEGURO", size=4)
    draw_text(left + w_imp*2 - 2, y + 4, _money_br(fields.get('vSeg')), size=8, bold=True, align="right")
    draw_text(left + w_imp*2 + 2, y + 14, "DESCONTO", size=4)
    draw_text(left + w_imp*3 - 2, y + 4, _money_br(fields.get('vDesc')), size=8, bold=True, align="right")
    draw_text(left + w_imp*3 + 2, y + 14, "OUTRAS DESPESAS", size=4)
    draw_text(left + w_imp*4 - 2, y + 4, _money_br(fields.get('vOutro')), size=8, bold=True, align="right")
    draw_text(left + w_imp*4 + 2, y + 14, "VALOR DO IPI", size=4)
    draw_text(left + w_imp*5 - 2, y + 4, _money_br(fields.get('vIPI')), size=8, bold=True, align="right")
    draw_text(left + w_imp*5 + 2, y + 14, "VALOR DA COFINS", size=4)
    draw_text(left + w_imp*6 - 2, y + 4, "0,00", size=8, bold=True, align="right")
    draw_text(left + w_imp*6 + 2, y + 14, "VALOR TOTAL DA NOTA", size=4)
    draw_text(right - 2, y + 4, _money_br(fields.get('vNF')), size=8, bold=True, align="right")

    # --- Transportador ---
    y -= 12
    draw_text(left, y, "TRANSPORTADOR / VOLUMES TRANSPORTADOS", size=7, bold=True)
    
    y -= 20
    draw_rect(left, y, width, 20)
    draw_line(right - 350, y, right - 350, y + 20)
    draw_line(right - 260, y, right - 260, y + 20)
    draw_line(right - 150, y, right - 150, y + 20)
    draw_line(right - 120, y, right - 120, y + 20)
    draw_text(left + 2, y + 14, "NOME / RAZÃO SOCIAL", size=5)
    draw_text(left + 2, y + 4, fields.get('transp_xNome'), size=8, bold=True, max_len=45)
    draw_text(right - 348, y + 14, "FRETE POR CONTA", size=5)
    
    mod_frete = fields.get('transp_modFrete')
    frete_str = ""
    if mod_frete == "0": frete_str = "0-Remetente"
    elif mod_frete == "1": frete_str = "1-Destinatário"
    elif mod_frete == "2": frete_str = "2-Terceiros"
    elif mod_frete == "9": frete_str = "9-Sem Frete"
    draw_text(right - 348, y + 4, frete_str, size=8, bold=True)
    
    draw_text(right - 258, y + 14, "CÓDIGO ANTT", size=5)
    draw_text(right - 148, y + 14, "PLACA DO VEÍCULO", size=5)
    draw_text(right - 118, y + 14, "UF", size=5)
    draw_text(right - 118, y + 4, fields.get('transp_UF'), size=8, bold=True)
    draw_text(right - 90, y + 14, "CNPJ / CPF", size=5)
    draw_text(right - 90, y + 4, format_doc(fields.get('transp_CNPJ'), fields.get('transp_CPF')), size=8, bold=True)

    y -= 20
    draw_rect(left, y, width, 20)
    draw_line(right - 260, y, right - 260, y + 20)
    draw_line(right - 120, y, right - 120, y + 20)
    draw_text(left + 2, y + 14, "ENDEREÇO", size=5)
    draw_text(left + 2, y + 4, fields.get('transp_xEnder'), size=8, bold=True, max_len=60)
    draw_text(right - 258, y + 14, "MUNICÍPIO", size=5)
    draw_text(right - 258, y + 4, fields.get('transp_xMun'), size=8, bold=True)
    draw_text(right - 118, y + 14, "UF", size=5)
    draw_text(right - 90, y + 14, "INSCRIÇÃO ESTADUAL", size=5)
    draw_text(right - 90, y + 4, fields.get('transp_IE'), size=8, bold=True)

    y -= 20
    draw_rect(left, y, width, 20)
    draw_line(left + 60, y, left + 60, y + 20)
    draw_line(left + 150, y, left + 150, y + 20)
    draw_line(left + 250, y, left + 250, y + 20)
    draw_line(left + 350, y, left + 350, y + 20)
    draw_line(left + 450, y, left + 450, y + 20)
    draw_text(left + 2, y + 14, "QUANTIDADE", size=5)
    draw_text(left + 2, y + 4, fields.get('vol_qVol'), size=8, bold=True)
    draw_text(left + 62, y + 14, "ESPÉCIE", size=5)
    draw_text(left + 62, y + 4, fields.get('vol_esp'), size=8, bold=True)
    draw_text(left + 152, y + 14, "MARCA", size=5)
    draw_text(left + 152, y + 4, fields.get('vol_marca'), size=8, bold=True)
    draw_text(left + 252, y + 14, "NUMERAÇÃO", size=5)
    draw_text(left + 252, y + 4, fields.get('vol_nVol'), size=8, bold=True)
    draw_text(left + 352, y + 14, "PESO BRUTO", size=5)
    draw_text(left + 352, y + 4, fields.get('vol_pesoB'), size=8, bold=True)
    draw_text(left + 452, y + 14, "PESO LÍQUIDO", size=5)
    draw_text(left + 452, y + 4, fields.get('vol_pesoL'), size=8, bold=True)

    # --- Itens ---
    y -= 12
    draw_text(left, y, "DADOS DOS PRODUTOS / SERVIÇOS", size=7, bold=True)
    
    # Header items
    y -= 15
    items_top = y + 15
    
    col_w = [55, 175, 40, 25, 25, 20, 35, 35, 40, 35, 35, 35] # Sum = 555
    cols_x = [left]
    for w in col_w:
        cols_x.append(cols_x[-1] + w)

    def draw_item_header():
        draw_rect(left, y, width, 15)
        for cx in cols_x[1:-1]:
            draw_line(cx, y, cx, y + 15)
        headers = ["CÓDIGO", "DESCRIÇÃO DO PRODUTO / SERVIÇO", "NCM/SH", "CST", "CFOP", "UN", "QUANT", "V. UNIT", "V. TOTAL", "BC ICMS", "V. ICMS", "V. IPI", "ALÍQ. ICMS", "ALÍQ. IPI"]
        # Simplified headers to fit columns
        h = ["CÓDIGO", "DESCRIÇÃO", "NCM/SH", "CST", "CFOP", "UN", "QUANT", "V. UNIT", "V. TOTAL", "BC ICMS", "V. ICMS", "V. IPI"]
        for i, text in enumerate(h):
            draw_text(cols_x[i] + (col_w[i]/2), y + 5, text, size=5, align="center")

    draw_item_header()
    
    items_area_bottom = 120
    items_y = y

    max_items = 35
    items = fields.get("items") or []
    
    for it in items[:max_items]:
        items_y -= 10
        if items_y < items_area_bottom:
            break # No page break implemented yet, just truncate
            
        draw_text(cols_x[0] + 2, items_y + 2, it.get("cProd"), size=6, max_len=14)
        draw_text(cols_x[1] + 2, items_y + 2, it.get("xProd"), size=6, max_len=45)
        draw_text(cols_x[2] + 2, items_y + 2, it.get("NCM") or "", size=6, max_len=8)
        draw_text(cols_x[3] + 2, items_y + 2, "", size=6) # CST not extracted
        draw_text(cols_x[4] + 2, items_y + 2, it.get("CFOP"), size=6)
        draw_text(cols_x[5] + 2, items_y + 2, it.get("uCom"), size=6)
        draw_text(cols_x[7] - 2, items_y + 2, _money_br(it.get("qCom")), size=6, align="right")
        draw_text(cols_x[8] - 2, items_y + 2, _money_br(it.get("vUnCom")), size=6, align="right")
        draw_text(cols_x[9] - 2, items_y + 2, _money_br(it.get("vProd")), size=6, align="right")

    # Draw vertical lines for items area
    draw_rect(left, items_area_bottom, width, items_top - items_area_bottom)
    for cx in cols_x[1:-1]:
        draw_line(cx, items_area_bottom, cx, items_top)

    # --- Dados Adicionais ---
    y = items_area_bottom - 12
    draw_text(left, y, "DADOS ADICIONAIS", size=7, bold=True)
    y -= 80
    draw_rect(left, y, width, 80)
    w_adic = width * 0.65
    draw_line(left + w_adic, y, left + w_adic, y + 80)
    draw_text(left + 2, y + 72, "INFORMAÇÕES COMPLEMENTARES", size=5)
    
    infcpl = fields.get("infCpl") or ""
    if infcpl:
        dy = y + 62
        for line in _wrap(infcpl, 90)[:8]:
            draw_text(left + 2, dy, line, size=6)
            dy -= 8

    draw_text(left + w_adic + 2, y + 72, "RESERVADO AO FISCO", size=5)
    fisco = fields.get("infAdicFisco") or ""
    if fisco:
        dy = y + 62
        for line in _wrap(fisco, 45)[:8]:
            draw_text(left + w_adic + 2, dy, line, size=6)
            dy -= 8

    return _build_pdf_bytes(ops), filename
