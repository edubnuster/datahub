from __future__ import annotations

import io
import os
import zipfile
from email.message import EmailMessage
from typing import Iterable, Optional, Sequence, Tuple

from .embedded_danfe_logo_kaninha import get_kaninha_danfe_logo_png_bytes


EMAIL_LOGO_CID = "clientlogo"


def attach_email_logo(msg: EmailMessage, *, cid: str = EMAIL_LOGO_CID) -> bool:
    try:
        logo = get_kaninha_danfe_logo_png_bytes()
    except Exception:
        logo = None
    if not logo:
        return False
    try:
        html_part = msg.get_body(preferencelist=("html",))
    except Exception:
        html_part = None
    if html_part is None:
        return False
    try:
        html_part.add_related(logo, maintype="image", subtype="png", cid=f"<{cid}>", filename="logo.png")
        return True
    except Exception:
        return False


def _sanitize_zip_entry_name(name: str, *, fallback: str) -> str:
    base = os.path.basename(str(name or "").strip()).strip()
    base = base.replace("\\", "_").replace("/", "_")
    if not base:
        base = fallback
    return base


def zip_named_files(
    files: Sequence[Tuple[bytes, str]],
    *,
    zip_filename: str,
    compression: int = zipfile.ZIP_DEFLATED,
) -> Tuple[bytes, str]:
    buf = io.BytesIO()
    used = set()
    with zipfile.ZipFile(buf, mode="w", compression=compression) as zf:
        for i, (data, name) in enumerate(files or [], start=1):
            if not data:
                continue
            safe = _sanitize_zip_entry_name(name, fallback=f"assinatura_{i}.bin")
            candidate = safe
            if candidate in used:
                root, ext = os.path.splitext(candidate)
                k = 2
                while True:
                    candidate = f"{root}_{k}{ext}"
                    if candidate not in used:
                        break
                    k += 1
            used.add(candidate)
            zf.writestr(candidate, data)
    return buf.getvalue(), zip_filename
