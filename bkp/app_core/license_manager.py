# -*- coding: utf-8 -*-
import hashlib
import hmac
import json
from datetime import datetime
from typing import Any, Dict, List
from .constants import LICENSE_FILENAME, LICENSE_PATH, LICENSE_SECRET
from .database import Database
from .helpers import AppError, is_valid_cnpj, normalize_document

class LicenseManager:
    @staticmethod
    def _payload(customer_document: str, expires_at: str) -> str:
        return f"{normalize_document(customer_document)}|{expires_at}"

    @staticmethod
    def signature(customer_document: str, expires_at: str) -> str:
        payload = LicenseManager._payload(customer_document, expires_at).encode("utf-8")
        return hmac.new(LICENSE_SECRET.encode("utf-8"), payload, hashlib.sha256).hexdigest()

    @staticmethod
    def load() -> Dict[str, Any]:
        if not LICENSE_PATH.exists():
            raise AppError(f"Arquivo de licença não encontrado: {LICENSE_FILENAME}")
        with open(LICENSE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)

    @staticmethod
    def validate_file() -> Dict[str, Any]:
        data = LicenseManager.load()
        customer_document = normalize_document(data.get("customer_document", ""))
        expires_at = str(data.get("expires_at", "")).strip()
        signature = str(data.get("signature", "")).strip()

        if not customer_document:
            raise AppError("Licença inválida: documento do cliente ausente.")
        if len(customer_document) != 14:
            raise AppError("Licença inválida: o CNPJ da chave deve conter 14 dígitos.")
        if not is_valid_cnpj(customer_document):
            raise AppError("Licença inválida: o CNPJ da chave é inválido.")
        if not expires_at:
            raise AppError("Licença inválida: data de expiração ausente.")
        if not signature:
            raise AppError("Licença inválida: assinatura ausente.")

        expected = LicenseManager.signature(customer_document, expires_at)
        if not hmac.compare_digest(signature, expected):
            raise AppError("Licença inválida: assinatura da chave não confere.")

        exp_dt = None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                parsed = datetime.strptime(expires_at, fmt)
                if fmt == "%Y-%m-%d":
                    parsed = parsed.replace(hour=23, minute=59, second=59)
                exp_dt = parsed
                break
            except ValueError:
                continue

        if exp_dt is None:
            raise AppError("Licença inválida: formato de data incorreto. Formatos aceitos: YYYY-MM-DD ou YYYY-MM-DD HH:MM:SS.")

        if datetime.now() > exp_dt:
            raise AppError(
                f"Licença expirada em {exp_dt.strftime('%d/%m/%Y %H:%M:%S')}. "
                "Entre em contato com o Suporte Databrev em www.databrev.com.br"
            )

        return {
            "customer_document": customer_document,
            "expires_at": expires_at,
            "customer_name": data.get("customer_name", ""),
        }

    @staticmethod
    def fetch_database_documents(config_data: Dict[str, Any]) -> List[str]:
        conn_cfg = config_data.get("connection", {})
        has_db = all(str(conn_cfg.get(k, "")).strip() for k in ("host", "port", "dbname", "user"))
        if not has_db:
            return []

        sql = "select cpf from empresa where grid in (select empresa from empresa_local)"
        with Database(config_data)._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql)
                rows = cur.fetchall()

        docs = sorted({normalize_document(r[0]) for r in rows if r and r[0] and normalize_document(r[0])})
        return docs

    @staticmethod
    def validate_against_database(config_data: Dict[str, Any], license_data: Dict[str, Any]) -> None:
        try:
            docs = LicenseManager.fetch_database_documents(config_data)
        except Exception as e:
            raise AppError(f"Falha ao validar a licença no banco de dados: {e}")

        if not docs:
            return

        if license_data["customer_document"] not in docs:
            valid_docs = ", ".join(docs)
            raise AppError(f"A licença não corresponde ao cliente deste banco de dados. CNPJ válido no sistema: {valid_docs}")
