# -*- coding: utf-8 -*-
from typing import Any, Dict, List
import psycopg2
import psycopg2.extras
from .helpers import AppError


class Database:
    def __init__(self, config: Dict[str, Any]):
        self.config = config

    def _connect(self):
        conn_cfg = self.config["connection"]
        conn = psycopg2.connect(
            host=conn_cfg["host"],
            port=conn_cfg["port"],
            dbname=conn_cfg["dbname"],
            user=conn_cfg["user"],
            password=conn_cfg["password"],
            connect_timeout=8,
        )
        enc = (conn_cfg.get("client_encoding") or "").strip()
        if enc:
            conn.set_client_encoding(enc)
        return conn

    def test_connection(self):
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute("select 1")
                cur.fetchone()

    def list_inactive_customers(self) -> List[Dict[str, Any]]:
        sql_text = (self.config.get("queries", {}).get("list_inactive_customers_sql") or "").strip()
        if not sql_text:
            raise AppError("A query de listagem de clientes inativos não está configurada.")

        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(sql_text)
                return [dict(r) for r in cur.fetchall()]

    def _escaped_open_invoices_sql(self) -> str:
        sql_text = (self.config.get("queries", {}).get("list_open_invoices_sql") or "").strip()
        if not sql_text:
            raise AppError("A query de listagem de faturas a receber não está configurada.")

        return sql_text.replace("%", "%%")

    def list_open_invoices(
        self,
        due_date_from=None,
        due_date_to=None,
        customer_id=None,
        account_code=None,
    ) -> List[Dict[str, Any]]:
        base_sql = self._escaped_open_invoices_sql()

        filters = []
        params: Dict[str, Any] = {}

        if due_date_from:
            filters.append("vencto >= %(due_date_from)s")
            params["due_date_from"] = due_date_from

        if due_date_to:
            filters.append("vencto <= %(due_date_to)s")
            params["due_date_to"] = due_date_to

        if customer_id not in (None, "", 0, "0"):
            filters.append("customer_id = %(customer_id)s")
            params["customer_id"] = customer_id

        if account_code not in (None, "", "0"):
            filters.append("conta = %(account_code)s")
            params["account_code"] = account_code

        outer_sql = f"SELECT * FROM ({base_sql}) base"
        if filters:
            outer_sql += " WHERE " + " AND ".join(filters)
        outer_sql += " ORDER BY empresa, cliente, vencto"

        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                if params:
                    cur.execute(outer_sql, params)
                else:
                    cur.execute(outer_sql)
                return [dict(r) for r in cur.fetchall()]

    def list_open_invoice_customers(self) -> List[Dict[str, Any]]:
        base_sql = self._escaped_open_invoices_sql()
        outer_sql = f"""
            SELECT DISTINCT
                customer_id,
                codigo_cliente,
                cliente
            FROM ({base_sql}) base
            ORDER BY cliente
        """

        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(outer_sql)
                return [dict(r) for r in cur.fetchall()]

    def list_open_invoice_accounts(self) -> List[Dict[str, Any]]:
        base_sql = self._escaped_open_invoices_sql()
        outer_sql = f"""
            SELECT DISTINCT
                conta,
                conta_nome
            FROM ({base_sql}) base
            WHERE conta IS NOT NULL
              AND trim(coalesce(conta, '')) <> ''
            ORDER BY conta, conta_nome
        """

        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(outer_sql)
                return [dict(r) for r in cur.fetchall()]

    def execute_action(self, sql_text: str, customer_ids: List[Any]) -> int:
        if not (sql_text or "").strip():
            raise AppError("SQL da ação não configurada.")

        total = 0
        with self._connect() as conn:
            try:
                with conn.cursor() as cur:
                    for customer_id in customer_ids:
                        cur.execute(sql_text, {"customer_id": customer_id})
                        if cur.rowcount and cur.rowcount > 0:
                            total += cur.rowcount
                conn.commit()
            except Exception:
                conn.rollback()
                raise

        return total
