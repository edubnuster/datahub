# -*- coding: utf-8 -*-
from typing import Any, Dict, List
import base64
import os
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


    def get_customer_email(self, customer_id) -> str:
        sql = "select email from cliente where grid = %s"
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, (customer_id,))
                row = cur.fetchone()
                return str(row[0]).strip() if row and row[0] else ""


    def get_boleto_email_payload(self, invoice_id) -> Dict[str, Any]:
        sql = """
            select
                b.grid as boleto_grid,
                b.movto,
                b.portador,
                b.nosso_numero,
                b.tipo_formulario,
                b.boleto_info,
                b.impresso,
                b.situacao,
                bi.documento,
                bi.vencto,
                bi.valor,
                bi.sacado_nome,
                bi.sacado_inscricao,
                bi.sacado_endereco,
                bi.sacado_cidade,
                bi.sacado_estado,
                bi.mensagem,
                bi.linha_digitavel,
                bi.codigo_barra,
                p.codigo as portador_codigo,
                p.nome as portador_nome,
                p.carteira as portador_carteira,
                p.convenio as portador_convenio,
                p.conta_corrente as portador_conta_corrente,
                p.contrato as portador_contrato
            from boleto b
            left join boleto_info bi on bi.grid = b.boleto_info
            left join portador p on p.grid = b.portador
            where b.movto = %s
            order by b.grid desc
            limit 1
        """
        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(sql, (invoice_id,))
                row = cur.fetchone()

        if not row:
            return {
                "exists": False,
                "attachment_data": None,
                "filename": "",
                "mime_type": "application/pdf",
                "email_note": "Observação: o boleto ainda não foi gerado.",
            }

        data = dict(row)
        attachment_data = None
        filename = f"boleto_{data.get('documento') or data.get('boleto_grid') or invoice_id}.pdf"

        open_banking_sql = """
            select
                pdf.arquivo_pdf,
                pdf.gerado,
                pdf.erro,
                ob.grid as open_banking_boleto_grid
            from open_banking_boleto ob
            left join open_banking_boleto_pdf pdf
                   on pdf.boleto = ob.grid
            where ob.boleto = %s
            order by ob.grid desc
            limit 1
        """
        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(open_banking_sql, (data.get("boleto_grid"),))
                ob = cur.fetchone()

        email_note = ""
        if ob:
            ob = dict(ob)
            arquivo_pdf = ob.get("arquivo_pdf")
            if isinstance(arquivo_pdf, memoryview):
                attachment_data = arquivo_pdf.tobytes()
            elif isinstance(arquivo_pdf, (bytes, bytearray)):
                attachment_data = bytes(arquivo_pdf)

            if attachment_data:
                email_note = "Observação: o boleto segue em anexo."
            elif ob.get("gerado"):
                email_note = "Observação: foi localizado um boleto, mas o PDF não pôde ser anexado automaticamente."
            elif ob.get("erro"):
                email_note = f"Observação: houve falha na geração automática do boleto: {ob.get('erro')}"
        if not email_note:
            email_note = "Observação: o boleto foi localizado e será gerado em PDF pelo app para envio em anexo."

        vencto = data.get("vencto")
        data["vencto_display"] = vencto.strftime("%d/%m/%Y") if hasattr(vencto, "strftime") else str(vencto or "")
        data["valor_display"] = self._money_display(data.get("valor"))
        cidade = str(data.get("sacado_cidade") or "").strip()
        estado = str(data.get("sacado_estado") or "").strip()
        data["sacado_cidade_uf"] = f"{cidade}/{estado}".strip("/ ")
        data["attachment_data"] = attachment_data
        data["filename"] = filename
        data["mime_type"] = "application/pdf"
        data["exists"] = True
        data["email_note"] = email_note
        return data

    @staticmethod
    def _money_display(value: Any) -> str:
        if value in (None, ""):
            return "0,00"
        try:
            num = float(value)
            return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return str(value)

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
