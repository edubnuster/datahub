# -*- coding: utf-8 -*-
from typing import Any, Dict, List, Optional
from datetime import date, datetime
import base64
import os
import psycopg2
import psycopg2.extras
from .helpers import AppError


class Database:
    _purchase_meta: Dict[str, Any] = {}

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

        filters = ["coalesce(pcli.tipo, pcli.tipo_pessoa::text, '') ILIKE '%%C%%'"]
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

        outer_sql = f"SELECT base.* FROM ({base_sql}) base JOIN pessoa pcli ON pcli.grid = base.customer_id"
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
            JOIN pessoa pcli
              ON pcli.grid = base.customer_id
            WHERE coalesce(pcli.tipo, pcli.tipo_pessoa::text, '') ILIKE '%%C%%'
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
            JOIN pessoa pcli
              ON pcli.grid = base.customer_id
            WHERE conta IS NOT NULL
              AND trim(coalesce(conta, '')) <> ''
              AND coalesce(pcli.tipo, pcli.tipo_pessoa::text, '') ILIKE '%%C%%'
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

    def list_grupos_pessoa(self) -> List[Dict[str, Any]]:
        sql = "select grid, nome from grupo_pessoa order by nome"
        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(sql)
                return [dict(r) for r in cur.fetchall()]

    def list_portadores(self) -> List[Dict[str, Any]]:
        sql = "select grid, nome from portador order by nome"
        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(sql)
                return [dict(r) for r in cur.fetchall()]

    def list_customer_options_tipo_c(self) -> List[Dict[str, Any]]:
        sql = """
            SELECT
                p.grid as customer_id,
                p.codigo as codigo_cliente,
                p.nome as cliente
            FROM pessoa p
            WHERE coalesce(p.tipo, p.tipo_pessoa::text, '') ILIKE '%%C%%'
            ORDER BY p.nome
        """
        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(sql)
                return [dict(r) for r in cur.fetchall()]


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
                p.contrato as portador_contrato,
                cc.banco as banco_codigo,
                bk.nome as banco_nome,
                cc.agencia,
                cc.agencia_digito,
                cc.nr_conta,
                cc.digito as conta_digito,
                cc.modelo_boleto,
                e.cpf as cedente_documento,
                pessoa_nome_f(e.grid) as cedente_nome
            from boleto b
            left join boleto_info bi on bi.grid = b.boleto_info
            left join portador p on p.grid = b.portador
            left join conta_corrente cc on cc.grid = p.conta_corrente
            left join banco bk on bk.codigo = cc.banco
            left join empresa e on e.grid = cc.empresa
            where b.movto = %s
            order by b.grid desc
            limit 1
        """
        ob = None
        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(sql, (invoice_id,))
                row = cur.fetchone()
                if row:
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
                    cur.execute(open_banking_sql, (row.get("boleto_grid"),))
                    ob = cur.fetchone()

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

    def get_boletos_email_payload_bulk(self, invoice_ids: List[Any]) -> Dict[Any, Dict[str, Any]]:
        invoice_ids = [i for i in (invoice_ids or []) if i not in (None, "", 0, "0")]
        invoice_ids_sql: List[int] = []
        for i in invoice_ids:
            if isinstance(i, bool):
                continue
            if isinstance(i, int):
                invoice_ids_sql.append(i)
                continue
            s = str(i).strip()
            if s.isdigit():
                try:
                    invoice_ids_sql.append(int(s))
                except Exception:
                    pass
        if not invoice_ids_sql:
            return {}

        sql = """
            SELECT DISTINCT ON (b.movto)
                b.movto,
                b.grid as boleto_grid,
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
                p.contrato as portador_contrato,
                cc.banco as banco_codigo,
                bk.nome as banco_nome,
                cc.agencia,
                cc.agencia_digito,
                cc.nr_conta,
                cc.digito as conta_digito,
                cc.modelo_boleto,
                e.cpf as cedente_documento,
                pessoa_nome_f(e.grid) as cedente_nome
            FROM boleto b
            LEFT JOIN boleto_info bi ON bi.grid = b.boleto_info
            LEFT JOIN portador p ON p.grid = b.portador
            LEFT JOIN conta_corrente cc ON cc.grid = p.conta_corrente
            LEFT JOIN banco bk ON bk.codigo = cc.banco
            LEFT JOIN empresa e ON e.grid = cc.empresa
            WHERE b.movto = ANY(%s)
            ORDER BY b.movto, b.grid DESC
        """

        out: Dict[Any, Dict[str, Any]] = {}
        boleto_grids: List[int] = []
        by_movto: Dict[int, Dict[str, Any]] = {}
        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(sql, (invoice_ids_sql,))
                rows = cur.fetchall() or []
                for r in rows:
                    d = dict(r)
                    movto = d.get("movto")
                    if movto in (None, "", 0, "0"):
                        continue
                    try:
                        movto_i = int(movto)
                    except Exception:
                        continue
                    by_movto[movto_i] = d
                    bg = d.get("boleto_grid")
                    try:
                        boleto_grids.append(int(bg))
                    except Exception:
                        pass

                ob_map: Dict[int, Dict[str, Any]] = {}
                if boleto_grids:
                    cur.execute(
                        """
                        SELECT DISTINCT ON (ob.boleto)
                            ob.boleto,
                            pdf.arquivo_pdf,
                            pdf.gerado,
                            pdf.erro,
                            ob.grid as open_banking_boleto_grid
                        FROM open_banking_boleto ob
                        LEFT JOIN open_banking_boleto_pdf pdf
                               ON pdf.boleto = ob.grid
                        WHERE ob.boleto = ANY(%s)
                        ORDER BY ob.boleto, ob.grid DESC
                        """,
                        (boleto_grids,),
                    )
                    for r in cur.fetchall() or []:
                        try:
                            ob_map[int(r.get("boleto"))] = dict(r)
                        except Exception:
                            continue

        for inv_id in invoice_ids_sql:
            row = by_movto.get(int(inv_id))
            if not row:
                out[inv_id] = {"exists": False, "attachment_data": None, "filename": "", "mime_type": "application/pdf", "email_note": "Boleto: ainda não foi gerado."}
                out[str(inv_id)] = out[inv_id]
                continue

            data = dict(row)
            attachment_data = None
            filename = f"boleto_{data.get('documento') or data.get('boleto_grid') or inv_id}.pdf"
            email_note = ""

            ob = None
            try:
                ob = (ob_map or {}).get(int(data.get("boleto_grid") or 0))
            except Exception:
                ob = None
            if ob:
                arquivo_pdf = ob.get("arquivo_pdf")
                if isinstance(arquivo_pdf, memoryview):
                    attachment_data = arquivo_pdf.tobytes()
                elif isinstance(arquivo_pdf, (bytes, bytearray)):
                    attachment_data = bytes(arquivo_pdf)
                if attachment_data:
                    email_note = "Boleto: segue em anexo (PDF)."
                elif ob.get("gerado"):
                    email_note = "Boleto: localizado, porém não foi possível anexar automaticamente."
                elif ob.get("erro"):
                    email_note = f"Boleto: houve falha na geração automática: {ob.get('erro')}"

            if not email_note:
                email_note = "Boleto: localizado e será gerado em PDF para envio."

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
            out[inv_id] = data
            out[str(inv_id)] = data

        return out

    def get_boletos_email_payload_by_boleto_grids(self, boleto_grids: List[Any]) -> Dict[Any, Dict[str, Any]]:
        boleto_grids = [b for b in (boleto_grids or []) if b not in (None, "", 0, "0")]
        bg_sql: List[int] = []
        for b in boleto_grids:
            if isinstance(b, bool):
                continue
            if isinstance(b, int):
                bg_sql.append(b)
                continue
            s = str(b).strip()
            if s.isdigit():
                try:
                    bg_sql.append(int(s))
                except Exception:
                    pass
        if not bg_sql:
            return {}

        sql = """
            SELECT
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
                p.contrato as portador_contrato,
                cc.banco as banco_codigo,
                bk.nome as banco_nome,
                cc.agencia,
                cc.agencia_digito,
                cc.nr_conta,
                cc.digito as conta_digito,
                cc.modelo_boleto,
                e.cpf as cedente_documento,
                pessoa_nome_f(e.grid) as cedente_nome
            FROM boleto b
            LEFT JOIN boleto_info bi ON bi.grid = b.boleto_info
            LEFT JOIN portador p ON p.grid = b.portador
            LEFT JOIN conta_corrente cc ON cc.grid = p.conta_corrente
            LEFT JOIN banco bk ON bk.codigo = cc.banco
            LEFT JOIN empresa e ON e.grid = cc.empresa
            WHERE b.grid = ANY(%s)
        """

        out: Dict[Any, Dict[str, Any]] = {}
        ob_map: Dict[int, Dict[str, Any]] = {}
        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(
                    """
                    SELECT DISTINCT ON (ob.boleto)
                        ob.boleto,
                        pdf.arquivo_pdf,
                        pdf.gerado,
                        pdf.erro,
                        ob.grid as open_banking_boleto_grid
                    FROM open_banking_boleto ob
                    LEFT JOIN open_banking_boleto_pdf pdf
                           ON pdf.boleto = ob.grid
                    WHERE ob.boleto = ANY(%s)
                    ORDER BY ob.boleto, ob.grid DESC
                    """,
                    (bg_sql,),
                )
                for r in cur.fetchall() or []:
                    try:
                        ob_map[int(r.get("boleto"))] = dict(r)
                    except Exception:
                        continue

                cur.execute(sql, (bg_sql,))
                for r in cur.fetchall() or []:
                    data = dict(r)
                    bg = data.get("boleto_grid")
                    try:
                        bg_i = int(bg)
                    except Exception:
                        continue
                    attachment_data = None
                    email_note = ""
                    ob = ob_map.get(bg_i)
                    if ob:
                        arquivo_pdf = ob.get("arquivo_pdf")
                        if isinstance(arquivo_pdf, memoryview):
                            attachment_data = arquivo_pdf.tobytes()
                        elif isinstance(arquivo_pdf, (bytes, bytearray)):
                            attachment_data = bytes(arquivo_pdf)
                        if attachment_data:
                            email_note = "Boleto: segue em anexo (PDF)."
                        elif ob.get("gerado"):
                            email_note = "Boleto: localizado, porém não foi possível anexar automaticamente."
                        elif ob.get("erro"):
                            email_note = f"Boleto: houve falha na geração automática: {ob.get('erro')}"
                    if not email_note:
                        email_note = "Boleto: localizado e será gerado em PDF para envio."
                    vencto = data.get("vencto")
                    data["vencto_display"] = vencto.strftime("%d/%m/%Y") if hasattr(vencto, "strftime") else str(vencto or "")
                    data["valor_display"] = self._money_display(data.get("valor"))
                    cidade = str(data.get("sacado_cidade") or "").strip()
                    estado = str(data.get("sacado_estado") or "").strip()
                    data["sacado_cidade_uf"] = f"{cidade}/{estado}".strip("/ ")
                    data["attachment_data"] = attachment_data
                    data["filename"] = f"boleto_{data.get('documento') or bg_i}.pdf"
                    data["mime_type"] = "application/pdf"
                    data["exists"] = True
                    data["email_note"] = email_note
                    out[bg_i] = data
                    out[str(bg_i)] = data
        return out

    def list_agenda_invoices(
        self,
        due_date: date,
        group_id=None,
        portador_id=None,
        customer_id: Optional[Any] = None,
    ) -> List[Dict[str, Any]]:
        sql = """
            WITH titulos AS (
                SELECT
                    m.grid AS movto_id,
                    m.empresa,
                    m.pessoa AS customer_id,
                    p.codigo AS codigo_cliente,
                    p.nome AS cliente,
                    m.conta_debitar AS conta,
                    conta_nome_f(m.conta_debitar) AS conta_nome,
                    m.data,
                    m.valor,
                    COALESCE((
                        SELECT SUM(fl.valor_desconto)
                        FROM fatura_lancto fl
                        WHERE fl.movto = m.grid
                    ), 0) AS valor_desconto,
                    COALESCE((
                        SELECT SUM(d.valor)
                        FROM movto_map mp
                        JOIN movto d ON d.grid = mp.child
                        WHERE mp.parent = m.grid
                          AND d.motivo = 155
                    ), 0) AS valor_baixado
                FROM movto m
                JOIN pessoa p
                  ON p.grid = m.pessoa
                WHERE m.pessoa IS NOT NULL
                  AND COALESCE(m.valor, 0) > 0
                  AND m.child = 0
                  AND m.conta_debitar LIKE '1.3.04%%'
                  AND EXISTS (SELECT 1 FROM boleto b2 WHERE b2.movto = m.grid)
            ),
            base AS (
                SELECT
                    movto_id,
                    pessoa_nome_f(t.empresa) AS empresa,
                    t.customer_id,
                    t.codigo_cliente,
                    t.cliente,
                    t.conta,
                    t.conta_nome,
                    t.data,
                    t.valor,
                    t.valor_desconto,
                    t.valor_baixado,
                    (t.valor - t.valor_desconto - t.valor_baixado) AS saldo_em_aberto
                FROM titulos t
                WHERE (t.valor - t.valor_desconto - t.valor_baixado) > 0
            )
            SELECT
                b.movto_id,
                b.empresa,
                b.customer_id,
                b.codigo_cliente,
                b.cliente,
                b.conta,
                b.conta_nome,
                b.data,
                bi.vencto as vencto,
                b.valor,
                b.valor_desconto,
                b.valor_baixado,
                b.saldo_em_aberto,
                coalesce(c.email, '') as customer_email,
                c.grupo as customer_group_id,
                coalesce(gp.nome, '') as customer_group_name,
                bol.portador_id,
                coalesce(po.nome, '') as portador_nome
            FROM base b
            JOIN pessoa pcli
              ON pcli.grid = b.customer_id
            LEFT JOIN cliente c ON c.grid = b.customer_id
            LEFT JOIN grupo_pessoa gp ON gp.grid = c.grupo
            LEFT JOIN LATERAL (
                SELECT
                    b2.portador as portador_id,
                    b2.boleto_info as boleto_info_id
                FROM boleto b2
                WHERE b2.movto = b.movto_id
                ORDER BY b2.grid DESC
                LIMIT 1
            ) bol ON true
            LEFT JOIN boleto_info bi
              ON bi.grid = bol.boleto_info_id
            LEFT JOIN portador po ON po.grid = bol.portador_id
            WHERE bi.vencto = %(due_date)s
              AND (%(group_id)s is null OR c.grupo = %(group_id)s)
              AND (%(portador_id)s is null OR bol.portador_id = %(portador_id)s)
              AND (%(customer_id)s is null OR b.customer_id = %(customer_id)s)
              AND coalesce(pcli.tipo, pcli.tipo_pessoa::text, '') ILIKE '%%C%%'
            ORDER BY b.empresa, b.cliente, bi.vencto
        """
        params: Dict[str, Any] = {
            "due_date": due_date,
            "group_id": group_id if group_id not in ("", "0") else None,
            "portador_id": portador_id if portador_id not in ("", "0") else None,
            "customer_id": customer_id if customer_id not in (None, "", 0, "0") else None,
        }
        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(sql, params)
                return [dict(r) for r in cur.fetchall()]

    def get_purchase_info_bulk(self, invoice_ids: List[Any]) -> Dict[Any, Dict[str, Any]]:
        invoice_ids = [i for i in (invoice_ids or []) if i not in (None, "", 0, "0")]
        if not invoice_ids:
            return {}

        invoice_ids_sql: List[int] = []
        for i in invoice_ids:
            if isinstance(i, bool):
                continue
            if isinstance(i, int):
                invoice_ids_sql.append(i)
                continue
            try:
                s = str(i).strip()
            except Exception:
                continue
            if s.isdigit():
                try:
                    invoice_ids_sql.append(int(s))
                except Exception:
                    pass
        if not invoice_ids_sql:
            return {}

        meta = self.__class__._purchase_meta or {}
        if not meta:
            has_hora = False
            has_produto_nome_f = False
            with self._connect() as conn:
                with conn.cursor() as cur:
                    try:
                        cur.execute("select 1 from pg_attribute where attrelid = 'movto'::regclass and attname = %s and not attisdropped limit 1", ("hora",))
                        has_hora = bool(cur.fetchone())
                    except Exception:
                        has_hora = False
                    try:
                        cur.execute("select 1 from pg_proc where proname = %s limit 1", ("produto_nome_f",))
                        has_produto_nome_f = bool(cur.fetchone())
                    except Exception:
                        has_produto_nome_f = False
            meta = {"has_hora": has_hora, "has_produto_nome_f": has_produto_nome_f}
            self.__class__._purchase_meta = meta

        dt_expr = "m.hora" if meta.get("has_hora") else "m.data"
        sale_dt_expr = "s.hora" if meta.get("has_hora") else "s.data"
        prod_expr = "produto_nome_f(l.produto)" if meta.get("has_produto_nome_f") else "l.produto::text"

        sql = f"""
            with inv as (
                select
                    m.grid as invoice_id,
                    m.valor as invoice_amount,
                    m.documento as invoice_documento,
                    {dt_expr} as invoice_dt,
                    m.mlid as invoice_mlid
                from movto m
                where m.grid = any(%s)
            ),
            sales as (
                select
                    inv.invoice_id,
                    inv.invoice_amount,
                    inv.invoice_documento as documento,
                    inv.invoice_dt as dt,
                    inv.invoice_mlid as mlid
                from inv
                union all
                select
                    inv.invoice_id,
                    inv.invoice_amount,
                    s.documento as documento,
                    {sale_dt_expr} as dt,
                    s.mlid as mlid
                from inv
                join movto_map mm
                  on mm.child = inv.invoice_id
                join movto s
                  on s.grid = mm.parent
            ),
            items as (
                select
                    sales.invoice_id,
                    sales.invoice_amount,
                    sales.documento,
                    sales.dt,
                    nullif(trim(coalesce({prod_expr}::text, '')), '') as product_name,
                    sum(coalesce(l.quantidade, 0)) as quantity,
                    sum(coalesce(l.valor, 0)) as item_total
                from sales
                join lancto l
                  on l.mlid = sales.mlid
                where l.operacao = 'V'
                group by
                    sales.invoice_id,
                    sales.invoice_amount,
                    sales.documento,
                    sales.dt,
                    product_name
            )
            select
                invoice_id,
                invoice_amount,
                documento,
                dt,
                product_name,
                quantity,
                item_total
            from items
            order by invoice_id, dt, documento, product_name
        """

        out: Dict[Any, Dict[str, Any]] = {}
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, (invoice_ids_sql,))
                rows = cur.fetchall() or []

        tmp: Dict[int, Dict[str, Any]] = {}
        for invoice_id, invoice_amount, documento, dt, product_name, quantity, item_total in rows:
            try:
                inv_num = float(invoice_amount) if invoice_amount not in (None, "") else None
            except Exception:
                inv_num = None
            slot = tmp.get(int(invoice_id))
            if not slot:
                slot = {"invoice_amount": inv_num, "documents": {}, "dt_start": None, "dt_end": None}
                tmp[int(invoice_id)] = slot
            if dt not in (None, ""):
                if slot["dt_start"] in (None, "") or dt < slot["dt_start"]:
                    slot["dt_start"] = dt
                if slot["dt_end"] in (None, "") or dt > slot["dt_end"]:
                    slot["dt_end"] = dt
            doc_key = str(documento or "").strip() or "N/A"
            doc = slot["documents"].get(doc_key)
            if not doc:
                doc = {"documento": doc_key, "dt": dt, "items": {}, "total": 0.0}
                slot["documents"][doc_key] = doc
            if doc.get("dt") in (None, "") and dt not in (None, ""):
                doc["dt"] = dt
            item_key = str(product_name or "").strip() or "Item"
            it = doc["items"].get(item_key)
            if not it:
                it = {"product": item_key, "quantity": 0.0, "item_total": 0.0}
                doc["items"][item_key] = it
            try:
                it["quantity"] += float(quantity or 0)
            except Exception:
                pass
            try:
                it["item_total"] += float(item_total or 0)
            except Exception:
                pass
            try:
                doc["total"] += float(item_total or 0)
            except Exception:
                pass

        for invoice_id in invoice_ids_sql:
            slot = tmp.get(int(invoice_id))
            if not slot:
                out[invoice_id] = {"purchase_dt": None, "purchase_dt_start": None, "purchase_dt_end": None, "invoice_amount": None, "items_total": None, "documents": []}
                out[str(invoice_id)] = out[invoice_id]
                continue
            documents = []
            items_total = 0.0
            for doc in slot["documents"].values():
                items = [v for _, v in sorted((doc.get("items") or {}).items(), key=lambda kv: kv[0].lower())]
                total = float(doc.get("total") or 0)
                items_total += total
                documents.append(
                    {
                        "documento": doc.get("documento"),
                        "dt": doc.get("dt"),
                        "total": total,
                        "items": items,
                    }
                )
            documents.sort(key=lambda d: (str(d.get("dt") or ""), str(d.get("documento") or "")))
            out[invoice_id] = {
                "purchase_dt": slot.get("dt_end") or slot.get("dt_start"),
                "purchase_dt_start": slot.get("dt_start"),
                "purchase_dt_end": slot.get("dt_end"),
                "invoice_amount": slot.get("invoice_amount"),
                "items_total": items_total,
                "documents": documents,
            }
            out[str(invoice_id)] = out[invoice_id]

        return out

    def get_sale_signature_pdf(self, movto_id: Any) -> Dict[str, Any]:
        movto_id = movto_id if movto_id not in (None, "", 0, "0") else None
        if not movto_id:
            return {"exists": False, "attachments": [], "attachment_data": None, "filename": "", "mime_type": "application/octet-stream"}
        return (self.get_sale_signatures_pdf_bulk([movto_id]).get(movto_id)) or {
            "exists": False,
            "attachments": [],
            "attachment_data": None,
            "filename": "",
            "mime_type": "application/octet-stream",
        }

    def get_sale_signatures_pdf_bulk(self, movto_ids: List[Any]) -> Dict[Any, Dict[str, Any]]:
        movto_ids = [i for i in (movto_ids or []) if i not in (None, "", 0, "0")]
        if not movto_ids:
            return {}

        movto_ids_sql: List[int] = []
        for i in movto_ids:
            if isinstance(i, bool):
                continue
            if isinstance(i, int):
                movto_ids_sql.append(i)
                continue
            try:
                s = str(i).strip()
            except Exception:
                continue
            if not s:
                continue
            if s.isdigit():
                try:
                    movto_ids_sql.append(int(s))
                except Exception:
                    continue
        if not movto_ids_sql:
            return {}

        sql = """
            WITH inv AS (
                SELECT unnest(%(movto_ids)s::bigint[]) AS invoice_id
            ),
            parents AS (
                SELECT
                    mm.child AS invoice_id,
                    mm.parent AS sale_id
                FROM movto_map mm
                JOIN inv ON inv.invoice_id = mm.child
                WHERE mm.parent IS NOT NULL
                  AND mm.child IS NOT NULL
                  AND mm.parent <> 0
                  AND mm.child <> 0
            ),
            candidates AS (
                SELECT
                    inv.invoice_id,
                    a.grid AS anexo_grid,
                    a.descricao,
                    a.extensao,
                    a.anexo,
                    a.ts
                FROM public.anexo a
                JOIN inv ON inv.invoice_id = a.movto
                WHERE lower(coalesce(a.descricao, '')) LIKE '%%assinatura%%'

                UNION ALL

                SELECT
                    p.invoice_id,
                    a.grid AS anexo_grid,
                    a.descricao,
                    a.extensao,
                    a.anexo,
                    a.ts
                FROM parents p
                JOIN public.anexo a
                  ON a.movto = p.sale_id
                WHERE lower(coalesce(a.descricao, '')) LIKE '%%assinatura%%'
            )
            SELECT
                c.invoice_id AS movto_id,
                c.anexo_grid AS grid,
                c.descricao,
                c.extensao,
                c.anexo,
                c.ts
            FROM candidates c
            ORDER BY c.invoice_id, c.ts DESC NULLS LAST, c.anexo_grid DESC
        """

        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(sql, {"movto_ids": movto_ids_sql})
                rows = [dict(r) for r in (cur.fetchall() or [])]

        def _normalize_ext(ext: str) -> str:
            e = str(ext or "").strip().lower()
            if e.startswith("."):
                e = e[1:]
            return e

        def _sniff_type(data: bytes, fallback_ext: str) -> tuple[str, str]:
            if not data:
                return _normalize_ext(fallback_ext), "application/octet-stream"
            if data.startswith(b"%PDF-"):
                return "pdf", "application/pdf"
            if data.startswith(b"\x89PNG\r\n\x1a\n"):
                return "png", "image/png"
            if data[:3] == b"\xff\xd8\xff":
                return "jpg", "image/jpeg"
            if data[:4] in (b"II*\x00", b"MM\x00*"):
                return "tiff", "image/tiff"
            ext_n = _normalize_ext(fallback_ext)
            if ext_n == "pdf":
                return "pdf", "application/pdf"
            if ext_n == "png":
                return "png", "image/png"
            if ext_n in ("jpg", "jpeg"):
                return "jpg", "image/jpeg"
            if ext_n in ("tif", "tiff"):
                return "tiff", "image/tiff"
            return ext_n, "application/octet-stream"

        def _maybe_decode_wrapped(data: bytes) -> bytes:
            if not data:
                return data
            if data.startswith(b"%PDF-") or data.startswith(b"\x89PNG\r\n\x1a\n"):
                return data

            try:
                s = data.decode("ascii", errors="strict").strip()
            except Exception:
                return data

            if s.startswith("\\x") and len(s) > 2:
                hex_part = s[2:].strip()
                if hex_part and all(c in "0123456789abcdefABCDEF" for c in hex_part):
                    try:
                        decoded = bytes.fromhex(hex_part)
                        if decoded:
                            return decoded
                    except Exception:
                        return data

            if len(s) >= 16 and all(c.isalnum() or c in "+/=\r\n-_" for c in s):
                compact = "".join(s.split())
                if compact:
                    pad = (-len(compact)) % 4
                    compact_padded = compact + ("=" * pad)
                    decoded = b""
                    try:
                        decoded = base64.b64decode(compact_padded, validate=True)
                    except Exception:
                        try:
                            decoded = base64.b64decode(compact_padded, validate=False)
                        except Exception:
                            decoded = b""
                    if decoded and (
                        decoded.startswith(b"%PDF-")
                        or decoded.startswith(b"\x89PNG\r\n\x1a\n")
                        or decoded[:3] == b"\xff\xd8\xff"
                        or decoded[:4] in (b"II*\x00", b"MM\x00*")
                    ):
                        return decoded

            return data

        grouped: Dict[Any, List[Dict[str, Any]]] = {}
        seen: Dict[Any, set] = {}
        for row in rows:
            inv_id = row.get("movto_id")
            if inv_id in (None, "", 0, "0"):
                continue

            raw = row.get("anexo")
            if isinstance(raw, memoryview):
                data = raw.tobytes()
            elif isinstance(raw, (bytes, bytearray)):
                data = bytes(raw)
            else:
                data = None
            if data:
                data = _maybe_decode_wrapped(data)
            if not data:
                continue

            g = row.get("grid")
            if inv_id not in seen:
                seen[inv_id] = set()
            if g in seen[inv_id]:
                continue
            seen[inv_id].add(g)

            ext = str(row.get("extensao") or "").strip()
            detected_ext, detected_mime = _sniff_type(data, ext)
            filename_ext = detected_ext or _normalize_ext(ext) or "bin"
            att = {
                "data": data,
                "filename": f"assinatura_{inv_id}_{g}.{filename_ext}",
                "mime_type": detected_mime,
                "descricao": row.get("descricao"),
                "grid": g,
                "ts": row.get("ts"),
            }
            grouped.setdefault(inv_id, []).append(att)

        out: Dict[Any, Dict[str, Any]] = {}
        for inv_id in movto_ids_sql:
            atts = grouped.get(inv_id) or []
            first = atts[0] if atts else {}
            out[inv_id] = {
                "exists": bool(atts),
                "attachments": atts,
                "attachment_data": first.get("data"),
                "filename": first.get("filename") or "",
                "mime_type": first.get("mime_type") or "application/octet-stream",
                "descricao": first.get("descricao"),
            }
            out[str(inv_id)] = out[inv_id]

        return out

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
