# -*- coding: utf-8 -*-
from typing import Any, Dict, List, Optional
from datetime import date, datetime
import base64
import os
import re
import psycopg2
import psycopg2.extras
from .helpers import AppError


class Database:
    _purchase_meta: Dict[str, Any] = {}
    _table_columns_cache: Dict[str, set] = {}

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

    def _get_table_columns(self, conn, table_name: str) -> set:
        t = str(table_name or "").strip().lower()
        if not t:
            return set()
        cached = self._table_columns_cache.get(t)
        if cached is not None:
            return cached
        cols: set = set()
        try:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT column_name
                      FROM information_schema.columns
                     WHERE table_schema = 'public'
                       AND table_name = %s
                    """,
                    (t,),
                )
                cols = {str(r[0]).strip().lower() for r in (cur.fetchall() or []) if r and r[0]}
        except Exception:
            cols = set()
        self._table_columns_cache[t] = cols
        return cols

    def _format_cep(self, cep: Any) -> str:
        s = "".join([c for c in str(cep or "") if c.isdigit()])
        if len(s) == 8:
            return f"{s[:5]}-{s[5:]}"
        return str(cep or "").strip()

    def _empresa_endereco_str(self, conn, empresa_id: Any) -> str:
        try:
            eid = int(empresa_id)
        except Exception:
            return ""
        cols = self._get_table_columns(conn, "empresa")
        if not cols:
            return ""

        def pick(*names: str) -> Optional[str]:
            for n in names:
                if str(n).lower() in cols:
                    return str(n).lower()
            return None

        c_log = pick("logradouro", "endereco", "rua")
        c_num = pick("numero", "nro", "num")
        c_bai = pick("bairro")
        c_cid = pick("cidade", "municipio")
        c_uf = pick("uf", "estado")
        c_cep = pick("cep")
        c_cpl = pick("complemento", "cpl")

        select_parts: List[str] = []
        if c_log:
            select_parts.append(f"{c_log} as logradouro")
        if c_num:
            select_parts.append(f"{c_num} as numero")
        if c_cpl:
            select_parts.append(f"{c_cpl} as complemento")
        if c_bai:
            select_parts.append(f"{c_bai} as bairro")
        if c_cid:
            select_parts.append(f"{c_cid} as cidade")
        if c_uf:
            select_parts.append(f"{c_uf} as uf")
        if c_cep:
            select_parts.append(f"{c_cep} as cep")
        if not select_parts:
            return ""

        sql = "SELECT " + ", ".join(select_parts) + " FROM empresa WHERE grid = %s"
        try:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(sql, (eid,))
                r = cur.fetchone() or {}
        except Exception:
            return ""

        logradouro = str(r.get("logradouro") or "").strip()
        numero = str(r.get("numero") or "").strip()
        complemento = str(r.get("complemento") or "").strip()
        bairro = str(r.get("bairro") or "").strip()
        cidade = str(r.get("cidade") or "").strip()
        uf = str(r.get("uf") or "").strip()
        cep = self._format_cep(r.get("cep"))

        first = logradouro
        if numero:
            first = (first + f", {numero}").strip(", ")
        if complemento:
            first = (first + f" - {complemento}").strip(" -")

        mid = bairro
        tail = ""
        if cidade and uf:
            tail = f"{cidade}/{uf}"
        elif cidade:
            tail = cidade
        elif uf:
            tail = uf
        if cep:
            tail = (tail + f" - CEP {cep}").strip(" -")

        parts = [p for p in [first, mid, tail] if p]
        return " - ".join(parts).strip()

    def _boleto_info_sacado_cep(self, conn, boleto_info_id: Any) -> str:
        try:
            bi_id = int(boleto_info_id)
        except Exception:
            return ""
        cols = self._get_table_columns(conn, "boleto_info")
        if not cols:
            return ""
        cep_col = None
        for n in ("sacado_cep", "cep", "cep_sacado"):
            if n in cols:
                cep_col = n
                break
        if not cep_col:
            return ""
        try:
            with conn.cursor() as cur:
                cur.execute(f"SELECT {cep_col} FROM boleto_info WHERE grid = %s", (bi_id,))
                r = cur.fetchone()
                v = r[0] if r else ""
        except Exception:
            v = ""
        return self._format_cep(v)

    def list_inactive_customers(self, inactive_months: int = 3) -> List[Dict[str, Any]]:
        sql_text = (self.config.get("queries", {}).get("list_inactive_customers_sql") or "").strip()
        if not sql_text:
            raise AppError("A query de listagem de clientes inativos não está configurada.")

        try:
            inactive_months = int(inactive_months or 0)
        except Exception:
            inactive_months = 0
        inactive_months = max(1, min(2400, inactive_months))

        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                if "%(inactive_months)" in sql_text:
                    cur.execute(sql_text, {"inactive_months": inactive_months})
                else:
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

    @staticmethod
    def _normalize_bigint_ids(ids: List[Any]) -> List[int]:
        out: List[int] = []
        for i in (ids or []):
            if i in (None, "", 0, "0"):
                continue
            if isinstance(i, bool):
                continue
            if isinstance(i, int):
                out.append(i)
                continue
            try:
                s = str(i).strip()
            except Exception:
                continue
            if not s:
                continue
            if s.isdigit():
                try:
                    out.append(int(s))
                except Exception:
                    continue
        return out

    @classmethod
    def _get_table_columns(cls, conn, table_name: str, *, schema: str = "public") -> set:
        key = f"{schema.lower().strip()}.{str(table_name or '').lower().strip()}"
        cached = cls._table_columns_cache.get(key)
        if cached is not None:
            return cached
        cols: set = set()
        try:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    select column_name
                    from information_schema.columns
                    where table_schema = %s
                      and table_name = %s
                    """,
                    (schema, str(table_name or "").strip()),
                )
                for (cname,) in (cur.fetchall() or []):
                    if cname:
                        cols.add(str(cname).strip().lower())
        except Exception:
            cols = set()
        cls._table_columns_cache[key] = cols
        return cols

    @staticmethod
    def _pick_existing_column(cols: set, candidates: List[str]) -> str:
        for c in (candidates or []):
            cc = str(c or "").strip().lower()
            if cc and cc in cols:
                return cc
        return ""

    @staticmethod
    def _maybe_decode_wrapped_bytes(data: bytes) -> bytes:
        if not data:
            return data
        if data.startswith(b"%PDF-") or data.startswith(b"\x89PNG\r\n\x1a\n") or data[:3] == b"\xff\xd8\xff":
            return data
        if data.lstrip().startswith(b"<"):
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
                    or decoded.lstrip().startswith(b"<")
                ):
                    return decoded

        return data

    @classmethod
    def _blob_to_bytes(cls, raw: Any) -> Optional[bytes]:
        if raw is None:
            return None
        data: Optional[bytes] = None
        if isinstance(raw, memoryview):
            data = raw.tobytes()
        elif isinstance(raw, (bytes, bytearray)):
            data = bytes(raw)
        elif isinstance(raw, str):
            s = raw.strip()
            if not s:
                data = None
            elif s.lstrip().startswith("<"):
                data = s.encode("utf-8", errors="replace")
            else:
                compact = "".join(s.split())
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
                if decoded and (decoded.startswith(b"%PDF-") or decoded.lstrip().startswith(b"<")):
                    data = decoded
                else:
                    data = s.encode("utf-8", errors="replace")
        if not data:
            return None
        return cls._maybe_decode_wrapped_bytes(data)

    def check_boleto_exists_bulk(self, movto_ids: List[Any]) -> Dict[Any, bool]:
        movto_ids_sql = self._normalize_bigint_ids(movto_ids)
        if not movto_ids_sql:
            return {}

        sql = """
            select distinct b.movto as movto_id
            from boleto b
            where b.movto = any(%(movto_ids)s::bigint[])
              and (b.situacao is null or b.situacao < 5)
        """
        found: set = set()
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, {"movto_ids": movto_ids_sql})
                for (movto_id,) in (cur.fetchall() or []):
                    if movto_id not in (None, "", 0, "0"):
                        found.add(int(movto_id))

        out: Dict[Any, bool] = {}
        for movto_id in movto_ids_sql:
            exists = int(movto_id) in found
            out[int(movto_id)] = exists
            out[str(int(movto_id))] = exists
        return out

    def check_nota_fiscal_exists_bulk(self, invoice_ids: List[Any]) -> Dict[Any, bool]:
        invoice_ids_sql = self._normalize_bigint_ids(invoice_ids)
        if not invoice_ids_sql:
            return {}

        sql = """
            select distinct inv.grid as invoice_id
            from movto inv
            left join movto_map mp
              on mp.child = inv.grid
            left join movto s
              on s.grid = mp.parent
            join nota_fiscal nf
              on nf.mlid = coalesce(nullif(inv.mlid, 0), nullif(s.mlid, 0))
            where inv.grid = any(%(invoice_ids)s::bigint[])
              and coalesce(nullif(inv.mlid, 0), nullif(s.mlid, 0)) is not null
        """

        found: set = set()
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, {"invoice_ids": invoice_ids_sql})
                for (invoice_id,) in (cur.fetchall() or []):
                    if invoice_id not in (None, "", 0, "0"):
                        found.add(int(invoice_id))

        out: Dict[Any, bool] = {}
        for invoice_id in invoice_ids_sql:
            exists = int(invoice_id) in found
            out[int(invoice_id)] = exists
            out[str(int(invoice_id))] = exists
        return out

    def get_nfe_attachments_bulk(self, invoice_ids: List[Any]) -> Dict[Any, Dict[str, Any]]:
        invoice_ids_sql = self._normalize_bigint_ids(invoice_ids)
        if not invoice_ids_sql:
            return {}

        out: Dict[Any, Dict[str, Any]] = {}
        for invoice_id in invoice_ids_sql:
            out[int(invoice_id)] = {"exists": False, "attachments": []}
            out[str(int(invoice_id))] = out[int(invoice_id)]

        with self._connect() as conn:
            nfe_xml_cols = self._get_table_columns(conn, "nfe_xml", schema="public")
            if not nfe_xml_cols:
                return out
            nf_cols = self._get_table_columns(conn, "nota_fiscal", schema="public")
            nfe_cols = self._get_table_columns(conn, "nfe", schema="public")
            nfs_cols = self._get_table_columns(conn, "nota_fiscal_situacao", schema="public")
            xml_col = self._pick_existing_column(
                nfe_xml_cols,
                [
                    "xml",
                    "xml_nfe",
                    "nfe_xml",
                    "fonte_xml",
                    "arquivo_xml",
                    "conteudo_xml",
                    "conteudo",
                    "arquivo",
                ],
            )
            danfe_col = self._pick_existing_column(
                nfe_xml_cols,
                [
                    "danfe_pdf",
                    "arquivo_pdf",
                    "pdf",
                    "danfe",
                    "danfe_arquivo",
                    "danfe_bytes",
                ],
            )

            xml_expr = f"x.{xml_col} as xml_raw" if xml_col else "NULL as xml_raw"
            danfe_expr = f"x.{danfe_col} as danfe_raw" if danfe_col else "NULL as danfe_raw"

            nf_num_col = self._pick_existing_column(nf_cols, ["numero", "num", "nr", "nro"])
            nf_serie_col = self._pick_existing_column(nf_cols, ["serie", "ser"])
            nf_chave_col = self._pick_existing_column(nf_cols, ["chave", "chave_nfe", "chave_acesso", "chave_acesso_nfe"])
            nfe_chave_col = self._pick_existing_column(nfe_cols, ["chave_acesso", "chave", "chave_nfe", "chave_acesso_nfe"]) if nfe_cols else ""
            nf_num_expr = f"nf.{nf_num_col} as nf_numero" if nf_num_col else "NULL as nf_numero"
            nf_serie_expr = f"nf.{nf_serie_col} as nf_serie" if nf_serie_col else "NULL as nf_serie"
            nf_chave_expr = f"nf.{nf_chave_col} as nf_chave" if nf_chave_col else "NULL as nf_chave"

            nfe_xml_fk_nf = self._pick_existing_column(nfe_xml_cols, ["nota_fiscal", "nota_fiscal_id"])
            nfe_xml_fk_nfe = self._pick_existing_column(nfe_xml_cols, ["nfe", "nfe_id"])
            nfe_fk_nf = self._pick_existing_column(nfe_cols, ["nota_fiscal", "nota_fiscal_id"]) if nfe_cols else ""

            join_sql = ""
            if nfe_xml_fk_nf:
                join_sql = f"left join nfe_xml x on x.{nfe_xml_fk_nf} = nf.grid"
            elif nfe_xml_fk_nfe and nfe_fk_nf:
                join_sql = f"""
                    left join nfe n
                      on n.{nfe_fk_nf} = nf.grid
                    left join nfe_xml x
                      on x.{nfe_xml_fk_nfe} = n.grid
                """
                if not nf_chave_col and nfe_chave_col:
                    nf_chave_expr = f"n.{nfe_chave_col} as nf_chave"
            else:
                join_sql = "left join nfe_xml x on 1=0"

            nfs_join_sql = ""
            if nfs_cols:
                nfs_nf_fk = self._pick_existing_column(nfs_cols, ["nota_fiscal", "nota_fiscal_id"])
                nfs_situacao_col = self._pick_existing_column(nfs_cols, ["situacao", "tipo_situacao", "codigo_situacao", "status"])
                if nfs_nf_fk and nfs_situacao_col:
                    nfs_join_sql = f"join nota_fiscal_situacao nfs on nfs.{nfs_nf_fk} = nf.grid and nfs.{nfs_situacao_col} = 310"

            sql = f"""
                select
                    l.invoice_id,
                    nf.grid as nota_fiscal_id,
                    {nf_num_expr},
                    {nf_serie_expr},
                    {nf_chave_expr},
                    {xml_expr},
                    {danfe_expr}
                from (
                    select
                        inv.grid as invoice_id,
                        coalesce(nullif(inv.mlid, 0), nullif(s.mlid, 0)) as mlid
                    from movto inv
                    left join movto_map mp
                      on mp.child = inv.grid
                    left join movto s
                      on s.grid = mp.parent
                    where inv.grid = any(%(invoice_ids)s::bigint[])
                ) l
                join nota_fiscal nf
                  on nf.mlid = l.mlid
                {nfs_join_sql}
                {join_sql}
                order by l.invoice_id, nf.grid desc
            """

            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(sql, {"invoice_ids": invoice_ids_sql})
                rows = [dict(r) for r in (cur.fetchall() or [])]

        seen_nf: Dict[int, set] = {}
        for r in rows:
            inv_id = r.get("invoice_id")
            if inv_id in (None, "", 0, "0"):
                continue
            try:
                inv_id_i = int(inv_id)
            except Exception:
                continue

            nf_id = r.get("nota_fiscal_id")
            try:
                nf_id_i = int(nf_id)
            except Exception:
                nf_id_i = 0

            if inv_id_i not in seen_nf:
                seen_nf[inv_id_i] = set()
            if nf_id_i and nf_id_i in seen_nf[inv_id_i]:
                continue
            if nf_id_i:
                seen_nf[inv_id_i].add(nf_id_i)

            out[inv_id_i]["exists"] = True
            out[str(inv_id_i)] = out[inv_id_i]

            nf_num = str(r.get("nf_numero") or "").strip()
            nf_serie = str(r.get("nf_serie") or "").strip()
            nf_chave = str(r.get("nf_chave") or "").strip()

            name_parts = []
            if nf_num:
                name_parts.append(nf_num)
            if nf_serie:
                name_parts.append(f"serie{nf_serie}")
            if nf_id_i and name_parts and len(name_parts) == 1 and name_parts[0].startswith("serie"):
                name_parts.append(str(nf_id_i))
            if not name_parts and nf_chave:
                name_parts.append(nf_chave[-12:])
            if not name_parts and nf_id_i:
                name_parts.append(str(nf_id_i))
            suffix = "_".join(name_parts) if name_parts else str(inv_id_i)

            xml_bytes = self._blob_to_bytes(r.get("xml_raw"))
            if xml_bytes and xml_bytes.strip():
                out[inv_id_i]["attachments"].append({"data": xml_bytes, "filename": f"nfe_{suffix}.xml", "mime_type": "application/xml"})
                out[str(inv_id_i)] = out[inv_id_i]

            danfe_bytes = self._blob_to_bytes(r.get("danfe_raw"))
            if danfe_bytes and danfe_bytes.strip():
                out[inv_id_i]["attachments"].append({"data": danfe_bytes, "filename": f"danfe_{suffix}.pdf", "mime_type": "application/pdf"})
                out[str(inv_id_i)] = out[inv_id_i]

        return out


    def get_boleto_email_payload(self, invoice_id) -> Dict[str, Any]:
        try:
            invoice_id_val = int(invoice_id)
        except (ValueError, TypeError):
            invoice_id_val = invoice_id

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
                bi.multa_prazo,
                bi.multa_valor,
                bi.multa_perc,
                bi.juros_valor_dia,
                bi.juros_perc_mes,
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
                pessoa_nome_f(e.grid) as cedente_nome,
                e.grid as cedente_empresa_id
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
                cur.execute(sql, (invoice_id_val,))
                row = cur.fetchone()
                if row:
                    try:
                        row["cedente_endereco"] = self._empresa_endereco_str(conn, row.get("cedente_empresa_id"))
                    except Exception:
                        row["cedente_endereco"] = ""
                    try:
                        row["sacado_cep"] = self._boleto_info_sacado_cep(conn, row.get("boleto_info"))
                    except Exception:
                        row["sacado_cep"] = ""
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
            email_note = "Observação: anexo boleto da fatura."

        try:
            sig = self.get_sale_signature_pdf(invoice_id)
            has_signature = (sig or {}).get("exists") or bool((sig or {}).get("attachments"))
        except Exception:
            has_signature = False

        if email_note in ("Observação: anexo boleto da fatura.", "Observação: o boleto segue em anexo.") and has_signature:
            email_note = "Observação: Anexo boleto e cupom assinado."

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
                bi.multa_prazo,
                bi.multa_valor,
                bi.multa_perc,
                bi.juros_valor_dia,
                bi.juros_perc_mes,
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
                pessoa_nome_f(e.grid) as cedente_nome,
                e.grid as cedente_empresa_id
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
                addr_cache: Dict[int, str] = {}
                cep_cache: Dict[int, str] = {}
                cur.execute(sql, (invoice_ids_sql,))
                rows = cur.fetchall() or []
                for r in rows:
                    d = dict(r)
                    try:
                        eid = int(d.get("cedente_empresa_id") or 0)
                    except Exception:
                        eid = 0
                    if eid:
                        if eid not in addr_cache:
                            addr_cache[eid] = self._empresa_endereco_str(conn, eid)
                        d["cedente_endereco"] = addr_cache.get(eid) or ""
                    try:
                        bi_id = int(d.get("boleto_info") or 0)
                    except Exception:
                        bi_id = 0
                    if bi_id:
                        if bi_id not in cep_cache:
                            cep_cache[bi_id] = self._boleto_info_sacado_cep(conn, bi_id)
                        d["sacado_cep"] = cep_cache.get(bi_id) or ""
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
                email_note = "Boleto: localizado."

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

    def list_generated_boletos(self, window_start: datetime, window_end: datetime) -> List[Dict[str, Any]]:
        sql = """
            SELECT 
                b.grid as boleto_grid,
                m.grid as movto_id,
                m.pessoa as customer_id,
                pessoa_nome_f(m.pessoa) as cliente,
                COALESCE(c.email, '') as customer_email,
                bi.documento,
                (bi.data_geracao::timestamp) as generated_at,
                bi.valor as valor
            FROM boleto b
            JOIN movto m ON m.grid = b.movto
            LEFT JOIN cliente c ON c.grid = m.pessoa
            LEFT JOIN boleto_info bi ON bi.grid = b.boleto_info
            WHERE (bi.data_geracao::timestamp) >= %s AND (bi.data_geracao::timestamp) <= %s
              AND (b.situacao IS NULL OR b.situacao < 5)
            ORDER BY (bi.data_geracao::timestamp) ASC
        """
        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(sql, (window_start, window_end))
                return [dict(r) for r in cur.fetchall()]

    def list_boletos_by_grids(self, boleto_grids: List[Any], *, include_closed: bool = False) -> List[Dict[str, Any]]:
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
            return []

        saldo_filter = "" if include_closed else " AND (t.valor - t.valor_desconto - t.valor_baixado) > 0"
        sql = f"""
            WITH titulos AS (
                SELECT
                    b.grid AS boleto_grid,
                    b.movto AS movto_id,
                    pessoa_nome_f(m.empresa) AS empresa,
                    m.pessoa AS customer_id,
                    p.codigo AS codigo_cliente,
                    p.nome AS cliente,
                    m.conta_debitar AS conta,
                    conta_nome_f(m.conta_debitar) AS conta_nome,
                    m.data AS data,
                    m.valor AS valor,
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
                    ), 0) AS valor_baixado,
                    b.boleto_info AS boleto_info_id
                FROM boleto b
                JOIN movto m
                  ON m.grid = b.movto
                JOIN pessoa p
                  ON p.grid = m.pessoa
                WHERE b.grid = ANY(%s)
            )
            SELECT
                t.movto_id,
                t.empresa,
                t.customer_id,
                t.codigo_cliente,
                t.cliente,
                t.conta,
                t.conta_nome,
                t.data,
                bi.vencto AS vencto,
                t.valor,
                t.valor_desconto,
                t.valor_baixado,
                (t.valor - t.valor_desconto - t.valor_baixado) AS saldo_em_aberto,
                COALESCE(c.email, '') AS customer_email,
                bi.documento AS documento,
                t.boleto_grid
            FROM titulos t
            JOIN pessoa pcli
              ON pcli.grid = t.customer_id
            LEFT JOIN cliente c
              ON c.grid = t.customer_id
            LEFT JOIN boleto_info bi
              ON bi.grid = t.boleto_info_id
            WHERE coalesce(pcli.tipo, pcli.tipo_pessoa::text, '') ILIKE '%%C%%'
            {saldo_filter}
            ORDER BY t.empresa, t.cliente, bi.vencto, t.boleto_grid
        """
        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(sql, (bg_sql,))
                return [dict(r) for r in cur.fetchall()]

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
                pessoa_nome_f(e.grid) as cedente_nome,
                e.grid as cedente_empresa_id
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
                    try:
                        eid = int(data.get("cedente_empresa_id") or 0)
                    except Exception:
                        eid = 0
                    if eid:
                        try:
                            data["cedente_endereco"] = self._empresa_endereco_str(conn, eid)
                        except Exception:
                            data["cedente_endereco"] = ""
                    try:
                        bi_id = int(data.get("boleto_info") or 0)
                    except Exception:
                        bi_id = 0
                    if bi_id:
                        try:
                            data["sacado_cep"] = self._boleto_info_sacado_cep(conn, bi_id)
                        except Exception:
                            data["sacado_cep"] = ""
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
                        email_note = "Boleto: localizado."
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

        dt_expr = "inv.hora" if meta.get("has_hora") else "inv.data"
        sale_dt_expr = "s.hora" if meta.get("has_hora") else "s.data"
        sale2_dt_expr = "s2.hora" if meta.get("has_hora") else "s2.data"
        prod_expr = "produto_nome_f(l.produto)" if meta.get("has_produto_nome_f") else "l.produto::text"

        sql = f"""
            with invoice_ids as (
                select unnest(%s::bigint[]) as invoice_id
            ),
            sale_links as (
                select
                    mm.child as invoice_id,
                    mm.parent as sale_id
                from movto_map mm
                join invoice_ids i
                  on i.invoice_id = mm.child
                where mm.parent is not null
                  and mm.child is not null
                  and mm.parent <> 0
                  and mm.child <> 0

                union

                select
                    mm1.child as invoice_id,
                    mm2.parent as sale_id
                from movto_map mm1
                join movto_map mm2
                  on mm2.child = mm1.parent
                join invoice_ids i
                  on i.invoice_id = mm1.child
                where mm1.parent is not null
                  and mm1.child is not null
                  and mm1.parent <> 0
                  and mm1.child <> 0
                  and mm2.parent is not null
                  and mm2.child is not null
                  and mm2.parent <> 0
                  and mm2.child <> 0
            ),
            sales_distinct as (
                select distinct on (sl.invoice_id, coalesce(nullif(s.mlid, 0), s.grid))
                    sl.invoice_id,
                    sl.sale_id
                from sale_links sl
                join movto s
                  on s.grid = sl.sale_id
                where sl.sale_id is not null
                  and sl.sale_id <> 0
                order by sl.invoice_id, coalesce(nullif(s.mlid, 0), s.grid), s.grid desc
            ),
            has_sales as (
                select distinct sd.invoice_id from sales_distinct sd
            )
            select
                inv.grid as invoice_id,
                inv.valor as invoice_amount,
                inv.documento as documento,
                inv.grid as doc_movto_id,
                {dt_expr} as dt,
                nullif(trim(coalesce({prod_expr}::text, '')), '') as product_name,
                sum(coalesce(l.quantidade, 0)) as quantity,
                sum(coalesce(l.valor, 0)) as item_total
            from movto inv
            join lancto l
              on l.mlid = inv.mlid
            where inv.grid = any(%s)
              and l.operacao = 'V'
              and not exists (select 1 from has_sales hs where hs.invoice_id = inv.grid)
            group by inv.grid, inv.valor, inv.documento, inv.grid, {dt_expr}, product_name

            union all

            select
                inv.grid as invoice_id,
                inv.valor as invoice_amount,
                s.documento as documento,
                s.grid as doc_movto_id,
                {sale_dt_expr} as dt,
                nullif(trim(coalesce({prod_expr}::text, '')), '') as product_name,
                sum(coalesce(l.quantidade, 0)) as quantity,
                sum(coalesce(l.valor, 0)) as item_total
            from sales_distinct sd
            join movto inv
              on inv.grid = sd.invoice_id
            join movto s
              on s.grid = sd.sale_id
            join lancto l
              on l.mlid = s.mlid
            where l.operacao = 'V'
            group by inv.grid, inv.valor, s.documento, s.grid, {sale_dt_expr}, product_name

            order by invoice_id, dt, documento, product_name
        """

        out: Dict[Any, Dict[str, Any]] = {}
        header_map: Dict[int, Dict[str, str]] = {}
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, (invoice_ids_sql, invoice_ids_sql))
                rows = cur.fetchall() or []

            try:
                movto_cols = self._get_table_columns(conn, "movto", schema="public")
                pessoa_cols = self._get_table_columns(conn, "pessoa", schema="public")
                doc_col = self._pick_existing_column(pessoa_cols, ["cpf", "cnpj", "cpf_cnpj", "cnpj_cpf", "documento"])
                obs_col = self._pick_existing_column(movto_cols, ["obs", "observacao", "observacoes", "historico", "hist", "obs1", "obs2"])
                doc_expr = f"p.{doc_col}::text as customer_doc" if doc_col else "NULL::text as customer_doc"
                obs_expr = f"inv.{obs_col}::text as invoice_obs" if obs_col else "NULL::text as invoice_obs"
                with conn.cursor() as cur:
                    cur.execute(
                        f"""
                        select
                            inv.grid as invoice_id,
                            {doc_expr},
                            {obs_expr}
                        from movto inv
                        left join pessoa p
                          on p.grid = inv.pessoa
                        where inv.grid = any(%(invoice_ids)s::bigint[])
                        """,
                        {"invoice_ids": invoice_ids_sql},
                    )
                    for invoice_id, customer_doc, invoice_obs in (cur.fetchall() or []):
                        if invoice_id in (None, "", 0, "0"):
                            continue
                        try:
                            inv_int = int(invoice_id)
                        except Exception:
                            continue
                        header_map[inv_int] = {
                            "customer_doc": str(customer_doc or "").strip(),
                            "invoice_obs": str(invoice_obs or "").strip(),
                        }
            except Exception:
                header_map = {}

        tmp: Dict[int, Dict[str, Any]] = {}
        for invoice_id, invoice_amount, documento, doc_movto_id, dt, product_name, quantity, item_total in rows:
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
                doc = {"documento": doc_key, "dt": dt, "items": {}, "total": 0.0, "movto_id": doc_movto_id}
                slot["documents"][doc_key] = doc
            if doc.get("dt") in (None, "") and dt not in (None, ""):
                doc["dt"] = dt
            if doc.get("movto_id") in (None, "", 0, "0") and doc_movto_id not in (None, "", 0, "0"):
                doc["movto_id"] = doc_movto_id
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

        movto_doc_ids: List[int] = []
        for slot in (tmp or {}).values():
            for d in (slot.get("documents") or {}).values():
                mid = d.get("movto_id")
                if mid in (None, "", 0, "0"):
                    continue
                try:
                    movto_doc_ids.append(int(mid))
                except Exception:
                    continue
        movto_doc_ids = sorted(list(set(movto_doc_ids)))

        placa_km_by_movto: Dict[int, Dict[str, str]] = {}
        if movto_doc_ids:
            try:
                with self._connect() as conn:
                    mi_cols = self._get_table_columns(conn, "movto_info", schema="public")
                    fk_col = self._pick_existing_column(mi_cols, ["movto", "movto_id"])
                    tipo_col = self._pick_existing_column(mi_cols, ["tipo", "tipo_info", "chave"])
                    info_col = self._pick_existing_column(mi_cols, ["info", "valor", "conteudo", "texto"])
                    if fk_col and tipo_col and info_col:
                        with conn.cursor() as cur:
                            cur.execute(
                                f"""
                                select
                                    mi.{fk_col} as movto_id,
                                    lower(mi.{tipo_col}::text) as tipo,
                                    mi.{info_col}::text as info
                                from movto_info mi
                                where mi.{fk_col} = any(%(ids)s::bigint[])
                                  and lower(mi.{tipo_col}::text) in ('placa','km','media_km','quilometragem','odometro','hodometro','nao_detalhar_km')
                                """,
                                {"ids": movto_doc_ids},
                            )
                            for movto_id, tipo, info in (cur.fetchall() or []):
                                if movto_id in (None, "", 0, "0"):
                                    continue
                                try:
                                    mid = int(movto_id)
                                except Exception:
                                    continue
                                slot = placa_km_by_movto.get(mid)
                                if not slot:
                                    slot = {}
                                    placa_km_by_movto[mid] = slot
                                tipo = str(tipo or "").strip().lower()
                                info = str(info or "").strip()
                                if not tipo or not info:
                                    continue
                                if tipo == "placa":
                                    cand = re.sub(r"[^A-Za-z0-9]+", "", info).upper()
                                    if not cand:
                                        continue
                                    is_valid = False
                                    if len(cand) == 7:
                                        if re.match(r"^[A-Z]{3}[0-9]{4}$", cand):
                                            is_valid = True
                                        elif re.match(r"^[A-Z]{3}[0-9][A-Z0-9][0-9]{2}$", cand):
                                            is_valid = True
                                    if is_valid:
                                        if not slot.get("placa_valid"):
                                            slot["placa"] = cand
                                            slot["placa_valid"] = True
                                    else:
                                        if not slot.get("placa"):
                                            slot["placa"] = cand
                                            slot["placa_valid"] = False
                                elif tipo == "nao_detalhar_km":
                                    slot["no_detail_km"] = True
                                else:
                                    if not re.search(r"\d", info):
                                        continue
                                    if slot.get("no_detail_km"):
                                        continue
                                    if tipo == "media_km":
                                        if not slot.get("km"):
                                            slot["km"] = info
                                        continue
                                    if not slot.get("km"):
                                        slot["km"] = info
            except Exception:
                placa_km_by_movto = {}

        for invoice_id in invoice_ids_sql:
            slot = tmp.get(int(invoice_id))
            hdr = header_map.get(int(invoice_id)) or {}
            if not slot:
                out[invoice_id] = {
                    "purchase_dt": None,
                    "purchase_dt_start": None,
                    "purchase_dt_end": None,
                    "invoice_amount": None,
                    "items_total": None,
                    "documents": [],
                    "customer_doc": hdr.get("customer_doc") or "",
                    "invoice_obs": hdr.get("invoice_obs") or "",
                }
                out[str(invoice_id)] = out[invoice_id]
                continue
            documents = []
            items_total = 0.0
            for doc in slot["documents"].values():
                mid = doc.get("movto_id")
                try:
                    mid_int = int(mid) if mid not in (None, "", 0, "0") else None
                except Exception:
                    mid_int = None
                pk = placa_km_by_movto.get(mid_int) if mid_int is not None else {}
                items = [v for _, v in sorted((doc.get("items") or {}).items(), key=lambda kv: kv[0].lower())]
                total = float(doc.get("total") or 0)
                items_total += total
                documents.append(
                    {
                        "documento": doc.get("documento"),
                        "dt": doc.get("dt"),
                        "total": total,
                        "items": items,
                        "movto_id": doc.get("movto_id"),
                        "placa": (pk or {}).get("placa") or "",
                        "km": (pk or {}).get("km") or "",
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
                "customer_doc": hdr.get("customer_doc") or "",
                "invoice_obs": hdr.get("invoice_obs") or "",
            }
            out[str(invoice_id)] = out[invoice_id]

        return out

    @staticmethod
    def _format_km_br(value: Any) -> str:
        s = str(value or "").strip()
        if not s:
            return ""
        cleaned = re.sub(r"[^\d,\.]", "", s)
        if not cleaned:
            return s
        try:
            if "," in cleaned and "." in cleaned:
                num = float(cleaned.replace(".", "").replace(",", "."))
            elif "," in cleaned:
                num = float(cleaned.replace(".", "").replace(",", "."))
            else:
                num = float(cleaned)
            return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return s

    @staticmethod
    def _normalize_placa(value: Any) -> str:
        cand = re.sub(r"[^A-Za-z0-9]+", "", str(value or "").strip()).upper()
        if not cand:
            return ""
        if len(cand) != 7:
            return cand
        if re.match(r"^[A-Z]{3}[0-9]{4}$", cand):
            return cand
        if re.match(r"^[A-Z]{3}[0-9][A-Z0-9][0-9]{2}$", cand):
            return cand
        return cand

    def get_placa_km_text_bulk(self, invoice_ids: List[Any]) -> Dict[Any, str]:
        invoice_ids_sql = self._normalize_bigint_ids(invoice_ids)
        if not invoice_ids_sql:
            return {}

        invoice_to_sales: Dict[int, List[int]] = {int(i): [int(i)] for i in invoice_ids_sql}

        links: List[tuple] = []
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    select
                        mm.child as invoice_id,
                        mm.parent as sale_id
                    from movto_map mm
                    where mm.child = any(%(invoice_ids)s::bigint[])
                      and mm.parent is not null
                      and mm.child is not null
                      and mm.parent <> 0
                      and mm.child <> 0

                    union all

                    select
                        mm1.child as invoice_id,
                        mm2.parent as sale_id
                    from movto_map mm1
                    join movto_map mm2
                      on mm2.child = mm1.parent
                    where mm1.child = any(%(invoice_ids)s::bigint[])
                      and mm1.parent is not null
                      and mm1.child is not null
                      and mm1.parent <> 0
                      and mm1.child <> 0
                      and mm2.parent is not null
                      and mm2.child is not null
                      and mm2.parent <> 0
                      and mm2.child <> 0
                    """,
                    {"invoice_ids": invoice_ids_sql},
                )
                links = cur.fetchall() or []

        for invoice_id, sale_id in links:
            try:
                inv_int = int(invoice_id)
                sale_int = int(sale_id)
            except Exception:
                continue
            slot = invoice_to_sales.get(inv_int)
            if not slot:
                slot = [inv_int]
                invoice_to_sales[inv_int] = slot
            slot.append(sale_int)

        sale_ids_all: List[int] = []
        for mids in invoice_to_sales.values():
            for m in mids:
                if m not in (None, 0):
                    sale_ids_all.append(int(m))
        sale_ids_all = sorted(list(set(sale_ids_all)))

        info_rows: List[tuple] = []
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    select
                        mi.movto as movto_id,
                        lower(mi.tipo::text) as tipo,
                        mi.info::text as info
                    from movto_info mi
                    where mi.movto = any(%(movto_ids)s::bigint[])
                      and lower(mi.tipo::text) in ('placa','km','media_km','nao_detalhar_km')
                    order by mi.movto, mi.grid
                    """,
                    {"movto_ids": sale_ids_all},
                )
                info_rows = cur.fetchall() or []

        sale_slot: Dict[int, Dict[str, Any]] = {}
        for movto_id, tipo, info in info_rows:
            try:
                mid = int(movto_id)
            except Exception:
                continue
            slot = sale_slot.get(mid)
            if not slot:
                slot = {"no_detail": False, "placa": "", "placa_valid": False, "km": ""}
                sale_slot[mid] = slot
            tipo = str(tipo or "").strip().lower()
            info = str(info or "").strip()
            if not tipo or not info:
                continue
            if tipo == "nao_detalhar_km":
                slot["no_detail"] = True
                continue
            if tipo == "placa":
                cand = self._normalize_placa(info)
                if not cand:
                    continue
                is_valid = bool(re.match(r"^[A-Z]{3}[0-9]{4}$", cand) or re.match(r"^[A-Z]{3}[0-9][A-Z0-9][0-9]{2}$", cand))
                if is_valid:
                    if not slot.get("placa_valid"):
                        slot["placa"] = cand
                        slot["placa_valid"] = True
                else:
                    if not slot.get("placa"):
                        slot["placa"] = cand
                continue
            if tipo == "km":
                if not slot.get("km") and re.search(r"\d", info):
                    slot["km"] = info
                continue
            if tipo == "media_km":
                continue

        out: Dict[Any, str] = {}
        for inv_id in invoice_ids_sql:
            inv_int = int(inv_id)
            sale_ids = invoice_to_sales.get(inv_int) or [inv_int]
            sale_ids = [int(s) for s in sale_ids if s not in (None, 0)]
            seen: Dict[str, float] = {}
            ordered: List[str] = []
            for s in sale_ids:
                slot = sale_slot.get(int(s)) or {}
                if slot.get("no_detail"):
                    continue
                placa = str(slot.get("placa") or "").strip().upper()
                if not placa:
                    continue
                km_raw = str(slot.get("km") or "").strip()
                km_txt = self._format_km_br(km_raw)
                km_num = -1.0
                if km_txt:
                    try:
                        km_num = float(km_txt.replace(".", "").replace(",", "."))
                    except Exception:
                        km_num = -1.0
                prev = seen.get(placa)
                if prev is None:
                    seen[placa] = km_num
                    ordered.append(placa)
                else:
                    if km_num > prev:
                        seen[placa] = km_num

            lines: List[str] = []
            for placa in ordered:
                km_num = seen.get(placa, -1.0)
                km_txt = ""
                if km_num >= 0:
                    km_txt = self._format_km_br(km_num)
                if km_txt:
                    lines.append(f" Placa: {placa} - KM: {km_txt}")
                else:
                    lines.append(f" Placa: {placa}")
            text = "\n".join(lines).rstrip()
            out[inv_int] = text
            out[str(inv_int)] = text
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
            select
                a.movto as movto_id,
                a.grid as grid,
                a.descricao,
                a.extensao,
                a.anexo,
                a.ts
            from public.anexo a
            where a.movto = any(%(movto_ids)s::bigint[])
              and lower(coalesce(a.descricao, '')) like '%%assinatura%%'

            union all

            select
                mm.child as movto_id,
                a.grid as grid,
                a.descricao,
                a.extensao,
                a.anexo,
                a.ts
            from movto_map mm
            join public.anexo a
              on a.movto = mm.parent
            where mm.child = any(%(movto_ids)s::bigint[])
              and mm.parent is not null
              and mm.child is not null
              and mm.parent <> 0
              and mm.child <> 0
              and lower(coalesce(a.descricao, '')) like '%%assinatura%%'

            order by movto_id, ts desc nulls last, grid desc
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
