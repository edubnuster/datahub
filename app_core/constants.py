# -*- coding: utf-8 -*-
from pathlib import Path
import sys

APP_TITLE = "DataHub"
CONFIG_FILENAME = "config.json"
AUDIT_FILENAME = "audit.log"
LOG_FOLDERNAME = "log"
LICENSE_FILENAME = "databrev.key"
LICENSE_SECRET = "DATABREV-LICENSE-2026"

MASTER_USERNAME = "databrev"
MASTER_PASSWORD = "270810"

DEFAULT_LIST_SQL = """
WITH last_purchase AS (
    SELECT DISTINCT ON (l.pessoa)
        l.pessoa AS customer_id,
        l.empresa AS last_company_id,
        l.data AS last_purchase_date
    FROM lancto l
    WHERE l.operacao = 'V'
    ORDER BY l.pessoa, l.data DESC
)
SELECT
    c.grid AS customer_id,
    pessoa_nome_f(lp.last_company_id) AS last_purchase_company,
    c.codigo AS customer_code,
    c.nome AS customer_name,
    COALESCE(co.nome, 'Sem conta') AS account_name,
    MAX(CASE WHEN pc.pessoa IS NOT NULL THEN 1 ELSE 0 END) AS has_account,
    COALESCE(MAX(pc.lim_credito), 0) AS credit_limit,
    lp.last_purchase_date,
    CASE c.flag
        WHEN 'A' THEN 'Ativo'
        WHEN 'I' THEN 'Inativo'
        WHEN 'D' THEN 'Deletado'
        ELSE COALESCE(c.flag, '')
    END AS customer_status
FROM cliente c
JOIN last_purchase lp
  ON lp.customer_id = c.grid
LEFT JOIN pessoa_conta pc
  ON pc.pessoa = c.grid
LEFT JOIN conta co
  ON co.codigo = pc.conta
WHERE lp.last_purchase_date < current_date - (interval '1 month' * %(inactive_months)s)
GROUP BY
    c.grid,
    lp.last_company_id,
    lp.last_purchase_date,
    c.codigo,
    c.nome,
    COALESCE(co.nome, 'Sem conta'),
    c.flag
ORDER BY lp.last_purchase_date
"""

DEFAULT_OPEN_INVOICES_SQL = """
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
        m.vencto,
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
      AND m.conta_debitar LIKE '1.3.04%'
      AND EXISTS (SELECT 1 FROM boleto b2 WHERE b2.movto = m.grid)
)
SELECT
    movto_id,
    pessoa_nome_f(t.empresa) AS empresa,
    t.customer_id,
    t.codigo_cliente,
    t.cliente,
    t.conta,
    t.conta_nome,
    t.data,
    t.vencto,
    t.valor,
    t.valor_desconto,
    t.valor_baixado,
    (t.valor - t.valor_desconto - t.valor_baixado) AS saldo_em_aberto
FROM titulos t
WHERE (t.valor - t.valor_desconto - t.valor_baixado) > 0
ORDER BY empresa, cliente, t.vencto
"""

DEFAULT_CONFIG = {
    "connection": {
        "host": "127.0.0.1",
        "port": 5432,
        "dbname": "",
        "user": "postgres",
        "password": "",
        "client_encoding": "LATIN1",
    },
    "security": {
        "users": []
    },
    "smtp": {
        "email": "app@databrev.com.br",
        "sender_name": "",
        "host": "smtp.zoho.com",
        "password": "",
        "port": 465,
        "delay_seconds": 5,
    },
    "financeiro_agendas": [],
    "queries": {
        "list_inactive_customers_sql": DEFAULT_LIST_SQL,
        "list_open_invoices_sql": DEFAULT_OPEN_INVOICES_SQL,
        "delete_customer_sql": """
            update cliente
               set flag = 'D'
             where grid = %(customer_id)s
        """,
        "inactivate_customer_sql": """
            update cliente
               set flag = 'I'
             where grid = %(customer_id)s
        """,
        "disable_credit_sql": """
            update pessoa_conta pc
               set lim_credito = 0
             where pc.pessoa = %(customer_id)s
        """
    }
}

def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent

def log_dir() -> Path:
    return app_dir() / LOG_FOLDERNAME

CONFIG_PATH = app_dir() / CONFIG_FILENAME
AUDIT_PATH = log_dir() / AUDIT_FILENAME
LICENSE_PATH = app_dir() / LICENSE_FILENAME
