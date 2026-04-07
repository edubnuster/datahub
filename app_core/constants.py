# -*- coding: utf-8 -*-
from pathlib import Path
import sys

APP_TITLE = "DataHub"
CONFIG_FILENAME = "config.json"
AUDIT_FILENAME = "audit.log"
LICENSE_FILENAME = "databrev.key"
LICENSE_SECRET = "DATABREV-LICENSE-2026"

MASTER_USERNAME = "databrev"
MASTER_PASSWORD = "270810"

DEFAULT_LIST_SQL = """
select
    c.grid as customer_id,
    pessoa_nome_f(l.empresa) as last_purchase_company,
    c.codigo as customer_code,
    c.nome as customer_name,
    coalesce(co.nome, 'Sem conta') as account_name,
    max(case when pc.pessoa is not null then 1 else 0 end) as has_account,
    coalesce(pc.lim_credito, 0) as credit_limit,
    max(l.data) as last_purchase_date,
    case c.flag
        when 'A' then 'Ativo'
        when 'I' then 'Inativo'
        when 'D' then 'Deletado'
        else coalesce(c.flag, '')
    end as customer_status
from cliente c
join lancto l
    on l.pessoa = c.grid
left join pessoa_conta pc
    on pc.pessoa = c.grid
left join conta co
    on co.codigo = pc.conta
where l.operacao = 'V'
group by
    c.grid,
    l.empresa,
    c.codigo,
    c.nome,
    coalesce(co.nome, 'Sem conta'),
    coalesce(pc.lim_credito, 0),
    c.flag,
    pessoa_nome_f(l.empresa)
having max(l.data) < current_date - interval '3 months'
order by pessoa_nome_f(l.empresa), max(l.data)
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
      AND m.vencto <= current_date
      AND COALESCE(m.valor, 0) > 0
      AND m.child = 0
      AND m.conta_debitar LIKE '1.3.04%'
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
        "host": "smtp.zoho.com",
        "password": "6UBwERuJiqJi",
        "port": 465
    },
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

CONFIG_PATH = app_dir() / CONFIG_FILENAME
AUDIT_PATH = app_dir() / AUDIT_FILENAME
LICENSE_PATH = app_dir() / LICENSE_FILENAME
