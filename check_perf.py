import json
import psycopg2
from app_core.config_manager import ConfigManager

def main():
    print("Starting script")
    import json
    with open("config.json", "r", encoding="utf-8") as f:
        cfg = json.load(f)
    print("Config loaded")
    db_cfg = cfg.get("connection", {})
    
    conn = psycopg2.connect(
        host=db_cfg.get("host", "localhost"),
        port=db_cfg.get("port", 5432),
        dbname=db_cfg.get("dbname"),
        user=db_cfg.get("user"),
        password=db_cfg.get("password")
    )
    conn.autocommit = True
    cursor = conn.cursor()

    queries = [
        "SELECT count(*) FROM movto;",
        "SELECT count(*) FROM lancto;",
        "SELECT count(*) FROM boleto;",
        "SELECT count(*) FROM nota_fiscal;",
        "SELECT count(*) FROM nfe;",
        "SELECT count(*) FROM nfe_xml;",
        "SELECT count(*) FROM movto_map;",
    ]
    for q in queries:
        try:
            cursor.execute(q)
            print(f"{q}: {cursor.fetchone()[0]}")
        except Exception as e:
            print(f"Erro em {q}: {e}")

    # Check indices on movto
    cursor.execute("""
        SELECT indexname, indexdef 
        FROM pg_indexes 
        WHERE tablename = 'movto';
    """)
    print("\n--- Indices in movto ---")
    for row in cursor.fetchall():
        print(row[0], ":", row[1])

    # Check indices on nota_fiscal
    cursor.execute("""
        SELECT indexname, indexdef 
        FROM pg_indexes 
        WHERE tablename = 'nota_fiscal';
    """)
    print("\n--- Indices in nota_fiscal ---")
    for row in cursor.fetchall():
        print(row[0], ":", row[1])

    # Check indices on nfe
    cursor.execute("""
        SELECT indexname, indexdef 
        FROM pg_indexes 
        WHERE tablename = 'nfe';
    """)
    print("\n--- Indices in nfe ---")
    for row in cursor.fetchall():
        print(row[0], ":", row[1])

    # Check indices on nfe_xml
    cursor.execute("""
        SELECT indexname, indexdef 
        FROM pg_indexes 
        WHERE tablename = 'nfe_xml';
    """)
    print("\n--- Indices in nfe_xml ---")
    for row in cursor.fetchall():
        print(row[0], ":", row[1])

    # Let's EXPLAIN ANALYZE the open invoices query
    explain_sql = """
        EXPLAIN ANALYZE
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
        SELECT * FROM titulos t WHERE vencto >= '2026-04-10' AND vencto <= '2026-04-10'
    """
    print("\n--- EXPLAIN ANALYZE list_open_invoices ---")
    try:
        cursor.execute(explain_sql)
        for row in cursor.fetchall():
            print(row[0])
    except Exception as e:
        print(e)
        
    cursor.close()
    conn.close()

if __name__ == '__main__':
    main()