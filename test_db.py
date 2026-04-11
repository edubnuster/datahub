import json
import psycopg2

def main():
    print("Starting")
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
    print("Connected")

    # Let's EXPLAIN ANALYZE the open invoices query
    print("Analyzing tables...")
    cursor.execute("ANALYZE nota_fiscal;")
    cursor.execute("ANALYZE nfe;")
    cursor.execute("ANALYZE nfe_xml;")
    cursor.execute("ANALYZE anexo;")
    cursor.execute("ANALYZE movto_map;")
    print("Analyze complete.")
    
    # Run the get_nfe_attachments_bulk query again
    explain_sql = """
        EXPLAIN ANALYZE
        select
            inv.grid as invoice_id,
            nx.grid as nfe_xml_id,
            length(nx.fonte_xml) as fonte_xml_len
        from movto inv
        join nota_fiscal nf on nf.mlid = inv.mlid
        join nfe n on n.nota_fiscal = nf.grid
        join nfe_xml nx on nx.nfe = n.grid
        where inv.grid = any(array[17210952210]::bigint[])
          and inv.mlid is not null
          and inv.mlid <> 0
    """
    print("\n--- EXPLAIN ANALYZE get_nfe_attachments_bulk (FIXED) ---")
    cursor.execute(explain_sql)
    for row in cursor.fetchall():
        print(row[0])
        
    explain_sql_2 = """
        EXPLAIN ANALYZE
        select
            a.movto as movto_id,
            a.grid as grid,
            a.descricao,
            a.extensao,
            a.anexo,
            a.ts
        from public.anexo a
        where a.movto = any(array[17210952210]::bigint[])
          and lower(coalesce(a.descricao, '')) like '%assinatura%'
        union all
        select
            mm.child as movto_id,
            a.grid as grid,
            a.descricao,
            a.extensao,
            a.anexo,
            a.ts
        from (
            select parent, child from movto_map 
            where child = any(array[17210952210]::bigint[]) 
              and parent is not null and parent <> 0
        ) mm
        join public.anexo a
          on a.movto = mm.parent
        where lower(coalesce(a.descricao, '')) like '%assinatura%'
        order by movto_id, ts desc nulls last, grid desc
    """
    print("\n--- EXPLAIN ANALYZE get_sale_signatures_pdf_bulk (FIXED) ---")
    cursor.execute(explain_sql_2)
    for row in cursor.fetchall():
        print(row[0])

    
    # Indexes
    for table in ['anexo', 'movto_map']:
        print(f"\n--- Indices on {table} ---")
        cursor.execute(f"SELECT indexname FROM pg_indexes WHERE tablename = '{table}'")
        for row in cursor.fetchall():
            print(row[0])
            
    cursor.close()
    conn.close()

if __name__ == '__main__':
    main()
