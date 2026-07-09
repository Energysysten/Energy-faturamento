import sqlite3, os
DB = os.environ.get("FAT_DB", "/data/faturamento.db")
conn = sqlite3.connect(DB)
r = conn.execute("""UPDATE folhas_recebidas SET n_contrato='4600027590',municipio='FORMOSA',
    periodo='2026-05',data_recebimento='2026-06-29',status='Processado',
    fornecedor='ENERGY SYSTEN SERVICOS ESPECIALIZAD' WHERE n_folha='1011850417'""")
conn.commit()
conn.close()
print(f"OK: {r.rowcount} folha corrigida")
