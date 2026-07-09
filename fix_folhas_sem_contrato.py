import sqlite3, os

DB = os.environ.get("FAT_DB", "/data/faturamento.db")
FIXES = [
    ("1011828415","4600027590","GOIÂNIA",61008.30),
    ("1011828417","4600027590","APARECIDA DE GOIÂNIA",55589.29),
    ("1011828462","4600027590","CRISTALINA",4906.18),
    ("1011828463","4600027590","INHUMAS",18035.98),
    ("1011828466","4600027590","ÁGUAS LINDAS DE GOIÁS",18035.98),
    ("1011828472","4600027590","LUZIÂNIA",5591.49),
    ("1011828476","4600027590","GOIÂNIA",9064.81),
    ("1011828478","4600027590","GOIÂNIA",9879.11),
    ("1011828481","4600027590","GOIÂNIA",9879.11),
    ("1011828572","4600027590","LUZIÂNIA",8161.33),
]
conn = sqlite3.connect(DB)
n = 0
for folha, contrato, municipio, valor in FIXES:
    r = conn.execute("""UPDATE folhas_recebidas SET n_contrato=?,municipio=?,valor_total=?,
        periodo='2026-05',data_recebimento='2026-06-29',status='Processado',
        fornecedor='ENERGY SYSTEN SERVICOS ESPECIALIZAD' WHERE n_folha=?""",
        (contrato, municipio, valor, folha))
    n += r.rowcount
conn.commit()
conn.close()
print(f"OK: {n} folhas corrigidas")
