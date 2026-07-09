"""
Remove vínculos duplicados: quando uma folha tem total rateado > valor_total,
mantém apenas os registros mais recentes até atingir o valor correto.
"""
import sqlite3, os

DB = os.environ.get("FAT_DB", "/data/faturamento.db")
conn = sqlite3.connect(DB)

# Encontra folhas com total rateado excedente
excessos = conn.execute("""
    SELECT mf.n_folha,
           fr.valor_total,
           SUM(mf.valor) as total_rateado,
           COUNT(*) as qtd_links
    FROM medicao_folhas mf
    JOIN folhas_recebidas fr ON fr.n_folha = mf.n_folha
    WHERE fr.valor_total > 0
    GROUP BY mf.n_folha
    HAVING SUM(mf.valor) > fr.valor_total + 0.01
    ORDER BY mf.n_folha
""").fetchall()

if not excessos:
    print("Nenhuma folha com rateio excedente encontrado.")
    conn.close()
    exit()

print(f"Folhas com rateio excedente: {len(excessos)}\n")
removidos = 0

for n_folha, valor_total, total_rateado, qtd in excessos:
    print(f"  Folha {n_folha}: valor_total={valor_total:.2f}, rateado={total_rateado:.2f}, links={qtd}")

    # Lista todos os vínculos desta folha, do mais antigo ao mais recente
    links = conn.execute("""
        SELECT mf.id, mf.medicao_id, mf.valor, mf.vinculado_em,
               m.obra, m.cod
        FROM medicao_folhas mf
        LEFT JOIN medicoes m ON m.id = mf.medicao_id
        WHERE mf.n_folha = ?
        ORDER BY mf.vinculado_em ASC
    """, (n_folha,)).fetchall()

    acumulado = 0.0
    for lk in links:
        lid, mid, lvalor, ldt, obra, cod = lk
        print(f"    [{ldt}] {obra or cod or mid}: R${lvalor:.2f}", end="")
        if acumulado + lvalor <= valor_total + 0.01:
            acumulado += lvalor
            print(" -> MANTÉM")
        else:
            conn.execute("DELETE FROM medicao_folhas WHERE id=?", (lid,))
            removidos += 1
            print(" -> REMOVE (duplicado)")

    # Recalcula medicao nas medicoes afetadas
    for lk in links:
        mid = lk[1]
        novo_total = conn.execute(
            "SELECT COALESCE(SUM(valor),0) FROM medicao_folhas WHERE medicao_id=?", (mid,)
        ).fetchone()[0]
        conn.execute("UPDATE medicoes SET medicao=? WHERE id=?", (novo_total, mid))

    print()

conn.commit()
conn.close()
print(f"OK: {removidos} vínculos duplicados removidos.")
