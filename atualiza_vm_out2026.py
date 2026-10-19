"""
Insere no banco do VM as folhas de registro recebidas por email entre 13/10 e 15/10/2026
(competência Set/2026) cujos número e valor constam no corpo do email.

- Só INSERE folhas que ainda não existem; folhas já cadastradas não são tocadas
  (preserva NF e status já vinculados).
- Faz backup do banco antes (backup_db.py), se estiver na mesma pasta.

Rodar no VM, na pasta do sistema:  python atualiza_vm_out2026.py
Para só conferir sem gravar:        python atualiza_vm_out2026.py --simular
"""
import os, sys, sqlite3, uuid
from pathlib import Path

BASE = Path(__file__).resolve().parent
DB_PATH = Path(os.environ.get("FAT_DB", BASE / "faturamento.db"))

CONSTR = "Energy Systen Construcoes E Manuten"
SERV = "Energy Systen Servicos Especializad"

FOLHAS = [
    # n_folha, n_contrato, valor_total, municipio/obra, fornecedor, data_recebimento, arquivo
    ("1012444372", "4600027637", 89719.24, "Porangatu", CONSTR, "2026-10-14", "1012444372 - SE PORANGATU - MED 13 CIVIL.pdf"),
    ("1012444418", "4600027637", 80687.73, "Porangatu", CONSTR, "2026-10-14", "1012444418 - SE PORANGATU - MED 13 ELET.pdf"),
    ("1012485721", "4600027637", 119090.90, "SE Milhã", CONSTR, "2026-10-14", "FOLHA 1012485721  SE  ITAIM MILHAO CIV 06 -14.10.2026.pdf"),
    ("1012485739", "4600027637", 103973.37, "SE Rio dos Bois", CONSTR, "2026-10-14", "FOLHA 1012485739 SE RIO DOS BOIS  ME  04 -14.10.2026.pdf"),
    ("1012485764", "4600027637", 106644.20, "SE Rio dos Bois", CONSTR, "2026-10-14", "FOLHA 1012485764 SE RIO DOS BOIS CIVIL 04 -14.10.2026.pdf"),
    ("1012485680", "4600027590", 12454.64, "SE Petrolina", SERV, "2026-10-14", "FOLHA 1012485680 SE PETROLINA  ME - 01- LV ENREGY 14.10.2026.pdf"),
    ("1012483903", "4600027637", 153884.21, "Porangatu", CONSTR, "2026-10-15", "1012483903 - SE PORANGATU BANCOS - MED. 01.pdf"),
]
PERIODO = "2026-09"  # competência = data de recebimento − 30 dias


def main():
    simular = "--simular" in sys.argv
    if not DB_PATH.exists():
        print(f"Banco não encontrado: {DB_PATH}")
        return 1

    if not simular and (BASE / "backup_db.py").exists():
        sys.path.insert(0, str(BASE))
        import backup_db
        if backup_db.main() == 1:
            print("Backup falhou; nada foi alterado.")
            return 1

    conn = sqlite3.connect(DB_PATH, timeout=30)
    existentes = {r[0] for r in conn.execute("SELECT n_folha FROM folhas_recebidas")}
    novas = [f for f in FOLHAS if f[0] not in existentes]
    for f in FOLHAS:
        print(f"  {'JA EXISTE' if f[0] in existentes else 'NOVA     '} {f[0]}  {f[1]}  R$ {f[2]:>12,.2f}  {f[3]}")

    if simular:
        print(f"\nSimulação: {len(novas)} seriam inseridas, {len(FOLHAS) - len(novas)} já existem.")
        return 0

    with conn:
        conn.executemany(
            "INSERT INTO folhas_recebidas(id,n_folha,n_contrato,periodo,municipio,fornecedor,"
            "valor_total,arquivo,data_recebimento,status,nf) VALUES(?,?,?,?,?,?,?,?,?,?,NULL)",
            [(str(uuid.uuid4()), f[0], f[1], PERIODO, f[3], f[4], f[2], f[6], f[5], "recebido")
             for f in novas],
        )
    conn.close()
    print(f"\nConcluído: {len(novas)} inseridas, {len(FOLHAS) - len(novas)} já existiam (não alteradas).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
