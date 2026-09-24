"""
Sincroniza o Controle_Medicoes.xlsx com o banco local do VM (localhost:8080).
Rodar diretamente no VM: python resync_vm.py
"""
import openpyxl
import json
import urllib.request
import datetime
import os

API_URL = os.environ.get("RENDER_API_URL", "http://localhost:8080")
API_KEY = os.environ.get("SYNC_API_KEY", "AUhb-dxNfR_EX54EmMR_8UibP8Ttt8Ujl6Wi7kPxL2Q")
XLSX_PATH = os.environ.get("XLSX_PATH", r"C:\Faturamento\Controle_Medicoes.xlsx")

if not os.path.exists(XLSX_PATH):
    print(f"Arquivo nao encontrado: {XLSX_PATH}")
    print("Use: set XLSX_PATH=caminho\\para\\Controle_Medicoes.xlsx")
    exit(1)

print(f"Lendo: {XLSX_PATH}")
wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
ws = wb["Medições"]

headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
col = {h: i+1 for i, h in enumerate(headers) if h}
print("Colunas:", list(col.keys()))

enviadas = 0
erros = 0
total = 0

for row in range(2, ws.max_row + 1):
    n_folha = ws.cell(row, col.get("Nº Folha", 2)).value
    if not n_folha:
        continue
    total += 1

    data_rec = ws.cell(row, col.get("Data Recebimento", 1)).value
    n_contrato = ws.cell(row, col.get("Nº Contrato", 3)).value
    municipio = ws.cell(row, col.get("Município", 6)).value
    fornecedor = ws.cell(row, col.get("Fornecedor", 7)).value
    valor = ws.cell(row, col.get("Valor Total (R$)", 8)).value
    arquivo = ws.cell(row, col.get("Arquivo PDF", 9)).value
    status = ws.cell(row, col.get("Status", 10)).value or "recebido"

    # Formata data
    if isinstance(data_rec, datetime.datetime):
        data_iso = data_rec.strftime("%Y-%m-%d")
        d = data_rec.date()
    elif isinstance(data_rec, str) and "/" in data_rec:
        parts = data_rec.split("/")
        data_iso = f"{parts[2]}-{parts[1]}-{parts[0]}"
        d = datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))
    else:
        data_iso = str(data_rec or "")
        d = None

    # Período = data_recebimento - 30 dias
    if d:
        mes = d.month - 1 or 12
        ano = d.year if d.month > 1 else d.year - 1
        periodo = f"{ano}-{mes:02d}"
    else:
        periodo = ""

    payload = json.dumps([{
        "n_folha":          str(n_folha),
        "n_contrato":       str(n_contrato or ""),
        "valor_total":      float(valor or 0),
        "municipio":        str(municipio or ""),
        "fornecedor":       str(fornecedor or ""),
        "periodo":          periodo,
        "data_recebimento": data_iso,
        "arquivo":          str(arquivo or ""),
        "status":           status,
    }]).encode("utf-8")

    try:
        req = urllib.request.Request(
            f"{API_URL}/api/folhas/sync",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "X-API-Key": API_KEY,
            },
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            json.loads(resp.read())
            enviadas += 1
            if enviadas % 100 == 0:
                print(f"  {enviadas}/{total} enviadas...")
    except Exception as e:
        erros += 1
        print(f"  ERRO folha {n_folha}: {e}")

print(f"\nConcluido: {enviadas} enviadas, {erros} erros, {total} total.")
input("Pressione Enter para fechar...")
