import json, urllib.request

API_URL = "http://localhost:8080"
API_KEY = "AUhb-dxNfR_EX54EmMR_8UibP8Ttt8Ujl6Wi7kPxL2Q"

FOLHAS = [{"n_folha": "1012486604", "n_contrato": "4600027637", "valor_total": 337203.08, "municipio": "Goiania", "fornecedor": "Energysysten", "periodo": "2026-09", "data_recebimento": "2026-10-19", "arquivo": "FR - TRM - ENERGY - CIV - BM10.pdf", "status": "Processado"}, {"n_folha": "1012486619", "n_contrato": "4600027637", "valor_total": 312590.77, "municipio": "Goiania", "fornecedor": "Energysysten", "periodo": "2026-09", "data_recebimento": "2026-10-19", "arquivo": "FR - TRM - ENERGY - EM - BM10.pdf", "status": "Processado"}]

enviadas = 0
erros = 0
for folha in FOLHAS:
    payload = json.dumps([folha]).encode("utf-8")
    try:
        req = urllib.request.Request(
            API_URL + "/api/folhas/sync",
            data=payload,
            headers={"Content-Type": "application/json", "X-API-Key": API_KEY},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            json.loads(resp.read())
            enviadas += 1
    except Exception as e:
        erros += 1
        print("  ERRO folha " + str(folha["n_folha"]) + ": " + str(e))

print("Concluido: " + str(enviadas) + " enviadas, " + str(erros) + " erros, " + str(len(FOLHAS)) + " total.")
input("Pressione Enter para fechar...")