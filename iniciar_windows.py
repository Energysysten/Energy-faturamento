"""
Script de inicialização do sistema de Faturamento para Windows.
Execute com: python iniciar_windows.py
"""
import os
import sys
from pathlib import Path

# ── Pasta onde este script está ──────────────────────────────────────────────
BASE = Path(__file__).parent

# ── Configurar variáveis de ambiente (ajuste os caminhos se necessário) ───────
os.environ.setdefault("FAT_BASE",  str(BASE))
os.environ.setdefault("FAT_TMPL",  str(BASE / "templates"))
os.environ.setdefault("FAT_DB",    str(BASE / "faturamento.db"))
os.environ.setdefault("FAT_CTRL",  str(BASE / "Controle_Medicoes.xlsx"))
os.environ.setdefault("FAT_CREDS", str(BASE / "credentials.json"))
os.environ.setdefault("FAT_TOKEN", str(BASE / "token_envio.json"))

HOST = "0.0.0.0"   # aceita conexões de qualquer IP da rede
PORT = 5000

# ── Iniciar com Waitress (servidor de produção para Windows) ──────────────────
try:
    from waitress import serve
    from app import app

    print(f"\n{'='*55}")
    print(f"  Sistema de Faturamento - Energy")
    print(f"  Servidor rodando em http://{HOST}:{PORT}")
    print(f"  Acesse pelo navegador: http://SEU_IP_DO_SERVIDOR:{PORT}")
    print(f"  Para parar: pressione Ctrl+C")
    print(f"{'='*55}\n")

    serve(app, host=HOST, port=PORT, threads=8)

except ImportError:
    # Fallback: Flask dev server (não recomendado para produção)
    from app import app
    print("\n[AVISO] waitress não encontrado. Usando servidor de desenvolvimento.")
    print(f"Acesse: http://SEU_IP:{PORT}\n")
    app.run(host=HOST, port=PORT, debug=False)
