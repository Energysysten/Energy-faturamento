#!/usr/bin/env python3
"""
Roda UMA VEZ para autorizar o envio de emails via Gmail.
Execute: python3 setup_envio_email.py
"""
from pathlib import Path
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
BASE   = Path(__file__).parent
CREDS  = BASE / "credentials.json"
TOKEN  = BASE / "token_envio.json"

if not CREDS.exists():
    print(f"✗ Arquivo não encontrado: {CREDS}")
    exit(1)

print("Abrindo navegador para autorizar envio de emails...")
flow  = InstalledAppFlow.from_client_secrets_file(str(CREDS), SCOPES)
creds = flow.run_local_server(port=8085)
TOKEN.write_text(creds.to_json())
print(f"\n✓ Autorizado com sucesso! Token salvo em:\n  {TOKEN}")
print("\nO sistema já pode enviar emails de reset de senha.")
