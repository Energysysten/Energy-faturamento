#!/bin/bash
# Script de execução automática — Faturamento Energy Systen
# Roda 2x ao dia nos dias 1-20 de cada mês via crontab do macOS

DIA=$(date +%-d)
LOG="/Users/leonardocarmo/Documents/Claude/Projects/Faturamento/processamento.log"

# Só executa nos primeiros 20 dias do mês
if [ "$DIA" -gt 20 ]; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Dia $DIA — fora da janela (1-20). Pulando." >> "$LOG"
    exit 0
fi

# Garante que usa o Python com os pacotes instalados
export PATH="/usr/local/bin:/usr/bin:/bin:/Library/Developer/CommandLineTools/usr/bin:$PATH"

cd "/Users/leonardocarmo/Documents/Claude/Projects/Faturamento"

# Tenta python3 do sistema, depois fallback para outros caminhos comuns
for PY in python3 /usr/bin/python3 /Library/Developer/CommandLineTools/usr/bin/python3; do
    if command -v "$PY" &>/dev/null; then
        "$PY" processar_medicoes.py >> "$LOG" 2>&1
        exit 0
    fi
done

echo "[$(date '+%Y-%m-%d %H:%M:%S')] ERRO: python3 não encontrado." >> "$LOG"
exit 1
