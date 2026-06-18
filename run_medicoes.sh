#!/bin/bash
# Script de execução automática - Energy Systen
# Agendado via LaunchAgent macOS - 15:00 diariamente

PROJECT_DIR="/Users/leonardocarmo/Documents/Claude/Projects/Faturamento"
LOG_FILE="$PROJECT_DIR/processamento.log"

echo "[$(date '+%Y-%m-%d %H:%M:%S')] Iniciando execução automática via LaunchAgent" >> "$LOG_FILE"

# Usa o python3 do sistema (ou Homebrew se disponível)
PYTHON=$(command -v python3 || command -v /opt/homebrew/bin/python3 || command -v /usr/local/bin/python3)

if [ -z "$PYTHON" ]; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] ERRO: python3 não encontrado" >> "$LOG_FILE"
    exit 1
fi

cd "$PROJECT_DIR" && "$PYTHON" processar_medicoes.py >> "$LOG_FILE" 2>&1
echo "[$(date '+%Y-%m-%d %H:%M:%S')] Execução automática finalizada" >> "$LOG_FILE"
