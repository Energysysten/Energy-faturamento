#!/bin/bash
echo "Encerrando Energy — Faturamento..."

# Parar Flask
if [ -f /tmp/energy-app.pid ]; then
    kill $(cat /tmp/energy-app.pid) 2>/dev/null
    rm /tmp/energy-app.pid
fi
lsof -ti:8080 | xargs kill -9 2>/dev/null

# Parar Cloudflare
if [ -f /tmp/cloudflared.pid ]; then
    kill $(cat /tmp/cloudflared.pid) 2>/dev/null
    rm /tmp/cloudflared.pid
fi

echo "✅ Encerrado."
