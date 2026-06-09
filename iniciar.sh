#!/bin/bash
cd /Users/leonardocarmo/Documents/Claude/Financeiro

echo "=========================================="
echo "  ⚡ Energy — Dashboard de Faturamento"
echo "=========================================="
echo ""

# Matar processos anteriores
lsof -ti:8080 | xargs kill -9 2>/dev/null
sleep 1

# Iniciar servidor em background (persiste mesmo com terminal fechado)
nohup /Library/Developer/CommandLineTools/Library/Frameworks/Python3.framework/Versions/3.9/bin/python3 app.py > /tmp/energy-app.log 2>&1 &
FLASK_PID=$!
echo $FLASK_PID > /tmp/energy-app.pid
sleep 3

echo "✅ Servidor iniciado (PID: $FLASK_PID)"
echo ""
echo "  🏠 Local:    http://localhost:8080"
echo "  🏠 Rede:     http://$(ipconfig getifaddr en0 2>/dev/null || echo 'SEU-IP'):8080"
echo ""
echo "  Login: energy / energy2026"
echo "  Admin: admin / admin123"
echo ""

# Iniciar túnel Cloudflare em background
~/cloudflared tunnel --url http://localhost:8080 >> /tmp/cloudflared.log 2>&1 &
echo $! > /tmp/cloudflared.pid

echo "🌐 Aguardando link externo (Cloudflare)..."
for i in {1..15}; do
    LINK=$(grep -o 'https://[a-z0-9-]*\.trycloudflare\.com' /tmp/cloudflared.log 2>/dev/null | tail -1)
    if [ -n "$LINK" ]; then
        echo "  🌍 Externo:  $LINK"
        break
    fi
    sleep 1
done

echo ""
echo "=========================================="
echo "  Para encerrar: bash parar.sh"
echo "=========================================="
