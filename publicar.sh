#!/bin/bash
# Publica as alterações no GitHub (Render faz o deploy automaticamente)
# Configure o token uma vez com: git config --global credential.helper store

echo "Publicando alterações..."
git add -A
git commit -m "Atualização $(date '+%d/%m/%Y %H:%M')" 2>/dev/null || echo "(nenhuma alteração nova)"
git push origin main

echo ""
echo "Pronto! O Render vai atualizar o site em ~2 minutos."
echo "Acesse: https://energy-faturamento.onrender.com"
