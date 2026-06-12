#!/bin/bash
# Uso: ./scripts/deploy.sh "Descripción del cambio"
# Sube a Google Apps Script (clasp push) y a GitHub main en un solo paso.

MSG="${1:-deploy $(date '+%Y-%m-%d %H:%M')}"

echo "→ Subiendo a Google Apps Script..."
npx clasp push || { echo "✗ Error en clasp push"; exit 1; }

echo "→ Commiteando y subiendo a GitHub..."
git add -A
git commit -m "$MSG" || { echo "✗ Nada que commitear"; exit 0; }
git push origin main || { echo "✗ Error en git push"; exit 1; }

echo "✓ Deploy completo: GAS + GitHub main"
