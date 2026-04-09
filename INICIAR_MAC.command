#!/bin/bash
# INICIAR_MAC.command — Pólizas de Provisión de Gastos
cd "$(dirname "$0")"

if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 no encontrado."
    read -p "Presiona Enter para cerrar..."
    exit 1
fi

echo "Verificando dependencias..."
python3 -m pip install flask openpyxl xlrd --quiet

echo ""
echo "═══════════════════════════════════════════"
echo "  Pólizas de Provisión de Gastos"
echo "  GBC Business Consulting"
echo "═══════════════════════════════════════════"
echo ""
echo "  Abre tu navegador en:"
echo "  → http://localhost:5051"
echo ""
echo "  Presiona Ctrl+C para detener"
echo ""

python3 app.py
