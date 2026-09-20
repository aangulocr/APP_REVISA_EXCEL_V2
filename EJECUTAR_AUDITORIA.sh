#!/usr/bin/env bash
# ==========================================================
# APP_REVISA_EXCEL_V2 — Lanzador CLI para Linux / macOS
# ==========================================================

set -e
cd "$(dirname "$0")"

echo "=========================================================="
echo "  APP_REVISA_EXCEL_V2 — AUDITORÍA POR RÚBRICA Y ESPEJO"
echo "=========================================================="
echo ""

PYTHON_BIN="python3"
if ! command -v $PYTHON_BIN &> /dev/null; then
    PYTHON_BIN="python"
fi

if ! command -v $PYTHON_BIN &> /dev/null; then
    echo "[ERROR] Python 3 no está instalado."
    exit 1
fi

VENV_DIR=".venv"
if [ ! -f "$VENV_DIR/bin/python" ]; then
    echo "[INFO] Creando entorno virtual en $VENV_DIR..."
    $PYTHON_BIN -m venv "$VENV_DIR"
fi

source "$VENV_DIR/bin/activate"

if [ -f "requirements.txt" ]; then
    echo "[INFO] Verificando dependencias..."
    pip install -q -r requirements.txt
fi

echo "[INFO] Iniciando auditoría CLI..."
python main.py "$@"
