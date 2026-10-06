#!/bin/bash
# Double-cliquez ce fichier pour démarrer la console BELVEDERE.
cd "$(dirname "$0")" || exit 1

echo "============================================================"
echo "   RÉSIDENCE BELVÉDÈRE — Console clients et lettres"
echo "============================================================"
echo

PY=""
command -v python3 >/dev/null 2>&1 && PY="python3"
[ -z "$PY" ] && command -v python >/dev/null 2>&1 && PY="python"

if [ -z "$PY" ]; then
  echo "[!] Python 3 n'est pas installé."
  echo "    Installez-le depuis https://www.python.org/downloads/ puis relancez."
  open "https://www.python.org/downloads/" 2>/dev/null
  read -r -p "Appuyez sur Entrée pour quitter."
  exit 1
fi

echo "[1/3] Python détecté."
echo "[2/3] Vérification des bibliothèques (Word, Excel)..."
if ! $PY -c "import docx, openpyxl" >/dev/null 2>&1; then
  echo "      Installation (une seule fois)..."
  if [ -d vendor ]; then
    $PY -m pip install --no-index --find-links vendor python-docx openpyxl >/dev/null 2>&1
  else
    $PY -m pip install --quiet python-docx openpyxl >/dev/null 2>&1
  fi
fi
echo "[3/3] Ouverture de l'interface..."
echo
echo "    Laissez cette fenêtre ouverte. Fermez-la pour quitter."
echo

# Adresse stable via Herd si le proxy "belvedere" existe, sinon adresse locale.
export PORT="${PORT:-8700}"
HERD_BIN="$HOME/Library/Application Support/Herd/bin/herd"
if [ -x "$HERD_BIN" ] && "$HERD_BIN" proxies 2>/dev/null | grep -qi "belvedere"; then
  export PUBLIC_URL="https://belvedere.test"
fi

$PY app.py
