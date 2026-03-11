#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

echo "================================"
echo "  Bilans - Global ou thematiques"
echo "================================"
echo ""
read -rp "Date de debut (YYYY-MM-DD) > " DATE_DEB
read -rp "Date de fin   (YYYY-MM-DD) > " DATE_FIN
read -rp "Code departement [21] > " DEPT
DEPT="${DEPT:-21}"

echo ""
echo "Periode : $DATE_DEB au $DATE_FIN - Departement $DEPT"
echo ""
echo "1. Bilan global uniquement"
echo "2. Bilan(s) thematique(s) - agrainage, chasse, piegeage, etc."
echo ""
read -rp "Choix (1 ou 2) > " CHOIX

if [ "$CHOIX" = "1" ]; then
    echo "=== Bilan global ==="
    python3 scripts/run_bilan.py --mode global --date-deb "$DATE_DEB" --date-fin "$DATE_FIN" --dept-code "$DEPT"
elif [ "$CHOIX" = "2" ]; then
    echo ""
    echo "Profils disponibles :"
    python3 scripts/run_bilan.py --list-themes
    echo ""
    read -rp "Profil(s) a lancer (numero(s) ou id, separes par des espaces) [1 2] > " PROFILS
    PROFILS="${PROFILS:-1 2}"
    ARGS=""
    for p in $PROFILS; do
        ARGS="$ARGS --profil $p"
    done
    echo "=== Bilans thematiques ==="
    python3 scripts/run_bilan.py --mode thematique $ARGS --date-deb "$DATE_DEB" --date-fin "$DATE_FIN" --dept-code "$DEPT"
else
    echo "Choix invalide."
    exit 1
fi

echo ""
echo "===== Termine ====="
