#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

echo "================================"
echo "  Generer des cartes - Periode libre"
echo "================================"
echo ""
read -rp "Date de debut (YYYY-MM-DD) > " DATE_DEB
read -rp "Date de fin   (YYYY-MM-DD) > " DATE_FIN
read -rp "Code departement [21] > " DEPT
DEPT="${DEPT:-21}"

echo ""
echo "1 = agrainage  2 = chasse  3 = piegeage  4 = types usagers  5 = procedures PVe  6 = toutes"
read -rp "Quelle(s) carte(s) (1 a 6) [6] > " CHOIX
CHOIX="${CHOIX:-6}"

case "$CHOIX" in
    1) MAP=agrainage ;;
    2) MAP=chasse ;;
    3) MAP=piegeage ;;
    4) MAP=global_usagers ;;
    5) MAP=procedures_pve ;;
    6) MAP=tous ;;
    *) MAP=tous ;;
esac

echo ""
echo "Periode : $DATE_DEB au $DATE_FIN - Departement $DEPT - Carte(s) : $MAP"
echo ""

python3 scripts/generateur_de_cartes/production_cartographique.py "$MAP" --date-deb "$DATE_DEB" --date-fin "$DATE_FIN" --dept-code "$DEPT"

echo ""
echo "===== Termine ====="
