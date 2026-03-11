#!/bin/bash

# Script de paramétrage des cartes pour Linux
# Équivalent à parametrer_cartes.bat sous Windows

# Chemin vers le répertoire du projet
PROJECT_DIR="/media/e357/4D6A54AF6C0849D8/_Activité police/_BASES DE DONNEES/Bilans_production"

# Changer de répertoire
cd "$PROJECT_DIR" || exit 1

# Vérifier si le script de paramétrage des cartes existe
if [ ! -f "scripts/generateur_de_cartes/gui_config_cartes.py" ]; then
    echo "Erreur: Le script gui_config_cartes.py n'existe pas."
    exit 1
fi

# Fonction pour demander une date valide avec une valeur par défaut
ask_date() {
    local prompt="$1"
    local default="$2"
    local date
    while true; do
        read -p "$prompt (format: YYYY-MM-DD, défaut: $default): " date
        if [ -z "$date" ]; then
            echo "$default"
            return
        fi
        if [[ $date =~ ^[0-9]{4}-[0-9]{2}-[0-9]{2}$ ]]; then
            echo "$date"
            return
        else
            echo "Format de date invalide. Veuillez utiliser le format YYYY-MM-DD."
        fi
    done
}

# Demander les dates avec des valeurs par défaut
default_date_deb="2025-01-01"
default_date_fin="2025-12-31"
date_deb=$(ask_date "Date de début" "$default_date_deb")
date_fin=$(ask_date "Date de fin" "$default_date_fin")

# Demander le code du département avec une valeur par défaut
default_dept_code="21"
read -p "Code du département (ex. 21, défaut: $default_dept_code): " dept_code
dept_code=${dept_code:-$default_dept_code}

# Lancer le script Python
echo "Lancement du script de paramétrage des cartes..."
python3 scripts/generateur_de_cartes/gui_config_cartes.py

# Vérifier le code de retour
exit_code=$?
if [ $exit_code -ne 0 ]; then
    echo "Erreur lors de l'exécution du script de paramétrage des cartes. Code de retour: $exit_code"
    exit $exit_code
fi

echo "Script de paramétrage des cartes terminé avec succès."
exit 0