# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.
#
# Ce programme est distribué dans l'espoir qu'il sera utile, mais SANS AUCUNE GARANTIE ;
# sans même la garantie implicite de QUALITÉ MARCHANDE ou D'ADÉQUATION À UN USAGE PARTICULIER.
# Voir la Licence Publique Générale GNU pour plus de détails.
#
# CONDITIONS SUPPLÉMENTAIRES D'ATTRIBUTION (SECTION 7(b) DE LA GPL v3) :
# Conformément à la section 7(b) de la GNU GPL v3, vous devez expressément conserver
# intactes et lisibles toutes les mentions d'auteur, notices de copyright et la présente
# clause dans chaque fichier source ou interface utilisateur redistribué. Toute version modifiée
# doit clairement indiquer qu'elle a été altérée et ne doit en aucun cas supprimer le nom
# de l'auteur original (Aguirre MAURIN).

#
import pandas as pd
from core.engine.orchestrateur_profils import _filter_point_ctrl

def test_filter_type_action_list():
    # Création d'un jeu de données factice avec des types d'action réels et modifiés
    data = pd.DataFrame({
        "type_actio": [
            "Eau et Nature;Gestion qualitative de la ressource en eau;Pollutions diffuses;Assurer le respect des conditions d'emplois des produits phytopharmaceutiques afin de préserver la qualité de l'eau et des milieux aquatiques (SNC 2.5);Procédures judiciaires liées aux PPP et en lien avec l'action 2.5 de la SNC",
            "Eau et Nature;Gestion qualitative de la ressource en eau;Pollutions diffuses;Assurer le respect des conditions d'emplois des produits phytopharmaceutiques afin de préserver la qualité de l'eau et des milieux aquatiques (SNC 2.5);Autres actions de police judiciaire liée aux PPP",
            "Eau et Nature;Gestion qualitative de la ressource en eau;Lutter contre les pollutions urbaines;Préserver la qualité des milieux aquatiques et la santé grâce à des systèmes d'assainissement conformes (SNC 2.1);Autres actions sur les systèmes d'assainissement",
            "Tronqué pour test: SNC 2.5;Autres actions"
        ]
    })

    # Profil configuré avec type_action_list
    profile = {
        "filter": {
            "type": "type_action_list",
            "type_actions": [
                "Procédures judiciaires liées aux PPP",
                "Autres actions de police judiciaire liée aux PPP",
                "SNC 2.5"
            ],
            "columns": ["type_actio"]
        }
    }

    # Application du filtre
    filtered = _filter_point_ctrl(data, profile)

    # Vérifications
    assert len(filtered) == 3
    # L'action 2.1 ne doit pas être conservée
    assert "SNC 2.1" not in "".join(filtered["type_actio"].values)
    # L'élément tronqué contenant "SNC 2.5" doit être conservé
    assert any("Tronqué pour test" in val for val in filtered["type_actio"].values)