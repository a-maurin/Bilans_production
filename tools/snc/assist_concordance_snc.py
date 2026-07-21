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
import csv
import os
from difflib import SequenceMatcher
from pathlib import Path

from complete_concordance_snc import normaliser_texte


BASE = Path(__file__).resolve().parents[2]

REF_PATH = (
    BASE
    / "ref"
    / "hors_programme"
    / "tables_reference"
    / "Plans de contrôle OSCEAN 2024.csv"
)
CONC_PATH = (
    BASE
    / "ref"
    / "hors_programme"
    / "tables_reference"
    / "concordance_snc_2023_vers_2024.csv"
)


def charger_referentiel_2024(path: str) -> list[dict]:
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f, delimiter=";")
        return list(reader)


def charger_concordance(path: str) -> tuple[list[dict], list[str]]:
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f, delimiter=";")
        rows = list(reader)
        fieldnames = reader.fieldnames or []
    return rows, fieldnames


def sauvegarder_concordance(path: str, rows: list[dict], fieldnames: list[str]) -> None:
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


def construire_index_anciens(rows: list[dict]) -> list[dict]:
    """Prépare la liste des anciens triplets avec versions normalisées."""
    result: list[dict] = []
    for row in rows:
        domaine = (row.get("domaine_ancien") or "").strip()
        theme = (row.get("theme_ancien") or "").strip()
        ta = (row.get("type_action_ancien") or "").strip()
        if not ta:
            continue
        result.append(
            {
                "domaine_ancien": domaine,
                "theme_ancien": theme,
                "type_action_ancien": ta,
                "type_action_ancien_norm": normaliser_texte(ta),
                "domaine_theme_ancien_norm": normaliser_texte(f"{domaine} {theme}"),
            }
        )
    return result


def deja_mappe(rows: list[dict], d: str, t: str, a: str) -> bool:
    for row in rows:
        if (
            (row.get("domaine_2024") or "").strip() == d
            and (row.get("theme_2024") or "").strip() == t
            and (row.get("type_action_2024") or "").strip() == a
        ):
            return True
    return False


def proposer_candidats(
    type_action_nouveau: str,
    anciens: list[dict],
    domaine_nouveau: str,
    theme_nouveau: str,
    max_candidats: int = 10,
) -> list[tuple[float, dict]]:
    """Retourne les meilleurs candidats anciens pour un type d'action nouveau.

    On filtre d'abord par recoupement de mots-clés sur le type d'action et le couple
    domaine/thème pour éviter les propositions hors sujet (ex : chasse vs eau).
    """
    norm_new = normaliser_texte(type_action_nouveau)
    scores: list[tuple[float, dict]] = []

    ctx_new_norm = normaliser_texte(f"{domaine_nouveau} {theme_nouveau}")
    ctx_new_tokens = set(ctx_new_norm.split()) if ctx_new_norm else set()
    ta_new_tokens = set(norm_new.split()) if norm_new else set()

    for r in anciens:
        norm_old = r["type_action_ancien_norm"]
        if not norm_old:
            continue

        ta_old_tokens = set(norm_old.split())
        if ta_new_tokens and ta_old_tokens and not (ta_new_tokens & ta_old_tokens):
            # Aucun mot-clé commun dans le libellé d'action -> pas pertinent
            continue

        ctx_old_tokens = set(
            (r.get("domaine_theme_ancien_norm") or "").split()
        )
        if ctx_new_tokens and ctx_old_tokens and not (ctx_new_tokens & ctx_old_tokens):
            # Aucun mot-clé commun sur domaine/thème -> probablement hors sujet
            continue

        score = SequenceMatcher(None, norm_new, norm_old).ratio()
        scores.append((score, r))

    scores.sort(key=lambda x: x[0], reverse=True)
    return scores[:max_candidats]


def boucle_assistance() -> None:
    ref_rows = charger_referentiel_2024(REF_PATH)
    conc_rows, fieldnames = charger_concordance(CONC_PATH)
    anciens_index = construire_index_anciens(conc_rows)

    print(f"{len(ref_rows)} lignes dans la nouvelle nomenclature.")
    print(f"{len(anciens_index)} triplets anciens disponibles pour mise en correspondance.")
    print(
        "Pour chaque type d'action 2024, les meilleurs candidats anciens seront proposés.\n"
        "Saisir le numéro choisi, 's' pour sauter, 'q' pour quitter.\n"
    )

    for i, ref in enumerate(ref_rows, start=1):
        domaine_new = (ref.get("domaine") or "").strip()
        theme_new = (ref.get("thème") or "").strip()
        type_action_new = (ref.get("type d'action") or "").strip()

        if not type_action_new:
            continue

        if deja_mappe(conc_rows, domaine_new, theme_new, type_action_new):
            continue

        print("=" * 80)
        print(
            f"[{i}/{len(ref_rows)}] Nouveau :\n"
            f"  Domaine : {domaine_new}\n"
            f"  Thème   : {theme_new}\n"
            f"  Action  : {type_action_new}\n"
        )

        candidats = proposer_candidats(
            type_action_new,
            anciens_index,
            domaine_nouveau=domaine_new,
            theme_nouveau=theme_new,
        )
        if not candidats:
            print("  Aucun candidat ancien trouvé.")
            continue

        print("Candidats anciens possibles :")
        for idx, (score, cand) in enumerate(candidats, start=1):
            print(
                f"  {idx}) score={score:.2f} | "
                f"{cand['domaine_ancien']} / {cand['theme_ancien']} / "
                f"{cand['type_action_ancien']}"
            )

        choix = input("Votre choix (numéro, 's' pour sauter, 'q' pour quitter) : ").strip()
        if choix.lower() == "q":
            break
        if choix.lower() == "s" or not choix:
            continue

        try:
            num = int(choix)
        except ValueError:
            print("Entrée non valide, ligne sautée.")
            continue

        if not (1 <= num <= len(candidats)):
            print("Numéro hors limites, ligne sautée.")
            continue

        _, selection = candidats[num - 1]

        # Mettre à jour la première ligne de concordance correspondant au triplet ancien choisi
        for row in conc_rows:
            if (
                (row.get("domaine_ancien") or "").strip()
                == selection["domaine_ancien"]
                and (row.get("theme_ancien") or "").strip()
                == selection["theme_ancien"]
                and (row.get("type_action_ancien") or "").strip()
                == selection["type_action_ancien"]
            ):
                row["domaine_2024"] = domaine_new
                row["theme_2024"] = theme_new
                row["type_action_2024"] = type_action_new
                print("  -> Concordance mise à jour.")
                break

        # Sauvegarde après chaque décision pour ne rien perdre
        sauvegarder_concordance(CONC_PATH, conc_rows, fieldnames)

    print("Fin de l'assistance au remplissage.")


if __name__ == "__main__":
    boucle_assistance()
