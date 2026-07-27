# Gabarits de présentation PDF

Ce dossier contient les gabarits de présentation officiels pour l'exportation des PDF (bilans et synthèses brochures).

## Concept & Cascade de résolution

Un gabarit est une surcharge de présentation (mise en page, sections actives/ordre, mode `brochure_custom`), distincte du profil de bilan qui définit les données analysées.

La cascade de résolution de présentation s'applique dans cet ordre :
1. `defaults` (définis dans `config/presentation/pdf_presentation.yaml`)
2. `scopes` (`global` ou `thematique`)
3. `profiles` (identifiant du profil de bilan)
4. **`gabarit`** (overlay final si spécifié)

## Structure d'un fichier gabarit YAML

Un fichier gabarit doit comporter les clés de métadonnées obligatoires suivantes :

```yaml
version: 1
gabarit_id: srp_r27
label: "R27 - Service Régional Police (BFC)"
description: "Gabarit de présentation brochure condensée pour le SRP BFC"

# Portée d'application
cible: bilan # bilan | brochure | les_deux
organisation:
  code_region: r27
  service: srp

profils_compatibles: # Optionnel (si absent : s'applique à tous les profils)
  - global
  - chasse
  - agrainage

# Mode de mise en page spécial (optionnel)
layout: brochure_custom # standard | brochure_custom

# Surcharges de présentation (syntaxe identique à pdf_presentation.yaml)
title:
  line2_mode: fixed
  line2_fixed: "Service Régional Police"

sections:
  order:
    - sec1
    - sec2
    - sec4
    - sec3
    - sec6
```
