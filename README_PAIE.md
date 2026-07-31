# Reconstruction IJSS — Partie PAIE

Reconstruit les IJSS côté **paie** (à partir de `PAIE_AUDIT.csv`), les compare au pipeline DSN déjà en place, et produit les 3 livrables demandés en réunion (Solange,
2026-07-31) : fichier "jours d'absence" côté paye, fichier des salariés réconciliés,
rapprochement des écarts. Ce document complète `README.md` (qui documente la partie DSN)
sans le remplacer.

## Sommaire

- [Vue d'ensemble](#vue-densemble)
- [Arborescence](#arborescence)
- [Données nécessaires](#données-nécessaires)
- [Ordre d'exécution et commandes](#ordre-dexécution-et-commandes)
- [Fichiers produits](#fichiers-produits)
- [Logique métier — points clés](#logique-métier--points-clés)
- [Limites connues / points de vigilance](#limites-connues--points-de-vigilance)
- [Livrables de la réunion du 2026-07-31](#livrables-de-la-réunion-du-2026-07-31)

## Vue d'ensemble

Trois pipelines s'enchaînent, chacun avec son point d'entrée (`python <fichier>.py`) :

```
main_paie.py          reconstruit la paie (grain rubrique) + jointure PPU
        │
        ├──► main_comparaison.py   compare les POPULATIONS de salariés PAIE ↔ DSN
        │
        └──► main_absences.py      compare les JOURS D'ABSENCE PAIE ↔ DSN
```

`main_comparaison.py` et `main_absences.py` ont chacun besoin des sorties de `main.py`
(DSN) **et** de `main_paie.py` (PAIE) — ce sont les deux reconstruits qui servent de base
à toute comparaison, jamais les fichiers source directement (cf. [Logique métier](#logique-métier--points-clés)).

## Arborescence

```
protection-social-v2/
├── config.py                    # paramètres centralisés (DSN + PAIE, un seul fichier)
├── main.py                      # (existant) reconstruction DSN
├── main_paie.py                 # reconstruction PAIE + jointure PPU
├── main_comparaison.py          # comparatif des populations de salariés PAIE ↔ DSN
├── main_absences.py             # comparatif des jours d'absence PAIE ↔ DSN + livrable 1
├── data/input/
│   ├── PAIE_AUDIT.csv                          # source paie (grain rubrique)
│   ├── Cartographie_ppu_v5_formules.xlsx       # référentiel PPU
│   └── CLTINL0008_..._v11-06-2026.csv          # source DSN (déjà utilisée par main.py)
├── output/xlsx/
│   ├── Reconstruit_PAIE_<code>.xlsx            # livrable final PAIE (1 par société)
│   └── Jours_Absence_PAIE_<code>.xlsx          # livrable 1 (miroir DSN, dérivé du reconstruit)
├── output/rapports/
│   ├── Comparatif_populations_PAIE_DSN.xlsx
│   └── Comparatif_jours_absence_PAIE_DSN.xlsx  # livrable 3
└── src/
    ├── chargement_paie.py       # lecture PAIE_AUDIT.csv + PPU, jointure, construction du reconstruit
    ├── export_paie.py           # écriture Reconstruit_PAIE_<code>.xlsx
    ├── comparaison_populations.py  # comparaison des salariés (4 niveaux, cf. plus bas)
    ├── export_comparatif.py     # écriture Comparatif_populations_PAIE_DSN.xlsx
    └── comparaison_absences.py  # identification des rubriques d'absence, comparaison,
                                  # écriture des 2 derniers livrables — tout est dans ce seul fichier
```

## Données nécessaires

À déposer dans `data/input/` avant de lancer quoi que ce soit :

| Fichier | Rôle |
|---|---|
| `PAIE_AUDIT.csv` | Source paie, grain rubrique (une ligne par rubrique/salarié/mois) |
| `Cartographie_ppu_v5_formules.xlsx` | Référentiel PPU (mapping rubrique → catégorie/thématique) |
| `CLTINL0008_ABS-detailed-SS..._v11-06-2026.csv` | Source DSN (déjà utilisée par `main.py`) |

## Ordre d'exécution et commandes

**L'ordre est impératif** — chaque étape lit les sorties de la précédente.

```bash
# 1. Reconstruction DSN (si pas déjà fait / à refaire si la source DSN a changé)
python main.py

# 2. Reconstruction PAIE (jointure PPU, colonnes Motif / Jours d'absence)
python main_paie.py

# 3. Comparatif des populations de salariés (a besoin de 1 et 2)
python main_comparaison.py

# 4. Comparatif des jours d'absence + livrable "miroir PAIE" (a besoin de 1 et 2)
python main_absences.py
```

Les étapes 3 et 4 sont indépendantes entre elles (aucune n'a besoin de l'autre), mais
toutes les deux ont besoin des étapes 1 et 2 déjà exécutées.

**Temps d'exécution approximatifs** (CSV paie ~1,3 Go) :
- `main.py` : quelques secondes
- `main_paie.py` : 3-5 min
- `main_comparaison.py` : ~2 min
- `main_absences.py` : ~2 min

**Un fichier de sortie ouvert dans Excel bloque son écriture** (`PermissionError`) —
fermer le fichier concerné puis relancer le script suffit, pas besoin de tout refaire.

## Fichiers produits

| Fichier | Produit par | Grain | Contenu |
|---|---|---|---|
| `output/xlsx/Reconstruit_PAIE_<code>.xlsx` (×4 : `52`, `ARF`, `GEF`, `SOG`) | `main_paie.py` | 1 ligne par rubrique | Colonnes PAIE_AUDIT + colonnes PPU (mapping/thématiques) + `Motif` + `Jours d'absence (retenu)` — **c'est le livrable final**, tout part de lui |
| `output/xlsx/Jours_Absence_PAIE_<code>.xlsx` (×4) | `main_absences.py` | 1 ligne par salarié/mois/motif | Livrable 1 — miroir du fichier DSN, lu depuis le reconstruit (pas recalculé) |
| `output/rapports/Comparatif_populations_PAIE_DSN.xlsx` | `main_comparaison.py` | 1 ligne par salarié | 4 niveaux (reconstruits / sources / période reconstr. / période sources), statut par salarié (présent des 2 côtés, DSN sans PAIE, PAIE sans arrêt DSN) |
| `output/rapports/Comparatif_jours_absence_PAIE_DSN.xlsx` | `main_absences.py` | 1 ligne par salarié/mois/motif | Livrable 3 — détail mensuel + total annuel + onglet anomalies, clé pivot, écart DSN-PAIE, filtres Excel |
| `output/logs/execution/execution_{paie,comparatif,absences}_<horodatage>.log` | chacun des 3 scripts | — | déroulé pas à pas + erreurs, jamais écrasés |

## Logique métier — points clés

**Jointure PAIE ↔ PPU** (`chargement_paie.py`) : sur la clé composite
`(rubrique, libelle)` ↔ `(Ppu Rubrique Code, Ppu Rubrique Libelle)` — le code seul n'est
pas unique dans le référentiel (ex. le code `1300` sert à des dizaines de choses
différentes selon le libellé). Toutes les colonnes du PPU sont reprises verbatim.

**Identification des rubriques "jours d'absence"** (`config.MOTIFS_PAIE_JOURS`,
utilisé par `comparaison_absences.identifier_motifs`) : la colonne `Base` de
`PAIE_AUDIT.csv` n'est **pas toujours un nombre de jours** — c'est un montant de
cotisation, un taux d'acquisition de CP ou un nombre d'épisodes selon la rubrique.
Seul un sous-ensemble validé (vérifié par la formule interne de la paie,
`Montant = Base × Taux`) est retenu comme "jours". Règles apprises :
- **`3350` compte pour "accident de travail"**, pas "maladie" — vérifié sur des cas
  réels où le motif DSN correspondant est "congé suite à accident de travail".
- **`3405` compte uniquement pour "paternité"**, pas "maternité" — sinon double-comptage.
- **`3680` est exclue** — elle doublonne `3350` pour la même absence.
- **`3300`/`3480` sont gardées ensemble** — elles se complètent (maintenue / non
  maintenue d'une même absence), pas un doublon.
- Accident de travail, accident de trajet et maladie professionnelle **ne sont pas
  distingués côté paie** — regroupés dans un seul panier "accident de travail" côté
  DSN aussi (`canoniser_motif_dsn`), pour comparer des choses comparables.

**Le reconstruit est la seule source de vérité** : `Reconstruit_PAIE_<code>.xlsx` porte
déjà `Motif`/`Jours d'absence (retenu)` — c'est lui qui sert de base à une éventuelle
récupération auprès de la Sécurité sociale. `main_absences.py` et
`main_comparaison.py` lisent ce fichier, ils **ne recalculent jamais** depuis
`PAIE_AUDIT.csv` (seule exception : la détection d'anomalies, volontairement sur la
source, car c'est un diagnostic de qualité de la donnée brute).

**Normalisation du matricule** : la DSN le stocke sur 6 chiffres avec zéros de tête
(`"005005"`), la paie sans (`"5005"`) — `normaliser_matricule()` retire les zéros de
tête pour pouvoir rapprocher les deux.

**Siren/Nic/Siret** : toujours à `0` dans `PAIE_AUDIT.csv` — repris depuis la source DSN
(`lire_identite_dsn_source`), qui a de vraies valeurs (1 SIREN par société, 1 NIC quasi
constant par salarié).

**Période couverte** (`config.DATE_DEBUT_PERIODE` / `DATE_FIN_PERIODE`) : actuellement
l'année 2026 complète. Le fichier source DSN contient réellement des déclarations
jusqu'en septembre 2026 (malgré un nom de fichier qui suggère "01 à 03") — vérifié
empiriquement, pas une supposition.

## Limites connues / points de vigilance

- **Onglet "Anomalies (jours)"** du comparatif absences : ~50 lignes où `Base` dépasse
  31 jours sur un mois ou est négative hors annulation — confirmé présentes dans le
  fichier source lui-même (pas un artefact de calcul), sans information exploitable
  (`num_bull`/`date_retro` vides) pour les expliquer. À vérifier avec l'éditeur de paie.
- **Cas GAUMET Lionel** (INLI/52) : en arrêt continu depuis 2025, quasi absent de la
  paie — seul écart de population résiduel identifié à ce jour.
- **Congés pathologiques prénataux** : parfois déclarés "maladie" côté DSN mais
  directement en rubrique "Absence Maternité" côté paie — écart résiduel documenté,
  pas un bug.
- **Temps partiel thérapeutique** : aucune rubrique paie dédiée retenue dans
  `MOTIFS_PAIE_JOURS` à ce jour — les jours DSN "temps partiel thérapeutique"
  n'ont donc pas d'équivalent PAIE dans le comparatif (colonne PAIE à 0, signalé
  explicitement, pas masqué).

## Livrables de la réunion du 2026-07-31

| # | Demandé | Fichier | Statut |
|---|---|---|---|
| 1 | Jours d'absence côté paye (miroir DSN) | `Jours_Absence_PAIE_<code>.xlsx` | ✅ Fait |
| 2 | Fiche d'identité salarié réconciliée (paye + DSN) | — | 🔴 Non fait — champs manquants dans les sources actuelles (adresse, salaire de base, dates d'entrée/sortie, date de naissance) ; disponibles dès aujourd'hui : SIREN, SIRET, Matricule, NIR, Nom/Prénom/Nom d'usage, Statut, Emploi |
| 3 | Rapprochement des écarts absence DSN vs paye (pivot salarié) | `Comparatif_jours_absence_PAIE_DSN.xlsx` | ✅ Fait |
