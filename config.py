# -*- coding: utf-8 -*-
"""Paramètres du projet de reconstruction IJSS (centralisés)."""
import os
import pandas as pd

# --- Arborescence ---
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR  = os.path.join(BASE_DIR, "data", "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output", "xlsx")
LOG_DIR    = os.path.join(BASE_DIR, "output", "logs")
LOG_EXEC_DIR = os.path.join(LOG_DIR, "execution")   # journaux d'exécution (déroulé + erreurs)
LOG_ANO_DIR  = os.path.join(LOG_DIR, "anomalies")   # journaux d'anomalies métier
RAPPORT_DIR = os.path.join(BASE_DIR, "output", "rapports")

# --- Fichier source DSN (déposer dans data/input/) ---
INPUT_FILE = "CLTINL0008_ABS-detailed-SS par mois calendaire_complete_2026 01 a 03_v11-06-2026.csv"

# --- Paramètres métier ---
# ===============================================================
# 0. PARAMÈTRES
# ===============================================================
DATE_DEBUT_PERIODE = pd.Timestamp("2026-01-01")
DATE_FIN_PERIODE   = pd.Timestamp("2026-12-25")

# Période de filtrage par défaut utilisée par les exécutions main*.py pour limiter
# les traitements aux trois premiers mois de l'année. Peut être surchargée si besoin.
PERIODE_FILTRE_DEBUT = DATE_DEBUT_PERIODE
PERIODE_FILTRE_FIN   = DATE_FIN_PERIODE
SEUIL_PREVOYANCE     = 91

# Sociétés à traiter — un fichier de sortie par société. AURA en tête (cas POUYET/GANDOUZ).
# Libellés normalisés (sans accent/apostrophe, majuscules). Mettre une seule entrée
# pour ne traiter qu'une société.
SOCIETES_CIBLES = ["AURA", "INLI", "INLI PM", "GRAND EST"]

# Motifs avec carence de 3 jours (choix CORRIGE12 : TPT exclu)
MOTIFS_CARENCE_3J = ["maladie"]

# Arbitrage maladie -> maternité, cas par cas (note chevauchement, règle 4)
# --- Priorité des motifs pour le départage au MÊME DJT (rang élevé = prime) ---
# maternité/paternité/adoption (congés légaux) > AT/MP/trajet (risques pro) > maladie > TPT
PRIORITE_MOTIF = {
    "maternité": 4, "paternité": 4, "adoption": 4,
    "accident de travail": 3, "accident du travail": 3,
    "maladie professionnelle": 3, "accident de trajet": 3,
    "temps partiel": 1,
    "maladie": 2,
}
RANG_MOTIF_DEFAUT = 2  # motif non listé : traité au niveau de la maladie

def rang_motif(motif):
    """Rang de priorité d'un motif (robuste à la casse / libellés partiels)."""
    m = str(motif).strip().lower()
    for cle, rang in PRIORITE_MOTIF.items():
        if cle in m:
            return rang
    return RANG_MOTIF_DEFAUT

ARBITRAGE_MATERNITE = {
    # "matricule": "AAAA-MM-JJ",
}

# ===============================================================
# PAIE — reconstruction IJSS "version paie" (étape 1 : structure uniquement)
# ===============================================================
# Source : data/input/PAIE_AUDIT.csv — grain rubrique (une ligne par rubrique de paie,
# un salarié a plusieurs lignes par mois). Découpage en sortie sur la colonne "entreprise".
PAIE_INPUT_FILE = "PAIE_AUDIT.csv"
PAIE_SEP = ";"
ANNEE_PAIE = 2026

# Colonnes d'origine PAIE_AUDIT reprises telles quelles : (colonne source -> libellé de sortie),
# dans l'ordre de sortie voulu. idclient/nom_prenom ajoutées en fin de bloc (pas de renommage,
# comme code_entite/num_bull) : on ne les exclut plus de la sortie.
PAIE_COLONNES_ORIGINE = [
    ("nir", "Nir"), ("matricule", "Matricule"), ("nom", "Nom"), ("prenom", "Prenom"),
    ("entreprise", "Entreprise"), ("etablissement", "Etablissement"), ("siren", "Siren"),
    ("nic", "Nic"), ("siret", "Siret"), ("code_entite", "code_entite"),
    ("periode", "Periode"), ("date_sous_periode", "Date_sous_periode"),
    ("num_sous_periode", "Num_sous_periode"), ("date_retro", "Date_retro"),
    ("num_contrat", "Num_contrat"), ("num_bull", "num_bull"),
    ("rubrique", "Rubrique"), ("libelle", "Libelle"), ("base", "Base"),
    ("base_salariale", "Base_salariale"), ("taux_salarial", "Taux_salarial"),
    ("montant_salarial", "Montant_salarial"), ("base_patronale", "Base_patronale"),
    ("taux_patronal", "Taux_patronal"), ("montant_patronal", "Montant_patronal"),
    ("idclient", "idclient"), ("nom_prenom", "nom_prenom"),
]

# ===============================================================
# PPU — référentiel de cartographie des rubriques (jointure avec PAIE_AUDIT)
# ===============================================================
# Source : data/input/Cartographie_ppu_v5_formules.xlsx, onglet "Cartographie".
# Jointure sur clé composite (rubrique, libelle) <-> (Ppu Rubrique Code, Ppu Rubrique Libelle) :
# le seul code de rubrique n'est pas unique dans le référentiel (ex. code 1300 réutilisé pour
# des dizaines de libellés différents avec des classifications différentes) ; le couple
# (code, libellé) l'est presque toujours (16 paires sur 618 restent ambiguës — on garde la
# première occurrence rencontrée dans le référentiel pour celles-ci).
# Toutes les colonnes du PPU sont conservées telles quelles (verbatim, ordre du fichier source) ;
# elles ne sont plus listées en dur ici pour rester synchronisées avec le référentiel.
PPU_INPUT_FILE = "Cartographie_ppu_v5_formules.xlsx"
PPU_SHEET = "Cartographie"
PPU_COL_CODE = "Ppu Rubrique Code"
PPU_COL_LIBELLE = "Ppu Rubrique Libelle"

# ===============================================================
# COMPARAISON PAIE <-> DSN — populations de salariés par société
# ===============================================================
# Correspondance établie empiriquement par recoupement des matricules (zéros de tête
# retirés côté DSN, cf. src/comparaison_populations.py) : {code société PAIE: société DSN}.
MAPPING_PAIE_DSN = {
    "52": "INLI",
    "ARF": "AURA",
    "GEF": "GRAND EST",
    "SOG": "INLI PM",
}
COMPARATIF_XLSX = "Comparatif_populations_PAIE_DSN.xlsx"

# ===============================================================
# MOTIFS D'ABSENCE PAIE — rubriques identifiées comme pertinentes pour le calcul des
# IJSS (absences prises en charge par la Sécurité Sociale), validées avec le métier.
# ===============================================================
# {motif: [(libellé, code rubrique), ...]} — liste et non dict, volontairement : un même
# libellé peut légitimement porter plusieurs codes (ex. "Accident du Travail" -> 4400/
# 4401/4420 selon le régime), et un même (rubrique, libellé) peut apparaître sous deux
# motifs à la fois (ex. "Reg Paternité"/1300 sous maternité ET paternité) — un dict
# écraserait ces doublons silencieusement. Clé d'identification côté PAIE_AUDIT :
# (rubrique, libelle), pas le code seul (réutilisé entre motifs, ex. 1300).
MOTIFS_PAIE = {
    "maladie": [
        ("Absence Maladie", 3300),
    ],
    "accident de travail": [
        ("Abs. Acc.Travail/Maladie Prof", 3350),
    ],
    "maternité": [
        ("Absence Maternite", 3390),
    ],
    "paternité": [
        ("Absence Paternite (Maintenue)", 3405),
    ],
}

# Sous-ensemble de MOTIFS_PAIE dont la colonne "Base" est un vrai nombre de jours
# d'absence (vérifié empiriquement sur PAIE_AUDIT 2026 — cf. échanges de validation) :
# la colonne "Base" est polymorphe, elle porte tantôt un nombre de jours, tantôt une
# assiette de cotisation (montant en €), tantôt un taux d'acquisition de CP ou un nombre
# d'épisodes, selon la rubrique. Seules les rubriques ci-dessous sont retenues pour le
# calcul des jours d'absence ; les rubriques de cotisation (SS Maladie*, Complement*,
# Accident du Travail RG/PIM/Apprenti, BS:SS*) et les cas ambigus (Maintien Maternité —
# valeurs > 31 sur certaines lignes, Maintien Paternité — toujours à 0, "Dont CP acquis
# sur maladie" — taux d'acquisition et non un nombre de jours, "Nombre d'arrêts
# Maladie-AT" — un nombre d'épisodes et non un nombre de jours) en sont exclues.
#
# "CP - Abs AT/Maladie Pro." / "CP-Abs AT/Maladie Pro. NUL" (3680) sont EXCLUES : elles
# doublonnent 3350 pour la même absence (ex. Janvier : 3350=31j + 3680=22j = 53j, plus
# que les 31 jours du mois — physiquement impossible si les deux étaient des jours
# indépendants). 3300/3480 en revanche sont conservées ensemble : leur somme colle au
# nombre de jours du mois (maintenue vs non-maintenue d'une même absence), ce n'est pas
# un doublon.
#
# 3350 ("Abs. Acc.Travail/Maladie Prof.") est classée ici sous "accident de travail",
# pas "maladie" : vérifié sur des cas réels (ex. matricules 8172/8504/20773/20122) où le
# motif DSN correspondant est "congé suite à accident de travail ou de service" (donc
# canonisé "accident de travail" côté DSN, cf. canoniser_motif_dsn) — la classer sous
# "maladie" créait un miroir artificiel : mêmes jours comptés "maladie" côté PAIE et
# "accident de travail" côté DSN pour la même absence, gonflant les deux écarts à la fois.
MOTIFS_PAIE_JOURS = {
    "maladie": [
        # ("Annulation Maladie", 3200),
        ("Absence Maladie", 3300),
        # ("Absence Maladie Covid", 3380),
        # ("Annul. Maladie", 3480),
        # ("Maladie non maintenue", 3480),
    ],
    "accident de travail": [
        # ("Annulation Accident travail", 3250),
        ("Abs. Acc.Travail/Maladie Prof.", 3350),
    ],
    "maternité": [
        ("Absence Maternite", 3390),
    ],
    "paternité": [
        ("Absence Paternite (Maintenue)", 3405),
        # ("Paternite (non maintenue)", 3415),
        # ("Paternité (non maintenue)", 3415),
    ],
}
