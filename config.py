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
DATE_FIN_PERIODE   = pd.Timestamp("2026-03-31")
SEUIL_PREVOYANCE   = 91

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
# dans l'ordre de sortie voulu.
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
]

# Colonnes de mapping PPU — ajoutées vides à cette étape ; seront remplies plus tard
# via la jointure rubrique <-> Ppu Rubrique Code du référentiel PPU.
PAIE_COLONNES_PPU = [
    "Ppu Categorie1", "Ppu Categorie 2", "Ppu Sous Categorie",
    "Ppu Affectation1", "Selection element PPU",
]

# Colonnes thématiques utiles à la reconstruction des IJSS — ajoutées vides à cette étape ;
# seront affectées selon l'affectation PPU de la rubrique (étape suivante).
PAIE_COLONNES_THEMATIQUES = [
    "Salaire de référence", "IJSS (assiette Brut)", "IJSS (nettes)",
    "IJSS subrogées", "Maintien", "Retenue pour absence",
    "IJ prévoyance (assiette Brut)",
]
