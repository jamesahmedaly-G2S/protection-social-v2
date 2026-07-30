# -*- coding: utf-8 -*-
"""Identification des rubriques d'absence PAIE par motif IJSS, et agrégation des
jours d'absence (colonne "Base") par salarié / mois / société / motif.

Clé d'identification : (rubrique, libelle) — comme pour la jointure PPU, le code de
rubrique seul est réutilisé entre motifs différents (ex. 1300 : maladie, maternité,
paternité selon le libellé), donc le libellé est indispensable pour lever l'ambiguïté.
Un même (rubrique, libelle) peut légitimement apparaître sous plusieurs motifs à la fois
(ex. "Reg Paternité"/1300 sous maternité ET paternité dans config.MOTIFS_PAIE) : dans ce
cas la ligne compte pour chacun des motifs concernés (pas d'arbitrage arbitraire).
"""
import re
import unicodedata
import pandas as pd
from config import MOTIFS_PAIE
from src.normalisation import to_num

_PONCTUATION = re.compile(r"[-/.:',]")


def _normaliser_libelle(libelle):
    """Minuscules, sans accents, ponctuation remplacée par des espaces, espaces
    multiples réduits à un seul — tolère les variantes de transcription du référentiel
    métier (ex. "CP-Abs AT/Maladie Pro. NUL" et "CP - Abs AT/Maladie Pro.   NUL" ou
    "Maternite"/"Maternité" doivent être reconnus comme identiques)."""
    s = unicodedata.normalize("NFD", str(libelle))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = _PONCTUATION.sub(" ", s.lower())
    return " ".join(s.split())


def _cle(rubrique, libelle):
    return (str(rubrique).strip(), _normaliser_libelle(libelle))


def construire_lookup_motifs():
    """{(rubrique, libellé normalisé): {motifs}} à partir de config.MOTIFS_PAIE."""
    lookup = {}
    for motif, entrees in MOTIFS_PAIE.items():
        for libelle, code in entrees:
            lookup.setdefault(_cle(code, libelle), set()).add(motif)
    return lookup


def identifier_motifs(df):
    """Ajoute une colonne "_motifs" (set des motifs reconnus, vide si aucune
    correspondance) au DataFrame PAIE (colonnes source : rubrique, libelle)."""
    lookup = construire_lookup_motifs()
    libelles_norm = df["libelle"].apply(_normaliser_libelle)
    cles = list(zip(df["rubrique"].astype(str).str.strip(), libelles_norm))
    df = df.copy()
    df["_motifs"] = [lookup.get(c, set()) for c in cles]
    return df


def jours_absence_par_mois(df):
    """Détail mensuel des jours d'absence identifiés : une ligne par
    (Entreprise, Matricule, Nom, Prenom, Mois, Motif), somme de la colonne "Base".
    Une ligne PAIE reconnue sous plusieurs motifs est comptée pour chacun d'eux."""
    df = identifier_motifs(df)
    df = df.loc[df["_motifs"].apply(len) > 0].copy()
    if df.empty:
        return pd.DataFrame(columns=["Entreprise", "Matricule", "Nom", "Prenom",
                                     "Mois", "Motif", "Base"])

    df["_motifs"] = df["_motifs"].apply(list)
    df = df.explode("_motifs")
    df["Base_num"] = to_num(df["base"])
    df["Mois"] = df["periode"].str[:7]

    detail = (df.groupby(["entreprise", "matricule", "nom", "prenom", "Mois", "_motifs"],
                         as_index=False)["Base_num"].sum())
    return detail.rename(columns={"entreprise": "Entreprise", "matricule": "Matricule",
                                  "nom": "Nom", "prenom": "Prenom",
                                  "_motifs": "Motif", "Base_num": "Base"})


def jours_absence_periode(detail_mensuel, mois_retenus):
    """Agrège le détail mensuel sur une période donnée (ex. le T1 : ["2026-01",
    "2026-02", "2026-03"]) par (Entreprise, Matricule, Nom, Prenom, Motif)."""
    sous = detail_mensuel.loc[detail_mensuel["Mois"].isin(mois_retenus)]
    return (sous.groupby(["Entreprise", "Matricule", "Nom", "Prenom", "Motif"],
                         as_index=False)["Base"].sum())
