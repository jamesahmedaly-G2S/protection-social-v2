# -*- coding: utf-8 -*-
"""Chargement PAIE_AUDIT.csv (étape 1 : structure uniquement).

Grain conservé tel quel (une ligne par rubrique) : on filtre l'année et on découpe
par société, sans rien sommer ni fusionner. Lecture par blocs (le fichier source
pèse plusieurs centaines de Mo / plusieurs millions de lignes) pour rester léger
en mémoire : seules les lignes de l'année demandée sont conservées à chaque bloc.
"""
import pandas as pd
from config import PAIE_COLONNES_ORIGINE, PAIE_COLONNES_PPU, PAIE_COLONNES_THEMATIQUES

TAILLE_BLOC = 200_000


def _colonnes_source():
    return [src for src, _ in PAIE_COLONNES_ORIGINE]


def lire_paie(path, annee, sep=";"):
    """Lit PAIE_AUDIT.csv par blocs, ne garde que les lignes dont l'année de
    la colonne "periode" correspond à `annee`. Renvoie un DataFrame (dtype=str,
    colonnes source non renommées)."""
    usecols = _colonnes_source()
    annee_str = str(annee)
    derniere_erreur = None
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            blocs = []
            lecteur = pd.read_csv(path, sep=sep, dtype=str, usecols=usecols,
                                  chunksize=TAILLE_BLOC, encoding=enc,
                                  keep_default_na=False, na_values=[])
            for bloc in lecteur:
                retenu = bloc.loc[bloc["periode"].str[:4] == annee_str]
                if not retenu.empty:
                    blocs.append(retenu)
            if not blocs:
                return pd.DataFrame(columns=usecols)
            return pd.concat(blocs, ignore_index=True)
        except Exception as e:
            derniere_erreur = e
            continue
    raise ValueError(f"CSV illisible : {path} ({derniere_erreur})")


def decouper_par_entreprise(df):
    """Découpe le DataFrame en {entreprise: sous-DataFrame}, une clé par valeur
    distincte de la colonne "entreprise" (une ligne par rubrique conservée)."""
    return {ent: g.copy() for ent, g in df.groupby("entreprise", sort=True)}


def construire_reconstruit(df):
    """Construit le DataFrame de sortie : colonnes d'origine renommées (telles
    quelles) + colonnes PPU et thématiques réservées, vides à cette étape."""
    sortie = pd.DataFrame(index=df.index)
    for src, dst in PAIE_COLONNES_ORIGINE:
        sortie[dst] = df[src].values
    for col in PAIE_COLONNES_PPU + PAIE_COLONNES_THEMATIQUES:
        sortie[col] = ""
    return sortie
