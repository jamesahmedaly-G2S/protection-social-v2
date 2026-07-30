# -*- coding: utf-8 -*-
"""Chargement PAIE_AUDIT.csv + référentiel PPU, jointure, découpage par société.

Grain conservé tel quel (une ligne par rubrique) : on filtre l'année et on découpe
par société, sans rien sommer ni fusionner. Lecture par blocs (le fichier source
pèse plusieurs centaines de Mo / plusieurs millions de lignes) pour rester léger
en mémoire : seules les lignes de l'année demandée sont conservées à chaque bloc.
"""
import pandas as pd
from config import PAIE_COLONNES_ORIGINE, PPU_COL_CODE, PPU_COL_LIBELLE

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


def lire_ppu(path, sheet):
    """Lit le référentiel PPU (toutes ses colonnes, telles quelles). dtype=str
    pour que le code de rubrique se compare directement à la colonne "rubrique"
    de PAIE_AUDIT (pas de suffixe ".0")."""
    return pd.read_excel(path, sheet_name=sheet, dtype=str,
                         keep_default_na=False, na_values=[])


def dedupliquer_ppu(ppu):
    """Le référentiel contient des lignes strictement dupliquées, et quelques
    couples (code, libellé) aux classifications divergentes (résidu marginal,
    cf. config.py) : on garde la première occurrence rencontrée pour chaque
    couple, pour que la jointure ne démultiplie jamais une ligne PAIE."""
    return ppu.drop_duplicates(subset=[PPU_COL_CODE, PPU_COL_LIBELLE], keep="first")


def joindre_ppu(paie, ppu):
    """Jointure gauche PAIE_AUDIT <-> PPU sur la clé composite (rubrique, libelle)
    <-> (Ppu Rubrique Code, Ppu Rubrique Libelle). "how=left" : une rubrique sans
    correspondance dans le PPU (hors périmètre IJSS — taxes, cotisations, etc.)
    reste dans le résultat, avec les colonnes PPU vides plutôt que la ligne PAIE
    supprimée."""
    gauche = paie.copy()
    gauche["_cle_rubrique"] = gauche["rubrique"].str.strip()
    gauche["_cle_libelle"] = gauche["libelle"].str.strip()
    droite = ppu.copy()
    droite["_cle_rubrique"] = droite[PPU_COL_CODE].str.strip()
    droite["_cle_libelle"] = droite[PPU_COL_LIBELLE].str.strip()

    fusion = gauche.merge(droite, on=["_cle_rubrique", "_cle_libelle"], how="left")
    return fusion.drop(columns=["_cle_rubrique", "_cle_libelle"])


def decouper_par_entreprise(df):
    """Découpe le DataFrame en {entreprise: sous-DataFrame}, une clé par valeur
    distincte de la colonne "entreprise" (une ligne par rubrique conservée)."""
    return {ent: g.copy() for ent, g in df.groupby("entreprise", sort=True)}


def construire_reconstruit(df, colonnes_ppu):
    """Construit le DataFrame de sortie : colonnes d'origine PAIE_AUDIT renommées
    (telles quelles) + toutes les colonnes du PPU, verbatim, dans l'ordre du
    référentiel. Une rubrique sans correspondance PPU (jointure "left") donne des
    colonnes PPU vides plutôt que NaN."""
    sortie = pd.DataFrame(index=df.index)
    for src, dst in PAIE_COLONNES_ORIGINE:
        sortie[dst] = df[src].values
    for col in colonnes_ppu:
        sortie[col] = df[col].values if col in df.columns else ""
    return sortie.fillna("")
