# -*- coding: utf-8 -*-
"""Chargement CSV DSN, mapping colonnes, normalisation, détection annulations."""
import pandas as pd
from src.normalisation import get_source_juridique, _norm_soc, to_dt, to_num


def lire_source(path):
    if str(path).lower().endswith((".csv", ".tsv", ".txt")):
        last = None
        for enc in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                return pd.read_csv(path, dtype=str, sep=None, engine="python",
                                   encoding=enc, keep_default_na=False, na_values=[])
            except Exception as e:
                last = e
        raise ValueError(f"CSV illisible : {path} ({last})")
    return pd.read_excel(path, engine="openpyxl", dtype=str)


def find_col(df, keyword):
    for c in df.columns:
        if c.strip().lower() == keyword.strip().lower():
            return c
    for c in df.columns:
        if keyword.lower() in c.lower():
            return c
    return None


def mapper_colonnes(df):
    """Repère les colonnes DSN par libellé. Renvoie un dict."""
    col_nir       = find_col(df, "ID")
    col_matricule = find_col(df, "Individu.Matricule")
    col_nom_usage = find_col(df, "NomUsage") or find_col(df, "Individu.NomUsage")
    col_mois      = find_col(df, "mois_absence")
    col_nom       = find_col(df, "S21.G00.30.002 Individu.NomFamille")
    col_prenom    = find_col(df, "S21.G00.30.004 Individu.Prenoms")
    col_statut    = find_col(df, "S21.G00.40.003_Contrat.StatutRC_Libelle")
    col_temps     = find_col(df, "S21.G00.40.014_Contrat.ModaliteTemps_libelle")
    col_motif     = find_col(df, "S21.G00.60.001_TravailArret.Motif_Libelle")
    col_djt       = find_col(df, "S21.G00.60.002_TravailArret.DernierJour")
    col_fin       = find_col(df, "S21.G00.60.003_TravailArret.DateFinPrevisionnelle")
    col_reprise   = find_col(df, "S21.G00.60.010_TravailArret.RepriseDate")
    col_jours     = find_col(df, "Nb_jours_absences_SS")
    col_societe   = find_col(df, "Nom usuel de la société")
    col_annul     = find_col(df, "flag_annulation_G2SA")
    col_decl      = find_col(df, "Declaration.Mois")
    return dict(col_nir=col_nir, col_matricule=col_matricule, col_nom_usage=col_nom_usage,
                col_mois=col_mois, col_nom=col_nom, col_prenom=col_prenom, col_statut=col_statut,
                col_temps=col_temps, col_motif=col_motif, col_djt=col_djt, col_fin=col_fin,
                col_reprise=col_reprise, col_jours=col_jours, col_societe=col_societe,
                col_annul=col_annul, col_decl=col_decl)


REQUIS = ["col_nom","col_prenom","col_motif","col_djt","col_jours","col_mois","col_societe"]


def colonnes_manquantes(cols):
    """Colonnes obligatoires absentes."""
    lib = {"col_nom":"nom","col_prenom":"prénom","col_motif":"motif","col_djt":"DJT",
           "col_jours":"jours","col_mois":"mois","col_societe":"société"}
    return [lib[k] for k in REQUIS if cols.get(k) is None]


def normaliser_et_detecter(df, c):
    """Normalise les champs et pose l'indicateur d'annulation (union 3 signaux)."""
    col_nir=c["col_nir"]; col_matricule=c["col_matricule"]; col_nom_usage=c["col_nom_usage"]
    col_mois=c["col_mois"]; col_nom=c["col_nom"]; col_prenom=c["col_prenom"]; col_statut=c["col_statut"]
    col_temps=c["col_temps"]; col_motif=c["col_motif"]; col_djt=c["col_djt"]; col_fin=c["col_fin"]
    col_reprise=c["col_reprise"]; col_jours=c["col_jours"]; col_societe=c["col_societe"]
    col_annul=c["col_annul"]; col_decl=c["col_decl"]
    # 3. NORMALISATION + (B1) DÉTECTION UNION DES 3 SIGNAUX
    # ===============================================================
    df = df.copy()
    df["NIR"]        = df[col_nir].astype(str).str.strip() if col_nir else ""
    df["Matricule"]  = df[col_matricule].astype(str).str.strip() if col_matricule else ""
    df["Nom"]        = df[col_nom].astype(str).str.strip()
    df["Nom usage"]  = df[col_nom_usage].astype(str).str.strip() if col_nom_usage else ""
    df["Prénom"]     = df[col_prenom].astype(str).str.strip()
    df["Motif arrêt"] = df[col_motif].astype(str).str.strip().str.lower()
    df["Source juridique"] = df["Motif arrêt"].apply(get_source_juridique)
    df["DJT"]        = to_dt(df[col_djt])
    df["FinPrev"]    = to_dt(df[col_fin]) if col_fin else pd.NaT
    df["Reprise"]    = to_dt(df[col_reprise]) if col_reprise else pd.NaT
    df["mois_p"]     = to_dt(df[col_mois]).dt.to_period("M")
    df["Jours src"]  = to_num(df[col_jours])
    df["Société"]    = df[col_societe]
    df["Statut"]     = df[col_statut] if col_statut else ""
    df["Temps de travail"] = df[col_temps] if col_temps else ""
    df["Decl"]       = to_dt(df[col_decl]) if col_decl else pd.NaT
    
    _flag = (df[col_annul].astype(str).str.strip().str.lower() == "true") if col_annul else False
    _motif_annul = df["Motif arrêt"] == "annulation"
    _jours_neg = df["Jours src"] < 0
    # (B1) UNION : aucune de ces conditions ne doit être ignorée
    df["_annul"] = (_flag | _motif_annul | _jours_neg).fillna(False)
    df["_soc"] = df["Société"].apply(_norm_soc)
    return df
