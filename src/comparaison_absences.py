# -*- coding: utf-8 -*-
"""Comparaison des jours d'absence PAIE <-> DSN, par société / salarié / mois / motif.

Un seul fichier pour toute la fonctionnalité (identification des rubriques PAIE,
extraction DSN, jointure, export) : à consulter de haut en bas, section par section.

  1. IDENTIFICATION DES RUBRIQUES PAIE — quelles rubriques comptent comme un motif
     d'absence, et lecture de la colonne "Base" (qui n'est un nombre de jours que pour
     certaines rubriques, cf. config.MOTIFS_PAIE_JOURS).
  2. EXTRACTION DSN — lecture des classeurs déjà reconstruits (output/xlsx/), qui
     portent déjà "Jours absence (mois)" par motif : pas de recalcul, seulement une
     lecture + un regroupement de motifs pour pouvoir comparer avec la PAIE.
  3. COMPARAISON — jointure PAIE <-> DSN sur (matricule normalisé, mois, motif).
  4. EXPORT — classeur Excel avec les colonnes d'identité demandées, une clé pivot par
     salarié, et un filtre Excel cliquable sur chaque colonne (dont l'écart).

Le calcul PAIE part directement du CSV source (via lire_paie(), déjà utilisé pour la
comparaison de populations) plutôt que des classeurs reconstruits : construire_reconstruit()
ne fait qu'un renommage/passage de colonnes, le résultat est donc rigoureusement identique
mais bien plus rapide à obtenir (le classeur "52" à lui seul prend plusieurs minutes à relire).
"""
import os
import re
import shutil
import tempfile
import unicodedata
import openpyxl
import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from config import MOTIFS_PAIE_JOURS, OUTPUT_DIR, RAPPORT_DIR
from src.normalisation import _norm_soc, to_num

NOM_FICHIER_SORTIE = "Comparatif_jours_absence_PAIE_DSN.xlsx"


# ===============================================================
# 0. UTILITAIRES PARTAGÉS
# ===============================================================
def normaliser_matricule(v):
    """Retire les zéros de tête (format DSN "005005" -> "5005", format PAIE déjà nu)."""
    v = str(v).strip()
    return v.lstrip("0") or "0"


def nom_fichier_dsn(societe):
    return f"Reconstruit_{_norm_soc(societe).replace(' ', '-')}_CORRIGE.xlsx"


def nom_fichier_paie(code):
    return f"Reconstruit_PAIE_{code}.xlsx"


def _lire_classeur_verrouille(path, parseur):
    """Ouvre un classeur en lecture seule et applique `parseur(wb)`. Contourne un
    éventuel verrou de synchronisation (OneDrive...) en retentant sur une copie
    temporaire — la lecture doit se terminer (et le classeur être fermé) avant la
    sortie du bloc `with`, sinon le nettoyage du dossier temporaire échoue lui-même."""
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        return parseur(wb)
    except PermissionError:
        with tempfile.TemporaryDirectory() as tmp:
            copie = os.path.join(tmp, os.path.basename(path))
            shutil.copy2(path, copie)
            wb = openpyxl.load_workbook(copie, read_only=True)
            return parseur(wb)


# ===============================================================
# 1. IDENTIFICATION DES RUBRIQUES PAIE ET JOURS D'ABSENCE
# ===============================================================
_PONCTUATION = re.compile(r"[-/.:',]")


def _normaliser_libelle(libelle):
    """Minuscules, sans accents, ponctuation remplacée par des espaces, espaces
    multiples réduits à un seul — tolère les variantes de transcription du référentiel
    métier (ex. "CP-Abs..." et "CP - Abs..." ou "Maternite"/"Maternité" sont identiques)."""
    s = unicodedata.normalize("NFD", str(libelle))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = _PONCTUATION.sub(" ", s.lower())
    return " ".join(s.split())


def _construire_lookup_motifs(mapping=MOTIFS_PAIE_JOURS):
    """{(rubrique, libellé normalisé): {motifs}} à partir d'un mapping
    motif -> [(libellé, code), ...] (par défaut config.MOTIFS_PAIE_JOURS)."""
    lookup = {}
    for motif, entrees in mapping.items():
        for libelle, code in entrees:
            cle = (str(code).strip(), _normaliser_libelle(libelle))
            lookup.setdefault(cle, set()).add(motif)
    return lookup


def _identifier_motifs(df):
    """Ajoute une colonne "_motifs" (set des motifs reconnus, vide si aucune
    correspondance) au DataFrame PAIE (colonnes source : rubrique, libelle)."""
    lookup = _construire_lookup_motifs()
    libelles_norm = df["libelle"].apply(_normaliser_libelle)
    cles = list(zip(df["rubrique"].astype(str).str.strip(), libelles_norm))
    df = df.copy()
    df["_motifs"] = [lookup.get(c, set()) for c in cles]
    return df


def charger_jours_paie(chemin_paie_source, annee, sep):
    """{code entreprise: DataFrame(Matricule, Nom, Prenom, Etablissement, Siren, Nic,
    Siret, Mois, Motif, Jours)} calculé depuis le CSV source PAIE_AUDIT.

    Un même (rubrique, libelle) peut légitimement compter pour plusieurs motifs à la
    fois (cf. config.MOTIFS_PAIE_JOURS, ex. "Absence Paternite (Maintenue)"/3405 sous
    maternité ET paternité) : la ligne est alors dupliquée (explode), une fois par motif."""
    from src.chargement_paie import lire_paie
    df = lire_paie(chemin_paie_source, annee=annee, sep=sep)
    df = _identifier_motifs(df)
    df = df.loc[df["_motifs"].apply(len) > 0].copy()
    if df.empty:
        cols = ["Entreprise", "Matricule", "Nom", "Prenom", "Etablissement", "Siren",
               "Nic", "Siret", "Mois", "Motif", "Jours"]
        return {}

    df["_motifs"] = df["_motifs"].apply(list)
    df = df.explode("_motifs")
    df["Jours"] = to_num(df["base"])
    df["Mois"] = df["periode"].str[:7]
    df["Matricule"] = df["matricule"].apply(normaliser_matricule)

    detail = (df.groupby(["entreprise", "Matricule", "nom", "prenom", "etablissement",
                         "siren", "nic", "siret", "Mois", "_motifs"], as_index=False)
                ["Jours"].sum())
    detail = detail.rename(columns={"entreprise": "Entreprise", "nom": "Nom", "prenom": "Prenom",
                                    "etablissement": "Etablissement", "siren": "Siren",
                                    "nic": "Nic", "siret": "Siret", "_motifs": "Motif"})
    return {ent: g.drop(columns=["Entreprise"]) for ent, g in detail.groupby("Entreprise")}


COLONNES_ANOMALIES = ["Entreprise", "Matricule", "Nom", "Prenom", "Mois", "Motif",
                      "Rubrique", "Libelle", "Base", "Anomalie"]


def detecter_anomalies_jours(chemin_paie_source, annee, sep):
    """Signale les lignes PAIE_AUDIT (sur les rubriques "jours" de MOTIFS_PAIE_JOURS)
    où "Base" est physiquement incohérente pour un mois calendaire : plus de 31 jours,
    ou négative sur une rubrique qui n'est pas une annulation (le libellé contient
    "annul"). Confirmé sur la donnée elle-même (pas un artefact d'agrégation) : voir
    échanges de validation — chaque cas correspond à une seule ligne brute du CSV
    source, sans information de sous-période/rétroactivité exploitable pour l'expliquer
    (date_sous_periode/num_sous_periode/date_retro sont à blanc sur ces lignes)."""
    from src.chargement_paie import lire_paie
    df = lire_paie(chemin_paie_source, annee=annee, sep=sep)
    df = _identifier_motifs(df)
    df = df.loc[df["_motifs"].apply(len) > 0].copy()
    if df.empty:
        return pd.DataFrame(columns=COLONNES_ANOMALIES)

    df["Base_num"] = to_num(df["base"])
    est_annulation = df["libelle"].apply(_normaliser_libelle).str.contains("annul")
    anormal = (df["Base_num"] > 31) | ((df["Base_num"] < 0) & ~est_annulation)
    df = df.loc[anormal].copy()
    if df.empty:
        return pd.DataFrame(columns=COLONNES_ANOMALIES)

    df["_motifs"] = df["_motifs"].apply(list)
    df = df.explode("_motifs")
    df["Matricule"] = df["matricule"].apply(normaliser_matricule)
    df["Mois"] = df["periode"].str[:7]
    df["Anomalie"] = df["Base_num"].apply(
        lambda b: "Base > 31 j sur un mois" if b > 31 else "Base négative (hors annulation)")

    sortie = pd.DataFrame({
        "Entreprise": df["entreprise"], "Matricule": df["Matricule"], "Nom": df["nom"],
        "Prenom": df["prenom"], "Mois": df["Mois"], "Motif": df["_motifs"],
        "Rubrique": df["rubrique"], "Libelle": df["libelle"], "Base": df["Base_num"],
        "Anomalie": df["Anomalie"],
    })
    return sortie.sort_values(["Entreprise", "Matricule", "Mois"]).reset_index(drop=True)


# ===============================================================
# 2. EXTRACTION DSN (classeurs déjà reconstruits)
# ===============================================================
_COLONNES_DSN = ("Matricule", "Nom", "Prénom", "Nom usage", "NIR", "Motif arrêt",
                "Mois DSN", "Jours absence (mois)", "Type de ligne")
_LIGNE_DONNEES = "📍 Données DSN"


def canoniser_motif_dsn(motif):
    """Ramène un motif DSN (texte libre) à l'un des paniers alignés sur le référentiel
    PAIE (config.MOTIFS_PAIE_JOURS) : regroupe accident de travail + accident de trajet
    + maladie professionnelle (non distingués côté paie). Les motifs sans équivalent
    PAIE identifié à ce jour (temps partiel thérapeutique, adoption) gardent leur
    propre étiquette plutôt que d'être rattachés arbitrairement à un autre."""
    m = str(motif).strip().lower()
    if "accident de travail" in m or "accident du travail" in m:
        return "accident de travail"
    if "accident de trajet" in m or "accident du trajet" in m:
        return "accident de travail"
    if "maladie professionnelle" in m:
        return "accident de travail"
    if "maternité" in m:
        return "maternité"
    if "paternité" in m:
        return "paternité"
    if "adoption" in m:
        return "adoption (pas de rubrique PAIE identifiée)"
    if "temps partiel" in m:
        return "temps partiel thérapeutique (pas de rubrique PAIE identifiée)"
    if "maladie" in m:
        return "maladie"
    return "autre / non classé"


def _parser_dsn(wb):
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = list(next(rows))
    idx = {c: header.index(c) for c in _COLONNES_DSN}
    lignes = []
    for r in rows:
        mat, mois, jours = r[idx["Matricule"]], r[idx["Mois DSN"]], r[idx["Jours absence (mois)"]]
        if not mat or not mois or jours in (None, "") or r[idx["Type de ligne"]] != _LIGNE_DONNEES:
            continue
        lignes.append({
            "Matricule": normaliser_matricule(mat), "Nom": r[idx["Nom"]], "Prenom": r[idx["Prénom"]],
            "Nom usage": r[idx["Nom usage"]] or "", "NIR": r[idx["NIR"]] or "",
            "Mois": str(mois), "Motif": canoniser_motif_dsn(r[idx["Motif arrêt"]]),
            "Jours": float(jours),
        })
    wb.close()
    return pd.DataFrame(lignes)


def charger_jours_dsn(mapping):
    """{société DSN: DataFrame(Matricule, Nom, Prenom, Nom usage, NIR, Mois, Motif,
    Jours)} depuis les classeurs DSN déjà reconstruits. Seules les lignes "📍 Données
    DSN" sont sommées : les lignes "❌ Annulation DSN" sont un historique de corrections
    déjà intégrées dans le résultat positif survivant (cf. pipeline.py, survivance par
    (emp, motif, DJT)) — les ressommer reviendrait à soustraire deux fois les mêmes
    corrections."""
    resultat = {}
    for societe_dsn in mapping.values():
        path = os.path.join(OUTPUT_DIR, nom_fichier_dsn(societe_dsn))
        detail = _lire_classeur_verrouille(path, _parser_dsn)
        if detail.empty:
            resultat[societe_dsn] = detail
            continue
        resultat[societe_dsn] = (detail.groupby(
            ["Matricule", "Nom", "Prenom", "Nom usage", "NIR", "Mois", "Motif"],
            as_index=False)["Jours"].sum())
    return resultat


_COL_SIREN_DSN = "S21.G00.06.001 Entreprise.Siren"
_COL_NIC_DSN = "S21.G00.11.001 Etablissement.Nic"


def lire_identite_dsn_source(chemin_dsn_source):
    """{matricule normalisé: (siren, nic, siret)} depuis le CSV source DSN (pas le
    classeur reconstruit, qui ne porte pas ces colonnes). Le SIREN est unique par
    société et le NIC constant par salarié (vérifié empiriquement sur 2026 : aucun
    matricule n'a plus d'un NIC) ; le SIRET est simplement SIREN+NIC concaténés (14
    chiffres). Bien plus fiable que la PAIE, où ces 3 champs sont à 0 pour toutes les
    lignes de PAIE_AUDIT.csv."""
    from src.chargement import lire_source, mapper_colonnes, normaliser_et_detecter
    df = lire_source(chemin_dsn_source)
    df.columns = df.columns.str.strip()
    cols = mapper_colonnes(df)
    df = normaliser_et_detecter(df, cols)
    df = df.loc[(df["_soc"] != "") & df[_COL_SIREN_DSN].notna() & df[_COL_NIC_DSN].notna()]
    df = df.drop_duplicates(subset=["Matricule"])

    identite = {}
    for mat, siren, nic in zip(df["Matricule"], df[_COL_SIREN_DSN], df[_COL_NIC_DSN]):
        siren, nic = str(siren).strip(), str(nic).strip()
        identite[normaliser_matricule(mat)] = (siren, nic, f"{siren}{nic}")
    return identite


# ===============================================================
# 3. COMPARAISON
# ===============================================================
COLONNES_SORTIE = [
    "Clé salarié (pivot)", "Nom", "Prénom", "Nom d'usage", "NIR", "Matricule",
    "Entreprise", "Etablissement", "Siren", "Nic", "Siret", "Motif",
    "Nombre de jours absence PAIE", "Nombre de jours absence DSN",
    "Mois absence (période)", "Écart constaté",
]


def comparer_jours(mapping, jours_dsn_par_societe, jours_paie_par_code, identite_dsn=None):
    """Fusionne DSN et PAIE par société sur (Matricule, Mois, Motif). Renvoie un
    DataFrame avec les colonnes COLONNES_SORTIE — une ligne par (société, salarié,
    mois, motif). La "Clé salarié (pivot)" (Entreprise + Matricule) permet de vérifier
    d'un coup d'œil qu'on compare bien le même salarié des deux côtés.

    `identite_dsn` (cf. lire_identite_dsn_source) : {matricule: (siren, nic, siret)}
    — si fourni, écrase le Siren/Nic/Siret de la PAIE (toujours à 0 dans PAIE_AUDIT.csv)
    par la valeur réelle issue de la DSN, quand le salarié y est identifié."""
    identite_dsn = identite_dsn or {}
    lignes = []
    cols_vides = ["Matricule", "Nom", "Prenom", "Mois", "Motif", "Jours"]
    for code_paie, societe_dsn in mapping.items():
        dsn = jours_dsn_par_societe.get(societe_dsn)
        if dsn is None or dsn.empty:
            dsn = pd.DataFrame(columns=cols_vides + ["Nom usage", "NIR"])
        paie = jours_paie_par_code.get(code_paie)
        if paie is None or paie.empty:
            paie = pd.DataFrame(columns=cols_vides + ["Etablissement", "Siren", "Nic", "Siret"])

        fusion = pd.merge(dsn, paie, on=["Matricule", "Mois", "Motif"], how="outer",
                          suffixes=("_DSN", "_PAIE"))
        fusion["Nom"] = fusion["Nom_DSN"].fillna(fusion["Nom_PAIE"])
        fusion["Prénom"] = fusion["Prenom_DSN"].fillna(fusion["Prenom_PAIE"])
        fusion["Jours_DSN"] = fusion["Jours_DSN"].fillna(0.0)
        fusion["Jours_PAIE"] = fusion["Jours_PAIE"].fillna(0.0)

        fusion["Clé salarié (pivot)"] = f"{code_paie}-" + fusion["Matricule"].astype(str)
        fusion["Entreprise"] = code_paie
        for col in ("Nom usage", "NIR", "Etablissement", "Siren", "Nic", "Siret"):
            if col not in fusion.columns:
                fusion[col] = ""
            else:
                fusion[col] = fusion[col].fillna("")

        # Nom usage/NIR/Etablissement ne viennent que d'un des deux côtés (DSN ou PAIE) :
        # sur un mois où ce côté n'a pas de ligne pour ce salarié, la valeur reste vide,
        # alors qu'un autre mois du même salarié peut l'avoir. Sans cette propagation,
        # comparer_jours_periode() (qui regroupe notamment sur ces colonnes) fractionne à
        # tort un même salarié en plusieurs lignes selon les mois où chaque source a parlé
        # (vu concrètement sur TEYSSEYRE Justine/AURA : Etablissement vide en janvier,
        # "2" en février, scindant un total par ailleurs strictement identique 21=21).
        for col in ("Nom usage", "NIR", "Etablissement"):
            premiere_valeur = fusion.groupby("Matricule")[col].transform(
                lambda s: next((v for v in s if v not in (None, "")), ""))
            fusion[col] = premiere_valeur

        if identite_dsn:
            reel = fusion["Matricule"].map(identite_dsn)
            a_identite = reel.notna()
            fusion.loc[a_identite, "Siren"] = reel[a_identite].apply(lambda t: t[0])
            fusion.loc[a_identite, "Nic"] = reel[a_identite].apply(lambda t: t[1])
            fusion.loc[a_identite, "Siret"] = reel[a_identite].apply(lambda t: t[2])

        lignes.append(pd.DataFrame({
            "Clé salarié (pivot)": fusion["Clé salarié (pivot)"],
            "Nom": fusion["Nom"], "Prénom": fusion["Prénom"],
            "Nom d'usage": fusion["Nom usage"], "NIR": fusion["NIR"],
            "Matricule": fusion["Matricule"], "Entreprise": fusion["Entreprise"],
            "Etablissement": fusion["Etablissement"], "Siren": fusion["Siren"],
            "Nic": fusion["Nic"], "Siret": fusion["Siret"], "Motif": fusion["Motif"],
            "Nombre de jours absence PAIE": fusion["Jours_PAIE"],
            "Nombre de jours absence DSN": fusion["Jours_DSN"],
            "Mois absence (période)": fusion["Mois"],
            "Écart constaté": fusion["Jours_DSN"] - fusion["Jours_PAIE"],
        }))

    return pd.concat(lignes, ignore_index=True) if lignes else pd.DataFrame(columns=COLONNES_SORTIE)


def comparer_jours_periode(detail_mensuel, mois_retenus, libelle_periode="T1"):
    """Agrège le détail mensuel sur une période (ex. le T1 : ["2026-01", "2026-02",
    "2026-03"]) par (Clé salarié, Motif) — même colonnes que le détail mensuel, avec
    "Mois absence (période)" remplacé par le libellé de la période agrégée."""
    sous = detail_mensuel.loc[detail_mensuel["Mois absence (période)"].isin(mois_retenus)]
    cols_identite = ["Clé salarié (pivot)", "Nom", "Prénom", "Nom d'usage", "NIR", "Matricule",
                     "Entreprise", "Etablissement", "Siren", "Nic", "Siret", "Motif"]
    agg = (sous.groupby(cols_identite, as_index=False)
              [["Nombre de jours absence PAIE", "Nombre de jours absence DSN"]].sum())
    agg["Mois absence (période)"] = libelle_periode
    agg["Écart constaté"] = agg["Nombre de jours absence DSN"] - agg["Nombre de jours absence PAIE"]
    return agg[COLONNES_SORTIE]


# ===============================================================
# 4. EXPORT EXCEL
# ===============================================================
NAVY = "1F3864"
REDF = "F4D6D2"
GREENF = "DDEBE2"
ORANGEF = "FBE7CE"


def _styler_entete_et_filtre(ws, df):
    for cell in ws[1]:
        cell.font = Font(bold=True, size=10, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
    ws.freeze_panes = "A2"
    # Filtre Excel cliquable sur chaque colonne : l'utilisateur ouvre le menu déroulant
    # de l'en-tête pour isoler ce qui l'intéresse (ex. "Écart constaté" différent de 0).
    ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"


def _ecrire_onglet(writer, df, nom_onglet):
    df = df if not df.empty else pd.DataFrame(columns=COLONNES_SORTIE)
    df.to_excel(writer, index=False, sheet_name=nom_onglet)
    ws = writer.sheets[nom_onglet]
    _styler_entete_et_filtre(ws, df)

    j_ecart = list(df.columns).index("Écart constaté") + 1
    for i, ecart in enumerate(df["Écart constaté"], start=2):
        ws.cell(i, j_ecart).fill = PatternFill("solid", fgColor=(GREENF if ecart == 0 else REDF))


def _ecrire_onglet_anomalies(writer, df, nom_onglet):
    df = df if not df.empty else pd.DataFrame(columns=COLONNES_ANOMALIES)
    df.to_excel(writer, index=False, sheet_name=nom_onglet)
    ws = writer.sheets[nom_onglet]
    _styler_entete_et_filtre(ws, df)
    for i in range(2, len(df) + 2):
        for j in range(1, len(df.columns) + 1):
            ws.cell(i, j).fill = PatternFill("solid", fgColor=ORANGEF)


def ecrire_comparatif_absences(detail_mensuel, detail_periode, anomalies, libelle_periode="T1"):
    out = os.path.join(RAPPORT_DIR, NOM_FICHIER_SORTIE)
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        _ecrire_onglet(writer, detail_mensuel, "1 - Détail mensuel")
        _ecrire_onglet(writer, detail_periode, f"2 - Total {libelle_periode}")
        _ecrire_onglet_anomalies(writer, anomalies, "3 - Anomalies (jours)")

    print(f"   ✅ écrit -> {out}")
    return out
