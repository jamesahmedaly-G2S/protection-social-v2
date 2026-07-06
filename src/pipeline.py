# -*- coding: utf-8 -*-
"""Cœur du pipeline : traitement d'une société (logique métier validée, inchangée).

NB — fonction gardée d'un seul tenant : ses étapes (survivance, qualification,
déchevauchement, ventilation) partagent un état local très couplé ; la scinder
introduirait un risque de régression. Les anomalies sont tracées via J (JournalAudit)."""
import pandas as pd
import numpy as np
from src.normalisation import _norm_soc, get_source_juridique, jours_cal_mois
from src.journal_audit import ANO
from config import (DATE_DEBUT_PERIODE, DATE_FIN_PERIODE, SEUIL_PREVOYANCE,
                    MOTIFS_CARENCE_3J, ARBITRAGE_MATERNITE, rang_motif)


def traiter_societe(SOCIETE_CIBLE, dsn_full, J=None):
    """Traite une société. dsn_full = DSN normalisée ; J = JournalAudit optionnel."""
    if J is not None: J.societe(SOCIETE_CIBLE)
    cible = _norm_soc(SOCIETE_CIBLE)
    dsn = dsn_full.loc[dsn_full["_soc"] == cible].copy()
    print(f"\n🏢 ===== {SOCIETE_CIBLE} : {len(dsn)} lignes =====")
    if dsn.empty:
        print("   ⚠️ aucune ligne. Société ignorée.")
        if J is not None: J.tracer(ANO.SOCIETE_VIDE)
        return None

    def cle_salarie(r):
        return r["Matricule"] if str(r["Matricule"]).strip() else f"{r['Nom']}|{r['Prénom']}|{r['NIR']}"
    dsn["_emp"] = dsn.apply(cle_salarie, axis=1)
    NB_SAL = dsn["_emp"].nunique()

    n_djt_invalide = int(dsn["DJT"].isna().sum())
    n_mois_invalide = int(dsn["mois_p"].isna().sum())
    dsn = dsn.dropna(subset=["DJT", "mois_p"]).copy()
    if J is not None and n_djt_invalide:
        J.tracer(ANO.DATE_INVALIDE_DJT, contexte=f"{n_djt_invalide} ligne(s) écartée(s)",
                 correction_auto="AUTO — lignes sans DJT exploitable exclues du calcul.")
    if J is not None and n_mois_invalide:
        J.tracer(ANO.DATE_INVALIDE_MOIS, contexte=f"{n_mois_invalide} ligne(s) écartée(s)",
                 correction_auto="AUTO — lignes sans mois d'absence exploitable exclues du calcul.")

    # ----------------------------------------------------------
    # (B1bis) RÉ-ATTRIBUTION DU MOTIF DES ANNULATIONS
    #   Une ligne d'annulation porte souvent le motif générique « annulation »
    #   (cas POUYET) : elle annule pourtant un épisode d'un motif réel. On la
    #   rattache au motif de la déclaration positive qu'elle vise, par
    #   correspondance (emp, DJT, FinPrev, mois) puis (emp, DJT, mois) puis
    #   (emp, DJT). Sans rattachement, l'annulation resterait dans un groupe
    #   « annulation » isolé et ne tuerait pas l'épisode réel.
    #   « Motif épisode » sert UNIQUEMENT à la survivance/qualification ;
    #   le reconstruit positif et la traçabilité gardent le motif d'origine.
    # ----------------------------------------------------------
    pos_lines = dsn.loc[(~dsn["_annul"]) & (dsn["Jours src"] > 0)]
    map_fpm, map_pm, map_p = {}, {}, {}
    for _, r in pos_lines.iterrows():
        map_fpm.setdefault((r["_emp"], r["DJT"], r["FinPrev"], r["mois_p"]), r["Motif arrêt"])
        map_pm.setdefault((r["_emp"], r["DJT"], r["mois_p"]), r["Motif arrêt"])
        map_p.setdefault((r["_emp"], r["DJT"]), r["Motif arrêt"])
    def motif_episode(r):
        if not r["_annul"] or r["Motif arrêt"] != "annulation":
            return r["Motif arrêt"]
        for key, m in [((r["_emp"], r["DJT"], r["FinPrev"], r["mois_p"]), map_fpm),
                       ((r["_emp"], r["DJT"], r["mois_p"]), map_pm),
                       ((r["_emp"], r["DJT"]), map_p)]:
            if key in m:
                return m[key]
        return r["Motif arrêt"]
    dsn["Motif épisode"] = dsn.apply(
        lambda r: motif_episode(r) if (r["_annul"] and r["Motif arrêt"] == "annulation")
        else r["Motif arrêt"], axis=1)

    # ----------------------------------------------------------
    # (B2) SURVIVANCE : un épisode (emp, motif épisode, DJT) est VIVANT si le
    #      netting de son dépôt DSN le plus récent (Decl max) est positif.
    # ----------------------------------------------------------
    epi_keys = ["_emp", "Motif épisode", "DJT"]
    vivant = {}            # (emp, motif épisode, DJT) -> bool
    finrep = {}            # (emp, motif épisode, DJT) -> (FinPrev, Reprise) du dépôt vivant
    for cle, g in dsn.groupby(epi_keys, dropna=False):
        # dépôt le plus récent (Decl max ; à Decl manquante, on retient l'ordre source)
        if g["Decl"].notna().any():
            decl_max = g["Decl"].max()
            dernier = g.loc[g["Decl"] == decl_max]
        else:
            dernier = g
        net_dernier = dernier["Jours src"].sum()
        est_vivant = net_dernier > 0
        vivant[cle] = est_vivant
        if est_vivant:
            pos = dernier.loc[(~dernier["_annul"]) & (dernier["Jours src"] > 0)]
            if pos.empty:
                pos = dernier.loc[dernier["Jours src"] > 0]
            fp = pos["FinPrev"].dropna().max() if pos["FinPrev"].notna().any() else pd.NaT
            # reprise : valeur de la ligne positive retenue (fin la plus tardive)
            rp = pd.NaT
            if not pos.empty and pos["Reprise"].notna().any():
                rp = pos.sort_values("FinPrev", na_position="first").iloc[-1]["Reprise"]
            finrep[cle] = (fp, rp)

    # ----------------------------------------------------------
    # (B4) QUALIFICATION par famille (emp, motif)
    # ----------------------------------------------------------
    interpretation = {}    # (emp, motif épisode) -> catégorie
    fam_meta = {}          # (emp, motif épisode) -> dict(n_annul, dead_decl)
    familles_avec_annul = set()
    for (emp, motif), g in dsn.groupby(["_emp", "Motif épisode"], dropna=False):
        djts = g["DJT"].dropna().unique()
        djts_vivants = [d for d in djts if vivant.get((emp, motif, d), False)]
        djts_morts   = [d for d in djts if not vivant.get((emp, motif, d), False)]
        if not g["_annul"].any():
            interpretation[(emp, motif)] = "RAS"
            continue
        familles_avec_annul.add((emp, motif))
        n_annul = int(g["_annul"].sum())
        _dd = g.loc[(~g["_annul"]) & (g["Jours src"] > 0) & (g["DJT"].isin(djts_morts)), "Jours src"].max()
        dead_decl = int(_dd) if pd.notna(_dd) else 0
        fam_meta[(emp, motif)] = {"n_annul": n_annul, "dead_decl": dead_decl}
        _who = str(g["Nom"].iloc[0]) + " " + str(g["Prénom"].iloc[0])
        if len(djts_vivants) == 0:
            interpretation[(emp, motif)] = "PURE"
            if J is not None: J.tracer(ANO.ANNULATION_PURE, salarie=_who, contexte=motif)
        elif len(djts_morts) > 0:
            if len(djts_vivants) == 1 and len(djts_morts) == 1 and dead_decl:
                interpretation[(emp, motif)] = "REDUITE"
                if J is not None: J.tracer(ANO.ANNULATION_REDUITE, salarie=_who, contexte=motif)
            else:
                interpretation[(emp, motif)] = "AVERIF"
                if J is not None: J.tracer(ANO.ANNULATION_AVERIF, salarie=_who, contexte=motif)
        else:
            interpretation[(emp, motif)] = "COMPENSEE"


    # ----------------------------------------------------------
    # (B3) ANNULATIONS BRUTES À TRACER (lignes négatives détectées)
    #      Conservées telles quelles, rattachées à leur famille.
    # ----------------------------------------------------------
    annul_raw = dsn.loc[dsn["_annul"]].copy()

    # ===============================================================
    # 6. CONSTRUCTION DES ÉPISODES VIVANTS UNIQUEMENT
    # ===============================================================
    cle_epi = ["_emp", "Nom", "Prénom", "NIR", "Matricule", "Nom usage", "Société",
               "Statut", "Temps de travail", "Motif arrêt", "Source juridique", "DJT"]

    def borne_fin_couple(fp, rp):
        end = fp
        if pd.notna(rp):
            er = rp - pd.Timedelta(days=1)
            end = er if pd.isna(end) else min(end, er)
        return end

    base = (dsn.loc[(~dsn["_annul"]) & (dsn["Jours src"] > 0)]
               .drop_duplicates(subset=cle_epi)[cle_epi].copy())
    episodes = []
    for _, rec in base.iterrows():
        cle = (rec["_emp"], rec["Motif arrêt"], rec["DJT"])
        if not vivant.get(cle, False):
            continue
        fp, rp = finrep.get(cle, (pd.NaT, pd.NaT))
        d = rec.to_dict()
        d["start"] = rec["DJT"] + pd.Timedelta(days=1)
        d["end"]   = borne_fin_couple(fp, rp)
        d["FinPrev_dsn"] = fp
        d["Reprise_dsn"] = rp
        episodes.append(d)
    episodes = pd.DataFrame(episodes)
    if not episodes.empty:
        episodes = episodes.dropna(subset=["start", "end"])
        episodes = episodes.loc[episodes["end"] >= episodes["start"]].copy()

    # ===============================================================
    # 7. (A) DÉCHEVAUCHEMENT PAR SALARIÉ  (correctif v2 conservé)
    # ===============================================================
    journal_chevauchement = []
    if not episodes.empty:
        episodes["start_corr"] = episodes["start"]
        episodes["end_corr"]   = episodes["end"]
        episodes["tronque"]    = False
        def is_maternite(m): return "maternité" in str(m)
        def is_tpt(m): return "temps partiel" in str(m).lower()
        _djtk = episodes["_emp"].astype(str) + "||" + episodes["DJT"].astype(str)
        episodes["_meme_djt"] = _djtk.map(episodes.groupby(_djtk)["Motif arrêt"].nunique() > 1).fillna(False)
        episodes["_rang"] = episodes["Motif arrêt"].map(rang_motif)
        for emp, grp in episodes.groupby("_emp"):
            # Tri : par début ; à début égal (même DJT), le motif PRIORITAIRE d'abord
            # (rang décroissant), puis le plus long — ainsi le motif prioritaire conserve
            # les jours en conflit (maternité/AT priment sur la maladie).
            idx = grp.sort_values(["start", "_rang", "end"],
                                  ascending=[True, False, False]).index.tolist()
            occupe_jusqua = pd.NaT
            motif_occupant = None; rang_occupant = None
            for i in idx:
                s = episodes.at[i, "start_corr"]; e = episodes.at[i, "end_corr"]
                motif = episodes.at[i, "Motif arrêt"]; mat_emp = episodes.at[i, "Matricule"]
                rang_i = episodes.at[i, "_rang"]
                forced = None
                if is_maternite(motif) and str(mat_emp) in ARBITRAGE_MATERNITE:
                    forced = pd.Timestamp(ARBITRAGE_MATERNITE[str(mat_emp)])
                nouveau_start = s
                if pd.notna(occupe_jusqua) and s <= occupe_jusqua:
                    nouveau_start = occupe_jusqua + pd.Timedelta(days=1)
                    nom_sal = str(episodes.at[i, "Nom"])
                    if episodes.at[i, "_meme_djt"]:
                        if rang_occupant is not None and rang_occupant > rang_i:
                            # départage automatique par priorité (maternité/AT retenu)
                            regle = f"Même DJT — priorité : {motif_occupant} retenu (à confirmer)"
                            if J is not None:
                                J.tracer(ANO.MEME_DJT, salarie=nom_sal,
                                         contexte=f"{motif_occupant} vs {motif}",
                                         detail=f"{motif_occupant} retenu par priorité — à confirmer")
                        else:
                            # rangs égaux : la machine ne tranche pas
                            regle = "Même DJT — rang égal — arbitrage"
                            if J is not None:
                                J.tracer(ANO.MEME_DJT, salarie=nom_sal,
                                         contexte=f"{motif_occupant} vs {motif}",
                                         detail="conflit à rang égal — arbitrage requis")
                    elif is_tpt(motif) or is_tpt(motif_occupant):
                        # TPT vs arrêt total (DJT différents) : comportement actuel conservé, signalé
                        regle = "Chronologique — chevauchement TPT (à vérifier)"
                        if J is not None:
                            J.tracer(ANO.TPT_CHEVAUCHEMENT, salarie=nom_sal,
                                     contexte=f"{motif_occupant} vs {motif}")
                    else:
                        regle = "Chronologique (motif précédent prioritaire)"
                    journal_chevauchement.append({
                        "Salarié": episodes.at[i, "Nom"], "Matricule": mat_emp, "Motif tronqué": motif,
                        "Début initial": s.date(), "Début corrigé": nouveau_start.date(), "Règle": regle})
                if forced is not None:
                    if forced != nouveau_start:
                        journal_chevauchement.append({
                            "Salarié": episodes.at[i, "Nom"], "Matricule": mat_emp, "Motif tronqué": motif,
                            "Début initial": s.date(), "Début corrigé": forced.date(),
                            "Règle": "Arbitrage légal maternité (cas par cas)"})
                    nouveau_start = forced
                    for j in idx:
                        if j == i: continue
                        if episodes.at[j, "end_corr"] >= forced > episodes.at[j, "start_corr"]:
                            episodes.at[j, "end_corr"] = forced - pd.Timedelta(days=1)
                            episodes.at[j, "tronque"] = True
                episodes.at[i, "start_corr"] = nouveau_start
                if nouveau_start != s:
                    episodes.at[i, "tronque"] = True
                borne_blocage = episodes.at[i, "end_corr"]   # correctif v2 : jamais la reprise
                if pd.isna(occupe_jusqua) or borne_blocage > occupe_jusqua:
                    occupe_jusqua = borne_blocage
                    motif_occupant = motif; rang_occupant = rang_i
        absorbes = episodes.loc[episodes["start_corr"] > episodes["end_corr"]].copy()
        for _, r in absorbes.iterrows():
            journal_chevauchement.append({
                "Salarié": r["Nom"], "Matricule": r["Matricule"], "Motif tronqué": r["Motif arrêt"],
                "Début initial": r["start"].date(), "Début corrigé": "— absorbé (0 j) —",
                "Règle": "Épisode entièrement recouvert"})
        episodes = episodes.loc[episodes["start_corr"] <= episodes["end_corr"]].copy()

    # ===============================================================
    # 8. VENTILATION MENSUELLE (lignes positives reconstruites)
    # ===============================================================
    mois_mission = pd.period_range(DATE_DEBUT_PERIODE, DATE_FIN_PERIODE, freq="M")
    def taux(m):
        if "maternité" in m or "paternité" in m: return "100%"
        if "accident" in m: return "60% J1-28 / 80% J29+"
        return "50%"
    pos_lignes = []
    if not episodes.empty:
        for _, ep in episodes.iterrows():
            s, e = ep["start_corr"], ep["end_corr"]
            motif = ep["Motif arrêt"]
            fam = (ep["_emp"], motif)
            a_carence = motif in MOTIFS_CARENCE_3J
            carence_start = s; carence_end = min(s + pd.Timedelta(days=2), e)
            mois_episode = pd.period_range(s.to_period("M"), e.to_period("M"), freq="M")
            cumul_ijss = 0; jours_abs_avant = 0; jours_ijss_avant = 0; epi_lignes = []
            for p in mois_episode:
                j_abs = jours_cal_mois(s, e, p)
                if j_abs == 0: continue
                j_car = jours_cal_mois(carence_start, carence_end, p) if a_carence else 0
                j_ijss = max(0, j_abs - j_car)
                reliquat = max(0, SEUIL_PREVOYANCE - 1 - cumul_ijss)
                j_cpam = min(j_ijss, reliquat); j_prev = j_ijss - j_cpam
                cumul_ijss += j_ijss
                if p in mois_mission:
                    epi_lignes.append({
                        "_fam": fam, "_mois": p,
                        "Nom": ep["Nom"], "Prénom": ep["Prénom"], "NIR": ep["NIR"],
                        "Matricule": ep["Matricule"], "Nom usage": ep["Nom usage"],
                        "Société": ep["Société"], "Statut": ep["Statut"],
                        "Temps de travail": ep["Temps de travail"],
                        "Motif arrêt": motif, "Source juridique": ep["Source juridique"],
                        "DJT": ep["DJT"], "FinPrev": ep["FinPrev_dsn"], "Reprise": ep["Reprise_dsn"],
                        "Mois": str(p), "j_abs": j_abs, "deb_car": carence_start if (a_carence and j_car>0) else pd.NaT,
                        "fin_car": carence_end if (a_carence and j_car>0) else pd.NaT, "j_car": j_car,
                        "j_ijss": j_ijss, "j_cpam": j_cpam, "j_prev": j_prev, "taux": taux(motif),
                        "bascule": ep["start_corr"] + pd.Timedelta(days=SEUIL_PREVOYANCE)})
                else:
                    jours_abs_avant += j_abs; jours_ijss_avant += j_ijss
            abs_m = sum(l["j_abs"] for l in epi_lignes); ijss_m = sum(l["j_ijss"] for l in epi_lignes)
            for l in epi_lignes:
                l["tot_abs_q1"] = abs_m; l["tot_ijss_q1"] = ijss_m
                l["abs_avant"] = jours_abs_avant; l["ijss_avant"] = jours_ijss_avant
                l["tot_abs_epi"] = abs_m + jours_abs_avant; l["tot_ijss_epi"] = ijss_m + jours_ijss_avant
            pos_lignes.extend(epi_lignes)

    # ===============================================================
    # 9. CONTRÔLES
    # ===============================================================
    n_inv = sum(1 for l in pos_lignes if l["j_abs"] != l["j_car"] + l["j_ijss"])
    print(f"   [Invariant jours=carence+IJSS] anomalies : {n_inv}")
    if J is not None and n_inv: J.tracer(ANO.INVARIANT, contexte=f"{n_inv} ligne(s)")
    recouv = 0
    if not episodes.empty:
        for emp, grp in episodes.groupby("_emp"):
            segs = sorted([(r["start_corr"], r["end_corr"], r["Motif arrêt"]) for _, r in grp.iterrows()])
            for a in range(len(segs)):
                for b in range(a + 1, len(segs)):
                    s1, e1, m1 = segs[a]; s2, e2, m2 = segs[b]
                    if s2 <= e1 and m1 != m2: recouv += 1
    print(f"   [Non-recouvrement] résiduels : {recouv}  | [Déchevauchement] journal : {len(journal_chevauchement)}")
    if J is not None and recouv: J.tracer(ANO.RECOUVREMENT, contexte=f"{recouv} paire(s)")
    cats = {}
    for fam in familles_avec_annul:
        cats[interpretation[fam]] = cats.get(interpretation[fam], 0) + 1
    print(f"   [Annulations] familles tracées : {len(familles_avec_annul)} -> {cats}")

    # ----------------------------------------------------------
    # SOLDE RETENU & TEXTES D'INTERPRÉTATION (depuis la reconstruction, pas la source brute)
    # ----------------------------------------------------------
    ret_fam = {}
    for l in pos_lignes:
        ret_fam[l["_fam"]] = ret_fam.get(l["_fam"], 0) + l["j_abs"]
    interp_texte = {}
    solde_retenu = {}
    for fam in familles_avec_annul:
        cat = interpretation[fam]
        n = fam_meta[fam]["n_annul"]; dead = fam_meta[fam]["dead_decl"]
        ret = ret_fam.get(fam, 0)
        solde_retenu[fam] = ret
        if cat == "PURE":
            interp_texte[fam] = (f"🔴 Annulation pure ({n} annul.) — aucune déclaration positive ne "
                f"subsiste sur le trimestre. Vérifier RH/paie : suppression légitime ou arrêt réel "
                f"non re-déclaré → IJSS potentiellement à récupérer.")
        elif cat == "REDUITE":
            interp_texte[fam] = (f"🟠 Compensée à durée réduite ({dead}j→{ret}j) — arrêt initial annulé "
                f"puis re-déclaré à durée/début différents ; {ret}j retenus (arbitre = paie, rubrique "
                f"3480). Pas de perte CPAM, vérifier reprise anticipée RH.")
        elif cat == "AVERIF":
            interp_texte[fam] = (f"🟠 Corrections multiples ({n} annul.) — épisode réel de {ret}j retenus ; "
                f"certains mois annulés sans re-déclaration claire dans le trimestre : vérifier qu'ils "
                f"sont re-signalés ailleurs. Pas de perte CPAM si arrêt continu.")
        else:
            interp_texte[fam] = (f"✅ Corrections DSN compensées ({n} annul.) — épisode réel de {ret}j "
                f"retenus, chaque mois annulé est re-déclaré. Aucune perte, aucune action.")

    # ===============================================================
    # 10. CONSTRUCTION DE L'ONGLET UNIQUE « Détail épisodes »
    # ===============================================================
    CODE_DSN = [
        ("temps partiel thérapeutique", "09"), ("accident de travail", "02"),
        ("accident de trajet", "03"), ("maladie professionnelle", "06"),
        ("maternité", "04"), ("paternité", "05"), ("adoption", "07"),
        ("maladie", "01"),
    ]
    def code_dsn(motif):
        m = str(motif).lower()
        for kw, c in CODE_DSN:
            if kw in m: return c
        return "—"
    def statut_prev(motif, j_prev):
        m = str(motif).lower()
        if "maternité" in m or "paternité" in m or "accident" in m:
            return "✅ SS seule (illimité)"
        return "⚠️ RELAI PRÉVOYANCE actif (J91+)" if j_prev > 0 else "✅ SS seule"

    def fdate(d):
        d = pd.to_datetime(d, errors="coerce")
        return "" if pd.isna(d) else d.strftime("%d/%m/%Y")

    # regrouper lignes positives par famille
    pos_by_fam = {}
    for l in pos_lignes:
        pos_by_fam.setdefault(l["_fam"], []).append(l)
    # annulations par famille
    annul_raw = annul_raw.assign(_fam=list(zip(annul_raw["_emp"], annul_raw["Motif épisode"])))
    ann_by_fam = {}
    for _, r in annul_raw.iterrows():
        ann_by_fam.setdefault(r["_fam"], []).append(r)

    # ordre des familles : par Nom, Prénom, motif
    def fam_label(fam):
        if fam in pos_by_fam:
            x = pos_by_fam[fam][0]; return (x["Nom"], x["Prénom"], fam[1])
        r = ann_by_fam[fam][0]; return (r["Nom"], r["Prénom"], fam[1])
    familles = sorted(set(list(pos_by_fam) + list(ann_by_fam)), key=fam_label)

    rows = []
    n_pos_total = 0; n_ann_total = 0
    for fam in familles:
        interp = interp_texte.get(fam, "RAS — aucune annulation")
        solde = solde_retenu.get(fam, sum(l["j_abs"] for l in pos_by_fam.get(fam, [])))
        first_line_done = False
        # --- lignes positives (par mois) ---
        for l in sorted(pos_by_fam.get(fam, []), key=lambda z: z["Mois"]):
            n_pos_total += 1
            rows.append({
                "Nom": l["Nom"], "Prénom": l["Prénom"], "NIR": l["NIR"], "Matricule": l["Matricule"],
                "Nom usage": l["Nom usage"], "Société": l["Société"], "Statut": l["Statut"],
                "Temps de travail": l["Temps de travail"], "Type de ligne": "📍 Données DSN",
                "Motif arrêt": l["Motif arrêt"], "Code DSN": code_dsn(l["Motif arrêt"]),
                "Source juridique": l["Source juridique"], "DJT": fdate(l["DJT"]),
                "Date fin prévisionnelle": fdate(l["FinPrev"]), "Date reprise": fdate(l["Reprise"]),
                "Mois DSN": l["Mois"], "Jours absence (mois)": l["j_abs"],
                "Début carence": fdate(l["deb_car"]) or "—", "Fin carence": fdate(l["fin_car"]) or "—",
                "Jours carence (mois)": l["j_car"], "Jours IJSS SS (mois)": l["j_ijss"],
                "Jours SS CPAM (J1-J90)": l["j_cpam"], "Jours Prévoyance J91+": l["j_prev"],
                "Total j abs Q1-2026": l["tot_abs_q1"], "Total IJSS Q1-2026": l["tot_ijss_q1"],
                "J avant Q1-2026": l["abs_avant"], "IJSS avant Q1-2026": l["ijss_avant"],
                "⭐ Total épisode complet": l["tot_abs_epi"], "⭐ Total IJSS épisode complet": l["tot_ijss_epi"],
                "Date début prévoyance": fdate(l["bascule"]),
                "Solde net épisode (pos+neg)": solde,
                "Statut prévoyance": statut_prev(l["Motif arrêt"], l["j_prev"]),
                "🔍 Interprétation épisode (auto)": ("" if first_line_done else interp),
            })
            first_line_done = True
        # --- lignes d'annulation (par mois) ---
        ann_sorted = sorted(ann_by_fam.get(fam, []), key=lambda z: str(z["mois_p"]))
        for k, r in enumerate(ann_sorted):
            n_ann_total += 1
            # si aucune ligne positive (famille pure), porter l'interprétation ici
            interp_cell = ""
            if not first_line_done and k == 0:
                interp_cell = interp
                first_line_done = True
            rows.append({
                "Nom": r["Nom"], "Prénom": r["Prénom"], "NIR": r["NIR"], "Matricule": r["Matricule"],
                "Nom usage": r["Nom usage"], "Société": r["Société"], "Statut": r["Statut"],
                "Temps de travail": r["Temps de travail"], "Type de ligne": "❌ Annulation DSN",
                "Motif arrêt": r["Motif arrêt"], "Code DSN": "00",
                "Source juridique": "", "DJT": fdate(r["DJT"]),
                "Date fin prévisionnelle": "", "Date reprise": "",
                "Mois DSN": str(r["mois_p"]), "Jours absence (mois)": int(r["Jours src"]) if pd.notna(r["Jours src"]) else "",
                "Début carence": "", "Fin carence": "", "Jours carence (mois)": "",
                "Jours IJSS SS (mois)": 0, "Jours SS CPAM (J1-J90)": "", "Jours Prévoyance J91+": "",
                "Total j abs Q1-2026": "", "Total IJSS Q1-2026": "", "J avant Q1-2026": "",
                "IJSS avant Q1-2026": "", "⭐ Total épisode complet": "", "⭐ Total IJSS épisode complet": "",
                "Date début prévoyance": "", "Solde net épisode (pos+neg)": solde,
                "Statut prévoyance": "", "🔍 Interprétation épisode (auto)": interp_cell,
            })
    detail = pd.DataFrame(rows)
    return {"societe": SOCIETE_CIBLE, "detail": detail, "journal": journal_chevauchement,
            "n_pos": n_pos_total, "n_ann": n_ann_total,
            "cats": cats, "n_familles": len(familles_avec_annul)}


