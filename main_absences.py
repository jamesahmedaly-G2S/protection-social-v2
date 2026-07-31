# -*- coding: utf-8 -*-
"""
main_absences.py — Comparaison des jours d'absence PAIE <-> DSN
=============================================================================
Compare, par société / salarié / mois / motif d'arrêt, le nombre de jours d'absence
côté DSN et côté PAIE — les deux lus depuis les classeurs déjà RECONSTRUITS
(output/xlsx/, produits par main.py et main_paie.py), jamais recalculés depuis les
fichiers source. Les reconstruits sont le livrable final : ils sont censés avoir déjà
corrigé les anomalies identifiées, donc toute analyse en aval doit en partir plutôt que
de recalculer indépendamment depuis la donnée brute (cf. échanges du 2026-07-31).

Les deux reconstruits sont désormais restreints à la même période — cf.
config.DATE_DEBUT_PERIODE/DATE_FIN_PERIODE, appliqué à la fois par main.py et par
main_paie.py, actuellement l'année 2026 complète (était : T1 seul) : pas de filtrage de
mois à refaire ici, les données lues sont déjà sur le bon périmètre. Le fichier source DSN
contient réellement des déclarations jusqu'en septembre 2026 (pas seulement Q1, malgré son
nom de fichier) : l'extension à l'année complète a donc un effet réel des deux côtés.

Produit deux livrables distincts (+ un journal) :

  - LIVRABLE "jours d'absence côté paye" (réunion Solange du 2026-07-31, point 1 des 3
    livrables demandés — miroir du fichier DSN déjà produit par main.py) :
      output/xlsx/Jours_Absence_PAIE_<code>.xlsx, un par société. Une ligne par
      salarié/mois/motif, lu depuis Reconstruit_PAIE_<code>.xlsx (donc la période
      DATE_DEBUT_PERIODE/DATE_FIN_PERIODE en vigueur).

  - COMPARATIF "jours d'absence PAIE <-> DSN" (livrable 3 de la même réunion —
    rapprochement DSN/paye via le pivot salarié) :
      output/rapports/Comparatif_jours_absence_PAIE_DSN.xlsx :
        "1 - Détail mensuel"       : une ligne par salarié / mois / motif ;
        "2 - Total Année 2026"     : une ligne par salarié / motif, sommée sur l'année 2026 ;
        "3 - Anomalies (jours)"    : lignes PAIE_AUDIT où "Base" est physiquement incohérente
          pour un mois (> 31 jours, ou négative hors annulation) — signalement, pas une
          correction automatique (aucune info exploitable en source pour arbitrer ; ce
          contrôle reste sur le CSV source, car c'est justement un diagnostic de qualité
          de la donnée brute, en amont du reconstruit).
        "4 - Synthèse OK par mois" : un tableau par société — pour chaque mois, combien de
          salariés communs DSN/PAIE ont un total de jours d'absence identique (OK) et
          combien ont un écart (Pas OK), avec le % OK (cf. réunion du 2026-07-31 : "pour
          chaque société et chaque mois, combien de salariés sont OK / pas OK").
        "5 - Traçabilité écarts"   : détail mensuel (grain motif) restreint aux salariés
          "Pas OK" d'un mois donné, avec le total DSN/PAIE/écart du mois — pour permettre
          la vérification / correction / arbitrage salarié par salarié.
      Colonnes (onglets 1/2/5) : Clé salarié (pivot), Nom, Prénom, Nom d'usage, NIR,
      Matricule, Entreprise, Etablissement, Siren, Nic, Siret, Motif, Nombre de jours
      absence PAIE/DSN, Mois absence (période), Écart constaté — avec filtre Excel
      cliquable sur chaque colonne. Siren/Nic/Siret sont recoupés avec la source DSN
      (toujours à 0 dans PAIE_AUDIT.csv).

  - output/logs/execution/execution_absences_<ts>.log : déroulé + erreurs.
  - output/logs/anomalies/ecarts_paie_dsn_<ts>.log : synthèse OK/Pas OK par société/mois
    + traçabilité salarié par salarié des écarts détectés (même contenu que les onglets
    "4"/"5", au format log pour un suivi / archivage indépendant du classeur Excel).

Lancement :
    python main_absences.py
"""
import os
from datetime import datetime

import config
from src.logger_execution import configurer_logger, etape
from src.comparaison_absences import (nom_fichier_dsn, nom_fichier_paie, charger_jours_dsn,
                                      charger_jours_paie, lire_identite_dsn_source, comparer_jours,
                                      comparer_jours_periode, detecter_anomalies_jours,
                                      construire_synthese_ok, ecrire_comparatif_absences,
                                      construire_absences_paie, ecrire_absences_paie)


def resoudre(nom_fichier):
    candidat = os.path.join(config.INPUT_DIR, nom_fichier)
    if os.path.exists(candidat):
        return candidat
    if os.path.exists(nom_fichier):
        return nom_fichier
    return candidat


def main():
    os.makedirs(config.RAPPORT_DIR, exist_ok=True)
    horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
    logger, chemin_exec = configurer_logger(config.LOG_EXEC_DIR, horodatage,
                                            prefix="execution_absences", logger_name="execution_absences")
    logger.info("=" * 70)
    logger.info(f"DÉBUT D'EXÉCUTION COMPARATIF ABSENCES — run {horodatage}")
    print(f"=== Comparaison jours d'absence PAIE <-> DSN — run {horodatage} ===")

    mapping = config.MAPPING_PAIE_DSN
    try:
        # 1. Vérification des classeurs reconstruits (DSN + PAIE) — on part d'eux, pas des sources
        etape(logger, 1, "Vérification des classeurs reconstruits (DSN + PAIE).")
        manquants = []
        for code_paie, societe_dsn in mapping.items():
            for chemin in (os.path.join(config.OUTPUT_DIR, nom_fichier_dsn(societe_dsn)),
                          os.path.join(config.OUTPUT_DIR, nom_fichier_paie(code_paie))):
                if not os.path.exists(chemin):
                    manquants.append(chemin)
        if manquants:
            logger.error(f"Classeurs manquants : {manquants}")
            raise FileNotFoundError(
                "Classeurs reconstruits manquants (lancer main.py et main_paie.py "
                f"au préalable) : {manquants}")

        # 2. Chargement des jours DSN (depuis les reconstruits)
        etape(logger, 2, "Chargement des jours d'absence DSN (classeurs reconstruits).")
        jours_dsn = charger_jours_dsn(mapping)
        for soc, df in jours_dsn.items():
            logger.info(f"— DSN {soc} : {len(df)} ligne(s) (matricule/mois/motif).")

        # 3. Chargement des jours PAIE (depuis les reconstruits, pas depuis PAIE_AUDIT.csv)
        etape(logger, 3, "Chargement des jours d'absence PAIE (classeurs Reconstruit_PAIE_<code>.xlsx).")
        jours_paie = charger_jours_paie(mapping)
        for code, df in jours_paie.items():
            logger.info(f"— PAIE {code} : {len(df)} ligne(s) (matricule/mois/motif).")

        # 4. Identité Siren/Nic/Siret (source DSN — toujours à 0 côté PAIE_AUDIT.csv)
        chemin_dsn_source = resoudre(config.INPUT_FILE)
        etape(logger, 4, f"Chargement de l'identité Siren/Nic/Siret depuis {chemin_dsn_source}.")
        identite_dsn = lire_identite_dsn_source(chemin_dsn_source) if os.path.exists(chemin_dsn_source) else {}
        logger.info(f"Identité DSN chargée pour {len(identite_dsn)} matricule(s).")

        # 5. Comparaison (les deux côtés sont déjà sur le même périmètre, cf. docstring)
        etape(logger, 5, "Comparaison DSN <-> PAIE (jointure matricule/mois/motif).")
        detail = comparer_jours(mapping, jours_dsn, jours_paie, identite_dsn=identite_dsn)
        mois_annee = [f"{config.ANNEE_PAIE}-{m:02d}" for m in range(1, 13)]
        libelle_annee = f"Année {config.ANNEE_PAIE}"
        detail_annee = comparer_jours_periode(detail, mois_annee, libelle_periode=libelle_annee)

        synthese = (detail_annee.groupby(["Entreprise", "Motif"], as_index=False)
                             [["Nombre de jours absence DSN", "Nombre de jours absence PAIE",
                               "Écart constaté"]].sum())
        for _, r in synthese.iterrows():
            logger.info(f"[{r['Entreprise']}] {r['Motif']} : "
                        f"DSN={r['Nombre de jours absence DSN']:.1f} j | "
                        f"PAIE={r['Nombre de jours absence PAIE']:.1f} j | "
                        f"écart={r['Écart constaté']:.1f} j")
            print(f"ℹ️  [{r['Entreprise']}] {r['Motif']} : "
                  f"DSN={r['Nombre de jours absence DSN']:.1f} j vs "
                  f"PAIE={r['Nombre de jours absence PAIE']:.1f} j "
                  f"(écart {r['Écart constaté']:.1f})")

        # 6. Synthèse OK/Pas OK par société et par mois + traçabilité des écarts salarié/mois
        etape(logger, 6, "Synthèse OK/Pas OK par société et par mois (population des salariés communs).")
        synthese_ok, tracabilite_ecarts = construire_synthese_ok(mapping, jours_dsn, jours_paie, detail)
        for _, r in synthese_ok.iterrows():
            logger.info(f"[{r['Entreprise']} / {r['Société DSN']}] {r['Mois']} : "
                        f"{r['OK']}/{r['Salariés communs']} OK, {r['Pas OK']} écart(s) "
                        f"({r['% OK']:.1%} OK).")
        print(f"ℹ️  Synthèse OK/Pas OK : {len(tracabilite_ecarts[['Entreprise', 'Matricule', 'Mois absence (période)']].drop_duplicates())} "
              f"salarié(s)/mois en écart sur {len(synthese_ok)} ligne(s) société/mois.")

        # 6bis. Log dédié "ecarts_paie_dsn" (traçabilité salarié par salarié, pour vérif/correction/arbitrage)
        logger_ecarts, chemin_ecarts = configurer_logger(config.LOG_ANO_DIR, horodatage,
                                                          prefix="ecarts_paie_dsn",
                                                          logger_name="ecarts_paie_dsn")
        logger_ecarts.info("=" * 70)
        logger_ecarts.info(f"SYNTHÈSE OK/PAS OK PAR SOCIÉTÉ ET PAR MOIS — run {horodatage}")
        for _, r in synthese_ok.iterrows():
            logger_ecarts.info(f"[{r['Entreprise']} / {r['Société DSN']}] {r['Mois']} : "
                               f"{r['OK']}/{r['Salariés communs']} OK, {r['Pas OK']} écart(s) "
                               f"({r['% OK']:.1%} OK).")
        logger_ecarts.info("-" * 70)
        logger_ecarts.info("TRAÇABILITÉ DES SALARIÉS EN ÉCART (un par mois, pour vérification/correction/arbitrage)")
        vus = set()
        for _, r in tracabilite_ecarts.iterrows():
            cle = (r["Entreprise"], r["Matricule"], r["Mois absence (période)"])
            if cle in vus:
                continue
            vus.add(cle)
            logger_ecarts.warning(
                f"[{r['Entreprise']}] Matricule {r['Matricule']} ({r['Nom']} {r['Prénom']}) — "
                f"{r['Mois absence (période)']} : DSN={r['Total jours DSN (mois)']:.1f} j, "
                f"PAIE={r['Total jours PAIE (mois)']:.1f} j, écart={r['Écart total (mois)']:.1f} j.")
        for h in list(logger_ecarts.handlers):
            h.flush()
        logger.info(f"Log écarts_paie_dsn écrit : {chemin_ecarts} ({len(vus)} salarié(s)/mois en écart).")
        print(f"⚠️  {len(vus)} salarié(s)/mois en écart -> détail dans {chemin_ecarts}")

        # 7. Détection des anomalies (diagnostic sur le CSV source, en amont du reconstruit)
        etape(logger, 7, "Détection des anomalies (Base > 31j ou négative hors annulation).")
        chemin_paie_source = resoudre(config.PAIE_INPUT_FILE)
        anomalies = detecter_anomalies_jours(chemin_paie_source, annee=config.ANNEE_PAIE, sep=config.PAIE_SEP)
        logger.info(f"{len(anomalies)} anomalie(s) détectée(s).")
        print(f"⚠️  {len(anomalies)} ligne(s) anormale(s) (jours incohérents pour un mois).")

        # 8. LIVRABLE "jours d'absence côté paye" (miroir autonome du fichier DSN)
        etape(logger, 8, "Écriture du LIVRABLE jours d'absence PAIE (Jours_Absence_PAIE_<code>.xlsx).")
        absences_paie = construire_absences_paie(jours_paie, identite_dsn=identite_dsn)
        fichiers_absences = ecrire_absences_paie(absences_paie)
        for f in fichiers_absences:
            logger.info(f"Livrable jours d'absence PAIE écrit : {f}")

        # 9. COMPARATIF "jours d'absence PAIE <-> DSN"
        etape(logger, 9, "Écriture du COMPARATIF PAIE <-> DSN (Comparatif_jours_absence_PAIE_DSN.xlsx).")
        fichier = ecrire_comparatif_absences(detail, detail_annee, anomalies, synthese_ok, tracabilite_ecarts,
                                             libelle_periode=libelle_annee)
        logger.info(f"Comparatif écrit : {fichier}")

        logger.info(f"FIN D'EXÉCUTION COMPARATIF ABSENCES — run {horodatage} — statut : SUCCÈS")
        print(f"\n✅ LIVRABLE jours d'absence PAIE — {len(fichiers_absences)} fichier(s) :")
        for f in fichiers_absences:
            print(f"   • {f}")
        print(f"✅ COMPARATIF PAIE <-> DSN : {fichier}")
        print(f"Journal d'exécution : {chemin_exec}")

    except Exception as exc:
        logger.exception(f"ERREUR D'EXÉCUTION COMPARATIF ABSENCES : {exc}")
        logger.error(f"FIN D'EXÉCUTION COMPARATIF ABSENCES — run {horodatage} — statut : ÉCHEC")
        print(f"❌ Erreur : {exc} (détail dans {chemin_exec})")

    finally:
        for h in list(logger.handlers):
            h.flush()


if __name__ == "__main__":
    main()
