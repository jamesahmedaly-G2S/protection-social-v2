# -*- coding: utf-8 -*-
"""
main_comparaison.py — Comparaison des populations de salariés PAIE <-> DSN
=============================================================================
Compare, société par société (cf. config.MAPPING_PAIE_DSN), les salariés du PAIE à
ceux du DSN, à quatre niveaux :
  1. reconstruits    : classeurs déjà reconstruits (output/xlsx/) — nécessite d'avoir
     lancé main.py (DSN) et main_paie.py (PAIE) au préalable ;
  2. sources         : fichiers sources bruts (CSV DSN, PAIE_AUDIT.csv), année complète —
     vérifie que les pipelines de reconstruction ne perdent aucun salarié en route ;
  3. période reconstr. : DSN reconstruit vs PAIE (source) restreinte à la période
     DATE_DEBUT_PERIODE/DATE_FIN_PERIODE (cf. config.py — même fenêtre que le DSN) ;
  4. période sources   : DSN source brute vs PAIE (source) restreinte à la même période.

DATE_DEBUT_PERIODE/DATE_FIN_PERIODE couvre actuellement l'année 2026 complète (était :
T1 seul) — les niveaux 1/3 et 2/4 sont alors quasi équivalents, le reconstruit PAIE et la
source PAIE couvrant désormais la même fenêtre. Le fichier source DSN contient réellement
des déclarations jusqu'en septembre 2026 (pas seulement Q1, malgré son nom de fichier),
donc le DSN profite lui aussi de l'extension de période, dans la limite de ce que couvre
sa source actuelle (jusqu'à septembre).

Produit :
  - output/rapports/Comparatif_populations_PAIE_DSN.xlsx (8 onglets, 2 par niveau :
    "n - Comparatif (...)" et "n+1 - Détail (...)") ;
  - output/logs/execution/execution_comparatif_<ts>.log : déroulé + erreurs.

Lancement :
    python main_comparaison.py
"""
import os
from datetime import datetime

import config
from src.logger_execution import configurer_logger, etape
from src.comparaison_populations import (comparer, charger_dsn_reconstruits, charger_paie_reconstruits,
                                         lire_population_dsn_source, lire_population_paie_source,
                                         nom_fichier_dsn, nom_fichier_paie)
from src.export_comparatif import ecrire_comparatif


def resoudre(nom_fichier):
    candidat = os.path.join(config.INPUT_DIR, nom_fichier)
    if os.path.exists(candidat):
        return candidat
    if os.path.exists(nom_fichier):
        return nom_fichier
    return candidat


def _log_synthese(logger, prefixe, synthese):
    for ligne in synthese:
        logger.info(f"[{prefixe}] {ligne['Société DSN']} / {ligne['Société PAIE']} : "
                    f"{ligne['Salariés communs']}/{ligne['Salariés DSN']} salariés DSN "
                    f"retrouvés dans PAIE.")
        print(f"ℹ️  [{prefixe}] {ligne['Société DSN']} / {ligne['Société PAIE']} : "
              f"{ligne['Salariés communs']}/{ligne['Salariés DSN']} retrouvés.")


def main():
    os.makedirs(config.RAPPORT_DIR, exist_ok=True)
    horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
    date_sortie = datetime.now().strftime("%d-%m-%Y")  # date de génération, ajoutée au nom du livrable
    logger, chemin_exec = configurer_logger(config.LOG_EXEC_DIR, horodatage,
                                            prefix="execution_comparatif",
                                            logger_name="execution_comparatif")
    logger.info("=" * 70)
    logger.info(f"DÉBUT D'EXÉCUTION COMPARATIF PAIE/DSN — run {horodatage}")
    print(f"=== Comparaison populations PAIE <-> DSN — run {horodatage} ===")

    mapping = config.MAPPING_PAIE_DSN
    try:
        # 1. Vérification + comparaison niveau "fichiers reconstruits"
        etape(logger, 1, "Vérification des classeurs reconstruits (PAIE + DSN).")
        manquants = []
        for code_paie, societe_dsn in mapping.items():
            for chemin in (os.path.join(config.OUTPUT_DIR, nom_fichier_paie(code_paie)),
                          os.path.join(config.OUTPUT_DIR, nom_fichier_dsn(societe_dsn))):
                if not os.path.exists(chemin):
                    manquants.append(chemin)
        if manquants:
            logger.error(f"Classeurs manquants : {manquants}")
            raise FileNotFoundError(
                "Classeurs reconstruits manquants (lancer main.py et main_paie.py "
                f"au préalable) : {manquants}")

        etape(logger, 2, "Comparaison niveau fichiers reconstruits.")
        dsn_reco_pop = charger_dsn_reconstruits(mapping)
        paie_reco_pop = charger_paie_reconstruits(mapping)
        synth_reco, detail_reco = comparer(mapping, dsn_reco_pop, paie_reco_pop)
        _log_synthese(logger, "Reconstruits", synth_reco)

        # 2. Comparaison niveau "fichiers sources bruts" (année complète)
        chemin_dsn_source = resoudre(config.INPUT_FILE)
        chemin_paie_source = resoudre(config.PAIE_INPUT_FILE)
        etape(logger, 3, f"Chargement des sources brutes : {chemin_dsn_source} / {chemin_paie_source}")
        if not os.path.exists(chemin_dsn_source) or not os.path.exists(chemin_paie_source):
            logger.error("Fichier(s) source(s) introuvable(s) pour la comparaison sources brutes.")
            raise FileNotFoundError(
                f"Source(s) introuvable(s) : {chemin_dsn_source} / {chemin_paie_source}")
        dsn_source_pop = lire_population_dsn_source(chemin_dsn_source)
        paie_source_pop = lire_population_paie_source(chemin_paie_source, annee=config.ANNEE_PAIE,
                                                       sep=config.PAIE_SEP)

        etape(logger, 4, "Comparaison niveau fichiers sources bruts (année complète).")
        synth_src, detail_src = comparer(mapping, dsn_source_pop, paie_source_pop)
        _log_synthese(logger, "Sources", synth_src)

        # 3. Comparaison à isopérimètre (PAIE source restreinte à la période DSN)
        etape(logger, 5, f"Population PAIE restreinte à la période ({config.DATE_DEBUT_PERIODE.date()} "
                         f"-> {config.DATE_FIN_PERIODE.date()}).")
        paie_periode_pop = lire_population_paie_source(chemin_paie_source, annee=config.ANNEE_PAIE,
                                                        sep=config.PAIE_SEP,
                                                        date_debut=config.DATE_DEBUT_PERIODE,
                                                        date_fin=config.DATE_FIN_PERIODE)

        etape(logger, 6, "Comparaison période : DSN reconstruit vs PAIE (source, période restreinte).")
        synth_periode_reco, detail_periode_reco = comparer(mapping, dsn_reco_pop, paie_periode_pop)
        _log_synthese(logger, "Période/Reconstruits", synth_periode_reco)

        etape(logger, 7, "Comparaison période : DSN source vs PAIE (source, période restreinte).")
        synth_periode_src, detail_periode_src = comparer(mapping, dsn_source_pop, paie_periode_pop)
        _log_synthese(logger, "Période/Sources", synth_periode_src)

        # 4. Écriture du classeur
        etape(logger, 8, "Écriture du classeur de comparaison.")
        fichier = ecrire_comparatif([
            ("reconstruits", synth_reco, detail_reco),
            ("sources", synth_src, detail_src),
            # Noms courts : limite Excel de 31 caractères par nom d'onglet (cf. "n - Comparatif (titre)").
            ("pér. reconstr.", synth_periode_reco, detail_periode_reco),
            ("pér. sources", synth_periode_src, detail_periode_src),
        ], date_sortie)
        logger.info(f"Classeur écrit : {fichier}")

        logger.info(f"FIN D'EXÉCUTION COMPARATIF PAIE/DSN — run {horodatage} — statut : SUCCÈS")
        print(f"\n✅ Terminé. Classeur : {fichier}")
        print(f"Journal d'exécution : {chemin_exec}")

    except Exception as exc:
        logger.exception(f"ERREUR D'EXÉCUTION COMPARATIF : {exc}")
        logger.error(f"FIN D'EXÉCUTION COMPARATIF PAIE/DSN — run {horodatage} — statut : ÉCHEC")
        print(f"❌ Erreur : {exc} (détail dans {chemin_exec})")

    finally:
        for h in list(logger.handlers):
            h.flush()


if __name__ == "__main__":
    main()
