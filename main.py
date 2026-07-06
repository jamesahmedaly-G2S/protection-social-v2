# -*- coding: utf-8 -*-
"""
main.py — Orchestrateur de la reconstruction IJSS
=================================================
Enchaîne : chargement DSN → mapping → normalisation/détection → traitement par
société → écriture Excel.

Trois traces sont produites à chaque exécution (toutes conservées, horodatées) :
  - output/logs/execution/execution_<ts>.log  : déroulé technique pas à pas + erreurs ;
  - output/logs/anomalies/journal_IJSS_<ts>.log : anomalies métier détectées ;
  - output/rapports/rapport_executions.(md|xlsx) : rapport cumulatif (append).

Lancement :
    python main.py
"""
import os
import sys
from datetime import datetime

import config
from src.logger_execution import configurer_logger, etape
from src.journal_audit import JournalAudit, ANO
from src.chargement import (lire_source, mapper_colonnes, colonnes_manquantes,
                            normaliser_et_detecter)
from src.pipeline import traiter_societe
from src.export_excel import ecrire_xlsx
from src.rapport_execution import RapportExecution


def resoudre_source():
    """Cherche le fichier DSN dans data/input/, sinon à côté du projet."""
    candidat = os.path.join(config.INPUT_DIR, config.INPUT_FILE)
    if os.path.exists(candidat):
        return candidat
    if os.path.exists(config.INPUT_FILE):
        return config.INPUT_FILE
    return candidat  # renvoyé tel quel pour le message d'erreur


def main():
    # --- dossiers ---
    for d in (config.OUTPUT_DIR, config.LOG_EXEC_DIR, config.LOG_ANO_DIR, config.RAPPORT_DIR):
        os.makedirs(d, exist_ok=True)

    # --- horodatage partagé par toutes les traces de ce run ---
    horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
    logger, chemin_exec = configurer_logger(config.LOG_EXEC_DIR, horodatage)
    chemin_ano = os.path.join(config.LOG_ANO_DIR, f"journal_IJSS_{horodatage}.log")
    J = JournalAudit(chemin_log=chemin_ano)

    R = RapportExecution(
        md_path=os.path.join(config.RAPPORT_DIR, "rapport_executions.md"),
        xlsx_path=os.path.join(config.RAPPORT_DIR, "rapport_executions.xlsx"),
        source=resoudre_source())
    R.demarrer()   # capture la sortie terminal

    fichiers_produits = []
    logger.info("=" * 70)
    logger.info(f"DÉBUT D'EXÉCUTION — run {horodatage}")
    logger.info(f"Journal anomalies : {chemin_ano}")

    try:
        # 1. Chargement
        chemin = resoudre_source()
        etape(logger, 1, f"Chargement de la source : {chemin}")
        if not os.path.exists(chemin):
            logger.error(f"Fichier source introuvable : {chemin}")
            J.tracer(ANO.FICHIER_SOURCE, contexte=chemin)
            raise FileNotFoundError(f"Fichier source introuvable : {chemin}")
        dsn_full = lire_source(chemin)
        dsn_full.columns = dsn_full.columns.str.strip()
        logger.info(f"Fichier chargé : {len(dsn_full)} lignes, {len(dsn_full.columns)} colonnes.")
        print("✅ Fichier chargé.")

        # 2. Mapping + contrôle des colonnes obligatoires
        etape(logger, 2, "Mapping des colonnes DSN.")
        cols = mapper_colonnes(dsn_full)
        manquant = colonnes_manquantes(cols)
        if manquant:
            logger.error(f"Colonnes DSN obligatoires manquantes : {manquant}")
            J.tracer(ANO.COLONNE_MANQUANTE, contexte=str(manquant))
            raise KeyError(f"Colonnes DSN introuvables : {manquant}")
        logger.info("Mapping OK — toutes les colonnes obligatoires sont présentes.")

        # 3. Normalisation + détection des annulations
        etape(logger, 3, "Normalisation et détection des annulations.")
        dsn_full = normaliser_et_detecter(dsn_full, cols)
        societes = sorted(dsn_full["_soc"].loc[dsn_full["_soc"] != ""].unique())
        n_annul = int(dsn_full["_annul"].sum())
        logger.info(f"Sociétés présentes : {societes}")
        logger.info(f"Annulations détectées (union 3 signaux) : {n_annul} lignes.")
        print(f"ℹ️  Sociétés présentes : {societes}")
        print(f"ℹ️  Annulations détectées (union 3 signaux) : {n_annul} lignes")
        J.tracer(ANO.ANNULATION_DETECTEE, contexte=f"{n_annul} ligne(s) sur l'ensemble")

        # 4. Traitement par société + écriture
        etape(logger, 4, f"Traitement des {len(config.SOCIETES_CIBLES)} sociétés cibles.")
        for soc in config.SOCIETES_CIBLES:
            logger.info(f"— Société '{soc}' : début du traitement.")
            res = traiter_societe(soc, dsn_full, J=J)
            if res is None or res["detail"].empty:
                logger.warning(f"— Société '{soc}' : aucune donnée, ignorée.")
                continue
            fichier = ecrire_xlsx(res, J=J)
            R.ajouter_societe(res, fichier or "")
            if fichier:
                fichiers_produits.append(fichier)
            logger.info(f"— Société '{soc}' : {res['n_pos']} positives + {res['n_ann']} "
                        f"annulations → {fichier}")

        # 5. Bilan
        etape(logger, 5, "Bilan de l'exécution.")
        resume_ano = J.resume()
        logger.info(f"Fichiers produits ({len(fichiers_produits)}) :")
        for f in fichiers_produits:
            logger.info(f"    • {f}")
        logger.info(f"Anomalies par code : {resume_ano}")
        logger.info(f"FIN D'EXÉCUTION — run {horodatage} — statut : SUCCÈS")
        print(f"\n✅ Terminé. Journal d'exécution : {chemin_exec}")

    except Exception as exc:
        # Toute erreur d'exécution est journalisée avec horodatage + traceback complet.
        logger.exception(f"ERREUR D'EXÉCUTION : {exc}")
        logger.error(f"FIN D'EXÉCUTION — run {horodatage} — statut : ÉCHEC")
        print(f"❌ Erreur : {exc} (détail dans {chemin_exec})")

    finally:
        # Les traces sont clôturées quoi qu'il arrive (succès comme échec).
        J.cloturer()
        R.finaliser(journal=J)
        for h in list(logger.handlers):
            h.flush()


if __name__ == "__main__":
    main()
