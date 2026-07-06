# -*- coding: utf-8 -*-
"""
logger_execution.py — Journal d'exécution technique (déroulé pas à pas)
=======================================================================
Trace, pour CHAQUE exécution, dans un fichier horodaté et conservé :
  - le début et la fin du run ;
  - chaque étape réalisée (chargement, mapping, normalisation, traitement par
    société, écriture) avec horodatage à la seconde ;
  - les fichiers générés (avec leur chemin absolu) ;
  - un rappel des anomalies ;
  - toute ERREUR d'exécution (exception) avec son traceback complet, horodatée,
    pour faciliter le diagnostic.

Ce journal est complémentaire du journal d'anomalies (journal_IJSS_*.log) :
  - execution_*.log  → déroulé technique + erreurs (ce module) ;
  - journal_IJSS_*.log → anomalies métier détectées (JournalAudit).

Les deux partagent le même horodatage de run, pour les retrouver ensemble.
"""
import logging
import os


def configurer_logger(log_dir, horodatage):
    """Crée un logger fichier (execution_<horodatage>.log) + console (warnings/erreurs).
    Renvoie (logger, chemin_fichier)."""
    os.makedirs(log_dir, exist_ok=True)
    chemin = os.path.join(log_dir, f"execution_{horodatage}.log")

    logger = logging.getLogger("execution_ijss")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S")
    fh = logging.FileHandler(chemin, mode="w", encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    # Console : seulement warnings/erreurs (les étapes détaillées restent dans le fichier ;
    # la progression lisible passe déjà par les print du pipeline).
    ch = logging.StreamHandler()
    ch.setLevel(logging.WARNING)
    ch.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
    logger.addHandler(ch)

    return logger, chemin


def etape(logger, numero, libelle):
    """Journalise une étape numérotée du pipeline."""
    logger.info(f"[Étape {numero}] {libelle}")
