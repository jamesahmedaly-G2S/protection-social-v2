# -*- coding: utf-8 -*-
"""
main_paie.py — Orchestrateur de la reconstruction IJSS "version paie" (étape 1 : structure)
=============================================================================================
Source : data/input/PAIE_AUDIT.csv (grain rubrique — un salarié a plusieurs lignes par mois).

Cette étape ne calcule et n'affecte rien : elle filtre l'année ANNEE_PAIE (colonne
"periode"), découpe par société (colonne "entreprise") et écrit un classeur par société
avec les colonnes d'origine reprises telles quelles, plus les colonnes de mapping PPU et
les colonnes thématiques IJSS réservées (vides). L'affectation des rubriques par thème et
les calculs viendront avec l'intégration du PPU (jointure rubrique <-> Ppu Rubrique Code).

Chaque exécution produit un journal technique horodaté et conservé :
  - output/logs/execution/execution_paie_<ts>.log : déroulé pas à pas + erreurs.

Lancement :
    python main_paie.py
"""
import os
from datetime import datetime

import config
from src.logger_execution import configurer_logger, etape
from src.chargement_paie import lire_paie, decouper_par_entreprise, construire_reconstruit
from src.export_paie import ecrire_xlsx_paie


def resoudre_source_paie():
    """Cherche PAIE_AUDIT.csv dans data/input/, sinon à côté du projet."""
    candidat = os.path.join(config.INPUT_DIR, config.PAIE_INPUT_FILE)
    if os.path.exists(candidat):
        return candidat
    if os.path.exists(config.PAIE_INPUT_FILE):
        return config.PAIE_INPUT_FILE
    return candidat


def main():
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)
    horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
    logger, chemin_exec = configurer_logger(config.LOG_EXEC_DIR, horodatage,
                                            prefix="execution_paie", logger_name="execution_paie")
    logger.info("=" * 70)
    logger.info(f"DÉBUT D'EXÉCUTION PAIE — run {horodatage}")
    print(f"=== Reconstruction IJSS PAIE — run {horodatage} ===")

    fichiers_produits = []
    try:
        # 1. Chargement
        chemin = resoudre_source_paie()
        etape(logger, 1, f"Chargement de la source : {chemin}")
        if not os.path.exists(chemin):
            logger.error(f"Fichier source introuvable : {chemin}")
            raise FileNotFoundError(f"Fichier source introuvable : {chemin}")
        print(f"Chargement de la source : {chemin}")

        # 2. Filtrage année
        etape(logger, 2, f"Filtrage de l'année {config.ANNEE_PAIE} (colonne periode).")
        paie = lire_paie(chemin, annee=config.ANNEE_PAIE, sep=config.PAIE_SEP)
        logger.info(f"{len(paie)} ligne(s) retenue(s) pour l'année {config.ANNEE_PAIE}.")
        print(f"✅ {len(paie)} ligne(s) retenue(s) pour l'année {config.ANNEE_PAIE}.")
        if paie.empty:
            logger.warning(f"Aucune ligne pour l'année {config.ANNEE_PAIE}. Aucun fichier produit.")
            print(f"⚠️  Aucune ligne pour l'année {config.ANNEE_PAIE}. Aucun fichier produit.")
            logger.info(f"FIN D'EXÉCUTION PAIE — run {horodatage} — statut : SUCCÈS (aucune donnée)")
            return

        # 3. Découpage par société
        etape(logger, 3, "Découpage par société (colonne entreprise).")
        par_entreprise = decouper_par_entreprise(paie)
        logger.info(f"Sociétés présentes : {sorted(par_entreprise)}")
        print(f"ℹ️  Sociétés présentes : {sorted(par_entreprise)}")

        # 4. Construction + écriture par société
        etape(logger, 4, f"Construction et écriture des {len(par_entreprise)} classeurs.")
        for entreprise, df in par_entreprise.items():
            detail = construire_reconstruit(df)
            fichier = ecrire_xlsx_paie(entreprise, detail)
            fichiers_produits.append(fichier)
            logger.info(f"— Société '{entreprise}' : {len(detail)} ligne(s) → {fichier}")

        # 5. Bilan
        etape(logger, 5, "Bilan de l'exécution.")
        logger.info(f"Fichiers produits ({len(fichiers_produits)}) :")
        for f in fichiers_produits:
            logger.info(f"    • {f}")
        logger.info(f"FIN D'EXÉCUTION PAIE — run {horodatage} — statut : SUCCÈS")
        print(f"\n✅ Terminé. {len(fichiers_produits)} fichier(s) produit(s) :")
        for f in fichiers_produits:
            print(f"   • {f}")
        print(f"\nJournal d'exécution : {chemin_exec}")

    except Exception as exc:
        logger.exception(f"ERREUR D'EXÉCUTION PAIE : {exc}")
        logger.error(f"FIN D'EXÉCUTION PAIE — run {horodatage} — statut : ÉCHEC")
        print(f"❌ Erreur : {exc} (détail dans {chemin_exec})")

    finally:
        for h in list(logger.handlers):
            h.flush()


if __name__ == "__main__":
    main()
