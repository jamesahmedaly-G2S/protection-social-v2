# -*- coding: utf-8 -*-
"""
main_absences.py — Comparaison des jours d'absence PAIE <-> DSN
=============================================================================
Compare, par société / salarié / mois / motif d'arrêt, le nombre de jours d'absence
reconstruit côté DSN (classeurs output/xlsx/, déjà reconstruits par main.py) et côté
PAIE (calculé depuis PAIE_AUDIT.csv, via les rubriques validées comme portant un vrai
nombre de jours — cf. config.MOTIFS_PAIE_JOURS). Toute la logique est dans un seul
fichier : src/comparaison_absences.py.

Le détail mensuel est volontairement restreint à janvier/février/mars 2026 : le DSN ne
couvre que le 1er trimestre, comparer des mois où le DSN n'a structurellement aucune
donnée n'apporterait qu'un faux écart (PAIE non nul vs DSN toujours à 0).

Produit :
  - output/rapports/Comparatif_jours_absence_PAIE_DSN.xlsx :
      "1 - Détail mensuel"            : une ligne par salarié / mois / motif, janvier à mars 2026 ;
      "2 - Total T1"                   : une ligne par salarié / motif, sommée sur le 1er trimestre ;
      "3 - Anomalies (jours)"          : lignes PAIE_AUDIT où "Base" est physiquement
        incohérente pour un mois (> 31 jours, ou négative hors annulation) — signalement,
        pas une correction automatique (aucune info exploitable en source pour arbitrer).
    Colonnes : Clé salarié (pivot), Nom, Prénom, Nom d'usage, NIR, Matricule, Entreprise,
    Etablissement, Siren, Nic, Siret, Motif, Nombre de jours absence PAIE/DSN, Mois
    absence (période), Écart constaté — avec filtre Excel cliquable sur chaque colonne.
    Siren/Nic/Siret proviennent de la source DSN (toujours à 0 dans PAIE_AUDIT.csv).
  - output/logs/execution/execution_absences_<ts>.log : déroulé + erreurs.

Lancement :
    python main_absences.py
"""
import os
from datetime import datetime

import config
from src.logger_execution import configurer_logger, etape
from src.comparaison_absences import (nom_fichier_dsn, charger_jours_dsn, charger_jours_paie,
                                      lire_identite_dsn_source, comparer_jours, comparer_jours_periode,
                                      detecter_anomalies_jours, ecrire_comparatif_absences)


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
        # 1. Vérification des classeurs DSN reconstruits
        etape(logger, 1, "Vérification des classeurs DSN reconstruits.")
        manquants = [os.path.join(config.OUTPUT_DIR, nom_fichier_dsn(soc))
                    for soc in mapping.values()
                    if not os.path.exists(os.path.join(config.OUTPUT_DIR, nom_fichier_dsn(soc)))]
        if manquants:
            logger.error(f"Classeurs DSN manquants : {manquants}")
            raise FileNotFoundError(f"Classeurs DSN reconstruits manquants (lancer main.py "
                                    f"au préalable) : {manquants}")

        # 2. Chargement des jours DSN
        etape(logger, 2, "Chargement des jours d'absence DSN (classeurs reconstruits).")
        jours_dsn = charger_jours_dsn(mapping)
        for soc, df in jours_dsn.items():
            logger.info(f"— DSN {soc} : {len(df)} ligne(s) (matricule/mois/motif).")

        # 3. Chargement des jours PAIE
        chemin_paie_source = resoudre(config.PAIE_INPUT_FILE)
        etape(logger, 3, f"Chargement des jours d'absence PAIE depuis {chemin_paie_source}.")
        if not os.path.exists(chemin_paie_source):
            logger.error(f"Source PAIE introuvable : {chemin_paie_source}")
            raise FileNotFoundError(f"Source PAIE introuvable : {chemin_paie_source}")
        jours_paie = charger_jours_paie(chemin_paie_source, annee=config.ANNEE_PAIE, sep=config.PAIE_SEP)
        for code, df in jours_paie.items():
            logger.info(f"— PAIE {code} : {len(df)} ligne(s) (matricule/mois/motif).")

        # 4. Identité Siren/Nic/Siret (source DSN — toujours à 0 côté PAIE_AUDIT.csv)
        chemin_dsn_source = resoudre(config.INPUT_FILE)
        etape(logger, 4, f"Chargement de l'identité Siren/Nic/Siret depuis {chemin_dsn_source}.")
        identite_dsn = lire_identite_dsn_source(chemin_dsn_source) if os.path.exists(chemin_dsn_source) else {}
        logger.info(f"Identité DSN chargée pour {len(identite_dsn)} matricule(s).")

        # 5. Comparaison — détail mensuel (janvier-mars, cohérent avec le périmètre DSN) et total T1
        etape(logger, 5, "Comparaison DSN <-> PAIE (jointure matricule/mois/motif).")
        mois_t1 = [f"{config.ANNEE_PAIE}-01", f"{config.ANNEE_PAIE}-02", f"{config.ANNEE_PAIE}-03"]
        detail_complet = comparer_jours(mapping, jours_dsn, jours_paie, identite_dsn=identite_dsn)
        detail = detail_complet.loc[detail_complet["Mois absence (période)"].isin(mois_t1)].copy()
        detail_t1 = comparer_jours_periode(detail, mois_t1, libelle_periode="T1 2026")

        synthese = (detail_t1.groupby(["Entreprise", "Motif"], as_index=False)
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

        # 6. Détection des anomalies (jours incohérents pour un mois)
        etape(logger, 6, "Détection des anomalies (Base > 31j ou négative hors annulation).")
        anomalies = detecter_anomalies_jours(chemin_paie_source, annee=config.ANNEE_PAIE, sep=config.PAIE_SEP)
        logger.info(f"{len(anomalies)} anomalie(s) détectée(s).")
        print(f"⚠️  {len(anomalies)} ligne(s) anormale(s) (jours incohérents pour un mois).")

        # 7. Écriture
        etape(logger, 7, "Écriture du classeur de comparaison.")
        fichier = ecrire_comparatif_absences(detail, detail_t1, anomalies, libelle_periode="T1 2026")
        logger.info(f"Classeur écrit : {fichier}")

        logger.info(f"FIN D'EXÉCUTION COMPARATIF ABSENCES — run {horodatage} — statut : SUCCÈS")
        print(f"\n✅ Terminé. Classeur : {fichier}")
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
