# -*- coding: utf-8 -*-
"""
journal_audit.py — Journalisation des erreurs et anomalies du pipeline IJSS
===========================================================================
Se branche sur reconstruction_IJSS_v4_detail_unique.py sans en changer la logique.

Il produit DEUX sorties :
  1) un fichier .log horodaté (toutes sociétés confondues) — trace lisible ;
  2) un onglet Excel « 3 - Journal anomalies » PAR société — restitution auditable,
     avec code, gravité, salarié, contexte, explication et correction.

Chaque anomalie porte :
  - un CODE stable (ANO-0XX), aligné sur la doc SPEC-0002 ;
  - une GRAVITÉ (🔴 bloquant / 🟠 arbitrage / 🟡 à vérifier / ℹ️ info / ✅ corrigé) ;
  - une EXPLICATION (pourquoi c'est signalé) ;
  - un TYPE DE CORRECTION : AUTO (le script corrige) ou SIGNALEMENT (action humaine).

Usage minimal :
    from journal_audit import JournalAudit, ANO
    J = JournalAudit(chemin_log="journal_IJSS.log")
    ...
    J.societe("AURA")
    J.tracer(ANO.ANNULATION_PURE, salarie="POUYET Marc", contexte="motif maladie",
             detail="aucun dépôt positif ne subsiste sur le trimestre")
    ...
    J.ajouter_onglet(wb)     # avant wb.save(...)
    J.cloturer()             # flush + résumé
"""

import logging
from datetime import datetime

# ===============================================================
# CATALOGUE DES ANOMALIES
#   code -> (titre, gravité, explication générique, type de correction)
#   gravités : "BLOQUANT" 🔴 | "ARBITRAGE" 🟠 | "VERIF" 🟡 | "INFO" ℹ️ | "CORRIGE" ✅
# ===============================================================
class ANO:
    # --- Chargement / structure (bloquants) ---
    COLONNE_MANQUANTE   = "ANO-001"
    FICHIER_SOURCE      = "ANO-002"
    SOCIETE_VIDE        = "ANO-010"
    # --- Dates / cohérence de ligne ---
    DATE_INVALIDE       = "ANO-020"
    REPRISE_INCOHERENTE = "ANO-021"
    FINPREV_MANQUANTE   = "ANO-022"
    # --- Annulations ---
    ANNULATION_DETECTEE = "ANO-100"
    ANNULATION_PURE     = "ANO-101"
    ANNULATION_REDUITE  = "ANO-102"
    ANNULATION_AVERIF   = "ANO-103"
    ANNULATION_ORPHELINE= "ANO-104"
    # --- Chevauchement ---
    CHEVAUCHEMENT       = "ANO-200"
    MEME_DJT            = "ANO-201"
    EPISODE_ABSORBE     = "ANO-202"
    ARBITRAGE_MATERNITE = "ANO-203"
    TPT_CHEVAUCHEMENT   = "ANO-205"
    # --- Contrôles finaux (ne devraient jamais arriver) ---
    INVARIANT           = "ANO-300"
    RECOUVREMENT        = "ANO-301"
    # --- Paramétrage ---
    MOTIF_NON_MAPPE     = "ANO-400"
    STATUT_MANQUANT     = "ANO-401"

CATALOGUE = {
    ANO.COLONNE_MANQUANTE:   ("Colonne DSN introuvable", "BLOQUANT",
        "Une colonne obligatoire du mapping est absente : le traitement ne peut pas se faire.",
        "SIGNALEMENT — vérifier le libellé dans find_col() ou renommer la colonne source."),
    ANO.FICHIER_SOURCE:      ("Fichier source illisible", "BLOQUANT",
        "Le CSV/Excel est introuvable ou l'encodage n'est pas reconnu.",
        "SIGNALEMENT — vérifier le chemin et l'encodage (utf-8-sig / latin-1)."),
    ANO.SOCIETE_VIDE:        ("Société sans ligne", "INFO",
        "Aucune ligne DSN pour cette société : elle est ignorée.",
        "AUTO — société sautée, aucun fichier produit."),
    ANO.DATE_INVALIDE:       ("Date invalide (DJT ou mois)", "VERIF",
        "DJT ou mois d'absence non convertible en date : la ligne est écartée du calcul.",
        "AUTO — ligne exclue ; à vérifier si la donnée source est corrigeable."),
    ANO.REPRISE_INCOHERENTE: ("Reprise antérieure ou égale au DJT", "VERIF",
        "La date de reprise est <= au dernier jour travaillé : incohérence de saisie DSN.",
        "AUTO — reprise ignorée pour le bornage ; à confirmer côté RH."),
    ANO.FINPREV_MANQUANTE:   ("Fin prévisionnelle manquante", "VERIF",
        "Un épisode vivant n'a pas de date de fin prévisionnelle exploitable.",
        "AUTO — bornage sur la reprise si dispo, sinon épisode non ventilé ; à vérifier."),
    ANO.ANNULATION_DETECTEE: ("Annulation détectée", "INFO",
        "Ligne marquée annulation (jours<0 OU motif=annulation OU flag_G2SA).",
        "AUTO — conservée et tracée ; ne supprime rien silencieusement."),
    ANO.ANNULATION_PURE:     ("Annulation pure — IJSS à récupérer", "BLOQUANT",
        "Aucun dépôt positif ne subsiste : l'arrêt est entièrement annulé.",
        "SIGNALEMENT — vérifier RH/paie : suppression légitime ou IJSS à récupérer."),
    ANO.ANNULATION_REDUITE:  ("Correction à durée réduite", "ARBITRAGE",
        "Un DJT initial est mort, un autre survit : arrêt re-déclaré à durée/début différents.",
        "AUTO — durée survivante retenue ; vérifier reprise anticipée (rubrique paie 3480)."),
    ANO.ANNULATION_AVERIF:   ("Corrections multiples à vérifier", "ARBITRAGE",
        "Plusieurs annulations ; certains mois annulés sans re-déclaration claire dans le trimestre.",
        "AUTO — épisode réel retenu ; vérifier qu'aucun mois n'est re-signalé ailleurs."),
    ANO.ANNULATION_ORPHELINE:("Annulation non rattachée", "ARBITRAGE",
        "Une ligne 'annulation' n'a pas pu être rattachée à un motif réel (emp+DJT+mois).",
        "SIGNALEMENT — reste dans un groupe 'annulation' isolé ; vérifier le dépôt visé."),
    ANO.CHEVAUCHEMENT:       ("Chevauchement tronqué", "INFO",
        "Deux motifs se chevauchent : le suivant est décalé (règle chronologique).",
        "AUTO — début du motif suivant décalé après le dernier jour absent du précédent."),
    ANO.MEME_DJT:            ("Même DJT sur deux motifs", "ARBITRAGE",
        "Deux motifs partagent le même dernier jour travaillé : découpage non automatique.",
        "SIGNALEMENT — arbitrage humain (ex. MAT/PAT, maladie/AT)."),
    ANO.EPISODE_ABSORBE:     ("Épisode entièrement recouvert", "VERIF",
        "Un épisode est intégralement couvert par un autre : 0 jour retenu.",
        "AUTO — épisode absorbé (0 j) et tracé ; vérifier s'il s'agit d'un doublon."),
    ANO.ARBITRAGE_MATERNITE: ("Arbitrage maternité forcé", "INFO",
        "Date de début maternité forcée manuellement (ARBITRAGE_MATERNITE).",
        "AUTO — date imposée appliquée, épisodes voisins tronqués."),
    ANO.TPT_CHEVAUCHEMENT:   ("Chevauchement temps partiel thérapeutique / arrêt total", "VERIF",
        "Un temps partiel thérapeutique chevauche un arrêt total (DJT différents). Traité par défaut "
        "en chronologie, en attente d'une règle validée (l'arrêt total suspend-il le TPT ?).",
        "SIGNALEMENT — cas à vérifier / arbitrer ultérieurement (décision méthodologie)."),
    ANO.INVARIANT:           ("Invariant jours = carence + IJSS violé", "BLOQUANT",
        "Sur une ligne mensuelle, jours d'absence != carence + IJSS : erreur de calcul.",
        "SIGNALEMENT — ne devrait jamais arriver ; investiguer le calcul de carence."),
    ANO.RECOUVREMENT:        ("Recouvrement résiduel entre motifs", "BLOQUANT",
        "Après déchevauchement, deux motifs se recouvrent encore : bug de bornage.",
        "SIGNALEMENT — ne devrait jamais arriver ; investiguer le déchevauchement."),
    ANO.MOTIF_NON_MAPPE:     ("Motif non paramétré", "VERIF",
        "Le motif n'a pas de code DSN / source juridique / taux dédié.",
        "AUTO — fallback (taux 50%, 'référence non définie') ; compléter le paramétrage."),
    ANO.STATUT_MANQUANT:     ("Statut ou temps de travail manquant", "VERIF",
        "Statut RC ou modalité de temps absent : impacte le futur maintien de salaire.",
        "SIGNALEMENT — récupérer le statut (bloc S21.G00.40)."),
}

_EMOJI = {"BLOQUANT": "🔴", "ARBITRAGE": "🟠", "VERIF": "🟡", "INFO": "ℹ️", "CORRIGE": "✅"}
_NIVEAU_LOG = {"BLOQUANT": logging.ERROR, "ARBITRAGE": logging.WARNING,
               "VERIF": logging.WARNING, "INFO": logging.INFO, "CORRIGE": logging.INFO}


class JournalAudit:
    def __init__(self, chemin_log="journal_IJSS.log"):
        self.chemin_log = chemin_log
        self.entrees = []          # liste de dicts (pour l'onglet Excel)
        self._societe = "—"
        self._logger = logging.getLogger("journal_ijss")
        self._logger.setLevel(logging.INFO)
        self._logger.handlers.clear()
        fh = logging.FileHandler(chemin_log, mode="w", encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s",
                                          datefmt="%Y-%m-%d %H:%M:%S"))
        self._logger.addHandler(fh)
        ch = logging.StreamHandler()
        ch.setFormatter(logging.Formatter("%(message)s"))
        self._logger.addHandler(ch)
        self._logger.info(f"=== Démarrage journal IJSS — {datetime.now():%Y-%m-%d %H:%M:%S} ===")

    def societe(self, nom):
        """Étiquette les anomalies suivantes avec cette société."""
        self._societe = str(nom)
        self._logger.info(f"----- Société : {self._societe} -----")

    def tracer(self, code, salarie="", contexte="", detail="", correction_auto=None):
        """Enregistre une anomalie du catalogue.
        - detail : précision libre sur le cas rencontré ;
        - correction_auto : si le script a appliqué une correction, passer son libellé
          (bascule la gravité affichée en ✅ « corrigé »)."""
        titre, gravite, explication, correction = CATALOGUE.get(
            code, ("Anomalie non cataloguée", "VERIF", detail, "SIGNALEMENT"))
        grav_aff = "CORRIGE" if correction_auto else gravite
        message = explication if not detail else f"{explication} — {detail}"
        corr_txt = correction_auto if correction_auto else correction
        self.entrees.append({
            "Code": code, "Gravité": _EMOJI[grav_aff], "Société": self._societe,
            "Salarié": salarie, "Contexte": contexte, "Titre": titre,
            "Explication": message, "Correction": corr_txt,
        })
        self._logger.log(_NIVEAU_LOG[grav_aff],
                         f"{code} {_EMOJI[grav_aff]} [{self._societe}] "
                         f"{salarie or '—'} | {titre} | {message} | {corr_txt}")

    def erreur_libre(self, message):
        """Erreur inattendue (exception) hors catalogue."""
        self.entrees.append({
            "Code": "ANO-999", "Gravité": "🔴", "Société": self._societe,
            "Salarié": "", "Contexte": "exception", "Titre": "Erreur inattendue",
            "Explication": str(message), "Correction": "SIGNALEMENT — investiguer.",
        })
        self._logger.error(f"ANO-999 🔴 [{self._societe}] {message}")

    def resume(self):
        """Compte par code, pour le log final."""
        cpt = {}
        for e in self.entrees:
            cpt[e["Code"]] = cpt.get(e["Code"], 0) + 1
        return dict(sorted(cpt.items()))

    def ajouter_onglet(self, wb, titre="3 - Journal anomalies"):
        """Ajoute au classeur openpyxl un onglet listant les anomalies de la société
        courante (ou toutes si société non filtrée). Style aligné sur le pipeline."""
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        NAVY = "1F3864"; REDF = "F4D6D2"; ORANGEF = "FBE7CE"; YEL = "FEF3CD"
        GREENF = "DDEBE2"; LIGHT = "F2F4F7"
        lignes = [e for e in self.entrees if e["Société"] in (self._societe, "—")] \
                 or self.entrees
        ws = wb.create_sheet(titre)
        cols = ["Code", "Gravité", "Salarié", "Contexte", "Titre", "Explication", "Correction"]
        thin = Side(style="thin", color="D0D5DD")
        bd = Border(left=thin, right=thin, top=thin, bottom=thin)
        for j, c in enumerate(cols, 1):
            cell = ws.cell(1, j, c)
            cell.font = Font(bold=True, size=9, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=NAVY)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = bd
        for i, e in enumerate(lignes, start=2):
            for j, c in enumerate(cols, 1):
                cell = ws.cell(i, j, e.get(c, ""))
                cell.border = bd
                cell.font = Font(size=9)
                cell.alignment = Alignment(horizontal="left", vertical="center",
                                           wrap_text=(c in ("Explication", "Correction")))
            g = e["Gravité"]
            fill = REDF if g == "🔴" else ORANGEF if g == "🟠" else YEL if g == "🟡" \
                   else GREENF if g == "✅" else LIGHT
            for j in range(1, len(cols) + 1):
                ws.cell(i, j).fill = PatternFill("solid", fgColor=fill)
        larg = {"Code": 10, "Gravité": 8, "Salarié": 20, "Contexte": 22,
                "Titre": 30, "Explication": 55, "Correction": 55}
        for j, c in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(j)].width = larg[c]
        ws.freeze_panes = "A2"
        return ws

    def cloturer(self):
        r = self.resume()
        self._logger.info(f"=== Fin — {len(self.entrees)} anomalie(s) : {r} ===")
        for h in self._logger.handlers:
            h.flush()
