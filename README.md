# Reconstruction IJSS — Projet Protection Sociale

Reconstruit les IJSS théoriques depuis la DSN (Q1 2026), avec déchevauchement des
motifs et traitement des annulations. Documentation : SPEC-0002 (version corrigée).

## Arborescence

```
reconstruction_ijss/
├── main.py              # orchestrateur (point d'entrée)
├── config.py            # paramètres (dates, seuil, sociétés, motifs, chemins)
├── requirements.txt
├── data/
│   └── input/           # ← déposer ici le CSV DSN source
├── output/
│   ├── xlsx/            # classeurs par société (générés)
│   └── logs/            # journaux d'anomalies horodatés (générés)
└── src/
    ├── normalisation.py # helpers (dates, normalisation, sources juridiques)
    ├── chargement.py    # lecture CSV, mapping colonnes, détection annulations
    ├── pipeline.py      # cœur métier : traiter_societe (logique validée)
    ├── export_excel.py  # écriture des classeurs
    ├── journal_audit.py # journalisation des anomalies (.log + onglet Excel)
    ├── logger_execution.py  # journal d'exécution (déroulé pas à pas + erreurs)
    └── rapport_execution.py # rapport cumulatif (append) : .md + .xlsx
```

## Utilisation

1. Déposer le fichier DSN dans `data/input/` (nom attendu dans `config.py`).
2. Installer les dépendances : `pip install -r requirements.txt`
3. Lancer : `python main.py`

Sorties :
- un classeur `Reconstruit_{SOCIETE}_CORRIGE.xlsx` par société dans `output/xlsx/`,
  chacun avec l'onglet « 3 - Journal anomalies » ;
- un journal global horodaté dans `output/logs/`.

## Journal des anomalies

Chaque anomalie porte un code (ANO-0XX), une gravité (🔴 bloquant / 🟠 arbitrage /
🟡 à vérifier / ℹ️ info / ✅ corrigé auto), une explication et une correction
(automatique ou action à mener). Voir le catalogue dans `src/journal_audit.py`.

## Rapport d'exécutions (cumulatif)

À chaque `python main.py`, deux fichiers de `output/rapports/` sont **complétés**
(jamais écrasés) :
- `rapport_executions.md` : une section datée par run (résumé + copie intégrale de
  la sortie terminal, repliée dans un bloc dépliable) ;
- `rapport_executions.xlsx` : classeur multi-onglets qui reprend TOUTES les infos
  du Markdown, empilées à chaque run —
    * « 1 - Exécutions » : 1 ligne par run (synthèse) ;
    * « 2 - Détail sociétés » : 1 ligne par (run, société) avec familles d'annulations ;
    * « 3 - Anomalies » : 1 ligne par anomalie (code, gravité, explication, correction) ;
    * « 4 - Sortie terminal » : capture intégrale par run.

Le `.md` peut être converti en Word à la demande via la chaîne `markdown_to_docx.py`.

## Journaux d'exécution (traçabilité & diagnostic)

Chaque run produit deux journaux horodatés et conservés (jamais écrasés), partageant
le même horodatage pour être retrouvés ensemble :

- `output/logs/execution/execution_<ts>.log` — **déroulé technique** : début/fin,
  chaque étape horodatée à la seconde, fichiers générés (chemins absolus), bilan des
  anomalies, et — en cas de problème — l'**erreur avec son traceback complet** et le
  statut ÉCHEC.
- `output/logs/anomalies/journal_IJSS_<ts>.log` — **anomalies métier** détectées.

On peut ainsi revenir sur une exécution d'une date donnée, consulter son déroulé,
identifier les fichiers produits et les retrouver. Les traces sont clôturées même si
le run échoue (bloc `try/finally`).
# protection-social-v2
