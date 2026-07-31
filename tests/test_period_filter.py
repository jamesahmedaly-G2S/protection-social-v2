import os
import tempfile
import unittest

import pandas as pd

import config
from src.chargement_paie import lire_paie
from src.chargement import lire_source, mapper_colonnes, normaliser_et_detecter


class PeriodFilterTests(unittest.TestCase):
    def test_lire_paie_filters_to_first_three_months(self):
        with tempfile.NamedTemporaryFile("w", suffix=".csv", delete=False, encoding="utf-8") as fh:
            fh.write(
                "periode;matricule;nom;prenom;entreprise;etablissement;siren;nic;siret;rubrique;libelle;base\n"
                "2026-01-01;1001;A;B;52;E1;111;222;33344455556666;3300;Absence Maladie;5\n"
                "2026-02-01;1001;A;B;52;E1;111;222;33344455556666;3300;Absence Maladie;7\n"
                "2026-04-01;1001;A;B;52;E1;111;222;33344455556666;3300;Absence Maladie;9\n"
            )
            path = fh.name

        try:
            df = lire_paie(path, annee=2026, sep=";",
                           date_debut=pd.Timestamp("2026-01-01"),
                           date_fin=pd.Timestamp("2026-03-31"))
            self.assertEqual(len(df), 2)
            self.assertEqual(df["periode"].tolist(), ["2026-01-01", "2026-02-01"])
        finally:
            os.remove(path)

    def test_normaliser_et_detecter_filters_to_first_three_months(self):
        with tempfile.NamedTemporaryFile("w", suffix=".csv", delete=False, encoding="utf-8") as fh:
            fh.write(
                "ID;Individu.Matricule;S21.G00.30.002 Individu.NomFamille;S21.G00.30.004 Individu.Prenoms;"
                "S21.G00.60.001_TravailArret.Motif_Libelle;S21.G00.60.002_TravailArret.DernierJour;"
                "Nb_jours_absences_SS;mois_absence;Nom usuel de la société;flag_annulation_G2SA\n"
                "1;1001;Doe;John;maladie;2026-01-15;10;2026-01;ACME;False\n"
                "2;1002;Smith;Jane;maladie;2026-02-15;12;2026-02;ACME;False\n"
                "3;1003;Brown;Bob;maladie;2026-04-15;14;2026-04;ACME;False\n"
            )
            path = fh.name

        try:
            raw = lire_source(path)
            raw.columns = raw.columns.str.strip()
            cols = mapper_colonnes(raw)
            df = normaliser_et_detecter(raw, cols,
                                        date_debut=pd.Timestamp("2026-01-01"),
                                        date_fin=pd.Timestamp("2026-03-31"))
            self.assertEqual(len(df), 2)
            self.assertEqual(df["mois_p"].tolist(), [pd.Period("2026-01", freq="M"), pd.Period("2026-02", freq="M")])
        finally:
            os.remove(path)

    def test_configured_period_filter_can_be_overridden(self):
        original = config.PERIODE_FILTRE_DEBUT, config.PERIODE_FILTRE_FIN
        try:
            config.PERIODE_FILTRE_DEBUT = pd.Timestamp("2026-02-01")
            config.PERIODE_FILTRE_FIN = pd.Timestamp("2026-02-28")
            with tempfile.NamedTemporaryFile("w", suffix=".csv", delete=False, encoding="utf-8") as fh:
                fh.write(
                    "periode;matricule;nom;prenom;entreprise;etablissement;siren;nic;siret;rubrique;libelle;base\n"
                    "2026-01-01;1001;A;B;52;E1;111;222;33344455556666;3300;Absence Maladie;5\n"
                    "2026-02-01;1002;A;B;52;E1;111;222;33344455556666;3300;Absence Maladie;7\n"
                    "2026-03-01;1003;A;B;52;E1;111;222;33344455556666;3300;Absence Maladie;9\n"
                )
                path = fh.name
            try:
                df = lire_paie(path, annee=2026, sep=";")
                self.assertEqual(len(df), 1)
                self.assertEqual(df.iloc[0]["periode"], "2026-02-01")
            finally:
                os.remove(path)
        finally:
            config.PERIODE_FILTRE_DEBUT, config.PERIODE_FILTRE_FIN = original


if __name__ == "__main__":
    unittest.main()
