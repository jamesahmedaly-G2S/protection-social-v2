# -*- coding: utf-8 -*-
"""Helpers de normalisation et de dérivation (sans état)."""
import pandas as pd
import unicodedata

# --- Sources juridiques par motif ---
# 0b. SOURCES JURIDIQUES
# ===============================================================
PARAM_SOURCE_JURIDIQUE = [
    ("temps partiel thérapeutique", "Art. L323-3 CSS - CCN IDCCC 1527"),
    ("accident de travail",         "Art. L.433-1 CSS"),
    ("accident de trajet",          "Art. L.433-1 CSS"),
    ("maladie professionnelle",     "Art. L.461-1 CSS"),
    ("maternité",                   "Art. L.331-3 CSS"),
    ("paternité",                   "Art. L.331-4 CSS"),
    ("adoption",                    "Art. L.331-7 CSS"),
    ("maladie",                     "Art. L323-1 CSS - CCN IDCCC 1527"),
]

def get_source_juridique(motif: str) -> str:
    m = str(motif).strip().lower()
    for kw, ref in PARAM_SOURCE_JURIDIQUE:
        if kw in m:
            return ref
    return "— Référence non définie —"

def _norm_soc(v) -> str:
    if pd.isna(v):
        return ""
    s = str(v).upper()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    s = s.replace("\u2019", "").replace("'", "")
    s = "".join(ch if (ch.isalnum() or ch == " ") else " " for ch in s)
    return " ".join(s.split())

def to_dt(serie):
    return pd.to_datetime(serie, errors="coerce", dayfirst=True)

def to_num(serie):
    s = serie.astype(str).str.strip().str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce")

def jours_cal_mois(start, end, periode):
    if pd.isna(start) or pd.isna(end) or start > end:
        return 0
    ms = periode.start_time
    me = periode.end_time.normalize()
    a = max(start, ms); b = min(end, me)
    if a > b:
        return 0
    return (b - a).days + 1
