#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
utils.py — Fonctions utilitaires partagées pour HammerPy Insight.
Contient le parsing numérique HAMMER et les constantes d'unités.
"""

import math
import pandas as pd

# =====================================================================
# CONSTANTES DE CLASSES DE PRESSION
# =====================================================================

PN_CLASSES = {
    "PN 6":   6.0,
    "PN 10":  10.0,
    "PN 16":  16.0,
    "PN 25":  25.0,
    "PN 40":  40.0,
}

PMIN_OPTIONS = {
    "0 bar (Recommandé – pas de dépression)": 0.0,
    "-0.1 bar (Faible dépression)":           -0.1,
    "-0.3 bar (Dépression modérée)":          -0.3,
    "-1.0 bar (Cavitation – Critique)":       -1.0,
}

# =====================================================================
# UNITÉS D'AFFICHAGE (volume + débit)
# =====================================================================

FLOW_UNITS = {
    "m³/h": 1.0,
    "L/s":  1.0 / 3.6,
}

VOLUME_UNITS = {
    "L":    1.0,
    "m³":   1000.0,
}

VOLUME_THRESHOLD_L_DEFAULT = 200.0

# Indices d'auto-détection depuis le nom de colonnes des CSV source
FLOW_UNIT_HINTS = {
    "m3/h": "m³/h", "m³/h": "m³/h", "l/s": "L/s", "l·s⁻¹": "L/s",
    "lps":  "L/s",  "l/h": "L/s",  "lph": "m³/h", "cms":  "m³/h",
    "m3s-1": "m³/h", "l·s−1": "L/s",
}

VOLUME_UNIT_HINTS = {
    "m3":   "m³",  "m³":  "m³",  "l":   "L",   "litres": "L", "liter": "L",
    "liters": "L", "litre": "L", "m^3":  "m³",  "1000l": "m³",
}


# =====================================================================
# PARSING NUMÉRIQUE HAMMER
# =====================================================================

def parse_number(value) -> float | None:
    """
    Parse une valeur numérique depuis un DataFrame HAMMER.
    Gère :
      - float / int natifs (pass-through)
      - Chaînes avec espace insécable \\xa0 comme séparateur de milliers
      - Virgule décimale française (1 029,00 → 1029.0)
      - Espaces normaux comme séparateurs de milliers (1 029.00)
      - None / NaN → None
    """
    if value is None:
        return None
    if isinstance(value, float):
        return None if math.isnan(value) else value
    if isinstance(value, int):
        return float(value)
    if not isinstance(value, str):
        try:
            return float(value)
        except (ValueError, TypeError):
            return None

    s = value.strip()
    if not s or s.lower() in ("nan", "none", "n/a", "(n/a)", ""):
        return None

    s = s.replace('\xa0', '').replace(' ', '')

    if ',' in s:
        s = s.replace(',', '.')

    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def find_col_in_df(df: pd.DataFrame, keywords: list[str]) -> str | None:
    """Recherche la première colonne dont le nom contient l'un des mots-clés."""
    for col in df.columns:
        col_lower = col.lower()
        if any(kw in col_lower for kw in keywords):
            return col
    return None
