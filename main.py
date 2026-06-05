#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
HammerPy Insight - Script Principal (Phase 3)
---------------------------------------------
Application d'automatisation et d'interprétation des résultats transitoires
exportés depuis le logiciel d'hydraulique Bentley HAMMER.

Fonctionnalités :
  - Parsing ultra-robuste des fichiers CSV/Excel (formats HAMMER francophones et anglophones)
  - Interface moderne avec configuration de l'étude (Projet, Ingénieur, PN conduite)
  - Visualisation graphique interactive avec seuils critiques (Matplotlib)
  - Export de rapport professionnel Word (.docx) avec tableaux et graphique intégré

Auteur  : HammerPy Insight / Expert Python & CustomTkinter
Version : 2.0 - Juin 2026
"""

import os
import io
import json
import tempfile
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox

import pandas as pd
import customtkinter as ctk

# Imports pour Matplotlib et son intégration dans Tkinter
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

# Import pour la génération du rapport Word
try:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# Configuration du thème visuel global (doit être en amont de toute création de widget)
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# Dictionnaire des classes de pression nominale (PN) → pression maximale admissible en bar
PN_CLASSES = {
    "PN 6":   6.0,
    "PN 10":  10.0,
    "PN 16":  16.0,
    "PN 25":  25.0,
    "PN 40":  40.0,
}

# Pression minimale admissible (dépression) selon les normes typiques PEHD/Acier
PMIN_OPTIONS = {
    "0 bar (Recommandé – pas de dépression)": 0.0,
    "-0.1 bar (Faible dépression)":           -0.1,
    "-0.3 bar (Dépression modérée)":          -0.3,
    "-1.0 bar (Cavitation – Critique)":       -1.0,
}

# =====================================================================
# UNITÉS D'AFFICHAGE (volume + débit)
# =====================================================================
# Stockage canonique interne : L (volume) et m³/h (débit).
# Les facteurs ci-dessous servent uniquement à la conversion d'affichage.
FLOW_UNITS = {
    "m³/h": 1.0,        # canonique
    "L/s":  1.0 / 3.6,  # 1 m³/h = 1000 L / 3600 s
}
VOLUME_UNITS = {
    "L":    1.0,        # canonique
    "m³":   1000.0,     # 1 m³ = 1000 L
}
VOLUME_THRESHOLD_L_DEFAULT = 200.0   # Seuil HPT, en L (canonique)

# ── Indices d'auto-détection depuis le nom de colonnes des CSV source ──
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
# 1. LOGIQUE DE TRAITEMENT DES DONNÉES (MÉTIER)
# =====================================================================

class HammerDataParser:
    """
    Classe responsable du chargement, de la validation et du parsing robuste
    des données hydrauliques exportées depuis Bentley HAMMER.

    Prend en charge :
      - CSV avec séparateur virgule, point-virgule ou tabulation
      - Décimales en format anglo-saxon (point) ou français (virgule)
      - Encodages UTF-8 et Latin-1 (ISO 8859-1)
      - Fichiers Excel .xlsx et .xls
    """

    def __init__(self):
        self.hpt_data: pd.DataFrame | None = None
        self.steady_state_data: pd.DataFrame | None = None

    # ------------------------------------------------------------------
    # Méthode utilitaire : chargement universel d'un fichier tabulaire
    # ------------------------------------------------------------------
    def _load_file(self, filepath: str) -> pd.DataFrame:
        """
        Charge n'importe quel fichier tabulaire supporté (CSV ou Excel)
        en gérant automatiquement les encodages et les séparateurs décimaux.

        Raises:
            ValueError: si l'extension n'est pas reconnue.
        """
        ext = os.path.splitext(filepath)[1].lower()

        if ext in ('.xlsx', '.xls'):
            return pd.read_excel(filepath)

        if ext == '.csv':
            # Essai successif avec différents encodages courants
            for enc in ('utf-8', 'latin-1', 'cp1252'):
                try:
                    df = pd.read_csv(filepath, sep=None, engine='python',
                                     encoding=enc, decimal='.', thousands=None)
                    # Si toutes les colonnes numériques sont des chaînes, réessayer avec la virgule comme décimale
                    numeric_cols = df.select_dtypes(include='number')
                    if len(numeric_cols.columns) == 0:
                        df = pd.read_csv(filepath, sep=None, engine='python',
                                         encoding=enc, decimal=',', thousands=None)
                    return df
                except UnicodeDecodeError:
                    continue
            raise ValueError(f"Impossible de lire le fichier CSV avec les encodages testés (UTF-8, Latin-1, CP1252).")

        raise ValueError(f"Extension '{ext}' non supportée. Utilisez .csv, .xlsx ou .xls")

    # ------------------------------------------------------------------
    # Méthode utilitaire : recherche flexible de colonnes
    # ------------------------------------------------------------------
    @staticmethod
    def _find_col(columns: list, keywords: list) -> str | None:
        """
        Recherche la première colonne dont le nom (en minuscules) contient
        l'un des mots-clés fournis.

        Returns:
            Nom de la colonne trouvée, ou None.
        """
        for col in columns:
            col_lower = col.lower()
            if any(kw in col_lower for kw in keywords):
                return col
        return None

    # ------------------------------------------------------------------
    # Parsing Régime Permanent (Station)
    # ------------------------------------------------------------------
    def parse_station_file(self, filepath: str, source_unit: str | None = None) -> dict:
        """
        Parse le fichier de régime permanent de la station hydraulique.
        Cherche les colonnes de Débit (Q) et de HMT dans le fichier.

        Args:
            filepath: chemin du fichier CSV/Excel.
            source_unit: unité des valeurs lues ("m³/h" par défaut, ou "L/s"
                         pour conversion automatique vers m³/h canonique).
                         Si None, suppose l'unité canonique (m³/h).

        Returns:
            dict avec les clés :
              success (bool), flow_rate_m3h, hmt_m, source_unit_detected,
              n_rows, columns, message
        """
        result = {
            "success": False,
            "flow_rate_m3h": None,
            "hmt_m": None,
            "source_unit_detected": None,
            "n_rows": 0,
            "columns": [],
            "message": "Fichier non traité"
        }

        try:
            df = self._load_file(filepath)
            df.columns = [str(c).strip() for c in df.columns]
            result["columns"] = list(df.columns)
            result["n_rows"] = len(df)

            # Recherche des colonnes Débit et HMT (mots-clés anglophones + francophones)
            q_col   = self._find_col(df.columns,
                ['debit', 'flow', 'flow_rate', 'q (m3', 'm3/h', 'discharge',
                 'q (l/s', 'l/s', 'lps', 'q (l·s'])
            hmt_col = self._find_col(df.columns, ['hmt', 'head', 'mce', 'hauteur', 'total_head', 'tdh'])

            flow_rate = float(df.iloc[0][q_col])   if q_col   else None
            hmt       = float(df.iloc[0][hmt_col]) if hmt_col else None

            # Conversion vers l'unité canonique (m³/h) si nécessaire
            if flow_rate is not None and source_unit == "L/s":
                flow_rate = flow_rate * 3.6   # L/s → m³/h

            self.steady_state_data = df
            result.update({
                "success": True,
                "flow_rate_m3h": round(flow_rate, 2) if flow_rate is not None else None,
                "hmt_m":         round(hmt, 2)        if hmt       is not None else None,
                "source_unit_detected": source_unit or "m³/h",
                "message": (
                    f"Chargement réussi : {len(df)} lignes, {len(df.columns)} colonnes.\n"
                    f"Colonnes trouvées : {', '.join(df.columns)}"
                )
            })

        except Exception as exc:
            result["message"] = f"Erreur de parsing station : {exc}"

        return result

    # ------------------------------------------------------------------
    # Parsing Analyse Transitoire (HPT)
    # ------------------------------------------------------------------
    def parse_hpt_file(self, filepath: str, source_unit: str | None = None) -> dict:
        """
        Parse le fichier des résultats transitoires (courbes enveloppes HPT).
        Recherche : Distance, Pressure (Minimum), Pressure (Maximum),
                    Volume of Gas (Maximum).

        Args:
            filepath: chemin du fichier CSV/Excel.
            source_unit: unité des valeurs lues ("L" par défaut, ou "m³"
                         pour conversion automatique vers L canonique).
                         Si None, suppose l'unité canonique (L).

        Returns:
            dict avec les clés :
              success, max_gas_volume_l, min_pressure_bar, max_pressure_bar,
              source_unit_detected, critical_columns_found,
              is_simulated (données placeholder), n_rows, message
        """
        result = {
            "success": False,
            "max_gas_volume_l": None,
            "min_pressure_bar": None,
            "max_pressure_bar": None,
            "source_unit_detected": None,
            "critical_columns_found": [],
            "is_simulated": False,
            "n_rows": 0,
            "message": "Fichier non traité"
        }

        try:
            df = self._load_file(filepath)
            df.columns = [str(c).strip() for c in df.columns]
            self.hpt_data = df
            result["n_rows"] = len(df)

            cols = list(df.columns)
            found = []

            # Mots-clés pour chaque grandeur physique (anglais + français HAMMER)
            vol_kw   = ['volume of gas', 'volume gaz', 'gas volume', 'vol gaz', 'vol. gaz', 'air volume']
            pmin_kw  = ['pressure (minimum)', 'pression (minimum)', 'pmin', 'p min', 'hgl min',
                        'min pressure', 'pressure min', 'pression min']
            pmax_kw  = ['pressure (maximum)', 'pression (maximum)', 'pmax', 'p max', 'hgl max',
                        'max pressure', 'pressure max', 'pression max']

            vol_col  = self._find_col(cols, vol_kw)
            pmin_col = self._find_col(cols, pmin_kw)
            pmax_col = self._find_col(cols, pmax_kw)

            is_simulated = False

            # --- Volume de gaz ---
            if vol_col:
                found.append(vol_col)
                max_vol = float(df[vol_col].max())
                # Conversion vers unité canonique (L) si nécessaire
                if source_unit == "m³":
                    max_vol = max_vol * 1000.0   # m³ → L
            else:
                max_vol = 124.5
                is_simulated = True

            # --- Pression minimale ---
            if pmin_col:
                found.append(pmin_col)
                min_press = float(df[pmin_col].min())
            else:
                min_press = -0.35
                is_simulated = True

            # --- Pression maximale ---
            if pmax_col:
                found.append(pmax_col)
                max_press = float(df[pmax_col].max())
            else:
                max_press = 12.4
                is_simulated = True

            result.update({
                "success": True,
                "max_gas_volume_l":    round(max_vol,   2),
                "min_pressure_bar":    round(min_press, 2),
                "max_pressure_bar":    round(max_press, 2),
                "source_unit_detected": source_unit or "L",
                "critical_columns_found": found,
                "is_simulated": is_simulated,
                "message": (
                    f"Succès : {len(df)} points d'enveloppe chargés.\n"
                    f"Colonnes identifiées : {', '.join(found) if found else 'Aucune — valeurs de simulation utilisées'}"
                )
            })

        except Exception as exc:
            result["message"] = f"Erreur de parsing transitoire : {exc}"

        return result


# =====================================================================
# 1b. CHARGEUR DE CLASSEUR HAMMER (Flex Tables .xlsx/.xls)
# =====================================================================

# Noms de feuilles reconnus (insensible casse/accents)
_SHEET_ALIASES = {
    "pipes":      ["pipes", "conduites", "pipe", "conduite"],
    "nodes":      ["noeuds", "nodes", "nœuds", "node", "nodo"],
    "reservoirs": ["reservoirs", "reservoir", "réservoirs", "réservoir", "tanks", "tank"],
    "pumps":      ["pumps", "pompe", "pump", "pompes"],
    "hpt":        ["hpt", "reservoir antideflagrant", "hydropneumatic", "hydropneumatique"],
    "air_valves": ["air valves", "air valve", "ventouses", "ventouse", "ventilateur"],
}

# Feuilles obligatoires (sinon le système est non-fonctionnel)
_SHEET_MANDATORY = {"pipes", "nodes", "pumps"}

# Colonnes attendues par parser (pour validation de colonnes critiques)
_SHEET_COLUMNS = {
    "pipes": [
        "length (user defined) (m)", "start node", "stop node", "diameter (mm)",
        "material", "hazen-williams c", "wave speed (m/s)",
        "pressure (maximum, transient) (bars)", "pressure (minimum, transient) (bars)",
    ],
    "nodes": [
        "id", "label", "elevation (m)",
        "pressure (maximum, transient) (bars)", "pressure (minimum, transient) (bars)",
        "x (m)", "y (m)",
    ],
    "pumps": [
        "id", "label", "elevation (m)", "flow (total) (l/s)", "pump head (m)",
        "pressure (maximum, transient) (bars)", "pressure (minimum, transient) (bars)",
    ],
}


def _parse_number(value) -> float | None:
    """
    Parse une valeur numérique depuis un DataFrame HAMMER.
    Gère :
      - float / int natifs (pass-through)
      - Chaînes avec espace insécable \xa0 comme séparateur de milliers
      - Virgule décimale française (1 029,00 → 1029.0)
      - Espaces normaux comme séparateurs de milliers (1 029.00)
      - None / NaN → None
    """
    if value is None:
        return None
    if isinstance(value, float):
        import math
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

    # Supprimer les espaces insécables (\xa0) et normaux (séparateurs milliers)
    s = s.replace('\xa0', '').replace(' ', '')

    # Remplacer la virgule décimale française par un point
    if ',' in s:
        s = s.replace(',', '.')

    try:
        return float(s)
    except (ValueError, TypeError):
        return None


class WorkbookManager:
    """
    Charge un classeur HAMMER (.xlsx/.xls) contenant les Flex Tables.
    Valide la présence des feuilles obligatoires (Pipes, Noeuds, Pumps).
    Extrait et nettoie les données de chaque feuille.
    """

    def __init__(self):
        self.filepath: str = ""
        self.sheet_names_found: list[str] = []
        self.sheet_map: dict[str, str] = {}   # canonical_name → actual_sheet_name
        self.sheets: dict[str, pd.DataFrame] = {}
        self.errors: list[str] = []
        self.warnings: list[str] = []

    def load(self, filepath: str) -> bool:
        """
        Charge le classeur et identifie les feuilles reconnues.
        Retourne True si au moins les feuilles obligatoires sont présentes.
        """
        self.filepath = filepath
        self.sheet_names_found = []
        self.sheet_map = {}
        self.sheets = {}
        self.errors = []
        self.warnings = []

        ext = os.path.splitext(filepath)[1].lower()
        if ext not in ('.xlsx', '.xls'):
            self.errors.append(f"Extension '{ext}' non supportée. Utilisez .xlsx ou .xls")
            return False

        try:
            xl = pd.ExcelFile(filepath)
        except Exception as exc:
            self.errors.append(f"Impossible d'ouvrir le classeur : {exc}")
            return False

        self.sheet_names_found = xl.sheet_names

        # Mapper les feuilles trouvées aux noms canoniques
        for canonical, aliases in _SHEET_ALIASES.items():
            for sheet_name in xl.sheet_names:
                sheet_lower = sheet_name.lower().strip()
                if any(a in sheet_lower for a in aliases):
                    self.sheet_map[canonical] = sheet_name
                    break

        # Vérifier les feuilles obligatoires
        for mandatory in _SHEET_MANDATORY:
            if mandatory not in self.sheet_map:
                self.errors.append(
                    f"Feuille obligatoire manquante : '{mandatory}'. "
                    f"Le système n'est pas fonctionnel sans cette feuille.")

        # Charger chaque feuille identifiée
        for canonical, sheet_name in self.sheet_map.items():
            try:
                df = pd.read_excel(filepath, sheet_name=sheet_name, header=0)
                # Nettoyer les noms de colonnes
                df.columns = [
                    str(c).replace('\r', ' ').replace('\n', ' ').strip()
                    for c in df.columns
                ]
                # Supprimer les lignes entièrement vides
                df = df.dropna(how='all')
                self.sheets[canonical] = df
            except Exception as exc:
                self.warnings.append(
                    f"Erreur lors du chargement de la feuille '{sheet_name}' : {exc}")

        return len(self.errors) == 0

    def validate(self) -> tuple[bool, list[str]]:
        """
        Valide la structure des données chargées.
        Retourne (is_valid, list_of_errors).
        """
        errors = list(self.errors)

        for canonical in _SHEET_MANDATORY:
            if canonical not in self.sheets:
                errors.append(f"Données manquantes : feuille '{canonical}' non chargée.")

        return len(errors) == 0, errors

    def get_sheet(self, name: str) -> pd.DataFrame | None:
        """Retourne une feuille par son nom canonique, ou None."""
        return self.sheets.get(name)

    def get_summary(self) -> dict:
        """
        Retourne un résumé des données chargées pour l'UI.
        {
          "pipes_count": int,
          "nodes_count": int,
          "pumps_count": int,
          "reservoirs_count": int,
          "hpt_count": int,
          "air_valves_count": int,
          "materials": list[str],
          "diameter_min_mm": float,
          "diameter_max_mm": float,
          "pmax_bar": float,
          "pmin_bar": float,
          "vmax_pump_ls": float,
          "vmax_hpt_l": float,
          "vmax_air_valve_l": float,
        }
        """
        summary = {
            "pipes_count": 0, "nodes_count": 0, "pumps_count": 0,
            "reservoirs_count": 0, "hpt_count": 0, "air_valves_count": 0,
            "materials": [], "diameter_min_mm": None, "diameter_max_mm": None,
            "pmax_bar": None, "pmin_bar": None,
            "vmax_pump_ls": None, "vmax_hpt_l": None, "vmax_air_valve_l": None,
        }

        # Pipes
        df = self.sheets.get("pipes")
        if df is not None and len(df) > 0:
            summary["pipes_count"] = len(df)
            mat_col = _find_col_in_df(df, ['material', 'matériau'])
            if mat_col:
                summary["materials"] = sorted(df[mat_col].dropna().unique().tolist())
            diam_col = _find_col_in_df(df, ['diameter', 'diamètre', 'ø'])
            if diam_col:
                diameters = df[diam_col].apply(_parse_number).dropna()
                if len(diameters) > 0:
                    summary["diameter_min_mm"] = round(float(diameters.min()), 1)
                    summary["diameter_max_mm"] = round(float(diameters.max()), 1)
            pmax_col = _find_col_in_df(df, ['pressure (maximum'])
            pmin_col = _find_col_in_df(df, ['pressure (minimum'])
            if pmax_col:
                vals = df[pmax_col].apply(_parse_number).dropna()
                if len(vals) > 0:
                    summary["pmax_bar"] = round(float(vals.max()), 4)
            if pmin_col:
                vals = df[pmin_col].apply(_parse_number).dropna()
                if len(vals) > 0:
                    v = float(vals.min())
                    summary["pmin_bar"] = round(v, 4)

        # Nodes
        df = self.sheets.get("nodes")
        if df is not None:
            summary["nodes_count"] = len(df)
            # Pressions globales depuis les nœuds si pas encore trouvées
            if summary["pmax_bar"] is None:
                pmax_col = _find_col_in_df(df, ['pressure (maximum'])
                if pmax_col:
                    vals = df[pmax_col].apply(_parse_number).dropna()
                    if len(vals) > 0:
                        summary["pmax_bar"] = round(float(vals.max()), 4)
            if summary["pmin_bar"] is None:
                pmin_col = _find_col_in_df(df, ['pressure (minimum'])
                if pmin_col:
                    vals = df[pmin_col].apply(_parse_number).dropna()
                    if len(vals) > 0:
                        summary["pmin_bar"] = round(float(vals.min()), 4)

        # Pumps
        df = self.sheets.get("pumps")
        if df is not None:
            summary["pumps_count"] = len(df)
            flow_col = _find_col_in_df(df, ['flow (total)', 'débit', 'flow'])
            if flow_col:
                vals = df[flow_col].apply(_parse_number).dropna()
                if len(vals) > 0:
                    summary["vmax_pump_ls"] = round(float(vals.max()), 2)

        # Reservoirs
        df = self.sheets.get("reservoirs")
        if df is not None:
            summary["reservoirs_count"] = len(df)

        # HPT
        df = self.sheets.get("hpt")
        if df is not None:
            summary["hpt_count"] = len(df)
            vol_col = _find_col_in_df(df, ['volume of gas', 'gas volume', 'volume gaz'])
            if vol_col:
                vals = df[vol_col].apply(_parse_number).dropna()
                if len(vals) > 0:
                    summary["vmax_hpt_l"] = round(float(vals.max()), 2)

        # Air valves
        df = self.sheets.get("air_valves")
        if df is not None:
            summary["air_valves_count"] = len(df)
            air_col = _find_col_in_df(df, ['air volume', 'volume air', 'volume (maximum'])
            if air_col:
                vals = df[air_col].apply(_parse_number).dropna()
                if len(vals) > 0:
                    summary["vmax_air_valve_l"] = round(float(vals.max()), 2)

        return summary


def _find_col_in_df(df: pd.DataFrame, keywords: list[str]) -> str | None:
    """Recherche la première colonne dont le nom contient l'un des mots-clés."""
    for col in df.columns:
        col_lower = col.lower()
        if any(kw in col_lower for kw in keywords):
            return col
    return None


# =====================================================================
# 1b. PARSER RAPPORT POMPE (RTF / TXT)
# =====================================================================

import re as _re

# Mots-clés RTF à ignorer lors de l'extraction du texte brut
_RTF_STRIP_RE = _re.compile(
    r'\\[a-zA-Z]+\d*\s?'
    r'|\\\{|\\\}'
    r'|\\pict[^}]*(?:\}|$)'
    r'|\{\\\*\\[^{}]+\}'
)


def _strip_rtf(rtf_text: str) -> str:
    """
    Extrait le texte lisible d'un fichier RTF brute.
    Approche : supprimer les groupes connus (fonttbl, pict, styles),
    puis extraire le texte restant.
    """
    text = rtf_text

    # 1. Supprimer le bloc fonttbl (contient toutes les définitions de polices)
    idx = text.find('{\\fonttbl')
    if idx >= 0:
        depth = 0
        for i in range(idx, len(text)):
            if text[i] == '{':
                depth += 1
            elif text[i] == '}':
                depth -= 1
                if depth == 0:
                    text = text[:idx] + text[i+1:]
                    break

    # 2. Supprimer les groupes {\*\...} (groupes de styles internes, info, etc.)
    #    On les détecte par {\* au début et on supprime jusqu'à l'accolade fermante
    while '{\\*' in text:
        start = text.find('{\\*')
        # Trouver l'accolade fermante correspondante
        depth = 0
        for i in range(start, len(text)):
            if text[i] == '{':
                depth += 1
            elif text[i] == '}':
                depth -= 1
                if depth == 0:
                    text = text[:start] + text[i+1:]
                    break

    # 3. Supprimer les groupes {\pict...} (images embarquées)
    while '{\\pict' in text:
        start = text.find('{\\pict')
        depth = 0
        for i in range(start, len(text)):
            if text[i] == '{':
                depth += 1
            elif text[i] == '}':
                depth -= 1
                if depth == 0:
                    text = text[:start] + text[i+1:]
                    break

    # 4. Supprimer les groupes {\background...} (contient des images)
    while '{\\background' in text:
        start = text.find('{\\background')
        depth = 0
        for i in range(start, len(text)):
            if text[i] == '{':
                depth += 1
            elif text[i] == '}':
                depth -= 1
                if depth == 0:
                    text = text[:start] + text[i+1:]
                    break

    # 5. Supprimer les contrôle mots RTF
    text = _re.sub(r'\\[a-zA-Z]+\d*\s?', ' ', text)

    # 6. Nettoyer les accolades restantes et le texte
    text = text.replace('{', '').replace('}', '')
    text = _re.sub(r'[ \t]+', ' ', text)
    text = _re.sub(r' \n', '\n', text)
    text = _re.sub(r'\n\s*\n', '\n', text)
    return text.strip()


class PumpReportParser:
    """
    Parse un rapport détaillé pompe Bentley HAMMER exporté en RTF (.rtf) ou TXT (.txt).
    Extrait les données d'identification, le point de fonctionnement actuel,
    et les grandeurs NPSH associées.

    Sortie typique :
    {
      "success": True,
      "pump_id": "131",
      "label": "PMP-2",
      "downstream_pipe": "P-4",
      "flow_lps": 100.0,       # débit en L/s (stockage interne)
      "pump_head_m": 110.80,
      "pressure_suction_bar": 0.54,
      "pressure_discharge_bar": 11.38,
      "npsh_required_m": None,  # (N/A) dans le rapport
      "npsh_available_m": 15.35,
      "speed_factor": 1.0,
      "status_initial": "On",
      "controlled": False,
      "hydraulic_grade_suction_m": 114.19,
      "hydraulic_grade_discharge_m": 224.99,
      "curve_points": [],       # Pour v3.x : points H(Q) saisis manuellement
    }
    """

    # Paires label→clé canoniques pour l'extraction
    _KEY_MAP = {
        "id":                  "pump_id",
        "label":               "label",
        "downstream pipe":     "downstream_pipe",
        "flow (total)":        "flow_lps",
        "flow (absolute)":     "_flow_abs_lps",
        "pump head":           "pump_head_m",
        "pressure (suction)":  "pressure_suction_bar",
        "pressure (discharge)": "pressure_discharge_bar",
        "npsh (required)":     "npsh_required_m",
        "npsh (available)":    "npsh_available_m",
        "relative speed factor (calculated)": "speed_factor",
        "relative speed factor (initial)":    "speed_factor",
        "status (initial)":    "status_initial",
        "status (calculated)": "_status_calc",
        "controlled?":         "controlled",
        "hydraulic grade (suction)":  "hydraulic_grade_suction_m",
        "hydraulic grade (discharge)": "hydraulic_grade_discharge_m",
    }

    _UNITS = {
        "flow_lps":    "L/s",
        "pump_head_m": "m",
        "pressure_suction_bar":   "bar",
        "pressure_discharge_bar": "bar",
        "npsh_required_m":  "m",
        "npsh_available_m": "m",
        "hydraulic_grade_suction_m":  "m",
        "hydraulic_grade_discharge_m": "m",
    }

    def __init__(self):
        self.filepath: str = ""
        self.raw_text: str = ""
        self.parsed: dict = {}
        self.errors: list[str] = []
        self.curve_points: list[dict] = []  # [{flow_lps: float, head_m: float}, ...]

    def load(self, filepath: str) -> bool:
        """
        Charge et parse un fichier rapport pompe (.rtf ou .txt).
        Retourne True si les données essentielles ont été extraites.
        """
        self.filepath = filepath
        self.raw_text = ""
        self.parsed = {}
        self.errors = []

        ext = os.path.splitext(filepath)[1].lower()
        if ext not in ('.rtf', '.txt'):
            self.errors.append(f"Extension '{ext}' non supportée. Utilisez .rtf ou .txt")
            return False

        try:
            encodings = ['utf-8', 'ansi', 'latin-1', 'cp1252']
            for enc in encodings:
                try:
                    with open(filepath, 'r', encoding=enc) as f:
                        self.raw_text = f.read()
                    break
                except (UnicodeDecodeError, LookupError):
                    continue
            else:
                self.errors.append("Impossible de lire le fichier (encodage non reconnu)")
                return False
        except Exception as exc:
            self.errors.append(f"Erreur de lecture : {exc}")
            return False

        # Extraire le texte lisible
        if ext == '.rtf':
            text = _strip_rtf(self.raw_text)
        else:
            text = self.raw_text

        # Parser les paires label → valeur
        self.parsed = self._extract_key_values(text)

        # Valider la présence des données essentielles
        essential = ["pump_id", "label", "flow_lps", "pump_head_m"]
        missing = [k for k in essential if k not in self.parsed]
        if missing:
            self.errors.append(f"Champs essentiels manquants : {', '.join(missing)}")
            return False

        return True

    def _extract_key_values(self, text: str) -> dict:
        """
        Extrait les paires Label / Valeur depuis le texte brut du rapport RTF.
        Deux phases :
          1. Section Scenario Summary → ID, Label, Downstream Pipe, Speed Factor, Status
          2. Section Pump Data → Flow, Head, Pressures, NPSH
        """
        result = {}
        lines = [l.strip() for l in text.split('\n') if l.strip()]

        # Patterns de valeurs connues (unités, etc.) à ignorer comme valeur
        _UNIT_PATTERNS = {
            "m", "mm", "m³", "l", "l/s", "m³/h", "bar", "bars",
            "sec", "hours", "kW", "kwh", "mg/l", "%", "°", "kg/day",
            "€", "€/kW", "€/kWh", "€/ML", "mL",
        }

        def _is_positioning_number(s: str) -> bool:
            """Vérifie si c'est un nombre de positionnement RTF (négatif, sans décimale)."""
            s = s.strip()
            if not s.startswith("-"):
                return False
            num_part = s[1:]
            if not num_part.isdigit():
                return False
            return int(num_part) > 20

        def _is_skip_line(s: str) -> bool:
            """Vérifie si une ligne doit être ignorée comme valeur."""
            s_lower = s.strip().lower()
            if not s_lower:
                return True
            if _is_positioning_number(s):
                return True
            if s_lower in _UNIT_PATTERNS:
                return True
            if s_lower in ("<none>", "<collection: 0 items>", "none"):
                return True
            if s_lower.startswith("<") and s_lower.endswith(">"):
                return True
            if any(p in s_lower for p in self._KEY_MAP.keys()):
                return True
            return False

        # Identifier les indices des deux sections
        first_report_idx = -1
        second_report_idx = -1
        for idx, line in enumerate(lines):
            if "pump detailed report" in line.lower():
                if first_report_idx == -1:
                    first_report_idx = idx
                else:
                    second_report_idx = idx
                    break

        # Phase 1 : Identifier les champs depuis la première section
        # Chercher spécifiquement la section <General> qui contient ID/Label
        _phase1_keys = {"pump_id", "label", "downstream_pipe", "speed_factor", "status_initial"}
        if first_report_idx >= 0:
            # Trouver la section <General>
            general_start = -1
            for idx in range(first_report_idx + 1, second_report_idx if second_report_idx > 0 else len(lines)):
                if lines[idx].strip().lower() == "<general>":
                    general_start = idx
                    break

            search_start = general_start if general_start >= 0 else first_report_idx + 1
            search_end = second_report_idx if second_report_idx > 0 else len(lines)
            i = search_start
            while i < len(lines) - 1 and i < search_end:
                line = lines[i]
                line_lower = line.lower()
                for pattern, key in self._KEY_MAP.items():
                    if key in _phase1_keys and pattern in line_lower and key not in result:
                        val_text = None
                        for j in range(i + 1, min(i + 6, len(lines))):
                            candidate = lines[j]
                            if not _is_skip_line(candidate):
                                val_text = candidate
                                break
                        if val_text is not None:
                            val_clean = val_text.strip()
                            if val_clean.startswith("(") and val_clean.endswith(")"):
                                val_clean = val_clean[1:-1]
                            if val_clean.upper() in ("N/A", "NONE", ""):
                                result[key] = None
                            elif key == "controlled":
                                result[key] = val_clean.lower() in ("true", "yes", "oui", "1")
                            elif key in ("pump_id", "label", "downstream_pipe", "status_initial"):
                                result[key] = val_clean
                            else:
                                parsed = _parse_number(val_clean)
                                if parsed is not None:
                                    result[key] = parsed
                        break
                i += 1

        # Phase 2 : Données opérées depuis la deuxième section
        if second_report_idx >= 0:
            i = second_report_idx + 1
            while i < len(lines) - 1:
                line = lines[i]
                line_lower = line.lower()
                for pattern, key in self._KEY_MAP.items():
                    if key not in _phase1_keys and pattern in line_lower and key not in result:
                        val_text = None
                        for j in range(i + 1, min(i + 6, len(lines))):
                            candidate = lines[j]
                            if not _is_skip_line(candidate):
                                val_text = candidate
                                break
                        if val_text is not None:
                            val_clean = val_text.strip()
                            if val_clean.startswith("(") and val_clean.endswith(")"):
                                val_clean = val_clean[1:-1]
                            if val_clean.upper() in ("N/A", "NONE", ""):
                                result[key] = None
                            elif key == "controlled":
                                result[key] = val_clean.lower() in ("true", "yes", "oui", "1")
                            else:
                                parsed = _parse_number(val_clean)
                                if parsed is not None:
                                    result[key] = parsed
                        break
                i += 1

        return {k: v for k, v in result.items() if not k.startswith("_")}

    def get_curve_points(self) -> list[dict]:
        """Retourne les points de courbe H(Q) saisis manuellement."""
        return list(self.curve_points)

    def add_curve_point(self, flow_lps: float, head_m: float):
        """Ajoute un point à la courbe H(Q)."""
        self.curve_points.append({"flow_lps": flow_lps, "head_m": head_m})
        self.curve_points.sort(key=lambda p: p["flow_lps"])

    def clear_curve_points(self):
        """Vide tous les points de courbe."""
        self.curve_points = []

    def interpolate_head(self, flow_lps: float) -> float | None:
        """
        Interpole la hauteur manométrique H pour un débit donné (L/s)
        en utilisant les points de courbe enregistrés.
        Retourne None si pas assez de points.
        """
        pts = self.curve_points
        if len(pts) < 2:
            return None

        flows = [p["flow_lps"] for p in pts]
        heads = [p["head_m"] for p in pts]

        # Extrapolation linéaire simple
        if flow_lps <= flows[0]:
            # Extrapolation vers le zéro (shutoff)
            if len(pts) >= 2:
                slope = (heads[1] - heads[0]) / (flows[1] - flows[0]) if flows[1] != flows[0] else 0
                return heads[0] + slope * (flow_lps - flows[0])
            return heads[0]

        if flow_lps >= flows[-1]:
            # Extrapolation au-delà du max
            if len(pts) >= 2:
                slope = (heads[-1] - heads[-2]) / (flows[-1] - flows[-2]) if flows[-1] != flows[-2] else 0
                return heads[-1] + slope * (flow_lps - flows[-1])
            return heads[-1]

        # Interpolation linéaire entre deux points
        for j in range(len(flows) - 1):
            if flows[j] <= flow_lps <= flows[j + 1]:
                t = ((flow_lps - flows[j]) / (flows[j + 1] - flows[j])
                     if flows[j + 1] != flows[j] else 0)
                return heads[j] + t * (heads[j + 1] - heads[j])

        return None

    def get_summary(self) -> dict:
        """Retourne un résumé des données extraites pour l'UI."""
        return {
            "label": self.parsed.get("label", "—"),
            "pump_id": self.parsed.get("pump_id", "—"),
            "flow_lps": self.parsed.get("flow_lps"),
            "pump_head_m": self.parsed.get("pump_head_m"),
            "pressure_suction_bar": self.parsed.get("pressure_suction_bar"),
            "pressure_discharge_bar": self.parsed.get("pressure_discharge_bar"),
            "npsh_available_m": self.parsed.get("npsh_available_m"),
            "npsh_required_m": self.parsed.get("npsh_required_m"),
            "downstream_pipe": self.parsed.get("downstream_pipe", "—"),
            "controlled": self.parsed.get("controlled", False),
            "n_curve_points": len(self.curve_points),
        }


# =====================================================================
# 2. GÉNÉRATION DU RAPPORT WORD
# =====================================================================

class WordReportGenerator:
    """
    Génère un rapport d'ingénierie hydraulique professionnel au format Word (.docx)
    à partir des données du projet et des résultats d'analyse.
    """

    # Palette couleur de la charte graphique du rapport
    COULEUR_TITRE_HEX   = (0x1f, 0x53, 0x8d)  # Bleu foncé corporate
    COULEUR_ALERTE_HEX  = (0xc0, 0x00, 0x00)  # Rouge foncé pour les alertes
    COULEUR_OK_HEX      = (0x1a, 0x7a, 0x1a)  # Vert foncé pour les validations

    def __init__(self):
        self.doc = Document()
        self._setup_styles()

    def _rgb(self, rgb_tuple: tuple) -> RGBColor:
        return RGBColor(*rgb_tuple)

    def _setup_styles(self):
        """Paramètre les marges du document."""
        for section in self.doc.sections:
            section.top_margin    = Cm(2.0)
            section.bottom_margin = Cm(2.0)
            section.left_margin   = Cm(2.5)
            section.right_margin  = Cm(2.0)

    def _add_title(self, text: str, level: int = 1, color_hex: tuple | None = None):
        """Ajoute un titre stylisé dans le document Word."""
        p = self.doc.add_heading(text, level=level)
        run = p.runs[0] if p.runs else p.add_run(text)
        run.font.color.rgb = self._rgb(color_hex or self.COULEUR_TITRE_HEX)

    def _add_separator(self):
        """Ajoute une ligne de séparation."""
        p = self.doc.add_paragraph()
        p.add_run("─" * 80).font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)

    def _add_key_value(self, key: str, value: str, bold_value: bool = False):
        """Ajoute une ligne 'Clé : Valeur' dans un paragraphe."""
        p = self.doc.add_paragraph()
        run_key = p.add_run(f"{key} : ")
        run_key.bold = True
        run_val = p.add_run(str(value))
        run_val.bold = bold_value

    def _add_alert(self, text: str, alert_type: str = "warning"):
        """
        Ajoute un paragraphe d'alerte (⚠️ ATTENTION ou ✔ OK).
        alert_type : 'warning' | 'ok' | 'error'
        """
        color_map = {
            "warning": self.COULEUR_ALERTE_HEX,
            "error":   self.COULEUR_ALERTE_HEX,
            "ok":      self.COULEUR_OK_HEX,
        }
        icon_map = {
            "warning": "⚠  ATTENTION : ",
            "error":   "✘  ERREUR : ",
            "ok":      "✔  CONFORME : ",
        }
        p = self.doc.add_paragraph()
        run = p.add_run(icon_map.get(alert_type, "• ") + text)
        run.font.color.rgb = self._rgb(color_map.get(alert_type, (0, 0, 0)))
        run.bold = (alert_type in ("warning", "error"))

    def generate(self,
                 metadata: dict,
                 steady: dict | None,
                 transient: dict | None,
                 pn_label: str,
                 pn_value: float,
                 pmin_label: str,
                 pmin_value: float,
                 chart_png_path: str | None,
                 flow_unit: str = "m³/h",
                 volume_unit: str = "L",
                 volume_threshold_disp: float = 200.0,
                 workbook_summary: dict | None = None,
                 pump_summary: dict | None = None) -> Document:
        """
        Génère le contenu complet du rapport Word.

        Args:
            metadata     : Informations du projet (nom_projet, ingenieur, date).
            steady       : Résultat du parser régime permanent.
            transient    : Résultat du parser analyse transitoire.
            pn_label     : Label de la classe PN (ex: "PN 16").
            pn_value     : Pression nominale en bar.
            pmin_label   : Label de la pression min admissible.
            pmin_value   : Valeur de la pression min en bar.
            flow_unit    : Unité d'affichage pour les débits ("m³/h" ou "L/s").
            volume_unit  : Unité d'affichage pour les volumes ("L" ou "m³").
            volume_threshold_disp : Seuil HPT affiché dans l'unité choisie.
            chart_png_path : Chemin vers l'image PNG du graphique.
            workbook_summary : Résumé du classeur HAMMER (dict du WorkbookManager.get_summary()).
            pump_summary    : Résumé du rapport pompe (dict du PumpReportParser.get_summary()).

        Returns:
            Le Document Word peuplé.
        """
        # ── Pré-formattage des valeurs avec les unités choisies ───────
        if steady and steady.get("success"):
            flow_raw = steady.get("flow_rate_m3h")
            if flow_raw is not None:
                flow_disp_val = flow_raw * FLOW_UNITS.get(flow_unit, 1.0)
                flow_disp_str = f"{flow_disp_val:.2f} {flow_unit}"
            else:
                flow_disp_str = "Non extrait – vérifier le fichier"
        else:
            flow_disp_str = "Non extrait – vérifier le fichier"

        if transient and transient.get("success"):
            vgas_raw = transient.get("max_gas_volume_l")
            if vgas_raw is not None:
                vgas_disp_val = vgas_raw / VOLUME_UNITS.get(volume_unit, 1.0)
                vgas_disp_str = f"{vgas_disp_val:.3f} {volume_unit}"
            else:
                vgas_disp_str = "—"
        else:
            vgas_raw = None
            vgas_disp_str = "—"

        threshold_disp_str = f"{volume_threshold_disp:.3f} {volume_unit}"

        now = metadata.get("date", datetime.now().strftime("%d/%m/%Y"))

        # ── En-tête du rapport ──────────────────────────────────────────
        p_title = self.doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_title.add_run("NOTE TECHNIQUE D'ANALYSE HYDRAULIQUE\nCOUP DE BÉLIER — PROTECTION PAR RÉSERVOIR HPT")
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = self._rgb(self.COULEUR_TITRE_HEX)
        self.doc.add_paragraph()  # Espacement

        # ── Tableau de métadonnées du projet ───────────────────────────
        self._add_title("1. Identification du Projet", level=1)
        meta_table = self.doc.add_table(rows=4, cols=2)
        meta_table.style = "Table Grid"
        meta_data = [
            ("Nom du projet",        metadata.get("nom_projet", "—")),
            ("Ingénieur responsable",metadata.get("ingenieur",  "—")),
            ("Date d'établissement", now),
            ("Logiciel source",      "Bentley HAMMER – Résultats transitoires"),
        ]
        for i, (k, v) in enumerate(meta_data):
            row = meta_table.rows[i]
            cell_k = row.cells[0]
            cell_v = row.cells[1]
            cell_k.paragraphs[0].add_run(k).bold = True
            cell_v.paragraphs[0].add_run(v)

        self.doc.add_paragraph()

        # ── Régime Permanent ───────────────────────────────────────────
        self._add_separator()
        self._add_title("2. Caractérisation du Régime Permanent", level=1)

        if steady and steady.get("success"):
            flow = steady.get("flow_rate_m3h")
            hmt  = steady.get("hmt_m")
            self._add_key_value("Débit nominal de dimensionnement (Q)",
                                 flow_disp_str,
                                 bold_value=True)
            self._add_key_value("Hauteur Manométrique Totale (HMT)",
                                 f"{hmt} mCE" if hmt is not None else "Non extraite – vérifier le fichier",
                                 bold_value=True)
            self._add_key_value("Fichier station analysé",
                                 os.path.basename(metadata.get("station_filepath", "—")))
            self._add_key_value("Nombre de lignes lues", str(steady.get("n_rows", "—")))
        else:
            msg = steady.get("message", "—") if steady else "Fichier non chargé"
            self._add_alert(f"Régime permanent non disponible : {msg}", "error")

        self.doc.add_paragraph()

        # ── Modèle hydraulique (classeur HAMMER) ─────────────────────
        if workbook_summary:
            self._add_separator()
            self._add_title("2. Modèle Hydraulique (Classeur HAMMER)", level=1)

            p_intro = self.doc.add_paragraph(
                "Le modèle hydraulique a été extrait du classeur HAMMER (Flex Tables). "
                "Les données ci-dessous résument la structure du réseau analysé."
            )
            p_intro.runs[0].font.color.rgb = RGBColor(0x44, 0x44, 0x44)
            self.doc.add_paragraph()

            # Tableau récapitulatif
            counts = [
                ("Conduites (Pipes)", workbook_summary.get("pipes_count", 0)),
                ("Nœuds (Nodes)", workbook_summary.get("nodes_count", 0)),
                ("Pompes (Pumps)", workbook_summary.get("pumps_count", 0)),
                ("Réservoirs (Reservoirs)", workbook_summary.get("reservoirs_count", 0)),
                ("Réservoirs HPT", workbook_summary.get("hpt_count", 0)),
                ("Ventouses (Air Valves)", workbook_summary.get("air_valves_count", 0)),
            ]

            tbl = self.doc.add_table(rows=len(counts) + 1, cols=2)
            tbl.style = "Table Grid"
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

            hdrs = ["Composant", "Nombre"]
            for j, h in enumerate(hdrs):
                run = tbl.rows[0].cells[j].paragraphs[0].add_run(h)
                run.bold = True
                run.font.color.rgb = self._rgb(self.COULEUR_TITRE_HEX)

            for i, (label, count) in enumerate(counts):
                tbl.rows[i + 1].cells[0].paragraphs[0].add_run(label)
                tbl.rows[i + 1].cells[1].paragraphs[0].add_run(str(count))

            self.doc.add_paragraph()

            # Propriétés des conduites
            materials = workbook_summary.get("materials", [])
            diam_min = workbook_summary.get("diameter_min_mm")
            diam_max = workbook_summary.get("diameter_max_mm")

            if materials or diam_min is not None:
                self._add_title("Propriétés des Conduites", level=2)
                if materials:
                    self._add_key_value("Matériaux détectés", ", ".join(materials))
                if diam_min is not None and diam_max is not None:
                    self._add_key_value("Diamètres",
                                        f"{diam_min} mm — {diam_max} mm")
                self.doc.add_paragraph()

            # Pressions globales du modèle
            pmax_model = workbook_summary.get("pmax_bar")
            pmin_model = workbook_summary.get("pmin_bar")

            if pmax_model is not None or pmin_model is not None:
                self._add_title("Pressions Transitoires (Modèle)", level=2)
                if pmax_model is not None:
                    pmax_ok = pmax_model <= pn_value
                    self._add_key_value("P max transitoire (modèle)",
                                        f"{pmax_model:.2f} bar")
                    self._add_alert(
                        f"P max ({pmax_model:.2f} bar) ≤ {pn_label} ({pn_value} bar)"
                        if pmax_ok else
                        f"P max ({pmax_model:.2f} bar) > {pn_label} ({pn_value} bar) — DÉPASSEMENT",
                        "ok" if pmax_ok else "warning"
                    )
                if pmin_model is not None:
                    pmin_ok = pmin_model >= pmin_value
                    self._add_key_value("P min transitoire (modèle)",
                                        f"{pmin_model:.2f} bar")
                    self._add_alert(
                        f"P min ({pmin_model:.2f} bar) ≥ limite ({pmin_value} bar)"
                        if pmin_ok else
                        f"P min ({pmin_model:.2f} bar) < limite ({pmin_value} bar) — DÉPRESSION CRITIQUE",
                        "ok" if pmin_ok else "warning"
                    )
                self.doc.add_paragraph()

            # Débit max pompe
            vmax_pump = workbook_summary.get("vmax_pump_ls")
            if vmax_pump is not None:
                self._add_key_value("Débit max pompe (modèle)",
                                    f"{vmax_pump:.1f} L/s")

            self.doc.add_paragraph()

        # ── Données Pompe (rapport détaillé) ─────────────────────────
        if pump_summary and pump_summary.get("label") and pump_summary["label"] != "—":
            self._add_separator()
            self._add_title("2b. Données Pompe (Rapport Détaillé)", level=1)

            p_intro = self.doc.add_paragraph(
                "Les données ci-dessous proviennent du rapport détaillé pompe exporté depuis "
                "Bentley HAMMER. Elles caractérisent le point de fonctionnement nominal et la "
                "courbe H(Q) de la pompe analysée."
            )
            p_intro.runs[0].font.color.rgb = RGBColor(0x44, 0x44, 0x44)
            self.doc.add_paragraph()

            # Tableau récapitulatif pompe
            pump_rows = [
                ("Identifiant",     str(pump_summary.get("pump_id", "—"))),
                ("Label",           str(pump_summary.get("label", "—"))),
                ("Conduite aval",   str(pump_summary.get("downstream_pipe", "—"))),
                ("Débit nominal (Q)",
                 f"{pump_summary['flow_lps']:.1f} L/s" if pump_summary.get("flow_lps") is not None else "—"),
                ("HMT pompe",
                 f"{pump_summary['pump_head_m']:.1f} m" if pump_summary.get("pump_head_m") is not None else "—"),
                ("Pression aspiration",
                 f"{pump_summary['pressure_suction_bar']:.2f} bar" if pump_summary.get("pressure_suction_bar") is not None else "—"),
                ("Pression refoulement",
                 f"{pump_summary['pressure_discharge_bar']:.2f} bar" if pump_summary.get("pressure_discharge_bar") is not None else "—"),
                ("NPSH disponible",
                 f"{pump_summary['npsh_available_m']:.1f} m" if pump_summary.get("npsh_available_m") is not None else "—"),
                ("NPSH requis",
                 f"{pump_summary['npsh_required_m']:.1f} m" if pump_summary.get("npsh_required_m") is not None else "N/D"),
                ("Points courbe H(Q)",
                 str(pump_summary.get("n_curve_points", 0))),
            ]

            tbl = self.doc.add_table(rows=len(pump_rows) + 1, cols=2)
            tbl.style = "Table Grid"
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

            hdrs = ["Paramètre", "Valeur"]
            for j, h in enumerate(hdrs):
                run = tbl.rows[0].cells[j].paragraphs[0].add_run(h)
                run.bold = True
                run.font.color.rgb = self._rgb(self.COULEUR_TITRE_HEX)

            for i, (label, val) in enumerate(pump_rows):
                tbl.rows[i + 1].cells[0].paragraphs[0].add_run(label)
                tbl.rows[i + 1].cells[1].paragraphs[0].add_run(val)

            self.doc.add_paragraph()

            # Alerte NPSH
            npsh_a = pump_summary.get("npsh_available_m")
            if npsh_a is not None:
                if npsh_a < 3:
                    self._add_alert(
                        f"NPSH disponible faible ({npsh_a:.1f} m). "
                        "Risque de cavitation. Vérifier l'altitude d'aspiration.",
                        "warning"
                    )
                else:
                    self._add_alert(
                        f"NPSH disponible ({npsh_a:.1f} m) suffisant.",
                        "ok"
                    )

            self.doc.add_paragraph()

        # ── Analyse Transitoire ────────────────────────────────────────
        self._add_separator()
        self._add_title("3. Synthèse de l'Analyse Transitoire (Coup de Bélier)", level=1)

        p_intro = self.doc.add_paragraph(
            "L'analyse transitoire par la méthode des caractéristiques a été réalisée via Bentley HAMMER. "
            "Les courbes enveloppes de pression maximale et minimale, ainsi que le volume d'air "
            "maximal dans le réservoir HPT, sont synthétisés ci-dessous."
        )
        p_intro.runs[0].font.color.rgb = RGBColor(0x44, 0x44, 0x44)
        self.doc.add_paragraph()

        if transient and transient.get("success"):
            pmax = transient.get("max_pressure_bar")
            pmin = transient.get("min_pressure_bar")
            vgas = transient.get("max_gas_volume_l")
            cols_found = transient.get("critical_columns_found", [])
            is_sim     = transient.get("is_simulated", False)

            # Tableau de résultats
            res_table = self.doc.add_table(rows=4, cols=3)
            res_table.style = "Table Grid"
            res_table.alignment = WD_TABLE_ALIGNMENT.CENTER

            # En-tête du tableau
            hdrs = ["Grandeur", "Valeur calculée", "Statut par rapport aux limites"]
            for j, h in enumerate(hdrs):
                cell = res_table.rows[0].cells[j]
                run  = cell.paragraphs[0].add_run(h)
                run.bold = True
                run.font.color.rgb = self._rgb(self.COULEUR_TITRE_HEX)

            # Ligne Pression Max
            pmax_ok = (pmax is not None and pmax <= pn_value)
            res_table.rows[1].cells[0].paragraphs[0].add_run("Pression max. (Surpression)").bold = True
            res_table.rows[1].cells[1].paragraphs[0].add_run(f"{pmax} bar" if pmax is not None else "—")
            pmax_status_run = res_table.rows[1].cells[2].paragraphs[0].add_run(
                f"✔ OK (< {pn_label} = {pn_value} bar)" if pmax_ok
                else f"✘ DÉPASSEMENT – {pmax} bar > {pn_label} ({pn_value} bar)"
            )
            pmax_status_run.font.color.rgb = self._rgb(self.COULEUR_OK_HEX if pmax_ok else self.COULEUR_ALERTE_HEX)
            pmax_status_run.bold = not pmax_ok

            # Ligne Pression Min
            pmin_ok = (pmin is not None and pmin >= pmin_value)
            res_table.rows[2].cells[0].paragraphs[0].add_run("Pression min. (Dépression)").bold = True
            res_table.rows[2].cells[1].paragraphs[0].add_run(f"{pmin} bar" if pmin is not None else "—")
            pmin_label_txt = f"✔ OK (> limite = {pmin_value} bar)" if pmin_ok else f"✘ DÉPRESSION CRITIQUE – {pmin} bar < limite ({pmin_value} bar)"
            pmin_status_run = res_table.rows[2].cells[2].paragraphs[0].add_run(pmin_label_txt)
            pmin_status_run.font.color.rgb = self._rgb(self.COULEUR_OK_HEX if pmin_ok else self.COULEUR_ALERTE_HEX)
            pmin_status_run.bold = not pmin_ok

            # Ligne Volume de Gaz HPT (avec unité et seuil dynamiques)
            vgas_ok = (vgas is not None and vgas <= volume_threshold_disp * VOLUME_UNITS.get(volume_unit, 1.0))
            res_table.rows[3].cells[0].paragraphs[0].add_run("Volume de gaz max. (HPT)").bold = True
            res_table.rows[3].cells[1].paragraphs[0].add_run(vgas_disp_str)
            vgas_status_run = res_table.rows[3].cells[2].paragraphs[0].add_run(
                f"✔ OK (≤ {threshold_disp_str})" if vgas_ok
                else f"✘ VOLUME ÉLEVÉ – {vgas_disp_str} > {threshold_disp_str}"
            )
            vgas_status_run.font.color.rgb = self._rgb(self.COULEUR_OK_HEX if vgas_ok else self.COULEUR_ALERTE_HEX)
            vgas_status_run.bold = not vgas_ok

            self.doc.add_paragraph()

            # Colonnes détectées et mention simulation
            if is_sim:
                self._add_alert(
                    "Les colonnes critiques n'ont pas été trouvées dans le fichier. "
                    "Les valeurs affichées sont des PLACEHOLDERS de démonstration. "
                    "Vérifiez le format de votre export Bentley HAMMER.",
                    "warning"
                )
            else:
                self._add_alert(
                    f"Colonnes HAMMER identifiées : {', '.join(cols_found)}.",
                    "ok"
                )

        else:
            msg = transient.get("message", "—") if transient else "Fichier non chargé"
            self._add_alert(f"Données transitoires non disponibles : {msg}", "error")

        self.doc.add_paragraph()

        # ── Interprétation et recommandations ─────────────────────────
        self._add_separator()
        self._add_title("4. Interprétation et Diagnostic de Sécurité", level=1)

        if transient and transient.get("success"):
            pmax = transient.get("max_pressure_bar")
            pmin = transient.get("min_pressure_bar")
            vgas = transient.get("max_gas_volume_l")

            # Diagnostic Surpression
            self._add_title("4.1  Surpression", level=2)
            if pmax is not None:
                if pmax > pn_value:
                    self._add_alert(
                        f"La surpression maximale ({pmax} bar) dépasse la pression nominale de la conduite "
                        f"({pn_label} = {pn_value} bar). Risque d'éclatement ou de détérioration accélérée des joints.\n"
                        "→ Recommandations : Augmenter la classe PN de la conduite ; réduire le débit de fermeture "
                        "des vannes (temps de fermeture plus long) ; installer une soupape de décharge.",
                        "warning"
                    )
                else:
                    self._add_alert(
                        f"Surpression maximale ({pmax} bar) inférieure à la PN choisie ({pn_label}). Situation sécurisée.",
                        "ok"
                    )

            # Diagnostic Dépression
            self._add_title("4.2  Dépression", level=2)
            if pmin is not None:
                if pmin < pmin_value:
                    self._add_alert(
                        f"Une dépression critique ({pmin} bar) est détectée dans la conduite, en-dessous de la "
                        f"limite de sécurité définie ({pmin_value} bar).\n"
                        "Risques associés : intrusion d'air ou d'eau contaminée, collapsus de conduite "
                        "(PEHD/PVC), désamorçage des pompes.\n"
                        "→ Recommandations : Vérifier le pré-gonflage et le volume nominal du réservoir HPT ; "
                        "installer des ventouses automatiques aux points hauts ; ajuster le débit de démarrage des pompes.",
                        "warning"
                    )
                else:
                    self._add_alert(
                        f"Dépression minimale ({pmin} bar) dans les limites admissibles (>{pmin_value} bar).",
                        "ok"
                    )

            # Diagnostic Volume HPT (avec unité et seuil dynamiques)
            self._add_title("4.3  Réservoir HPT (Volume de Gaz)", level=2)
            if vgas is not None:
                threshold_l = volume_threshold_disp * VOLUME_UNITS.get(volume_unit, 1.0)
                if vgas > threshold_l:
                    self._add_alert(
                        f"Le volume d'air maximal dans le réservoir HPT ({vgas_disp_str}) est très élevé. "
                        "Cela peut indiquer un réservoir sous-dimensionné qui se vide totalement lors du transitoire, "
                        "perdant ainsi son efficacité.\n"
                        "→ Recommandation : Augmenter le volume nominal du réservoir HPT.",
                        "warning"
                    )
                else:
                    self._add_alert(
                        f"Volume de gaz maximal au HPT ({vgas_disp_str}) dans les limites admissibles (≤ {threshold_disp_str}).",
                        "ok"
                    )
        else:
            self.doc.add_paragraph("Aucune analyse disponible – charger un fichier transitoire valide.")

        self.doc.add_paragraph()

        # ── Graphique intégré ──────────────────────────────────────────
        if chart_png_path and os.path.exists(chart_png_path):
            self._add_separator()
            self._add_title("5. Visualisation Graphique — Courbes Enveloppes", level=1)
            self.doc.add_picture(chart_png_path, width=Cm(16))
            # Centrer l'image
            last_paragraph = self.doc.paragraphs[-1]
            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        self.doc.add_paragraph()

        # ── Pied de page ───────────────────────────────────────────────
        self._add_separator()
        footer_p = self.doc.add_paragraph()
        footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = footer_p.add_run(
            f"Document généré automatiquement par HammerPy Insight v3.0 — {now}"
        )
        run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        run.font.size = Pt(9)

        return self.doc


# =====================================================================
# 3. INTERFACE GRAPHIQUE (GUI)
# =====================================================================

class HammerPyApp(ctk.CTk):
    """
    Fenêtre principale de l'application de bureau HammerPy Insight.
    Structure moderne basée sur CustomTkinter avec Sidebar de navigation
    et Tabview central.
    """

    def __init__(self):
        super().__init__()

        # ── Propriétés de la fenêtre ────────────────────────────────────
        self.title("HammerPy Insight v3 — Analyse Transitoire Hydraulique")
        self.geometry("1200x750")
        self.minsize(1000, 650)

        # ── Icône de l'application ─────────────────────────────────────
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "hammerpy_icon.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)

        # ── Initialisation du parser ────────────────────────────────────
        self.parser = HammerDataParser()
        self.workbook_manager = WorkbookManager()
        self.pump_parser = PumpReportParser()

        # ── État de l'application ───────────────────────────────────────
        self.station_filepath: str = ""
        self.hpt_filepath: str = ""
        self.workbook_filepath: str = ""
        self.pump_filepath: str = ""
        self.steady_state_status: dict | None = None
        self.transient_status: dict | None = None

        # ── Gestion du projet (fichier .hpi) ───────────────────────────
        self.current_project_path: str | None = None
        self.is_dirty: bool = False
        self.project_creation_date: str = datetime.now().isoformat(timespec="seconds")
        self._suppress_dirty: bool = False  # Anti-boucle lors du chargement

        # ── Unités d'affichage (préférence UI, pas de "dirty") ────────
        self.var_flow_unit       = tk.StringVar(value="m³/h")
        self.var_volume_unit     = tk.StringVar(value="L")
        self.var_volume_threshold = tk.StringVar(value=str(VOLUME_THRESHOLD_L_DEFAULT))
        # Source units (override manuel ; "Auto" = auto-détection)
        self.var_station_src_unit   = tk.StringVar(value="Auto-détection")
        self.var_transient_src_unit = tk.StringVar(value="Auto-détection")

        # ── Objets Matplotlib ───────────────────────────────────────────
        self.canvas = None
        self.toolbar = None
        self.current_fig = None  # Référence à la figure active pour l'export

        # ── Layout de la fenêtre principale ────────────────────────────
        self.grid_columnconfigure(0, weight=0)   # Sidebar fixe
        self.grid_columnconfigure(1, weight=1)   # Zone principale extensible
        self.grid_rowconfigure(0, weight=0)      # Barre de menu supérieure (hauteur fixe)
        self.grid_rowconfigure(1, weight=1)      # Zone principale extensible

        # ── Construction des composants ─────────────────────────────────
        self._create_top_bar()
        self._create_sidebar()
        self._create_main_content()
        self._update_title()

    # ================================================================
    # BARRE DE MENU SUPÉRIEURE (Ouvrir / Enregistrer / Enregistrer sous / Quitter)
    # ================================================================

    def _create_top_bar(self):
        """Crée la barre horizontale de boutons de gestion de projet."""
        self.top_bar = ctk.CTkFrame(self, height=50, corner_radius=0, fg_color=("gray90", "gray17"))
        self.top_bar.grid(row=0, column=0, columnspan=2, sticky="ew", padx=0, pady=0)
        self.top_bar.grid_propagate(False)
        self.top_bar.grid_columnconfigure(4, weight=1)  # Espace central extensible

        # ── Bouton Ouvrir ──────────────────────────────────────────────
        self.btn_open = ctk.CTkButton(
            self.top_bar,
            text="📂  Ouvrir",
            width=130,
            fg_color="#1f538d", hover_color="#14375e",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._open_project,
        )
        self.btn_open.grid(row=0, column=0, padx=(14, 6), pady=8, sticky="w")

        # ── Bouton Enregistrer ────────────────────────────────────────
        self.btn_save = ctk.CTkButton(
            self.top_bar,
            text="💾  Enregistrer",
            width=140,
            fg_color="#2d7d3a", hover_color="#1f5a28",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._save_project_quick,
        )
        self.btn_save.grid(row=0, column=1, padx=6, pady=8, sticky="w")

        # ── Bouton Enregistrer sous ───────────────────────────────────
        self.btn_save_as = ctk.CTkButton(
            self.top_bar,
            text="💾  Enregistrer sous…",
            width=160,
            fg_color="transparent", border_width=1,
            text_color=("gray10", "gray90"),
            hover_color=("gray75", "gray30"),
            font=ctk.CTkFont(size=12),
            command=self._save_project_as,
        )
        self.btn_save_as.grid(row=0, column=2, padx=6, pady=8, sticky="w")

        # ── Séparateur visuel ─────────────────────────────────────────
        sep = ctk.CTkFrame(self.top_bar, width=2, height=30, fg_color="gray40")
        sep.grid(row=0, column=3, padx=14, pady=10, sticky="w")

        # ── Titre central ─────────────────────────────────────────────
        self.lbl_top_title = ctk.CTkLabel(
            self.top_bar,
            text="⚡ HammerPy Insight  v3.0",
            font=ctk.CTkFont(family="Consolas", size=13, weight="bold"),
            text_color="gray",
        )
        self.lbl_top_title.grid(row=0, column=4, padx=10, pady=8, sticky="w")

        # ── Indicateur d'état + chemin projet (colonne extensible droite)
        self.top_bar.grid_columnconfigure(5, weight=1)
        self.lbl_project_state = ctk.CTkLabel(
            self.top_bar,
            text="● Nouveau projet (non sauvegardé)",
            font=ctk.CTkFont(size=11),
            text_color="orange",
            anchor="e",
        )
        self.lbl_project_state.grid(row=0, column=5, padx=10, pady=8, sticky="e")

        # ── Bouton Quitter ────────────────────────────────────────────
        self.btn_quit = ctk.CTkButton(
            self.top_bar,
            text="🚪  Quitter",
            width=110,
            fg_color="#a52828", hover_color="#7a1d1d",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._quit_app,
        )
        self.btn_quit.grid(row=0, column=6, padx=(6, 14), pady=8, sticky="e")

    # ================================================================
    # SIDEBAR
    # ================================================================

    def _create_sidebar(self):
        """Crée la barre latérale gauche de navigation."""
        self.sidebar_frame = ctk.CTkFrame(self, width=240, corner_radius=0)
        self.sidebar_frame.grid(row=1, column=0, sticky="nsew")
        self.sidebar_frame.grid_propagate(False)
        self.sidebar_frame.grid_rowconfigure(15, weight=1)  # Espace élastique en bas

        # Logo
        ctk.CTkLabel(
            self.sidebar_frame,
            text="⚡ HammerPy\nInsight",
            font=ctk.CTkFont(family="Consolas", size=22, weight="bold"),
            justify="center"
        ).grid(row=0, column=0, padx=20, pady=(30, 6))

        ctk.CTkLabel(
            self.sidebar_frame,
            text="v3.0 — Bentley HAMMER",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        ).grid(row=1, column=0, padx=20, pady=(0, 24))

        # Séparateur
        ctk.CTkFrame(self.sidebar_frame, height=1, fg_color="gray40").grid(
            row=2, column=0, padx=15, sticky="ew")

        # Section Paramètres de la conduite
        ctk.CTkLabel(
            self.sidebar_frame,
            text="CONDUITE",
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color="gray"
        ).grid(row=3, column=0, padx=20, pady=(18, 4), sticky="w")

        # Classe PN
        ctk.CTkLabel(self.sidebar_frame, text="Classe PN :", anchor="w").grid(
            row=4, column=0, padx=20, pady=(4, 0), sticky="w")
        self.var_pn = tk.StringVar(value="PN 16")
        self.opt_pn = ctk.CTkOptionMenu(
            self.sidebar_frame,
            values=list(PN_CLASSES.keys()),
            variable=self.var_pn,
            command=self._on_pn_change
        )
        self.opt_pn.grid(row=5, column=0, padx=20, pady=(2, 8), sticky="ew")

        # Pression min admissible
        ctk.CTkLabel(self.sidebar_frame, text="P min admissible :", anchor="w").grid(
            row=6, column=0, padx=20, pady=(4, 0), sticky="w")
        self.var_pmin = tk.StringVar(value=list(PMIN_OPTIONS.keys())[1])
        self.opt_pmin = ctk.CTkOptionMenu(
            self.sidebar_frame,
            values=list(PMIN_OPTIONS.keys()),
            variable=self.var_pmin,
            command=self._on_pmin_change,
            width=200
        )
        self.opt_pmin.grid(row=7, column=0, padx=20, pady=(2, 8), sticky="new")

        # Séparateur
        ctk.CTkFrame(self.sidebar_frame, height=1, fg_color="gray40").grid(
            row=8, column=0, padx=15, pady=8, sticky="ew")

        # Thème
        ctk.CTkLabel(
            self.sidebar_frame,
            text="AFFICHAGE",
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color="gray"
        ).grid(row=9, column=0, padx=20, pady=(4, 4), sticky="w")

        ctk.CTkLabel(self.sidebar_frame, text="Mode :", anchor="w").grid(
            row=10, column=0, padx=20, pady=(0, 0), sticky="w")
        self.theme_opt = ctk.CTkOptionMenu(
            self.sidebar_frame,
            values=["Dark", "Light", "System"],
            command=self._change_appearance_mode
        )
        self.theme_opt.grid(row=11, column=0, padx=20, pady=(2, 12), sticky="ew")

        # Séparateur
        ctk.CTkFrame(self.sidebar_frame, height=1, fg_color="gray40").grid(
            row=12, column=0, padx=15, pady=4, sticky="ew")

        # Section Unités
        ctk.CTkLabel(
            self.sidebar_frame,
            text="UNITÉS D'AFFICHAGE",
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color="gray"
        ).grid(row=13, column=0, padx=20, pady=(8, 4), sticky="w")

        ctk.CTkLabel(self.sidebar_frame, text="Débit :", anchor="w").grid(
            row=14, column=0, padx=20, pady=(2, 0), sticky="w")
        self.opt_flow_unit = ctk.CTkOptionMenu(
            self.sidebar_frame,
            values=list(FLOW_UNITS.keys()),
            variable=self.var_flow_unit,
            command=self._on_flow_unit_change,
            width=200
        )
        self.opt_flow_unit.grid(row=15, column=0, padx=20, pady=(2, 6), sticky="new")

        ctk.CTkLabel(self.sidebar_frame, text="Volume :", anchor="w").grid(
            row=16, column=0, padx=20, pady=(4, 0), sticky="w")
        self.opt_volume_unit = ctk.CTkOptionMenu(
            self.sidebar_frame,
            values=list(VOLUME_UNITS.keys()),
            variable=self.var_volume_unit,
            command=self._on_volume_unit_change,
            width=200
        )
        self.opt_volume_unit.grid(row=17, column=0, padx=20, pady=(2, 6), sticky="new")

        ctk.CTkLabel(self.sidebar_frame, text="Seuil HPT :", anchor="w").grid(
            row=18, column=0, padx=20, pady=(4, 0), sticky="w")
        self.entry_volume_threshold = ctk.CTkEntry(
            self.sidebar_frame,
            textvariable=self.var_volume_threshold,
            width=200
        )
        self.entry_volume_threshold.grid(row=19, column=0, padx=20, pady=(2, 4), sticky="new")
        self.lbl_threshold_unit = ctk.CTkLabel(
            self.sidebar_frame,
            text="(en L)",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        )
        self.lbl_threshold_unit.grid(row=20, column=0, padx=20, pady=(0, 6), sticky="new")

        # Hooks de modification
        self.var_volume_threshold.trace_add("write",
            lambda *_: self._on_threshold_change())

        # Bouton aide
        ctk.CTkButton(
            self.sidebar_frame,
            text="? Aide",
            fg_color="transparent",
            border_width=1,
            text_color=("gray20", "gray80"),
            hover_color=("gray75", "gray30"),
            command=self._show_help
        ).grid(row=21, column=0, padx=20, pady=(4, 20), sticky="s")

    # ================================================================
    # TABVIEW PRINCIPAL
    # ================================================================

    def _create_main_content(self):
        """Crée la zone centrale avec les onglets."""
        self.tabview = ctk.CTkTabview(self, corner_radius=10)
        self.tabview.grid(row=1, column=1, padx=(10, 18), pady=15, sticky="nsew")

        for name in ["Régime Permanent", "Analyse Transitoire", "Rapport Technique"]:
            self.tabview.add(name)

        self.tabview._segmented_button.configure(font=ctk.CTkFont(size=13, weight="bold"))

        self._setup_steady_state_tab()
        self._setup_transient_tab()
        self._setup_report_tab()

        # ── Hooks de détection de modifications utilisateur ──────────
        # Entry du projet / ingénieur → dirty à chaque frappe
        self.entry_projet.bind("<KeyRelease>",     lambda _e: self._mark_dirty())
        self.entry_projet.bind("<FocusOut>",       lambda _e: self._mark_dirty())
        self.entry_ingenieur.bind("<KeyRelease>",  lambda _e: self._mark_dirty())
        self.entry_ingenieur.bind("<FocusOut>",    lambda _e: self._mark_dirty())

        # Textbox du rapport → dirty à chaque édition
        def _on_report_modified(_event=None):
            if self.txt_report.edit_modified():
                self._mark_dirty()
                self.txt_report.edit_modified(False)
        self.txt_report.bind("<<Modified>>", _on_report_modified)

    # ──────────────────────────────────────────────────────────────────
    # Onglet 1 : Régime Permanent
    # ──────────────────────────────────────────────────────────────────
    def _setup_steady_state_tab(self):
        tab = self.tabview.tab("Régime Permanent")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(tab, text="Caractérisation du Régime Permanent",
                     font=ctk.CTkFont(size=18, weight="bold"), anchor="w"
                     ).grid(row=0, column=0, padx=20, pady=(18, 4), sticky="w")

        # ── Panneau de configuration du projet ───────────────────────
        cfg_frame = ctk.CTkFrame(tab)
        cfg_frame.grid(row=1, column=0, padx=20, pady=8, sticky="ew")
        cfg_frame.grid_columnconfigure((1, 3), weight=1)

        ctk.CTkLabel(cfg_frame, text="Nom du projet :").grid(
            row=0, column=0, padx=(15, 6), pady=12, sticky="w")
        self.entry_projet = ctk.CTkEntry(cfg_frame, placeholder_text="Ex : AEP Ville-Nord — Adduction principale")
        self.entry_projet.grid(row=0, column=1, padx=(0, 20), pady=12, sticky="ew")

        ctk.CTkLabel(cfg_frame, text="Ingénieur :").grid(
            row=0, column=2, padx=(0, 6), pady=12, sticky="w")
        self.entry_ingenieur = ctk.CTkEntry(cfg_frame, placeholder_text="Nom de l'ingénieur")
        self.entry_ingenieur.grid(row=0, column=3, padx=(0, 15), pady=12, sticky="ew")

        # ── Import du fichier station ─────────────────────────────────
        import_frame = ctk.CTkFrame(tab)
        import_frame.grid(row=2, column=0, padx=20, pady=8, sticky="nsew")
        import_frame.grid_columnconfigure(1, weight=1)
        import_frame.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(
            import_frame,
            text="Chargez le fichier de la station pour extraire le débit nominal (Q) et la HMT.",
            text_color="gray", anchor="w"
        ).grid(row=0, column=0, columnspan=3, padx=18, pady=(16, 6), sticky="w")

        self.btn_import_steady = ctk.CTkButton(
            import_frame, text="  Charger fichier station  (.csv / .xlsx)",
            command=self._import_steady_state_file
        )
        self.btn_import_steady.grid(row=1, column=0, padx=18, pady=8, sticky="w")

        self.lbl_steady_file = ctk.CTkLabel(import_frame, text="Aucun fichier sélectionné",
                                            text_color="gray", anchor="w")
        self.lbl_steady_file.grid(row=1, column=1, padx=8, pady=8, sticky="ew")

        # Sélecteur d'unité source pour le débit
        ctk.CTkLabel(import_frame, text="Unité :", anchor="w",
                     text_color="gray").grid(row=1, column=2, padx=(0, 4), pady=8, sticky="e")
        self.opt_station_src_unit = ctk.CTkOptionMenu(
            import_frame,
            values=["Auto-détection", "m³/h", "L/s"],
            variable=self.var_station_src_unit,
            width=150
        )
        self.opt_station_src_unit.grid(row=1, column=3, padx=(0, 18), pady=8, sticky="e")

        # ── Carte de résultats ────────────────────────────────────────
        kpi_frame = ctk.CTkFrame(import_frame, fg_color=("gray88", "gray14"))
        kpi_frame.grid(row=2, column=0, columnspan=3, padx=18, pady=16, sticky="ew")
        kpi_frame.grid_columnconfigure((0, 1), weight=1)

        for col, title, attr in [(0, "Débit Nominal (Q)", "lbl_flow_val"),
                                  (1, "Hauteur Manométrique Totale (HMT)", "lbl_hmt_val")]:
            ctk.CTkLabel(kpi_frame, text=title, font=ctk.CTkFont(size=12, weight="bold")
                         ).grid(row=0, column=col, padx=24, pady=(18, 2))
            lbl = ctk.CTkLabel(kpi_frame, text="-- --",
                               font=ctk.CTkFont(size=28, weight="bold"), text_color="#1f8ecf")
            lbl.grid(row=1, column=col, padx=24, pady=(2, 18))
            setattr(self, attr, lbl)

        # Zone de détails (colonnes du fichier)
        self.lbl_steady_details = ctk.CTkLabel(
            import_frame, text="", text_color="gray", anchor="w", justify="left")
        self.lbl_steady_details.grid(row=3, column=0, columnspan=3, padx=18, pady=(4, 12), sticky="ew")

    # ──────────────────────────────────────────────────────────────────
    # Onglet 2 : Analyse Transitoire
    # ──────────────────────────────────────────────────────────────────
    def _setup_transient_tab(self):
        tab = self.tabview.tab("Analyse Transitoire")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(6, weight=1)

        ctk.CTkLabel(tab, text="Analyse des Pressions & Volumes Transitoires",
                     font=ctk.CTkFont(size=18, weight="bold"), anchor="w"
                     ).grid(row=0, column=0, padx=20, pady=(18, 4), sticky="w")

        # ── Section 1 : Import HPT (fichier CSV/Excel existant) ──────
        frame = ctk.CTkFrame(tab)
        frame.grid(row=1, column=0, padx=20, pady=8, sticky="ew")
        frame.grid_columnconfigure(1, weight=1)

        self.btn_import_transient = ctk.CTkButton(
            frame, text="  Importer données HPT  (.csv / .xlsx)",
            fg_color="#1f538d", hover_color="#14375e",
            command=self._import_transient_file
        )
        self.btn_import_transient.grid(row=0, column=0, padx=14, pady=12, sticky="w")

        self.lbl_transient_file = ctk.CTkLabel(frame, text="Aucun fichier HPT chargé",
                                               text_color="gray", anchor="w")
        self.lbl_transient_file.grid(row=0, column=1, padx=8, pady=12, sticky="ew")

        ctk.CTkLabel(frame, text="Unité :", anchor="e",
                     text_color="gray").grid(row=0, column=2, padx=(0, 4), pady=12, sticky="e")
        self.opt_transient_src_unit = ctk.CTkOptionMenu(
            frame,
            values=["Auto-détection", "L", "m³"],
            variable=self.var_transient_src_unit,
            width=150
        )
        self.opt_transient_src_unit.grid(row=0, column=3, padx=(0, 14), pady=12, sticky="e")

        # ── KPI strip (pressions + volume) ───────────────────────────
        kpi_strip = ctk.CTkFrame(tab, fg_color=("gray88", "gray14"))
        kpi_strip.grid(row=2, column=0, padx=20, pady=(0, 8), sticky="ew")
        kpi_strip.grid_rowconfigure(0, weight=0)
        kpi_strip.grid_columnconfigure((0, 1, 2), weight=1)

        for col, title, attr, unit in [
            (0, "Pression Min (Dépression)", "lbl_pmin_val", "bar"),
            (1, "Pression Max (Surpression)", "lbl_pmax_val", "bar"),
            (2, "Volume Gaz Max (HPT)", "lbl_vgas_val", "L"),
        ]:
            ctk.CTkLabel(kpi_strip, text=title, font=ctk.CTkFont(size=11)
                         ).grid(row=0, column=col, pady=(12, 0))
            lbl = ctk.CTkLabel(kpi_strip, text=f"-- {unit}",
                               font=ctk.CTkFont(size=20, weight="bold"), text_color="#1f8ecf")
            lbl.grid(row=1, column=col, pady=(0, 12))
            setattr(self, attr, lbl)

        # ── Section 2 : Classeur HAMMER (Flex Tables .xlsx/.xls) ─────
        wb_frame = ctk.CTkFrame(tab, fg_color=("gray88", "gray14"))
        wb_frame.grid(row=3, column=0, padx=20, pady=8, sticky="ew")
        wb_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(wb_frame, text="MODÈLE HAMMER",
                     font=ctk.CTkFont(size=11, weight="bold"), text_color="gray"
                     ).grid(row=0, column=0, columnspan=4, padx=18, pady=(12, 4), sticky="w")

        self.btn_import_workbook = ctk.CTkButton(
            wb_frame, text="  Charger classeur HAMMER  (.xlsx / .xls)",
            fg_color="#2d7d3a", hover_color="#1f5a28",
            command=self._import_workbook
        )
        self.btn_import_workbook.grid(row=1, column=0, padx=14, pady=8, sticky="w")

        self.lbl_workbook_file = ctk.CTkLabel(wb_frame, text="Aucun classeur chargé",
                                              text_color="gray", anchor="w")
        self.lbl_workbook_file.grid(row=1, column=1, padx=8, pady=8, sticky="ew")

        self.btn_reset_workbook = ctk.CTkButton(
            wb_frame, text="Réinitialiser", width=100,
            fg_color="transparent", border_width=1,
            text_color=("gray10", "gray90"),
            hover_color=("gray75", "gray30"),
            state="disabled",
            command=self._reset_workbook
        )
        self.btn_reset_workbook.grid(row=1, column=2, padx=(0, 4), pady=8, sticky="e")

        # ── Compteurs de feuilles ────────────────────────────────────
        counter_frame = ctk.CTkFrame(wb_frame, fg_color="transparent")
        counter_frame.grid(row=2, column=0, columnspan=4, padx=14, pady=(0, 8), sticky="ew")
        counter_frame.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1)

        self._wb_counter_labels = {}
        for col, (key, label) in enumerate([
            ("pipes", "Pipes"), ("nodes", "Nœuds"), ("pumps", "Pompes"),
            ("reservoirs", "Réservoirs"), ("hpt", "HPT"), ("air_valves", "Ventouses"),
        ]):
            ctk.CTkLabel(counter_frame, text=label, font=ctk.CTkFont(size=10),
                         text_color="gray").grid(row=0, column=col, pady=(4, 0))
            lbl = ctk.CTkLabel(counter_frame, text="—",
                               font=ctk.CTkFont(size=14, weight="bold"), text_color="#1f8ecf")
            lbl.grid(row=1, column=col, pady=(0, 4))
            self._wb_counter_labels[key] = lbl

        # ── Mini-stats du modèle ─────────────────────────────────────
        stats_frame = ctk.CTkFrame(wb_frame, fg_color="transparent")
        stats_frame.grid(row=3, column=0, columnspan=4, padx=14, pady=(0, 12), sticky="ew")
        stats_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self._wb_stats = {}
        for col, (key, label, unit) in enumerate([
            ("pmax", "P max", "bar"),
            ("pmin", "P min", "bar"),
            ("vmax_pump", "Q max pompe", "L/s"),
            ("materials", "Matériaux", ""),
        ]):
            ctk.CTkLabel(stats_frame, text=label, font=ctk.CTkFont(size=10),
                         text_color="gray").grid(row=0, column=col, pady=(4, 0))
            lbl = ctk.CTkLabel(stats_frame, text=f"— {unit}" if unit else "—",
                               font=ctk.CTkFont(size=12, weight="bold"), text_color="#1f8ecf")
            lbl.grid(row=1, column=col, pady=(0, 4))
            self._wb_stats[key] = lbl

        # ── Section 3 : Courbe H(Q) Pompe (Rapport détaillé RTF/TXT) ─
        pump_frame = ctk.CTkFrame(tab, fg_color=("gray88", "gray14"))
        pump_frame.grid(row=4, column=0, padx=20, pady=8, sticky="ew")
        pump_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(pump_frame, text="COURBE H(Q) POMPE",
                     font=ctk.CTkFont(size=11, weight="bold"), text_color="gray"
                     ).grid(row=0, column=0, columnspan=4, padx=18, pady=(12, 4), sticky="w")

        self.btn_import_pump = ctk.CTkButton(
            pump_frame, text="  Charger rapport pompe  (.rtf / .txt)",
            fg_color="#8b5cf6", hover_color="#6d28d9",
            command=self._import_pump_report
        )
        self.btn_import_pump.grid(row=1, column=0, padx=14, pady=8, sticky="w")

        self.lbl_pump_file = ctk.CTkLabel(pump_frame, text="Aucun rapport pompe chargé",
                                          text_color="gray", anchor="w")
        self.lbl_pump_file.grid(row=1, column=1, padx=8, pady=8, sticky="ew")

        self.btn_reset_pump = ctk.CTkButton(
            pump_frame, text="Réinitialiser", width=100,
            fg_color="transparent", border_width=1,
            text_color=("gray10", "gray90"),
            hover_color=("gray75", "gray30"),
            state="disabled",
            command=self._reset_pump_report
        )
        self.btn_reset_pump.grid(row=1, column=2, padx=(0, 4), pady=8, sticky="e")

        # ── KPI strip pompe ──────────────────────────────────────────
        pump_kpi = ctk.CTkFrame(pump_frame, fg_color="transparent")
        pump_kpi.grid(row=2, column=0, columnspan=4, padx=14, pady=(0, 8), sticky="ew")
        pump_kpi.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        self._pump_kpi = {}
        for col, (key, label, unit) in enumerate([
            ("label", "Pompe", ""),
            ("flow", "Q nom.", "L/s"),
            ("head", "HMT pompe", "m"),
            ("npsh_a", "NPSH dispo.", "m"),
            ("n_pts", "Pts courbe", ""),
        ]):
            ctk.CTkLabel(pump_kpi, text=label, font=ctk.CTkFont(size=10),
                         text_color="gray").grid(row=0, column=col, pady=(4, 0))
            lbl = ctk.CTkLabel(pump_kpi, text=f"— {unit}" if unit else "—",
                               font=ctk.CTkFont(size=12, weight="bold"), text_color="#8b5cf6")
            lbl.grid(row=1, column=col, pady=(0, 4))
            self._pump_kpi[key] = lbl

        # ── Saisie des points de courbe H(Q) ──────────────────────
        pts_label = ctk.CTkLabel(pump_frame, text="SAISIE DES POINTS DE COURBE",
                                 font=ctk.CTkFont(size=10, weight="bold"), text_color="gray")
        pts_label.grid(row=3, column=0, columnspan=4, padx=18, pady=(8, 2), sticky="w")

        pts_row = ctk.CTkFrame(pump_frame, fg_color="transparent")
        pts_row.grid(row=4, column=0, columnspan=4, padx=14, pady=(0, 4), sticky="ew")
        pts_row.grid_columnconfigure(0, weight=0)
        pts_row.grid_columnconfigure(1, weight=0)
        pts_row.grid_columnconfigure(2, weight=0)
        pts_row.grid_columnconfigure(3, weight=0)
        pts_row.grid_columnconfigure(4, weight=0)
        pts_row.grid_columnconfigure(5, weight=1)

        ctk.CTkLabel(pts_row, text="Q (L/s) :", font=ctk.CTkFont(size=10),
                     text_color="gray").grid(row=0, column=0, padx=(0, 4), pady=4, sticky="e")
        self.entry_pump_q = ctk.CTkEntry(pts_row, width=80, placeholder_text="ex: 75")
        self.entry_pump_q.grid(row=0, column=1, padx=(0, 12), pady=4, sticky="w")

        ctk.CTkLabel(pts_row, text="H (m) :", font=ctk.CTkFont(size=10),
                     text_color="gray").grid(row=0, column=2, padx=(0, 4), pady=4, sticky="e")
        self.entry_pump_h = ctk.CTkEntry(pts_row, width=80, placeholder_text="ex: 150")
        self.entry_pump_h.grid(row=0, column=3, padx=(0, 12), pady=4, sticky="w")

        self.btn_add_pump_point = ctk.CTkButton(
            pts_row, text="+ Ajouter", width=90,
            fg_color="#8b5cf6", hover_color="#6d28d9",
            state="disabled",
            command=self._on_add_pump_point
        )
        self.btn_add_pump_point.grid(row=0, column=4, padx=(0, 6), pady=4, sticky="w")

        self.btn_clear_pump_points = ctk.CTkButton(
            pts_row, text="Effacer tout", width=90,
            fg_color="transparent", border_width=1,
            text_color=("gray10", "gray90"),
            hover_color=("gray75", "gray30"),
            state="disabled",
            command=self._on_clear_pump_points
        )
        self.btn_clear_pump_points.grid(row=0, column=5, padx=0, pady=4, sticky="w")

        # ── Liste des points saisis ────────────────────────────────
        self.txt_pump_points = ctk.CTkTextbox(
            pump_frame, height=60, font=ctk.CTkFont(family="Consolas", size=10),
            state="disabled", fg_color=("gray92", "gray18")
        )
        self.txt_pump_points.grid(row=5, column=0, columnspan=4, padx=14, pady=(0, 8), sticky="ew")

        # ── Graphique H(Q) ─────────────────────────────────────────
        self.pump_curve_frame = ctk.CTkFrame(tab, fg_color=("gray85", "gray12"))
        self.pump_curve_frame.grid(row=5, column=0, padx=20, pady=(0, 8), sticky="nsew")
        self.pump_curve_frame.grid_rowconfigure(0, weight=1)
        self.pump_curve_frame.grid_columnconfigure(0, weight=1)

        self.lbl_pump_curve_placeholder = ctk.CTkLabel(
            self.pump_curve_frame,
            text="📈   La courbe H(Q) s'affichera ici après saisie de ≥ 2 points.",
            text_color="gray", justify="center"
        )
        self.lbl_pump_curve_placeholder.grid(row=0, column=0, padx=20, pady=30, sticky="nsew")

        self.pump_curve_canvas = None
        self.pump_curve_toolbar = None

        # ── Zone graphique Matplotlib (HPT) ────────────────────────
        self.plot_placeholder_frame = ctk.CTkFrame(tab, fg_color=("gray85", "gray12"))
        self.plot_placeholder_frame.grid(row=6, column=0, padx=20, pady=(0, 12), sticky="nsew")
        self.plot_placeholder_frame.grid_rowconfigure(0, weight=1)
        self.plot_placeholder_frame.grid_columnconfigure(0, weight=1)

        self.lbl_graph_placeholder = ctk.CTkLabel(
            self.plot_placeholder_frame,
            text="📈   Le graphique interactif s'affichera ici après l'importation du fichier HPT.",
            text_color="gray", justify="center"
        )
        self.lbl_graph_placeholder.grid(row=0, column=0, padx=20, pady=40, sticky="nsew")

    # ──────────────────────────────────────────────────────────────────
    # Onglet 3 : Rapport Technique
    # ──────────────────────────────────────────────────────────────────
    def _setup_report_tab(self):
        tab = self.tabview.tab("Rapport Technique")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(tab, text="Rapport Technique — Prévisualisation & Export",
                     font=ctk.CTkFont(size=18, weight="bold"), anchor="w"
                     ).grid(row=0, column=0, padx=20, pady=(18, 4), sticky="w")

        frame = ctk.CTkFrame(tab)
        frame.grid(row=1, column=0, padx=20, pady=8, sticky="nsew")
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            frame,
            text="Prévisualisation du rapport.  Vous pouvez éditer le texte avant export.",
            text_color="gray", anchor="w"
        ).grid(row=0, column=0, padx=18, pady=(12, 4), sticky="w")

        self.txt_report = ctk.CTkTextbox(frame, font=ctk.CTkFont(family="Consolas", size=12), wrap="word")
        self.txt_report.grid(row=1, column=0, padx=18, pady=(4, 8), sticky="nsew")
        self._update_report_preview()

        # Barre de boutons d'export
        btn_bar = ctk.CTkFrame(frame, fg_color="transparent")
        btn_bar.grid(row=2, column=0, padx=18, pady=(4, 16), sticky="ew")
        btn_bar.grid_columnconfigure(0, weight=1)

        self.btn_export_txt = ctk.CTkButton(
            btn_bar, text="Export .txt",
            fg_color="transparent", border_width=1,
            text_color=("gray10", "gray90"),
            command=self._export_txt
        )
        self.btn_export_txt.grid(row=0, column=1, padx=6)

        self.btn_export_docx = ctk.CTkButton(
            btn_bar,
            text="  Générer Rapport Word (.docx)  ",
            fg_color="#1f538d", hover_color="#14375e",
            font=ctk.CTkFont(weight="bold"),
            state="normal" if DOCX_AVAILABLE else "disabled",
            command=self._export_word
        )
        self.btn_export_docx.grid(row=0, column=2, padx=6)

        if not DOCX_AVAILABLE:
            ctk.CTkLabel(btn_bar, text="python-docx non installé",
                         text_color="orange").grid(row=0, column=3, padx=6)

    # ================================================================
    # ACTIONS – Handlers événements
    # ================================================================

    def _get_pn_value(self) -> float:
        return PN_CLASSES.get(self.var_pn.get(), 16.0)

    def _get_pmin_value(self) -> float:
        return PMIN_OPTIONS.get(self.var_pmin.get(), -0.1)

    def _on_pn_change(self, _=None):
        """Mise à jour du graphique quand la PN change."""
        self._update_chart()
        self._update_report_preview()
        self._mark_dirty()

    def _on_pmin_change(self, _=None):
        """Mise à jour du graphique quand la pression min change."""
        self._update_chart()
        self._update_report_preview()
        self._mark_dirty()

    # ------------------------------------------------------------------
    # Gestion des unités d'affichage
    # ------------------------------------------------------------------
    def _on_flow_unit_change(self, _=None):
        """Mise à jour visuelle après changement d'unité de débit."""
        self._update_chart()
        self._update_report_preview()
        # Pas de _mark_dirty : préférence d'affichage, pas une donnée métier.

    def _on_volume_unit_change(self, _=None):
        """Mise à jour visuelle après changement d'unité de volume.
        Met aussi à jour le label d'unité du seuil HPT."""
        self._refresh_threshold_unit_label()
        self._update_chart()
        self._update_report_preview()

    def _on_threshold_change(self, _=None):
        """Mise à jour après édition du seuil HPT (donnée métier → dirty)."""
        self._update_chart()
        self._update_report_preview()
        self._mark_dirty()

    # ------------------------------------------------------------------
    # Helpers de conversion et formatage
    # ------------------------------------------------------------------
    def _format_flow(self, value_m3h):
        """Formate un débit stocké en m³/h (canonique) vers l'unité d'affichage."""
        if value_m3h is None:
            return "—"
        unit = self.var_flow_unit.get()
        factor = FLOW_UNITS.get(unit, 1.0)
        return f"{value_m3h * factor:.2f} {unit}"

    def _format_volume(self, value_l):
        """Formate un volume stocké en L (canonique) vers l'unité d'affichage."""
        if value_l is None:
            return "—"
        unit = self.var_volume_unit.get()
        factor = VOLUME_UNITS.get(unit, 1.0)
        return f"{value_l / factor:.3f} {unit}"

    def _threshold_internal_l(self) -> float:
        """Renvoie la valeur du seuil HPT stockée canoniquement en L."""
        try:
            displayed = float(self.var_volume_threshold.get())
            return displayed * VOLUME_UNITS.get(self.var_volume_unit.get(), 1.0)
        except (ValueError, TypeError):
            return VOLUME_THRESHOLD_L_DEFAULT

    def _volume_threshold_display(self) -> float:
        """Renvoie le seuil HPT dans l'unité d'affichage choisie."""
        return self._threshold_internal_l() / VOLUME_UNITS.get(
            self.var_volume_unit.get(), 1.0)

    def _refresh_threshold_unit_label(self):
        """Met à jour le label '(en L)' ou '(en m³)' à côté du champ seuil."""
        if hasattr(self, "lbl_threshold_unit"):
            self.lbl_threshold_unit.configure(
                text=f"(en {self.var_volume_unit.get()})")

    def _detect_source_unit(self, df, kind: str) -> str:
        """Auto-détecte l'unité source d'un DataFrame à partir de ses colonnes.
        kind = 'flow' ou 'volume'. Retourne l'unité devinée (fallback canonique)."""
        hints = FLOW_UNIT_HINTS if kind == "flow" else VOLUME_UNIT_HINTS
        default = "m³/h" if kind == "flow" else "L"
        if df is None or df.columns is None:
            return default
        for col in df.columns:
            col_lower = str(col).lower()
            for pattern, unit in hints.items():
                if pattern in col_lower:
                    return unit
        return default

    def _resolve_station_src_unit(self, df=None) -> str:
        """Renvoie l'unité source effective (override > auto > canonique)."""
        sel = self.var_station_src_unit.get()
        if sel == "Auto-détection":
            if df is None:
                try:
                    df = self.parser.steady_state_data
                except Exception:
                    df = None
            return self._detect_source_unit(df, "flow")
        return sel

    def _resolve_transient_src_unit(self, df=None) -> str:
        """Renvoie l'unité source effective pour le volume HPT."""
        sel = self.var_transient_src_unit.get()
        if sel == "Auto-détection":
            if df is None:
                try:
                    df = self.parser.hpt_data
                except Exception:
                    df = None
            return self._detect_source_unit(df, "volume")
        return sel

    def _change_appearance_mode(self, mode: str):
        ctk.set_appearance_mode(mode)
        self._update_chart()

    def _show_help(self):
        messagebox.showinfo(
            "Aide — HammerPy Insight v3",
            "WORKFLOW :\n\n"
            "1. Renseignez le nom du projet et l'ingénieur en charge (Onglet 1).\n"
            "2. Choisissez la classe PN de la conduite et la pression min admissible (Sidebar).\n"
            "3. Chargez le fichier station CSV/Excel (Onglet 1 — Régime Permanent).\n"
            "4. Chargez le fichier d'enveloppe HPT CSV/Excel (Onglet 2 — Analyse Transitoire).\n"
            "   → Le graphique interactif avec les seuils critiques s'affiche automatiquement.\n"
            "5. Consultez le rapport dans l'Onglet 3 et exportez en Word (.docx) ou en texte (.txt).\n\n"
            "FORMATS SUPPORTÉS :\n"
            "  • CSV : séparateurs , / ; / Tab – décimales . ou ,\n"
            "  • Excel : .xlsx, .xls\n"
            "  • Encodages : UTF-8, Latin-1, CP1252"
        )

    # ── Import Station ────────────────────────────────────────────────
    def _import_steady_state_file(self):
        filepath = filedialog.askopenfilename(
            title="Sélectionner le fichier de la Station",
            filetypes=[("CSV/Excel", "*.csv *.xlsx *.xls"), ("Tous", "*.*")]
        )
        if not filepath:
            return

        self.station_filepath = filepath
        self.lbl_steady_file.configure(text=os.path.basename(filepath),
                                       text_color=("gray20", "gray80"))
        src_unit = self._resolve_station_src_unit()
        data = self.parser.parse_station_file(filepath, source_unit=src_unit)
        self.steady_state_status = data

        if data["success"]:
            q   = data.get("flow_rate_m3h")
            hmt = data.get("hmt_m")
            self.lbl_flow_val.configure(
                text=self._format_flow(q),
                text_color="#1f8ecf" if q is not None else "orange"
            )
            self.lbl_hmt_val.configure(
                text=f"{hmt} mCE" if hmt is not None else "—",
                text_color="#1f8ecf" if hmt is not None else "orange"
            )
            # Détails des colonnes
            cols_txt = "Colonnes détectées : " + ", ".join(data.get("columns", []))
            self.lbl_steady_details.configure(text=cols_txt[:120] + ("…" if len(cols_txt) > 120 else ""))
            self._update_report_preview()
            self._mark_dirty()
            messagebox.showinfo("Import réussi",
                                f"Fichier chargé : {len(data.get('columns', []))} colonne(s) — {data['n_rows']} ligne(s).\n"
                                f"{'✔ Q et HMT extraits.' if q and hmt else '⚠ Colonnes Q/HMT non trouvées automatiquement.'}")
        else:
            self.lbl_flow_val.configure(text="Erreur", text_color="tomato")
            self.lbl_hmt_val.configure(text="Erreur", text_color="tomato")
            self.lbl_steady_details.configure(text="")
            self._update_report_preview()
            messagebox.showerror("Erreur de chargement",
                                 f"Parsing échoué :\n{data['message']}")

    # ── Import HPT ────────────────────────────────────────────────────
    def _import_transient_file(self):
        filepath = filedialog.askopenfilename(
            title="Sélectionner le fichier transitoire HPT",
            filetypes=[("CSV/Excel", "*.csv *.xlsx *.xls"), ("Tous", "*.*")]
        )
        if not filepath:
            return

        self.hpt_filepath = filepath
        self.lbl_transient_file.configure(text=os.path.basename(filepath),
                                          text_color=("gray20", "gray80"))
        src_unit = self._resolve_transient_src_unit()
        data = self.parser.parse_hpt_file(filepath, source_unit=src_unit)
        self.transient_status = data

        if data["success"]:
            self.lbl_pmin_val.configure(text=f"{data['min_pressure_bar']} bar",
                                        text_color=self._pression_color(data['min_pressure_bar'], "min"))
            self.lbl_pmax_val.configure(text=f"{data['max_pressure_bar']} bar",
                                        text_color=self._pression_color(data['max_pressure_bar'], "max"))
            self.lbl_vgas_val.configure(
                text=self._format_volume(data['max_gas_volume_l']),
                text_color="#33a02c" if data['max_gas_volume_l'] <= 200 else "orange")
            self._update_chart()
            self._update_report_preview()
            self._mark_dirty()
            sim_note = "\n⚠ Les valeurs sont des placeholders (colonnes non détectées)." if data['is_simulated'] else ""
            messagebox.showinfo("Import réussi",
                                f"Fichier chargé : {data['n_rows']} point(s).\n"
                                f"{data['message']}{sim_note}")
        else:
            for lbl in (self.lbl_pmin_val, self.lbl_pmax_val, self.lbl_vgas_val):
                lbl.configure(text="Erreur", text_color="tomato")
            self.lbl_graph_placeholder.configure(
                text="❌  Échec du chargement — Consultez le rapport pour le diagnostic.",
                text_color="tomato"
            )
            if self.canvas:
                self.canvas.get_tk_widget().destroy()
                self.canvas = None
            self.lbl_graph_placeholder.grid(row=0, column=0, padx=20, pady=40, sticky="nsew")
            self._update_report_preview()
            messagebox.showerror("Erreur de chargement", data["message"])

    def _pression_color(self, value: float, kind: str) -> str:
        """Retourne la couleur d'un KPI selon qu'il dépasse la limite ou non."""
        if kind == "min":
            return "#e87c2a" if value < self._get_pmin_value() else "#33a02c"
        if kind == "max":
            return "#e87c2a" if value > self._get_pn_value() else "#33a02c"
        return "#1f8ecf"

    # ------------------------------------------------------------------
    # Import du classeur HAMMER (Flex Tables)
    # ------------------------------------------------------------------
    def _import_workbook(self):
        """Charge un classeur HAMMER (.xlsx/.xls) et affiche le résumé."""
        filepath = filedialog.askopenfilename(
            title="Sélectionner le classeur HAMMER (Flex Tables)",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Tous", "*.*")]
        )
        if not filepath:
            return

        ok = self.workbook_manager.load(filepath)
        valid, errors = self.workbook_manager.validate()

        if not valid:
            self.lbl_workbook_file.configure(
                text="❌ " + "; ".join(errors[:2]),
                text_color="tomato"
            )
            messagebox.showerror(
                "Classeur invalide",
                "Le classeur ne contient pas les feuilles obligatoires :\n\n"
                + "\n".join(errors))
            return

        self.workbook_filepath = filepath
        self.lbl_workbook_file.configure(
            text=os.path.basename(filepath),
            text_color=("gray20", "gray80")
        )
        self.btn_reset_workbook.configure(state="normal")

        # Mettre à jour les compteurs
        summary = self.workbook_manager.get_summary()
        for key, lbl in self._wb_counter_labels.items():
            count = summary.get(f"{key}_count", 0)
            lbl.configure(text=str(count) if count > 0 else "—",
                          text_color="#33a02c" if count > 0 else "gray")

        # Mettre à jour les mini-stats
        pmax = summary.get("pmax_bar")
        pmin = summary.get("pmin_bar")
        vmax = summary.get("vmax_pump_ls")
        mats = summary.get("materials", [])

        self._wb_stats["pmax"].configure(
            text=f"{pmax:.2f} bar" if pmax is not None else "— bar",
            text_color="#e87c2a" if pmax and pmax > self._get_pn_value() else "#33a02c")
        self._wb_stats["pmin"].configure(
            text=f"{pmin:.2f} bar" if pmin is not None else "— bar",
            text_color="#e87c2a" if pmin and pmin < self._get_pmin_value() else "#33a02c")
        self._wb_stats["vmax_pump"].configure(
            text=f"{vmax:.1f} L/s" if vmax is not None else "— L/s")
        self._wb_stats["materials"].configure(
            text=", ".join(mats) if mats else "—")

        self._mark_dirty()

        # Avertissements
        warn_msg = ""
        if self.workbook_manager.warnings:
            warn_msg = "\n\nAvertissements :\n" + "\n".join(self.workbook_manager.warnings)

        messagebox.showinfo(
            "Classeur chargé",
            f"Feuilles trouvées : {len(self.workbook_manager.sheet_map)}\n"
            f"Pipes: {summary['pipes_count']} | Nœuds: {summary['nodes_count']} | "
            f"Pompes: {summary['pumps_count']}\n"
            f"Réservoirs: {summary['reservoirs_count']} | HPT: {summary['hpt_count']} | "
            f"Ventouses: {summary['air_valves_count']}"
            + warn_msg
        )

    def _reset_workbook(self):
        """Réinitialise le classeur chargé."""
        self.workbook_manager = WorkbookManager()
        self.workbook_filepath = ""
        self.lbl_workbook_file.configure(text="Aucun classeur chargé", text_color="gray")
        self.btn_reset_workbook.configure(state="disabled")
        for key, lbl in self._wb_counter_labels.items():
            lbl.configure(text="—", text_color="#1f8ecf")
        for key, lbl in self._wb_stats.items():
            unit = {"pmax": "bar", "pmin": "bar", "vmax_pump": "L/s", "materials": ""}.get(key, "")
            lbl.configure(text=f"— {unit}" if unit else "—", text_color="#1f8ecf")
        self._mark_dirty()

    # ------------------------------------------------------------------
    # Import du rapport pompe (RTF / TXT)
    # ------------------------------------------------------------------
    def _import_pump_report(self):
        """Charge un rapport pompe et affiche le résumé."""
        filepath = filedialog.askopenfilename(
            title="Sélectionner le rapport pompe détaillé",
            filetypes=[("Rapport pompe", "*.rtf *.txt"), ("Tous", "*.*")]
        )
        if not filepath:
            return

        ok = self.pump_parser.load(filepath)
        if not ok:
            self.lbl_pump_file.configure(
                text="❌ " + "; ".join(self.pump_parser.errors[:2]),
                text_color="tomato"
            )
            messagebox.showerror(
                "Rapport invalide",
                "Impossible d'extraire les données pompe :\n\n"
                + "\n".join(self.pump_parser.errors))
            return

        self.pump_filepath = filepath
        self.lbl_pump_file.configure(
            text=os.path.basename(filepath),
            text_color=("gray20", "gray80")
        )
        self.btn_reset_pump.configure(state="normal")
        self.btn_add_pump_point.configure(state="normal")
        self.btn_clear_pump_points.configure(state="normal")

        # Mettre à jour les KPI pompe
        summary = self.pump_parser.get_summary()
        self._pump_kpi["label"].configure(text=summary["label"])
        self._pump_kpi["flow"].configure(
            text=f"{summary['flow_lps']:.1f} L/s" if summary["flow_lps"] is not None else "— L/s")
        self._pump_kpi["head"].configure(
            text=f"{summary['pump_head_m']:.1f} m" if summary["pump_head_m"] is not None else "— m")
        npsh_a = summary["npsh_available_m"]
        self._pump_kpi["npsh_a"].configure(
            text=f"{npsh_a:.1f} m" if npsh_a is not None else "— m",
            text_color="#33a02c" if npsh_a is not None and npsh_a > 3 else "#e87c2a")
        self._pump_kpi["n_pts"].configure(text=str(summary["n_curve_points"]))

        self._mark_dirty()
        messagebox.showinfo(
            "Rapport pompe chargé",
            f"Pompe : {summary['label']} (ID {summary['pump_id']})\n"
            f"Débit nominal : {summary['flow_lps']} L/s\n"
            f"HMT pompe : {summary['pump_head_m']} m\n"
            f"NPSH disponible : {summary['npsh_available_m']} m\n\n"
            "Vous pouvez maintenant saisir des points de courbe H(Q) manuellement\n"
            "si vous disposez de la courbe constructeur."
        )

    def _reset_pump_report(self):
        """Réinitialise le rapport pompe chargé."""
        self.pump_parser = PumpReportParser()
        self.pump_filepath = ""
        self.lbl_pump_file.configure(text="Aucun rapport pompe chargé", text_color="gray")
        self.btn_reset_pump.configure(state="disabled")
        self.btn_add_pump_point.configure(state="disabled")
        self.btn_clear_pump_points.configure(state="disabled")
        for key, lbl in self._pump_kpi.items():
            unit = {"flow": "L/s", "head": "m", "npsh_a": "m", "n_pts": ""}.get(key, "")
            lbl.configure(text=f"— {unit}" if unit else "—", text_color="#8b5cf6")
        self._update_pump_points_display()
        self._clear_pump_curve_chart()
        self._mark_dirty()

    def _on_add_pump_point(self):
        """Ajoute un point (Q, H) à la courbe."""
        q_str = self.entry_pump_q.get().strip()
        h_str = self.entry_pump_h.get().strip()
        if not q_str or not h_str:
            messagebox.showwarning("Champs requis", "Veuillez saisir Q (L/s) et H (m).")
            return
        try:
            q = float(q_str.replace(",", "."))
            h = float(h_str.replace(",", "."))
        except ValueError:
            messagebox.showerror("Erreur de saisie",
                                 f"Valeurs invalides :\nQ = « {q_str} »\nH = « {h_str} »\n\n"
                                 "Entrez des nombres décimaux (ex: 75 ou 150,5).")
            return
        if q < 0 or h < 0:
            messagebox.showwarning("Valeur négative", "Q et H doivent être ≥ 0.")
            return

        self.pump_parser.add_curve_point(q, h)
        self.entry_pump_q.delete(0, "end")
        self.entry_pump_h.delete(0, "end")
        self.entry_pump_q.focus_set()
        self._update_pump_points_display()
        self._update_pump_curve_chart()
        self._pump_kpi["n_pts"].configure(text=str(len(self.pump_parser.curve_points)))
        self._mark_dirty()

    def _on_clear_pump_points(self):
        """Efface tous les points de courbe."""
        if self.pump_parser.curve_points:
            if not messagebox.askyesno("Confirmer",
                                       "Effacer tous les points de courbe saisis ?"):
                return
        self.pump_parser.clear_curve_points()
        self._update_pump_points_display()
        self._clear_pump_curve_chart()
        self._pump_kpi["n_pts"].configure(text="0")
        self._mark_dirty()

    def _update_pump_points_display(self):
        """Rafraîchit la zone texte avec la liste des points."""
        self.txt_pump_points.configure(state="normal")
        self.txt_pump_points.delete("1.0", "end")
        pts = self.pump_parser.curve_points
        if not pts:
            self.txt_pump_points.insert("1.0", "  (aucun point saisi)")
        else:
            header = f"  {'#':>3}   {'Q (L/s)':>10}   {'H (m)':>10}\n"
            self.txt_pump_points.insert("1.0", header + "  " + "─" * 30 + "\n")
            for i, p in enumerate(pts, 1):
                line = f"  {i:>3}   {p['flow_lps']:>10.1f}   {p['head_m']:>10.1f}\n"
                self.txt_pump_points.insert("end", line)
        self.txt_pump_points.configure(state="disabled")

    def _update_pump_curve_chart(self):
        """Trace / rafraîchit le graphique H(Q) dans l'onglet Analyse Transitoire."""
        # Nettoyage de l'ancien canvas
        if self.pump_curve_canvas:
            self.pump_curve_canvas.get_tk_widget().destroy()
            self.pump_curve_canvas = None
        if self.pump_curve_toolbar:
            self.pump_curve_toolbar.destroy()
            self.pump_curve_toolbar = None

        pts = self.pump_parser.curve_points
        if len(pts) < 2:
            self.lbl_pump_curve_placeholder.grid(
                row=0, column=0, padx=20, pady=30, sticky="nsew")
            return

        self.lbl_pump_curve_placeholder.grid_forget()

        import numpy as np
        from matplotlib.figure import Figure

        q_pts = np.array([p["flow_lps"] for p in pts])
        h_pts = np.array([p["head_m"] for p in pts])

        # Interpolation polynomiale (degré max 3)
        deg = min(len(pts) - 1, 3)
        coeffs = np.polyfit(q_pts, h_pts, deg)
        poly = np.poly1d(coeffs)

        q_min, q_max = q_pts.min(), q_pts.max()
        margin = max((q_max - q_min) * 0.15, 5.0)
        q_smooth = np.linspace(max(0, q_min - margin), q_max + margin, 200)
        h_smooth = np.maximum(poly(q_smooth), 0)

        # Couleurs thème
        mode = ctk.get_appearance_mode()
        if mode == "Dark":
            bg, ax_bg = "#1a1a2e", "#16213e"
            text_c, grid_c, spine_c = "#e0e0e0", "#333355", "#555577"
        else:
            bg, ax_bg = "#f8f9fa", "#ffffff"
            text_c, grid_c, spine_c = "#333333", "#dddddd", "#cccccc"
        curve_c  = "#8b5cf6"
        point_c  = "#f59e0b"

        fig = Figure(figsize=(7, 3.5), dpi=100)
        fig.patch.set_facecolor(bg)
        ax = fig.add_subplot(111)
        ax.set_facecolor(ax_bg)

        ax.plot(q_smooth, h_smooth, color=curve_c, lw=2.5, label="H(Q) interpolée", zorder=3)
        ax.scatter(q_pts, h_pts, color=point_c, s=90, zorder=5, edgecolors="white",
                   linewidths=1.2, label="Points saisis")

        # Point nominal si disponible
        flow_nom = self.pump_parser.parsed.get("flow_lps")
        head_nom = self.pump_parser.parsed.get("pump_head_m")
        if flow_nom and head_nom:
            ax.scatter([flow_nom], [head_nom], color="#ef4444", s=120, zorder=6,
                       edgecolors="white", linewidths=1.5, marker="D",
                       label=f"Nominal ({flow_nom:.0f} L/s, {head_nom:.1f} m)")

        ax.set_xlabel("Débit Q (L/s)", color=text_c, fontweight="bold", fontsize=10)
        ax.set_ylabel("HMT H (m)", color=text_c, fontweight="bold", fontsize=10)
        label_pump = self.pump_parser.parsed.get("label", "Pompe")
        ax.set_title(f"Courbe H(Q) — {label_pump}", color=text_c, fontsize=11, fontweight="bold")
        ax.tick_params(colors=text_c, labelsize=9)
        ax.grid(True, color=grid_c, ls=":", alpha=0.7)
        ax.legend(fontsize=8, loc="upper right", framealpha=0.85,
                  facecolor=ax_bg, labelcolor=text_c, edgecolor=spine_c)
        for sp in ax.spines.values():
            sp.set_color(spine_c)

        fig.tight_layout(pad=1.5)

        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
        import tkinter as tk

        self.pump_curve_canvas = FigureCanvasTkAgg(fig, master=self.pump_curve_frame)
        self.pump_curve_canvas.get_tk_widget().grid(
            row=0, column=0, sticky="nsew", padx=4, pady=4)

        toolbar_bg = bg.replace("#", "")
        self.pump_curve_toolbar = tk.Frame(self.pump_curve_frame, bg=toolbar_bg)
        self.pump_curve_toolbar.grid(row=1, column=0, sticky="ew", padx=4)
        self.pump_curve_toolbar = NavigationToolbar2Tk(
            self.pump_curve_canvas, self.pump_curve_toolbar)
        self.pump_curve_toolbar.update()
        self.pump_curve_canvas.draw()

    def _clear_pump_curve_chart(self):
        """Efface le graphique H(Q)."""
        if self.pump_curve_canvas:
            self.pump_curve_canvas.get_tk_widget().destroy()
            self.pump_curve_canvas = None
        if self.pump_curve_toolbar:
            self.pump_curve_toolbar.destroy()
            self.pump_curve_toolbar = None
        self.lbl_pump_curve_placeholder.grid(
            row=0, column=0, padx=20, pady=30, sticky="nsew")

    def _update_chart(self):
        """Trace / rafraîchit le graphique interactif dans l'onglet Analyse Transitoire."""
        # Nettoyage de l'ancien canvas
        if self.canvas:
            self.canvas.get_tk_widget().destroy()
            self.canvas = None
        if self.toolbar:
            self.toolbar.destroy()
            self.toolbar = None
        self.current_fig = None

        if self.parser.hpt_data is None:
            self.lbl_graph_placeholder.grid(row=0, column=0, padx=20, pady=40, sticky="nsew")
            return

        self.lbl_graph_placeholder.grid_forget()

        df   = self.parser.hpt_data
        cols = list(df.columns)

        dist_col = self.parser._find_col(cols, ['distance', 'abscisse', 'position', 'station', 'chainage', 'x (m)'])
        vol_col  = self.parser._find_col(cols, ['volume of gas', 'volume gaz', 'gas volume', 'vol gaz'])
        pmin_col = self.parser._find_col(cols, ['pressure (minimum)', 'pression (minimum)', 'pmin', 'p min',
                                                 'min pressure', 'pressure min', 'pression min'])
        pmax_col = self.parser._find_col(cols, ['pressure (maximum)', 'pression (maximum)', 'pmax', 'p max',
                                                 'max pressure', 'pressure max', 'pression max'])

        x      = df[dist_col] if dist_col else df.index
        y_pmin = df[pmin_col] if pmin_col else ([0] * len(df))
        y_pmax = df[pmax_col] if pmax_col else ([0] * len(df))
        # Volume de gaz : données stockées en L (canonique) → conversion affichage
        vol_unit_factor = VOLUME_UNITS.get(self.var_volume_unit.get(), 1.0)
        y_vgas = (df[vol_col] / vol_unit_factor) if vol_col else ([0] * len(df))
        threshold_disp = self._volume_threshold_display()
        vol_unit_label = self.var_volume_unit.get()

        pn_val   = self._get_pn_value()
        pmin_val = self._get_pmin_value()

        # Palette en fonction du thème
        mode = ctk.get_appearance_mode()
        if mode == "Dark":
            bg, ax_bg = "#2b2b2b", "#1e1e1e"
            text_c, grid_c, spine_c = "#dcdcdc", "#3d3d3d", "#555555"
        else:
            bg, ax_bg = "#ebebeb", "#ffffff"
            text_c, grid_c, spine_c = "#202020", "#dedede", "#bbbbbb"

        fig = Figure(figsize=(8, 4.5), dpi=95)
        fig.patch.set_facecolor(bg)

        ax1 = fig.add_subplot(211)
        ax2 = fig.add_subplot(212, sharex=ax1)

        # ── Pression ────────────────────────────────────────────────
        ax1.set_facecolor(ax_bg)
        ax1.plot(x, y_pmax, color="#e3211b", lw=2.0, label="Pression Max")
        ax1.plot(x, y_pmin, color="#1f78b4", lw=2.0, label="Pression Min")
        ax1.axhline(0,        color="#ff9f00", lw=1.2, ls="--", label="P. Atm. (0 bar)")
        ax1.axhline(pn_val,   color="#ff4444", lw=1.5, ls="-.", alpha=0.8,
                    label=f"Limite PN ({self.var_pn.get()} = {pn_val} bar)")
        ax1.axhline(pmin_val, color="#4488ff", lw=1.5, ls="-.", alpha=0.8,
                    label=f"P min sécurité ({pmin_val} bar)")
        ax1.fill_between(x, y_pmax, pn_val, where=[v > pn_val for v in y_pmax],
                         color="#ff4444", alpha=0.12, label="Zone dépassement PN")
        ax1.fill_between(x, y_pmin, pmin_val, where=[v < pmin_val for v in y_pmin],
                         color="#4488ff", alpha=0.12, label="Zone dépression critique")
        ax1.set_ylabel("Pression (bar)", color=text_c, fontweight="bold", fontsize=9)
        ax1.set_title("Enveloppes de Pression — Coup de Bélier", color=text_c, fontsize=10, fontweight="bold")
        ax1.tick_params(colors=text_c, labelsize=8)
        ax1.grid(True, color=grid_c, ls=":")
        ax1.legend(fontsize=7, loc="upper right", framealpha=0.85,
                   facecolor=ax_bg, labelcolor=text_c, edgecolor=spine_c)
        for sp in ax1.spines.values(): sp.set_color(spine_c)

        # ── Volume de gaz ────────────────────────────────────────────
        ax2.set_facecolor(ax_bg)
        ax2.fill_between(x, y_vgas, color="#33a02c", alpha=0.25)
        ax2.plot(x, y_vgas, color="#33a02c", lw=2.0, label="Vol. Gaz (HPT)")
        ax2.axhline(threshold_disp, color="#ff9f00", lw=1.5, ls="-.", alpha=0.8,
                    label=f"Seuil {threshold_disp:.3f} {vol_unit_label}")
        ax2.set_xlabel("Distance (m)", color=text_c, fontweight="bold", fontsize=9)
        ax2.set_ylabel(f"Vol. Gaz ({vol_unit_label})",
                       color=text_c, fontweight="bold", fontsize=9)
        ax2.set_title("Volume de Gaz Maximum dans le Réservoir HPT", color=text_c, fontsize=10, fontweight="bold")
        ax2.tick_params(colors=text_c, labelsize=8)
        ax2.grid(True, color=grid_c, ls=":")
        ax2.legend(fontsize=7, loc="upper right", framealpha=0.85,
                   facecolor=ax_bg, labelcolor=text_c, edgecolor=spine_c)
        for sp in ax2.spines.values(): sp.set_color(spine_c)

        fig.tight_layout(pad=1.8)
        self.current_fig = fig

        # Intégration dans Tkinter
        self.canvas = FigureCanvasTkAgg(fig, master=self.plot_placeholder_frame)
        self.canvas.get_tk_widget().grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        self.plot_placeholder_frame.grid_rowconfigure(0, weight=1)
        self.plot_placeholder_frame.grid_rowconfigure(1, weight=0)
        self.plot_placeholder_frame.grid_columnconfigure(0, weight=1)

        toolbar_bg = bg.replace("#", "")
        self.toolbar_frame = tk.Frame(self.plot_placeholder_frame, bg=bg)
        self.toolbar_frame.grid(row=1, column=0, sticky="ew", padx=4)
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.toolbar_frame)
        self.toolbar.update()
        self.canvas.draw()

    # ================================================================
    # RAPPORT TEXTUEL (Prévisualisation)
    # ================================================================

    def _update_report_preview(self):
        """Génère et affiche la note technique textuelle dans l'onglet Rapport."""
        projet    = self.entry_projet.get().strip()    if hasattr(self, 'entry_projet')    else ""
        ingenieur = self.entry_ingenieur.get().strip() if hasattr(self, 'entry_ingenieur') else ""
        now       = datetime.now().strftime("%d/%m/%Y  %H:%M")
        pn_lbl    = self.var_pn.get()   if hasattr(self, 'var_pn')   else "PN 16"
        pmin_lbl  = self.var_pmin.get() if hasattr(self, 'var_pmin') else ""
        pn_val    = self._get_pn_value()
        pmin_val  = self._get_pmin_value()

        q   = "--"; hmt  = "--"
        pmax = "--"; pmin_ = "--"; vgas = None
        station_file  = os.path.basename(self.station_filepath) if self.station_filepath else "Non chargé"
        hpt_file      = os.path.basename(self.hpt_filepath)     if self.hpt_filepath     else "Non chargé"

        if self.steady_state_status and self.steady_state_status.get("success"):
            q   = self.steady_state_status.get("flow_rate_m3h") or "--"
            hmt = self.steady_state_status.get("hmt_m")         or "--"
        if self.transient_status and self.transient_status.get("success"):
            pmax  = self.transient_status.get("max_pressure_bar")
            pmin_ = self.transient_status.get("min_pressure_bar")
            vgas  = self.transient_status.get("max_gas_volume_l")

        # Pré-formattage des valeurs avec les unités d'affichage
        flow_disp       = self._format_flow(q)        if q != "--" else "--"
        vgas_disp       = self._format_volume(vgas)   if vgas is not None else "--"
        threshold_disp  = self._volume_threshold_display()
        vol_unit_lbl    = self.var_volume_unit.get()

        lines = [
            "═" * 70,
            "   NOTE TECHNIQUE D'ANALYSE HYDRAULIQUE",
            "   Coup de Bélier — Protection par Réservoir HPT",
            "═" * 70,
            "",
            f"  Projet    : {projet or '(non renseigné)'}",
            f"  Ingénieur : {ingenieur or '(non renseigné)'}",
            f"  Date      : {now}",
            "",
            "─" * 70,
            "  PARAMÈTRES DE DIMENSIONNEMENT (Sidebar)",
            "─" * 70,
            f"  Classe PN de la conduite          : {pn_lbl} → P max admissible = {pn_val} bar",
            f"  Pression minimale de sécurité     : {pmin_val} bar  ({pmin_lbl})",
            f"  Seuil HPT (volume de gaz max)     : {threshold_disp:.3f} {vol_unit_lbl}",
            "",
            "─" * 70,
            "  1. RÉGIME PERMANENT",
            "─" * 70,
            f"  Fichier station          : {station_file}",
            f"  Débit nominal (Q)        : {flow_disp}",
            f"  HMT                      : {hmt} mCE",
            "",
        ]

        # Section Pompe (si disponible)
        pump_summary = self.pump_parser.get_summary() if self.pump_parser.parsed else None
        if pump_summary and pump_summary.get("label") and pump_summary["label"] != "—":
            pump_file = os.path.basename(self.pump_filepath) if self.pump_filepath else "Non chargé"
            pump_flow = f"{pump_summary['flow_lps']:.1f} L/s" if pump_summary.get("flow_lps") is not None else "—"
            pump_head = f"{pump_summary['pump_head_m']:.1f} m" if pump_summary.get("pump_head_m") is not None else "—"
            npsh_a = f"{pump_summary['npsh_available_m']:.1f} m" if pump_summary.get("npsh_available_m") is not None else "—"
            lines += [
                "─" * 70,
                "  1b. DONNÉES POMPE (Rapport Détaillé)",
                "─" * 70,
                f"  Fichier rapport pompe    : {pump_file}",
                f"  Pompe                    : {pump_summary['label']} (ID {pump_summary['pump_id']})",
                f"  Conduite aval            : {pump_summary.get('downstream_pipe', '—')}",
                f"  Débit nominal (Q)        : {pump_flow}",
                f"  HMT pompe                : {pump_head}",
                f"  NPSH disponible          : {npsh_a}",
                f"  Points courbe H(Q)       : {pump_summary.get('n_curve_points', 0)}",
                "",
            ]

        lines += [
            "─" * 70,
            "  2. ANALYSE TRANSITOIRE — COURBES ENVELOPPES HPT",
            "─" * 70,
            f"  Fichier HPT analysé      : {hpt_file}",
            f"  Pression maximale        : {pmax} bar",
            f"  Pression minimale        : {pmin_} bar",
            f"  Volume de gaz max (HPT)  : {vgas_disp}",
            "",
            "─" * 70,
            "  3. DIAGNOSTIC DE SÉCURITÉ",
            "─" * 70,
        ]

        # Diagnostic Surpression
        if pmax != "--":
            ok = float(pmax) <= pn_val
            icon = "✔" if ok else "⚠"
            lines.append(f"  {icon} Surpression : {pmax} bar  {'<= ' + str(pn_val) + ' bar [OK]' if ok else '> ' + str(pn_val) + ' bar [DÉPASSEMENT PN – ATTENTION]'}")
            if not ok:
                lines.append("    → Augmenter la classe PN ou allonger le temps de fermeture des vannes.")
        else:
            lines.append("  [–] Surpression : données HPT non chargées")

        # Diagnostic Dépression
        if pmin_ != "--":
            ok = float(pmin_) >= pmin_val
            icon = "✔" if ok else "⚠"
            lines.append(f"  {icon} Dépression : {pmin_} bar  {'>= ' + str(pmin_val) + ' bar [OK]' if ok else '< ' + str(pmin_val) + ' bar [DÉPRESSION CRITIQUE]'}")
            if not ok:
                lines.append("    → Vérifier le pré-gonflage HPT. Installer des ventouses automatiques.")
        else:
            lines.append("  [–] Dépression : données HPT non chargées")

        # Diagnostic Volume Gaz (seuil dynamique en unité d'affichage)
        if vgas is not None:
            threshold_l = self._threshold_internal_l()
            ok = float(vgas) <= threshold_l
            icon = "✔" if ok else "⚠"
            lines.append(
                f"  {icon} Volume gaz HPT : {vgas_disp}  "
                f"{'<= ' + f'{threshold_disp:.3f} ' + vol_unit_lbl + ' [OK]' if ok else '> ' + f'{threshold_disp:.3f} ' + vol_unit_lbl + ' [VOLUME ÉLEVÉ]'}")
            if not ok:
                lines.append("    → Augmenter le volume nominal du réservoir HPT.")
        else:
            lines.append("  [–] Volume gaz : données HPT non chargées")

        # Section erreurs de parsing
        errors = []
        if self.steady_state_status and not self.steady_state_status.get("success"):
            errors.append(f"  ✘ Station  : {self.steady_state_status.get('message')}")
        if self.transient_status and not self.transient_status.get("success"):
            errors.append(f"  ✘ HPT      : {self.transient_status.get('message')}")
        if errors:
            lines += ["", "─" * 70, "  ERREURS DE CHARGEMENT", "─" * 70] + errors

        lines += ["", "═" * 70,
                   "  Document généré par HammerPy Insight v3.0",
                  "═" * 70]

        self.txt_report.delete("1.0", tk.END)
        self.txt_report.insert("1.0", "\n".join(lines))

    # ================================================================
    # GESTION DE PROJET (.hpi)  —  Ouvrir / Enregistrer / Enregistrer sous / Quitter
    # ================================================================

    PROJECT_FORMAT_VERSION = "3.0"

    def _update_title(self):
        """Met à jour la barre de titre + l'indicateur d'état en haut de la fenêtre."""
        if self.current_project_path:
            base = os.path.basename(self.current_project_path)
            self.lbl_project_state.configure(text=f"● {base}", text_color="#33a02c")
            self.title(f"HammerPy Insight v3 — {base}"
                       + ("  •  *" if self.is_dirty else ""))
        else:
            if self.is_dirty:
                self.lbl_project_state.configure(
                    text="● Nouveau projet (modifications non sauvegardées)",
                    text_color="orange")
            else:
                self.lbl_project_state.configure(
                    text="● Nouveau projet (non sauvegardé)",
                    text_color="gray")
            self.title("HammerPy Insight v3 — Analyse Transitoire Hydraulique"
                       + ("  •  *" if self.is_dirty else ""))

    def _mark_dirty(self):
        """Marque l'état comme modifié et rafraîchit l'indicateur visuel."""
        if self._suppress_dirty:
            return
        if not self.is_dirty:
            self.is_dirty = True
            self._update_title()

    def _mark_clean(self):
        """Marque l'état comme sauvegardé."""
        self.is_dirty = False
        self._update_title()

    # ------------------------------------------------------------------
    # Construction de l'état complet du projet pour sérialisation
    # ------------------------------------------------------------------
    def _build_project_payload(self) -> dict:
        """Sérialise l'intégralité de l'état de la session en dictionnaire."""
        try:
            station_data = (
                self.parser.steady_state_data.to_dict(orient="records")
                if self.parser.steady_state_data is not None else None
            )
        except Exception:
            station_data = None

        try:
            hpt_data = (
                self.parser.hpt_data.to_dict(orient="records")
                if self.parser.hpt_data is not None else None
            )
        except Exception:
            hpt_data = None

        # Données du classeur HAMMER (6 feuilles, orient='records' pour compacité)
        workbook_data = {}
        for canonical in ["pipes", "nodes", "reservoirs", "pumps", "hpt", "air_valves"]:
            df = self.workbook_manager.sheets.get(canonical)
            if df is not None and len(df) > 0:
                try:
                    # Convertir les types numpy en types natifs Python pour JSON
                    records = df.to_dict(orient="records")
                    for rec in records:
                        for k, v in rec.items():
                            if hasattr(v, 'item'):
                                rec[k] = v.item()
                    workbook_data[canonical] = records
                except Exception:
                    workbook_data[canonical] = []

        return {
            "version": self.PROJECT_FORMAT_VERSION,
            "metadata": {
                "nom_projet":       self.entry_projet.get().strip()
                                     if hasattr(self, "entry_projet") else "",
                "ingenieur":        self.entry_ingenieur.get().strip()
                                     if hasattr(self, "entry_ingenieur") else "",
                "date_creation":    self.project_creation_date,
                "date_modification": datetime.now().isoformat(timespec="seconds"),
            },
            "config": {
                "pn_label":   self.var_pn.get(),
                "pn_value":   self._get_pn_value(),
                "pmin_label": self.var_pmin.get(),
                "pmin_value": self._get_pmin_value(),
                "flow_unit":             self.var_flow_unit.get(),
                "volume_unit":           self.var_volume_unit.get(),
                "volume_threshold_l":    self._threshold_internal_l(),
            },
            "station": {
                "filepath": self.station_filepath,
                "data":     station_data,
                "status":   self.steady_state_status,
            },
            "hpt": {
                "filepath": self.hpt_filepath,
                "data":     hpt_data,
                "status":   self.transient_status,
            },
            "workbook": {
                "filepath": self.workbook_filepath,
                "sheets":   workbook_data,
            },
            "pump": {
                "filepath":     self.pump_filepath,
                "parsed":       self.pump_parser.parsed,
                "curve_points": self.pump_parser.curve_points,
            },
            "report_text": (self.txt_report.get("1.0", tk.END)
                            if hasattr(self, "txt_report") else ""),
        }

    # ------------------------------------------------------------------
    # Sauvegarde / chargement bas niveau
    # ------------------------------------------------------------------
    def _save_project(self, filepath: str):
        """Sérialise l'état du projet dans un fichier .hpi (JSON)."""
        payload = self._build_project_payload()
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)
        self.current_project_path = filepath
        self._mark_clean()
        messagebox.showinfo(
            "Projet enregistré",
            f"Projet sauvegardé avec succès :\n{filepath}")

    def _load_project(self, filepath: str):
        """Charge un fichier .hpi et restaure l'état complet de l'application."""
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                payload = json.load(f)
        except json.JSONDecodeError as exc:
            messagebox.showerror("Fichier invalide",
                                 f"Le fichier .hpi est corrompu ou non lisible :\n{exc}")
            return
        except Exception as exc:
            messagebox.showerror("Erreur de chargement", str(exc))
            return

        version = payload.get("version", "1.0")
        if version not in ("2.0", "3.0"):
            messagebox.showwarning(
                "Version incompatible",
                f"Ce projet a été créé avec une version différente "
                f"({version}).\nLa tentative de chargement peut échouer.")

        # Anti-boucle : empêche les handlers de marquer dirty pendant la restauration
        self._suppress_dirty = True
        try:
            # ── Métadonnées ──────────────────────────────────────────
            meta = payload.get("metadata", {})
            self.project_creation_date = (meta.get("date_creation")
                                          or datetime.now().isoformat(timespec="seconds"))
            if hasattr(self, "entry_projet"):
                self.entry_projet.delete(0, tk.END)
                self.entry_projet.insert(0, meta.get("nom_projet", ""))
            if hasattr(self, "entry_ingenieur"):
                self.entry_ingenieur.delete(0, tk.END)
                self.entry_ingenieur.insert(0, meta.get("ingenieur", ""))

            # ── Configuration PN / Pmin ──────────────────────────────
            cfg = payload.get("config", {})
            if "pn_label" in cfg and cfg["pn_label"] in PN_CLASSES:
                self.var_pn.set(cfg["pn_label"])
            if "pmin_label" in cfg and cfg["pmin_label"] in PMIN_OPTIONS:
                self.var_pmin.set(cfg["pmin_label"])
            # Unités d'affichage
            if "flow_unit" in cfg and cfg["flow_unit"] in FLOW_UNITS:
                self.var_flow_unit.set(cfg["flow_unit"])
            if "volume_unit" in cfg and cfg["volume_unit"] in VOLUME_UNITS:
                self.var_volume_unit.set(cfg["volume_unit"])
            # Seuil HPT : stocké en L, on l'affiche dans l'unité choisie
            if "volume_threshold_l" in cfg:
                threshold_disp = cfg["volume_threshold_l"] / VOLUME_UNITS.get(
                    self.var_volume_unit.get(), 1.0)
                self.var_volume_threshold.set(f"{threshold_disp:.3f}")
            self._refresh_threshold_unit_label()

            # ── Station (régime permanent) ───────────────────────────
            station = payload.get("station", {})
            self.station_filepath = station.get("filepath", "") or ""
            self.steady_state_status = station.get("status")
            try:
                self.parser.steady_state_data = (
                    pd.DataFrame.from_records(station["data"])
                    if station.get("data") is not None else None
                )
            except Exception:
                self.parser.steady_state_data = None

            # Mise à jour visuelle des KPI station
            self.lbl_steady_file.configure(
                text=os.path.basename(self.station_filepath)
                     if self.station_filepath else "Aucun fichier sélectionné",
                text_color=("gray20", "gray80")
                          if self.station_filepath else "gray")
            s_ok = (self.steady_state_status or {}).get("success")
            if s_ok:
                q   = self.steady_state_status.get("flow_rate_m3h")
                hmt = self.steady_state_status.get("hmt_m")
                self.lbl_flow_val.configure(
                    text=f"{q} m³/h" if q is not None else "—",
                    text_color="#1f8ecf" if q is not None else "orange")
                self.lbl_hmt_val.configure(
                    text=f"{hmt} mCE" if hmt is not None else "—",
                    text_color="#1f8ecf" if hmt is not None else "orange")
                cols = self.steady_state_status.get("columns", [])
                self.lbl_steady_details.configure(
                    text=("Colonnes détectées : " + ", ".join(cols))[:160])
            else:
                self.lbl_flow_val.configure(text="-- --", text_color="#1f8ecf")
                self.lbl_hmt_val.configure(text="-- --", text_color="#1f8ecf")
                self.lbl_steady_details.configure(text="")

            # ── HPT (transitoire) ────────────────────────────────────
            hpt = payload.get("hpt", {})
            self.hpt_filepath = hpt.get("filepath", "") or ""
            self.transient_status = hpt.get("status")
            try:
                self.parser.hpt_data = (
                    pd.DataFrame.from_records(hpt["data"])
                    if hpt.get("data") is not None else None
                )
            except Exception:
                self.parser.hpt_data = None

            self.lbl_transient_file.configure(
                text=os.path.basename(self.hpt_filepath)
                     if self.hpt_filepath else "Aucun fichier HPT chargé",
                text_color=("gray20", "gray80")
                          if self.hpt_filepath else "gray")
            t_ok = (self.transient_status or {}).get("success")
            if t_ok:
                self.lbl_pmin_val.configure(
                    text=f"{self.transient_status.get('min_pressure_bar')} bar",
                    text_color=self._pression_color(
                        self.transient_status.get("min_pressure_bar"), "min"))
                self.lbl_pmax_val.configure(
                    text=f"{self.transient_status.get('max_pressure_bar')} bar",
                    text_color=self._pression_color(
                        self.transient_status.get("max_pressure_bar"), "max"))
                vgas = self.transient_status.get("max_gas_volume_l")
                self.lbl_vgas_val.configure(
                    text=f"{vgas} L",
                    text_color="#33a02c" if vgas is not None and vgas <= 200 else "orange")
            else:
                self.lbl_pmin_val.configure(text="-- bar", text_color="#1f8ecf")
                self.lbl_pmax_val.configure(text="-- bar", text_color="#1f8ecf")
                self.lbl_vgas_val.configure(text="-- L", text_color="#1f8ecf")

            # ── Classeur HAMMER (v3.0+) ───────────────────────────────
            wb = payload.get("workbook", {})
            self.workbook_filepath = wb.get("filepath", "") or ""
            sheets_data = wb.get("sheets", {})

            if sheets_data:
                # Reconstruire les DataFrames à partir des records JSON
                self.workbook_manager = WorkbookManager()
                self.workbook_manager.filepath = self.workbook_filepath
                for canonical, records in sheets_data.items():
                    if records:
                        try:
                            self.workbook_manager.sheets[canonical] = pd.DataFrame.from_records(records)
                        except Exception:
                            self.workbook_manager.sheets[canonical] = pd.DataFrame()

                # Mettre à jour l'UI
                self.lbl_workbook_file.configure(
                    text=os.path.basename(self.workbook_filepath)
                         if self.workbook_filepath else "Aucun classeur chargé",
                    text_color=("gray20", "gray80")
                              if self.workbook_filepath else "gray")
                self.btn_reset_workbook.configure(
                    state="normal" if self.workbook_filepath else "disabled")

                summary = self.workbook_manager.get_summary()
                for key, lbl in self._wb_counter_labels.items():
                    count = summary.get(f"{key}_count", 0)
                    lbl.configure(text=str(count) if count > 0 else "—",
                                  text_color="#33a02c" if count > 0 else "gray")

                pmax = summary.get("pmax_bar")
                pmin = summary.get("pmin_bar")
                vmax = summary.get("vmax_pump_ls")
                mats = summary.get("materials", [])
                self._wb_stats["pmax"].configure(
                    text=f"{pmax:.2f} bar" if pmax is not None else "— bar",
                    text_color="#e87c2a" if pmax and pmax > self._get_pn_value() else "#33a02c")
                self._wb_stats["pmin"].configure(
                    text=f"{pmin:.2f} bar" if pmin is not None else "— bar",
                    text_color="#e87c2a" if pmin and pmin < self._get_pmin_value() else "#33a02c")
                self._wb_stats["vmax_pump"].configure(
                    text=f"{vmax:.1f} L/s" if vmax is not None else "— L/s")
                self._wb_stats["materials"].configure(
                    text=", ".join(mats) if mats else "—")
            else:
                self._reset_workbook()

            # ── Pompe (rapport détaillé, v3.0+) ───────────────────────
            pump_data = payload.get("pump", {})
            self.pump_filepath = pump_data.get("filepath", "") or ""
            if pump_data.get("parsed"):
                self.pump_parser = PumpReportParser()
                self.pump_parser.filepath = self.pump_filepath
                self.pump_parser.parsed = pump_data["parsed"]
                self.pump_parser.curve_points = pump_data.get("curve_points", [])

                self.lbl_pump_file.configure(
                    text=os.path.basename(self.pump_filepath)
                         if self.pump_filepath else "Aucun rapport pompe chargé",
                    text_color=("gray20", "gray80")
                              if self.pump_filepath else "gray")
                self.btn_reset_pump.configure(
                    state="normal" if self.pump_filepath else "disabled")

                summary = self.pump_parser.get_summary()
                self._pump_kpi["label"].configure(text=summary["label"])
                self._pump_kpi["flow"].configure(
                    text=f"{summary['flow_lps']:.1f} L/s" if summary["flow_lps"] is not None else "— L/s")
                self._pump_kpi["head"].configure(
                    text=f"{summary['pump_head_m']:.1f} m" if summary["pump_head_m"] is not None else "— m")
                npsh_a = summary["npsh_available_m"]
                self._pump_kpi["npsh_a"].configure(
                    text=f"{npsh_a:.1f} m" if npsh_a is not None else "— m",
                    text_color="#33a02c" if npsh_a is not None and npsh_a > 3 else "#e87c2a")
                self._pump_kpi["n_pts"].configure(text=str(summary["n_curve_points"]))
            else:
                self._reset_pump_report()

            # ── Rapport ──────────────────────────────────────────────
            report_text = payload.get("report_text", "")
            if hasattr(self, "txt_report") and report_text:
                self.txt_report.delete("1.0", tk.END)
                self.txt_report.insert("1.0", report_text)

            # ── Graphique (régénéré à partir des données) ───────────
            self._update_chart()
            self._update_report_preview()

        finally:
            self._suppress_dirty = False

        self.current_project_path = filepath
        self._mark_clean()
        messagebox.showinfo(
            "Projet chargé",
            f"Projet ouvert avec succès :\n{filepath}")

    # ------------------------------------------------------------------
    # Commandes du menu (boutons)
    # ------------------------------------------------------------------
    def _open_project(self):
        """Ouvre un fichier .hpi existant après confirmation éventuelle."""
        if not self._confirm_discard_changes("ouvrir un autre projet"):
            return
        filepath = filedialog.askopenfilename(
            title="Ouvrir un projet HammerPy Insight",
            defaultextension=".hpi",
            filetypes=[("Projet HammerPy Insight", "*.hpi"),
                       ("Tous les fichiers", "*.*")]
        )
        if not filepath:
            return
        self._load_project(filepath)

    def _save_project_as(self):
        """Enregistre le projet sous un nouveau chemin."""
        default_name = "projet_hammerpy.hpi"
        if hasattr(self, "entry_projet"):
            proj = self.entry_projet.get().strip()
            if proj:
                default_name = (proj.replace(" ", "_").replace("/", "-")
                                + ".hpi")
        filepath = filedialog.asksaveasfilename(
            title="Enregistrer le projet sous…",
            defaultextension=".hpi",
            initialfile=default_name,
            filetypes=[("Projet HammerPy Insight", "*.hpi"),
                       ("Tous les fichiers", "*.*")]
        )
        if not filepath:
            return
        try:
            self._save_project(filepath)
        except Exception as exc:
            messagebox.showerror("Erreur d'enregistrement", str(exc))

    def _save_project_quick(self):
        """Enregistre rapidement (au dernier chemin connu, ou bascule sur 'Enregistrer sous')."""
        if self.current_project_path and os.path.exists(self.current_project_path):
            try:
                payload = self._build_project_payload()
                with open(self.current_project_path, "w", encoding="utf-8") as f:
                    json.dump(payload, f, indent=2, ensure_ascii=False)
                self._mark_clean()
                messagebox.showinfo(
                    "Projet enregistré",
                    f"Projet sauvegardé :\n{self.current_project_path}")
            except Exception as exc:
                messagebox.showerror("Erreur d'enregistrement", str(exc))
        else:
            self._save_project_as()

    def _quit_app(self):
        """Quitte l'application après confirmation éventuelle."""
        if self.is_dirty:
            choice = messagebox.askyesnocancel(
                "Modifications non sauvegardées",
                "Le projet courant contient des modifications non sauvegardées.\n\n"
                "Voulez-vous les enregistrer avant de quitter ?\n\n"
                "• Oui  : enregistrer puis quitter\n"
                "• Non  : quitter sans enregistrer\n"
                "• Annuler : revenir à l'application")
            if choice is None:   # Annuler
                return
            if choice:            # Oui → enregistrer
                if self.current_project_path and os.path.exists(self.current_project_path):
                    try:
                        payload = self._build_project_payload()
                        with open(self.current_project_path, "w", encoding="utf-8") as f:
                            json.dump(payload, f, indent=2, ensure_ascii=False)
                    except Exception as exc:
                        messagebox.showerror(
                            "Erreur d'enregistrement",
                            f"Impossible d'enregistrer avant de quitter :\n{exc}")
                        return
                else:
                    filepath = filedialog.asksaveasfilename(
                        title="Enregistrer le projet avant de quitter",
                        defaultextension=".hpi",
                        initialfile="projet_hammerpy.hpi",
                        filetypes=[("Projet HammerPy Insight", "*.hpi")]
                    )
                    if not filepath:
                        return
                    try:
                        self._save_project(filepath)
                    except Exception as exc:
                        messagebox.showerror("Erreur d'enregistrement", str(exc))
                        return
        self.destroy()

    def _confirm_discard_changes(self, action_label: str) -> bool:
        """Demande confirmation si modifications non sauvegardées. Retourne True si OK."""
        if not self.is_dirty:
            return True
        choice = messagebox.askyesnocancel(
            "Modifications non sauvegardées",
            f"Vous allez {action_label}, mais le projet courant contient "
            f"des modifications non sauvegardées.\n\n"
            f"Voulez-vous les enregistrer d'abord ?\n\n"
            f"• Oui  : enregistrer puis continuer\n"
            f"• Non  : continuer sans enregistrer (perte des modifications)\n"
            f"• Annuler : revenir à l'application")
        if choice is None:
            return False
        if choice:
            if self.current_project_path and os.path.exists(self.current_project_path):
                try:
                    payload = self._build_project_payload()
                    with open(self.current_project_path, "w", encoding="utf-8") as f:
                        json.dump(payload, f, indent=2, ensure_ascii=False)
                    self._mark_clean()
                except Exception as exc:
                    messagebox.showerror(
                        "Erreur d'enregistrement",
                        f"Impossible d'enregistrer : {exc}\nOpération annulée.")
                    return False
            else:
                filepath = filedialog.asksaveasfilename(
                    title="Enregistrer le projet",
                    defaultextension=".hpi",
                    initialfile="projet_hammerpy.hpi",
                    filetypes=[("Projet HammerPy Insight", "*.hpi")]
                )
                if not filepath:
                    return False
                try:
                    self._save_project(filepath)
                except Exception as exc:
                    messagebox.showerror("Erreur d'enregistrement", str(exc))
                    return False
        return True

    # ================================================================
    # EXPORT
    # ================================================================

    def _export_txt(self):
        """Export de la note textuelle brute au format .txt."""
        filepath = filedialog.asksaveasfilename(
            title="Exporter la note (.txt)",
            defaultextension=".txt",
            filetypes=[("Texte", "*.txt"), ("Tous", "*.*")]
        )
        if filepath:
            try:
                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(self.txt_report.get("1.0", tk.END))
                messagebox.showinfo("Export réussi", f"Fichier enregistré :\n{filepath}")
            except Exception as exc:
                messagebox.showerror("Erreur", str(exc))

    def _export_word(self):
        """Export du rapport professionnel au format Word .docx avec graphique intégré."""
        if not DOCX_AVAILABLE:
            messagebox.showerror("python-docx manquant",
                                 "Installez python-docx :\n  pip install python-docx")
            return

        filepath = filedialog.asksaveasfilename(
            title="Enregistrer le rapport Word",
            defaultextension=".docx",
            filetypes=[("Word", "*.docx"), ("Tous", "*.*")]
        )
        if not filepath:
            return

        # Sauvegarde temporaire du graphique Matplotlib (PNG)
        chart_path = None
        if self.current_fig:
            try:
                tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                self.current_fig.savefig(tmp.name, dpi=150, bbox_inches="tight",
                                         facecolor=self.current_fig.get_facecolor())
                tmp.close()
                chart_path = tmp.name
            except Exception:
                chart_path = None

        metadata = {
            "nom_projet":       self.entry_projet.get().strip()    if hasattr(self, 'entry_projet')    else "",
            "ingenieur":        self.entry_ingenieur.get().strip() if hasattr(self, 'entry_ingenieur') else "",
            "date":             datetime.now().strftime("%d/%m/%Y"),
            "station_filepath": self.station_filepath,
            "hpt_filepath":     self.hpt_filepath,
        }

        try:
            # Récupérer le résumé du classeur si disponible
            wb_summary = None
            if self.workbook_manager.sheets:
                wb_summary = self.workbook_manager.get_summary()

            pump_summary = None
            if self.pump_parser.parsed:
                pump_summary = self.pump_parser.get_summary()

            gen = WordReportGenerator()
            doc = gen.generate(
                metadata              = metadata,
                steady                = self.steady_state_status,
                transient             = self.transient_status,
                pn_label              = self.var_pn.get(),
                pn_value              = self._get_pn_value(),
                pmin_label            = self.var_pmin.get(),
                pmin_value            = self._get_pmin_value(),
                flow_unit             = self.var_flow_unit.get(),
                volume_unit           = self.var_volume_unit.get(),
                volume_threshold_disp = self._volume_threshold_display(),
                chart_png_path        = chart_path,
                workbook_summary      = wb_summary,
                pump_summary          = pump_summary,
            )
            doc.save(filepath)
            messagebox.showinfo("Export réussi",
                                f"Rapport Word enregistré :\n{filepath}")
        except Exception as exc:
            messagebox.showerror("Erreur génération Word", str(exc))
        finally:
            # Nettoyage du fichier temporaire PNG
            if chart_path and os.path.exists(chart_path):
                try:
                    os.unlink(chart_path)
                except Exception:
                    pass


# =====================================================================
# 4. POINT D'ENTRÉE DU SCRIPT
# =====================================================================

if __name__ == "__main__":
    app = HammerPyApp()
    app.mainloop()
