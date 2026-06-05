#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
pump_parser.py — Parser du rapport pompe détaillé HAMMER (RTF/TXT).
Contient PumpReportParser et la fonction _strip_rtf.
"""

import os
import re as _re

from utils import parse_number


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

    # 1. Supprimer le bloc fonttbl
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

    # 2. Supprimer les groupes {\*\...}
    while '{\\*' in text:
        start = text.find('{\\*')
        depth = 0
        for i in range(start, len(text)):
            if text[i] == '{':
                depth += 1
            elif text[i] == '}':
                depth -= 1
                if depth == 0:
                    text = text[:start] + text[i+1:]
                    break

    # 3. Supprimer les groupes {\pict...}
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

    # 4. Supprimer les groupes {\background...}
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
    """

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
        self.curve_points: list[dict] = []

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

        if ext == '.rtf':
            text = _strip_rtf(self.raw_text)
        else:
            text = self.raw_text

        self.parsed = self._extract_key_values(text)

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

        _UNIT_PATTERNS = {
            "m", "mm", "m³", "l", "l/s", "m³/h", "bar", "bars",
            "sec", "hours", "kW", "kwh", "mg/l", "%", "°", "kg/day",
            "€", "€/kW", "€/kWh", "€/ML", "mL",
        }

        def _is_positioning_number(s: str) -> bool:
            s = s.strip()
            if not s.startswith("-"):
                return False
            num_part = s[1:]
            if not num_part.isdigit():
                return False
            return int(num_part) > 20

        def _is_skip_line(s: str) -> bool:
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

        first_report_idx = -1
        second_report_idx = -1
        for idx, line in enumerate(lines):
            if "pump detailed report" in line.lower():
                if first_report_idx == -1:
                    first_report_idx = idx
                else:
                    second_report_idx = idx
                    break

        _phase1_keys = {"pump_id", "label", "downstream_pipe", "speed_factor", "status_initial"}
        if first_report_idx >= 0:
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
                                parsed = parse_number(val_clean)
                                if parsed is not None:
                                    result[key] = parsed
                        break
                i += 1

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
                                parsed = parse_number(val_clean)
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

        if flow_lps <= flows[0]:
            if len(pts) >= 2:
                slope = (heads[1] - heads[0]) / (flows[1] - flows[0]) if flows[1] != flows[0] else 0
                return heads[0] + slope * (flow_lps - flows[0])
            return heads[0]

        if flow_lps >= flows[-1]:
            if len(pts) >= 2:
                slope = (heads[-1] - heads[-2]) / (flows[-1] - flows[-2]) if flows[-1] != flows[-2] else 0
                return heads[-1] + slope * (flow_lps - flows[-1])
            return heads[-1]

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
