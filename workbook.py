#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
workbook.py — Chargeur de classeur HAMMER (Flex Tables .xlsx/.xls).
Contient WorkbookManager et la configuration des feuilles.
"""

import os
import pandas as pd

from utils import parse_number, find_col_in_df


# =====================================================================
# CONFIGURATION DES FEUILLES HAMMER
# =====================================================================

# Noms de feuilles reconnus (insensible casse/accents)
SHEET_ALIASES = {
    "pipes":      ["pipes", "conduites", "pipe", "conduite"],
    "nodes":      ["noeuds", "nodes", "nœuds", "node", "nodo"],
    "reservoirs": ["reservoirs", "reservoir", "réservoirs", "réservoir", "tanks", "tank"],
    "pumps":      ["pumps", "pompe", "pump", "pompes"],
    "hpt":        ["hpt", "reservoir antideflagrant", "hydropneumatic", "hydropneumatique"],
    "air_valves": ["air valves", "air valve", "ventouses", "ventouse", "ventilateur"],
}

# Feuilles obligatoires (sinon le système est non-fonctionnel)
SHEET_MANDATORY = {"pipes", "nodes", "pumps"}

# Colonnes attendues par parser (pour validation de colonnes critiques)
SHEET_COLUMNS = {
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

        for canonical, aliases in SHEET_ALIASES.items():
            for sheet_name in xl.sheet_names:
                sheet_lower = sheet_name.lower().strip()
                if any(a in sheet_lower for a in aliases):
                    self.sheet_map[canonical] = sheet_name
                    break

        for mandatory in SHEET_MANDATORY:
            if mandatory not in self.sheet_map:
                self.errors.append(
                    f"Feuille obligatoire manquante : '{mandatory}'. "
                    f"Le système n'est pas fonctionnel sans cette feuille.")

        for canonical, sheet_name in self.sheet_map.items():
            try:
                df = pd.read_excel(filepath, sheet_name=sheet_name, header=0)
                df.columns = [
                    str(c).replace('\r', ' ').replace('\n', ' ').strip()
                    for c in df.columns
                ]
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

        for canonical in SHEET_MANDATORY:
            if canonical not in self.sheets:
                errors.append(f"Données manquantes : feuille '{canonical}' non chargée.")

        return len(errors) == 0, errors

    def get_sheet(self, name: str) -> pd.DataFrame | None:
        """Retourne une feuille par son nom canonique, ou None."""
        return self.sheets.get(name)

    def get_summary(self) -> dict:
        """
        Retourne un résumé des données chargées pour l'UI.
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
            mat_col = find_col_in_df(df, ['material', 'matériau'])
            if mat_col:
                summary["materials"] = sorted(df[mat_col].dropna().unique().tolist())
            diam_col = find_col_in_df(df, ['diameter', 'diamètre', 'ø'])
            if diam_col:
                diameters = df[diam_col].apply(parse_number).dropna()
                if len(diameters) > 0:
                    summary["diameter_min_mm"] = round(float(diameters.min()), 1)
                    summary["diameter_max_mm"] = round(float(diameters.max()), 1)
            length_col = find_col_in_df(df, ['length (user defined)', 'length', 'longueur'])
            # Métré par DN + Matériau
            pipe_detail = []
            if diam_col is not None:
                group_cols = [diam_col]
                group_names = ["dn_mm"]
                if mat_col is not None:
                    group_cols.insert(0, mat_col)
                    group_names.insert(0, "material")
                for keys, grp in df.groupby(group_cols):
                    if mat_col is not None and diam_col is not None:
                        mat_val = str(keys[0]) if pd.notna(keys[0]) else "N/C"
                        dn_val = parse_number(keys[1])
                    elif diam_col is not None:
                        dn_val = parse_number(keys) if not isinstance(keys, tuple) else keys
                        mat_val = "N/C"
                    else:
                        continue
                    if dn_val is None:
                        continue
                    dn_int = int(round(dn_val))
                    total_len = None
                    if length_col is not None:
                        lens = grp[length_col].apply(parse_number).dropna()
                        if len(lens) > 0:
                            total_len = round(float(lens.sum()), 1)
                    pipe_detail.append({
                        "dn_mm": dn_int,
                        "material": mat_val,
                        "count": len(grp),
                        "total_length_m": total_len,
                    })
                pipe_detail.sort(key=lambda r: (r["dn_mm"], r["material"]))
            summary["pipes_by_dn_material"] = pipe_detail

            pmax_col = find_col_in_df(df, ['pressure (maximum'])
            pmin_col = find_col_in_df(df, ['pressure (minimum'])
            if pmax_col:
                vals = df[pmax_col].apply(parse_number).dropna()
                if len(vals) > 0:
                    summary["pmax_bar"] = round(float(vals.max()), 4)
            if pmin_col:
                vals = df[pmin_col].apply(parse_number).dropna()
                if len(vals) > 0:
                    v = float(vals.min())
                    summary["pmin_bar"] = round(v, 4)

        # Nodes
        df = self.sheets.get("nodes")
        if df is not None:
            summary["nodes_count"] = len(df)
            if summary["pmax_bar"] is None:
                pmax_col = find_col_in_df(df, ['pressure (maximum'])
                if pmax_col:
                    vals = df[pmax_col].apply(parse_number).dropna()
                    if len(vals) > 0:
                        summary["pmax_bar"] = round(float(vals.max()), 4)
            if summary["pmin_bar"] is None:
                pmin_col = find_col_in_df(df, ['pressure (minimum'])
                if pmin_col:
                    vals = df[pmin_col].apply(parse_number).dropna()
                    if len(vals) > 0:
                        summary["pmin_bar"] = round(float(vals.min()), 4)

        # Pumps
        df = self.sheets.get("pumps")
        if df is not None:
            summary["pumps_count"] = len(df)
            flow_col = find_col_in_df(df, ['flow (total)', 'débit', 'flow'])
            if flow_col:
                vals = df[flow_col].apply(parse_number).dropna()
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
            vol_col = find_col_in_df(df, ['volume of gas', 'gas volume', 'volume gaz'])
            if vol_col:
                vals = df[vol_col].apply(parse_number).dropna()
                if len(vals) > 0:
                    summary["vmax_hpt_l"] = round(float(vals.max()), 2)

        # Air valves
        df = self.sheets.get("air_valves")
        if df is not None:
            summary["air_valves_count"] = len(df)
            air_col = find_col_in_df(df, ['air volume', 'volume air', 'volume (maximum'])
            if air_col:
                vals = df[air_col].apply(parse_number).dropna()
                if len(vals) > 0:
                    summary["vmax_air_valve_l"] = round(float(vals.max()), 2)

        return summary
