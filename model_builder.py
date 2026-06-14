#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
model_builder.py — Import de classeur Excel libre avec mapping interactif
(type HAMMER Model Builder).
"""
import os
import pandas as pd
from utils import parse_number

PIPE_COLUMNS = {
    "label": "Label",
    "start_node": "N\u0153ud d\u00e9part",
    "end_node": "N\u0153ud arriv\u00e9e",
    "length_m": "Longueur (m)",
    "diameter_mm": "Diam\u00e8tre (mm)",
    "material": "Mat\u00e9riau",
    "roughness_c": "Rugosit\u00e9 (C)",
    "wave_speed": "C\u00e9l\u00e9rit\u00e9 (m/s)",
}

NODE_COLUMNS = {
    "label": "Label",
    "x_m": "X (m)",
    "y_m": "Y (m)",
    "z_m": "Z (m)",
}

PIPE_MANDATORY = ["label", "start_node", "end_node", "length_m", "diameter_mm"]
NODE_MANDATORY = ["label", "x_m", "y_m"]

FILE_TYPE_MODEL_PIPES = "model_builder_pipes"
FILE_TYPE_MODEL_NODES = "model_builder_nodes"


class ModelData:
    def __init__(self):
        self.nodes = []
        self.pipes = []
        self.unlinked_nodes = set()
        self.errors = []
        self.warnings = []

    @property
    def is_valid(self):
        return len(self.pipes) > 0 and len(self.nodes) > 0 and not self.errors

    def get_node_map(self):
        return {n["label"]: n for n in self.nodes if n.get("label")}

    def get_summary(self):
        diameters = set()
        materials = set()
        total_length = 0.0
        for p in self.pipes:
            d = p.get("diameter_mm")
            if d:
                diameters.add(int(round(d)))
            m = p.get("material")
            if m and str(m).strip():
                materials.add(str(m).strip())
            lv = p.get("length_m")
            if lv:
                total_length += lv
        return {
            "pipes_count": len(self.pipes),
            "nodes_count": len(self.nodes),
            "diameters": sorted(diameters) if diameters else [],
            "materials": sorted(materials) if materials else [],
            "total_length_m": round(total_length, 1),
            "unlinked_nodes": sorted(self.unlinked_nodes) if self.unlinked_nodes else [],
        }


class ModelBuilder:
    def __init__(self):
        self.data = ModelData()
        self.filepath = ""

    def open_workbook(self, filepath):
        self.filepath = filepath
        try:
            xl = pd.ExcelFile(filepath)
            return xl.sheet_names
        except Exception:
            return []

    def load_sheet(self, filepath, sheet_name):
        try:
            df = pd.read_excel(filepath, sheet_name=sheet_name, header=0)
            df.columns = [str(c).replace('\r', ' ').replace('\n', ' ').strip() for c in df.columns]
            df = df.dropna(how='all')
            return df
        except Exception:
            return None

    def build_model(self, pipe_df, node_df, pipe_map, node_map):
        data = ModelData()

        node_label_col = node_map.get("label")
        node_x_col = node_map.get("x_m")
        node_y_col = node_map.get("y_m")
        node_z_col = node_map.get("z_m")

        if not all([node_label_col, node_x_col, node_y_col]):
            data.errors.append("Colonnes obligatoires des n\u0153uds manquantes (label, X, Y).")
            return data

        seen_labels = set()
        for _, row in node_df.iterrows():
            label = str(row.get(node_label_col, "")).strip()
            if not label or label in ("nan", "", "None"):
                continue
            if label in seen_labels:
                data.warnings.append("N\u0153ud dupliqu\u00e9 : '{0}' \u2014 ignor\u00e9.".format(label))
                continue
            seen_labels.add(label)
            x = parse_number(row.get(node_x_col))
            y = parse_number(row.get(node_y_col))
            z = parse_number(row.get(node_z_col)) if node_z_col else None
            if x is None or y is None:
                data.warnings.append("N\u0153ud '{0}' : X ou Y invalide \u2014 ignor\u00e9.".format(label))
                continue
            data.nodes.append({
                "label": label,
                "x_m": float(x),
                "y_m": float(y),
                "z_m": float(z) if z is not None else None,
            })

        pipe_label_col = pipe_map.get("label")
        start_col = pipe_map.get("start_node")
        end_col = pipe_map.get("end_node")
        length_col = pipe_map.get("length_m")
        diam_col = pipe_map.get("diameter_mm")
        material_col = pipe_map.get("material")
        roughness_col = pipe_map.get("roughness_c")
        wave_col = pipe_map.get("wave_speed")

        if not all([pipe_label_col, start_col, end_col, length_col, diam_col]):
            data.errors.append("Colonnes obligatoires des canalisations manquantes (label, d\u00e9part, arriv\u00e9e, longueur, \u00d8).")
            return data

        node_positions = data.get_node_map()
        referenced_nodes = set()

        for _, row in pipe_df.iterrows():
            label = str(row.get(pipe_label_col, "")).strip()
            if not label or label in ("nan", "", "None"):
                continue
            start_n = str(row.get(start_col, "")).strip()
            end_n = str(row.get(end_col, "")).strip()
            if not start_n or not end_n:
                data.warnings.append("Canalisation '{0}' : n\u0153ud d\u00e9part ou arriv\u00e9e vide \u2014 ignor\u00e9e.".format(label))
                continue
            length_val = parse_number(row.get(length_col))
            diam_val = parse_number(row.get(diam_col))
            if length_val is None or float(length_val) <= 0:
                data.warnings.append("Canalisation '{0}' : longueur invalide \u2014 ignor\u00e9e.".format(label))
                continue
            if diam_val is None or float(diam_val) <= 0:
                data.warnings.append("Canalisation '{0}' : diam\u00e8tre invalide \u2014 ignor\u00e9.".format(label))
                continue
            mat_val = str(row.get(material_col, "")).strip() if material_col else ""
            rough_val = parse_number(row.get(roughness_col)) if roughness_col else None
            wave_val = parse_number(row.get(wave_col)) if wave_col else None

            data.pipes.append({
                "label": label,
                "start_node": start_n,
                "end_node": end_n,
                "length_m": float(length_val),
                "diameter_mm": float(diam_val),
                "material": mat_val if mat_val and mat_val.upper() != "NAN" else "",
                "roughness_c": float(rough_val) if rough_val is not None else None,
                "wave_speed": float(wave_val) if wave_val is not None else None,
            })
            referenced_nodes.add(start_n)
            referenced_nodes.add(end_n)

        for ref_node in referenced_nodes:
            if ref_node not in node_positions:
                data.unlinked_nodes.add(ref_node)

        return data
