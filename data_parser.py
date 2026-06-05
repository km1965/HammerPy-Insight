#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
data_parser.py — Parser robuste des données hydrauliques HAMMER.
Contient HammerDataParser pour le chargement des fichiers CSV/Excel
(régime permanent + analyse transitoire).
"""

import os
import pandas as pd


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
            for enc in ('utf-8', 'latin-1', 'cp1252'):
                try:
                    df = pd.read_csv(filepath, sep=None, engine='python',
                                     encoding=enc, decimal='.', thousands=None)
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

            q_col   = self._find_col(df.columns,
                ['debit', 'flow', 'flow_rate', 'q (m3', 'm3/h', 'discharge',
                 'q (l/s', 'l/s', 'lps', 'q (l·s'])
            hmt_col = self._find_col(df.columns, ['hmt', 'head', 'mce', 'hauteur', 'total_head', 'tdh'])

            flow_rate = float(df.iloc[0][q_col])   if q_col   else None
            hmt       = float(df.iloc[0][hmt_col]) if hmt_col else None

            if flow_rate is not None and source_unit == "L/s":
                flow_rate = flow_rate * 3.6

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

            vol_kw   = ['volume of gas', 'volume gaz', 'gas volume', 'vol gaz', 'vol. gaz', 'air volume']
            pmin_kw  = ['pressure (minimum)', 'pression (minimum)', 'pmin', 'p min', 'hgl min',
                        'min pressure', 'pressure min', 'pression min']
            pmax_kw  = ['pressure (maximum)', 'pression (maximum)', 'pmax', 'p max', 'hgl max',
                        'max pressure', 'pressure max', 'pression max']

            vol_col  = self._find_col(cols, vol_kw)
            pmin_col = self._find_col(cols, pmin_kw)
            pmax_col = self._find_col(cols, pmax_kw)

            is_simulated = False

            if vol_col:
                found.append(vol_col)
                max_vol = float(df[vol_col].max())
                if source_unit == "m³":
                    max_vol = max_vol * 1000.0
            else:
                max_vol = 124.5
                is_simulated = True

            if pmin_col:
                found.append(pmin_col)
                min_press = float(df[pmin_col].min())
            else:
                min_press = -0.35
                is_simulated = True

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
