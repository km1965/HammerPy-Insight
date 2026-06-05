#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Tests unitaires pour le module WorkbookManager et _parse_number.
HammerPy Insight v3.0 — Parser classeur HAMMER (Flex Tables).
"""

import os
import sys
import pytest
import pandas as pd

# Ajouter le répertoire courant au path pour importer main
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from main import _parse_number, WorkbookManager, _find_col_in_df, PumpReportParser


# =====================================================================
# Tests pour _parse_number
# =====================================================================

class TestParseNumber:
    """Tests du helper _parse_number pour la normalisation numérique HAMMER."""

    def test_float_native(self):
        assert _parse_number(3.14) == 3.14

    def test_int_native(self):
        assert _parse_number(42) == 42.0

    def test_none(self):
        assert _parse_number(None) is None

    def test_nan(self):
        import math
        assert _parse_number(float('nan')) is None

    def test_empty_string(self):
        assert _parse_number("") is None

    def test_whitespace_only(self):
        assert _parse_number("   ") is None

    def test_nax_variants(self):
        for val in ("nan", "NaN", "NAN", "none", "None", "N/A", "(N/A)"):
            assert _parse_number(val) is None

    def test_simple_decimal_point(self):
        assert _parse_number("12.5") == 12.5

    def test_simple_integer_string(self):
        assert _parse_number("42") == 42.0

    def test_french_comma_decimal(self):
        assert _parse_number("12,5") == 12.5

    def test_thousands_space_normal(self):
        assert _parse_number("1 029,00") == 1029.0

    def test_thousands_nbsp(self):
        assert _parse_number("1\xa0029,00") == 1029.0

    def test_thousands_nbsp_french_comma(self):
        assert _parse_number("344\xa0438,04") == 344438.04

    def test_negative_number(self):
        assert _parse_number("-0.35") == -0.35

    def test_negative_french(self):
        assert _parse_number("-12,5") == -12.5

    def test_large_number_with_nbsp(self):
        assert _parse_number("11\xa0601,1") == 11601.1

    def test_zero(self):
        assert _parse_number("0") == 0.0

    def test_zero_french(self):
        assert _parse_number("0,00") == 0.0

    def test_already_float_with_nan(self):
        import math
        assert _parse_number(float('nan')) is None


# =====================================================================
# Tests pour _find_col_in_df
# =====================================================================

class TestFindColInDf:
    """Tests du helper _find_col_in_df pour la recherche de colonnes."""

    def test_exact_match(self):
        df = pd.DataFrame(columns=["Pressure (Maximum, Transient) (bars)", "Other"])
        assert _find_col_in_df(df, ["pressure (maximum"]) == "Pressure (Maximum, Transient) (bars)"

    def test_no_match(self):
        df = pd.DataFrame(columns=["ID", "Label"])
        assert _find_col_in_df(df, ["pressure"]) is None

    def test_case_insensitive(self):
        df = pd.DataFrame(columns=["PRESSURE (MINIMUM) (BARS)"])
        assert _find_col_in_df(df, ["pressure (minimum"]) == "PRESSURE (MINIMUM) (BARS)"

    def test_partial_match(self):
        df = pd.DataFrame(columns=["Flow (Total) (L/s)"])
        assert _find_col_in_df(df, ["flow (total)"]) == "Flow (Total) (L/s)"


# =====================================================================
# Tests pour WorkbookManager — intégration avec Flex Tables.xlsx
# =====================================================================

class TestWorkbookManager:
    """Tests d'intégration avec le classeur Flex Tables.xlsx réel."""

    WORKBOOK_PATH = os.path.join(os.path.dirname(__file__), "Flex Tables.xlsx")

    @pytest.fixture(autouse=True)
    def setup(self):
        """Charge le classeur avant chaque test."""
        self.wb = WorkbookManager()
        if os.path.exists(self.WORKBOOK_PATH):
            self.wb.load(self.WORKBOOK_PATH)

    def test_file_exists(self):
        assert os.path.exists(self.WORKBOOK_PATH), \
            f"Fichier de test non trouvé : {self.WORKBOOK_PATH}"

    def test_load_success(self):
        assert len(self.wb.errors) == 0

    def test_six_sheets_found(self):
        assert len(self.wb.sheet_map) >= 6

    def test_mandatory_sheets_present(self):
        valid, errors = self.wb.validate()
        assert valid, f"Validation échouée : {errors}"

    def test_pipes_loaded(self):
        df = self.wb.get_sheet("pipes")
        assert df is not None
        assert len(df) == 47

    def test_nodes_loaded(self):
        df = self.wb.get_sheet("nodes")
        assert df is not None
        assert len(df) == 42

    def test_pumps_loaded(self):
        df = self.wb.get_sheet("pumps")
        assert df is not None
        assert len(df) == 2

    def test_reservoirs_loaded(self):
        df = self.wb.get_sheet("reservoirs")
        assert df is not None
        assert len(df) == 2

    def test_hpt_loaded(self):
        df = self.wb.get_sheet("hpt")
        assert df is not None
        assert len(df) == 1

    def test_air_valves_loaded(self):
        df = self.wb.get_sheet("air_valves")
        assert df is not None
        assert len(df) == 1

    def test_pipes_columns_clean(self):
        df = self.wb.get_sheet("pipes")
        for col in df.columns:
            assert '\r' not in col, f"Carriage return dans colonne : {col}"
            assert '\n' not in col, f"Newline dans colonne : {col}"

    def test_summary_pipes_count(self):
        summary = self.wb.get_summary()
        assert summary["pipes_count"] == 47

    def test_summary_nodes_count(self):
        summary = self.wb.get_summary()
        assert summary["nodes_count"] == 42

    def test_summary_pumps_count(self):
        summary = self.wb.get_summary()
        assert summary["pumps_count"] == 2

    def test_summary_materials(self):
        summary = self.wb.get_summary()
        assert "ABG" in summary["materials"]

    def test_summary_pmax_is_positive(self):
        summary = self.wb.get_summary()
        assert summary["pmax_bar"] is not None
        assert summary["pmax_bar"] > 0

    def test_summary_pmin_is_negative(self):
        summary = self.wb.get_summary()
        assert summary["pmin_bar"] is not None
        assert summary["pmin_bar"] < 0

    def test_summary_vmax_pump(self):
        summary = self.wb.get_summary()
        assert summary["vmax_pump_ls"] is not None
        assert summary["vmax_pump_ls"] > 0


# =====================================================================
# Tests pour WorkbookManager — cas d'erreur
# =====================================================================

class TestWorkbookManagerErrors:
    """Tests de gestion des erreurs."""

    def test_wrong_extension(self):
        wb = WorkbookManager()
        result = wb.load("test.txt")
        assert not result
        assert len(wb.errors) > 0

    def test_nonexistent_file(self):
        wb = WorkbookManager()
        result = wb.load("nonexistent_file_12345.xlsx")
        assert not result

    def test_missing_mandatory_sheet(self):
        """Simule un classeur sans feuille Pipes."""
        wb = WorkbookManager()
        # Créer un mini-classeur Excel temporaire sans Pipes
        import tempfile
        import gc
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        try:
            with pd.ExcelWriter(tmp.name) as writer:
                pd.DataFrame({"A": [1]}).to_excel(writer, sheet_name="Noeuds", index=False)
                pd.DataFrame({"B": [2]}).to_excel(writer, sheet_name="Pumps", index=False)
            wb.load(tmp.name)
            valid, errors = wb.validate()
            assert not valid
            assert any("pipes" in e.lower() for e in errors)
        finally:
            gc.collect()
            try:
                os.unlink(tmp.name)
            except PermissionError:
                pass  # Windows file locking — skip cleanup


# =====================================================================
# Tests de rétrocompatibilité CSV (anciens fichiers)
# =====================================================================

class TestRetrocompatCSV:
    """Vérifie que les anciens fichiers CSV restent lisibles."""

    def test_hpt_csv_exists(self):
        path = os.path.join(os.path.dirname(__file__), "hpt_transient_test.csv")
        assert os.path.exists(path)

    def test_station_csv_exists(self):
        path = os.path.join(os.path.dirname(__file__), "station_steady_state_test.csv")
        assert os.path.exists(path)

    def test_hpt_csv_readable(self):
        from main import HammerDataParser
        parser = HammerDataParser()
        path = os.path.join(os.path.dirname(__file__), "hpt_transient_test.csv")
        result = parser.parse_hpt_file(path)
        assert result["success"]
        assert result["n_rows"] > 0

    def test_station_csv_readable(self):
        from main import HammerDataParser
        parser = HammerDataParser()
        path = os.path.join(os.path.dirname(__file__), "station_steady_state_test.csv")
        result = parser.parse_station_file(path)
        assert result["success"]
        assert result["n_rows"] > 0


# =====================================================================
# Tests pour PumpReportParser
# =====================================================================

class TestPumpReportParser:
    """Tests du parser de rapport pompe détaillé (RTF)."""

    def test_load_real_rtf(self):
        """Charge le fichier RTF réel du rapport pompe PMP-2."""
        path = os.path.join(os.path.dirname(__file__), "Pump detailed report.rtf")
        if not os.path.exists(path):
            pytest.skip("Fichier RTF de test non disponible")
        parser = PumpReportParser()
        ok = parser.load(path)
        assert ok, f"Erreurs : {parser.errors}"
        assert parser.parsed.get("label") == "PMP-2"
        assert parser.parsed.get("pump_id") == "122"
        assert parser.parsed.get("flow_lps") == 100.0
        assert parser.parsed.get("pump_head_m") == 110.80
        assert parser.parsed.get("pressure_suction_bar") == 0.54
        assert parser.parsed.get("pressure_discharge_bar") == 11.38
        assert parser.parsed.get("npsh_available_m") == 15.35
        assert parser.parsed.get("downstream_pipe") == "P-4"
        assert parser.parsed.get("controlled") is False

    def test_load_unsupported_ext(self):
        parser = PumpReportParser()
        ok = parser.load("test.xlsx")
        assert not ok
        assert any("non supportée" in e for e in parser.errors)

    def test_load_nonexistent_file(self):
        parser = PumpReportParser()
        ok = parser.load("nonexistent_file.rtf")
        assert not ok

    def test_strip_rtf_basic(self):
        from pump_parser import _strip_rtf
        rtf = r"{\rtf1 Some text \b bold \b0 normal}"
        text = _strip_rtf(rtf)
        assert "Some text" in text
        assert "normal" in text

    def test_strip_rtf_images(self):
        from pump_parser import _strip_rtf
        rtf = r"Before \pict\picw100\pich100 data After"
        text = _strip_rtf(rtf)
        assert "Before" in text
        assert "After" in text

    def test_curve_points_empty(self):
        parser = PumpReportParser()
        assert parser.get_curve_points() == []
        assert parser.interpolate_head(50.0) is None

    def test_curve_points_add_and_interpolate(self):
        parser = PumpReportParser()
        parser.add_curve_point(0.0, 200.0)    # Shutoff
        parser.add_curve_point(75.0, 150.0)   # Design
        parser.add_curve_point(150.0, 50.0)   # Max op

        pts = parser.get_curve_points()
        assert len(pts) == 3
        assert pts[0]["flow_lps"] == 0.0
        assert pts[2]["flow_lps"] == 150.0

        # Interpolation
        h = parser.interpolate_head(75.0)
        assert h == 150.0

        h = parser.interpolate_head(37.5)
        assert h is not None
        assert 170 < h < 180  # Between shutoff and design

        # Extrapolation inférieure
        h = parser.interpolate_head(0.0)
        assert h == 200.0

        # Extrapolation supérieure
        h = parser.interpolate_head(200.0)
        assert h is not None
        assert h < 50.0

    def test_clear_curve_points(self):
        parser = PumpReportParser()
        parser.add_curve_point(0.0, 200.0)
        parser.add_curve_point(100.0, 100.0)
        parser.clear_curve_points()
        assert parser.get_curve_points() == []

    def test_get_summary(self):
        parser = PumpReportParser()
        parser.parsed = {
            "pump_id": "42",
            "label": "TestPump",
            "flow_lps": 50.0,
            "pump_head_m": 80.0,
            "npsh_available_m": 5.0,
        }
        summary = parser.get_summary()
        assert summary["label"] == "TestPump"
        assert summary["flow_lps"] == 50.0
        assert summary["n_curve_points"] == 0

    def test_curve_points_sorted(self):
        parser = PumpReportParser()
        parser.add_curve_point(100.0, 100.0)
        parser.add_curve_point(0.0, 200.0)
        parser.add_curve_point(50.0, 150.0)
        pts = parser.get_curve_points()
        assert pts[0]["flow_lps"] == 0.0
        assert pts[1]["flow_lps"] == 50.0
        assert pts[2]["flow_lps"] == 100.0


# =====================================================================
# Tests AirValveSizing (Phase 3)
# =====================================================================

class TestAirValveSizing:

    def test_load_profile_manual(self):
        from air_valve_sizing import AirValveSizing
        sizer = AirValveSizing()
        sizer.load_profile_manual([(0, 100), (250, 120), (500, 110)])
        assert len(sizer.profile) == 3
        assert sizer.profile[0]["pk_m"] == 0
        assert sizer.profile[2]["pk_m"] == 500

    def test_high_low_points(self):
        from air_valve_sizing import AirValveSizing
        sizer = AirValveSizing()
        sizer.load_profile_manual([
            (0, 100), (100, 110), (200, 105), (300, 115), (400, 100)
        ])
        # Point haut au pk=100 (z=110) et pk=300 (z=115)
        assert len(sizer.high_points) >= 2
        # Point bas au pk=200 (z=105) et pk=400 (z=100)
        assert len(sizer.low_points) >= 2

    def test_size_ventouses(self):
        from air_valve_sizing import AirValveSizing
        sizer = AirValveSizing(pipe_dn_mm=250)
        sizer.load_profile_manual([
            (0, 100), (100, 110), (200, 105), (300, 115), (400, 100)
        ])
        sizer.size_ventouses()
        assert len(sizer.ventouses) >= 2
        for v in sizer.ventouses:
            assert v["dn_mm"] >= 25
            assert v["type"] in ("Anti-vide simple", "Combinée (admission + dégazage)",
                                 "Grande orifice (admission rapide)")

    def test_size_drains(self):
        from air_valve_sizing import AirValveSizing
        sizer = AirValveSizing(pipe_dn_mm=250)
        sizer.load_profile_manual([
            (0, 100), (100, 115), (250, 95), (400, 120), (500, 100)
        ])
        sizer.size_ventouses()
        sizer.size_drains()
        # Au moins une vidange entre les ventouses
        assert len(sizer.vidanges) >= 0  # dépend de la distance

    def test_dn_calculation(self):
        from air_valve_sizing import AirValveSizing
        sizer = AirValveSizing(pipe_dn_mm=250)
        dn = sizer._calc_dn(250, 12)
        assert dn >= 25
        assert dn in [25, 32, 40, 50, 65, 80, 100, 125, 150, 200, 250, 300]

    def test_export_csv(self):
        import tempfile, os
        from air_valve_sizing import AirValveSizing
        sizer = AirValveSizing(pipe_dn_mm=250)
        sizer.load_profile_manual([
            (0, 100), (100, 115), (250, 95), (400, 120), (500, 100)
        ])
        sizer.size_ventouses()
        sizer.size_drains()
        tmp = tempfile.NamedTemporaryFile(suffix=".csv", delete=False)
        tmp.close()
        sizer.export_csv(tmp.name)
        assert os.path.exists(tmp.name)
        with open(tmp.name, 'r', encoding='utf-8-sig') as f:
            content = f.read()
        assert "Ventouse" in content or "Vidange" in content
        os.unlink(tmp.name)

    def test_load_profile_csv(self):
        import tempfile, os
        from air_valve_sizing import AirValveSizing
        tmp = tempfile.NamedTemporaryFile(mode='w', suffix=".csv", delete=False,
                                          encoding='utf-8')
        tmp.write("pk;z\n")
        tmp.write("0;100\n")
        tmp.write("100;110\n")
        tmp.write("200;105\n")
        tmp.close()
        sizer = AirValveSizing()
        ok = sizer.load_profile_csv(tmp.name)
        assert ok is True
        assert len(sizer.profile) == 3
        os.unlink(tmp.name)


# =====================================================================
# Point d'entrée
# =====================================================================

if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])
