#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
test_diagnostics_charts.py — Tests pour les 4 fonctions de
diagnostics_charts.py (KPI donut, catégorie, conformité, profil).
"""

import os
import tempfile
import unittest

import matplotlib
matplotlib.use("Agg")  # Backend non-interactif pour les tests

from diagnostics_charts import (
    build_kpi_donut,
    build_category_stack,
    build_compliance_bars,
    build_profile_chart,
    COLOR_OK, COLOR_WARN, COLOR_FAIL, COLOR_NA,
    STATUS_NUM, STATUS_COLOR,
)


def _is_valid_png(path: str) -> bool:
    """Vérifie qu'un fichier est un PNG valide (signature 8 octets)."""
    if not path or not os.path.isfile(path):
        return False
    try:
        with open(path, "rb") as f:
            sig = f.read(8)
        # Signature PNG : 89 50 4E 47 0D 0A 1A 0A
        return sig == b"\x89PNG\r\n\x1a\n"
    except OSError:
        return False


# ── Données de test ────────────────────────────────────────────────

def _make_checks(cat_assign=None):
    """Construit 16 checks fictifs avec statuts répartis."""
    if cat_assign is None:
        # Par défaut : mix réaliste
        cat_assign = [
            ("A1", "A. Pompe ↔ Réseau",        "OK"),
            ("A2", "A. Pompe ↔ Réseau",        "OK"),
            ("A3", "A. Pompe ↔ Réseau",        "WARN"),
            ("B1", "B. Pompe ↔ HPT",           "OK"),
            ("B2", "B. Pompe ↔ HPT",           "FAIL"),
            ("C1", "C. Réseau ↔ HPT",          "OK"),
            ("C2", "C. Réseau ↔ HPT",          "NA"),
            ("C3", "C. Réseau ↔ HPT",          "NA"),
            ("D1", "D. HPT ↔ Ventouses / Vidanges", "OK"),
            ("D2", "D. HPT ↔ Ventouses / Vidanges", "WARN"),
            ("D3", "D. HPT ↔ Ventouses / Vidanges", "NA"),
            ("E1", "E. Cohérence globale",      "FAIL"),
            ("E2", "E. Cohérence globale",      "OK"),
            ("E3", "E. Cohérence globale",      "OK"),
            ("E4", "E. Cohérence globale",      "OK"),
            ("E5", "E. Cohérence globale",      "OK"),
        ]
    return [
        {"code": c, "category": cat, "status": st, "name": f"Check {c}",
         "message": f"Test {c}", "value": None, "threshold": None}
        for c, cat, st in cat_assign
    ]


def _make_summary(checks):
    s = {"OK": 0, "WARN": 0, "FAIL": 0, "NA": 0}
    for c in checks:
        s[c["status"]] += 1
    s["total"] = len(checks)
    return s


# ── Tests : constantes & imports ───────────────────────────────────

class TestConstants(unittest.TestCase):

    def test_colors_defined(self):
        for c in (COLOR_OK, COLOR_WARN, COLOR_FAIL, COLOR_NA):
            self.assertIsInstance(c, str)
            self.assertTrue(c.startswith("#"))
            self.assertEqual(len(c), 7)

    def test_status_num(self):
        # 4 statuts distincts
        self.assertEqual(len(set(STATUS_NUM.values())), 4)
        self.assertIn("OK", STATUS_NUM)
        self.assertIn("WARN", STATUS_NUM)
        self.assertIn("FAIL", STATUS_NUM)
        self.assertIn("NA", STATUS_NUM)

    def test_status_color(self):
        for s, color in STATUS_COLOR.items():
            self.assertTrue(color.startswith("#"))


# ── Tests : build_kpi_donut ────────────────────────────────────────

class TestBuildKpiDonut(unittest.TestCase):

    def test_returns_none_when_no_summary(self):
        self.assertIsNone(build_kpi_donut(None))
        self.assertIsNone(build_kpi_donut({}))

    def test_returns_none_when_zero_total(self):
        self.assertIsNone(build_kpi_donut({"OK": 0, "WARN": 0, "FAIL": 0, "NA": 0}))

    def test_returns_png_with_valid_data(self):
        checks = _make_checks()
        summary = _make_summary(checks)
        path = build_kpi_donut(summary)
        try:
            self.assertIsNotNone(path)
            self.assertTrue(os.path.isfile(path))
            self.assertTrue(_is_valid_png(path))
            # Taille > 5 KB (réaliste pour un donut à 150 dpi)
            self.assertGreater(os.path.getsize(path), 5000)
        finally:
            if path and os.path.isfile(path):
                os.unlink(path)

    def test_partial_summary(self):
        # Summary ne contenant que certaines clés
        path = build_kpi_donut({"OK": 5, "FAIL": 2})
        try:
            self.assertIsNotNone(path)
            self.assertTrue(_is_valid_png(path))
        finally:
            if path and os.path.isfile(path):
                os.unlink(path)


# ── Tests : build_category_stack ───────────────────────────────────

class TestBuildCategoryStack(unittest.TestCase):

    def test_returns_none_when_no_checks(self):
        self.assertIsNone(build_category_stack(None))
        self.assertIsNone(build_category_stack([]))

    def test_returns_png_with_valid_checks(self):
        checks = _make_checks()
        path = build_category_stack(checks)
        try:
            self.assertIsNotNone(path)
            self.assertTrue(_is_valid_png(path))
            self.assertGreater(os.path.getsize(path), 5000)
        finally:
            if path and os.path.isfile(path):
                os.unlink(path)

    def test_single_category(self):
        checks = [
            {"code": "A1", "category": "A. Pompe", "status": "OK"},
            {"code": "A2", "category": "A. Pompe", "status": "FAIL"},
        ]
        path = build_category_stack(checks)
        try:
            self.assertIsNotNone(path)
            self.assertTrue(_is_valid_png(path))
        finally:
            if path and os.path.isfile(path):
                os.unlink(path)

    def test_unknown_status(self):
        checks = [{"code": "X1", "category": "A. Pompe", "status": "UNKNOWN"}]
        path = build_category_stack(checks)
        try:
            # Ne doit pas planter — Unknown tombe dans la catégorie par défaut
            self.assertIsNotNone(path)
        finally:
            if path and os.path.isfile(path):
                os.unlink(path)


# ── Tests : build_compliance_bars ───────────────────────────────────

class TestBuildComplianceBars(unittest.TestCase):

    def test_returns_none_when_no_transient(self):
        self.assertIsNone(build_compliance_bars(None, 16.0, 0.5, 200.0))
        self.assertIsNone(build_compliance_bars({}, 16.0, 0.5, 200.0))
        self.assertIsNone(build_compliance_bars(
            {"success": False}, 16.0, 0.5, 200.0))

    def test_returns_none_when_missing_thresholds(self):
        ts = {"success": True, "max_pressure_bar": 12.0}
        self.assertIsNone(build_compliance_bars(ts, None, 0.5, 200.0))
        self.assertIsNone(build_compliance_bars(ts, 16.0, None, 200.0))
        self.assertIsNone(build_compliance_bars(ts, 16.0, 0.5, None))

    def test_returns_none_when_no_data(self):
        ts = {"success": True}  # pas de max/min/vgas
        self.assertIsNone(build_compliance_bars(ts, 16.0, 0.5, 200.0))

    def test_returns_png_with_full_data(self):
        ts = {
            "success": True,
            "max_pressure_bar": 12.0,
            "min_pressure_bar": 2.0,
            "max_gas_volume_l": 100.0,
        }
        path = build_compliance_bars(ts, 16.0, 0.5, 200.0)
        try:
            self.assertIsNotNone(path)
            self.assertTrue(_is_valid_png(path))
            self.assertGreater(os.path.getsize(path), 5000)
        finally:
            if path and os.path.isfile(path):
                os.unlink(path)

    def test_pmax_exceeds_pn(self):
        # Cas pathologique : Pmax > PN
        ts = {
            "success": True,
            "max_pressure_bar": 20.0,
            "min_pressure_bar": -0.5,
            "max_gas_volume_l": 300.0,
        }
        path = build_compliance_bars(ts, 16.0, 0.0, 200.0)
        try:
            self.assertIsNotNone(path)
            self.assertTrue(_is_valid_png(path))
        finally:
            if path and os.path.isfile(path):
                os.unlink(path)

    def test_partial_data(self):
        # Seulement Pmax
        ts = {"success": True, "max_pressure_bar": 12.0}
        path = build_compliance_bars(ts, 16.0, 0.5, 200.0)
        try:
            self.assertIsNotNone(path)
            self.assertTrue(_is_valid_png(path))
        finally:
            if path and os.path.isfile(path):
                os.unlink(path)


# ── Tests : build_profile_chart ────────────────────────────────────

class TestBuildProfileChart(unittest.TestCase):

    def test_returns_none_when_no_profile(self):
        self.assertIsNone(build_profile_chart(None))
        self.assertIsNone(build_profile_chart([]))

    def test_returns_none_with_single_point(self):
        self.assertIsNone(build_profile_chart([{"pk_m": 0, "z_m": 10}]))

    def test_returns_png_with_valid_profile(self):
        profile = [
            {"pk_m": 0, "z_m": 10},
            {"pk_m": 100, "z_m": 12},
            {"pk_m": 200, "z_m": 8},
            {"pk_m": 300, "z_m": 11},
        ]
        path = build_profile_chart(profile, pipe_dn_mm=250.0)
        try:
            self.assertIsNotNone(path)
            self.assertTrue(_is_valid_png(path))
            self.assertGreater(os.path.getsize(path), 5000)
        finally:
            if path and os.path.isfile(path):
                os.unlink(path)

    def test_with_ventouses_vidanges(self):
        profile = [
            {"pk_m": 0, "z_m": 10},
            {"pk_m": 100, "z_m": 12},
            {"pk_m": 200, "z_m": 8},
        ]
        ventouses = [{"pk_m": 100, "z_m": 12, "type": "combinée"}]
        vidanges  = [{"pk_m": 200, "z_m": 8, "type": "DN 50"}]
        path = build_profile_chart(profile, ventouses, vidanges, pipe_dn_mm=300.0)
        try:
            self.assertIsNotNone(path)
            self.assertTrue(_is_valid_png(path))
        finally:
            if path and os.path.isfile(path):
                os.unlink(path)

    def test_no_dn(self):
        profile = [{"pk_m": 0, "z_m": 10}, {"pk_m": 100, "z_m": 12}]
        path = build_profile_chart(profile, pipe_dn_mm=None)
        try:
            self.assertIsNotNone(path)
            self.assertTrue(_is_valid_png(path))
        finally:
            if path and os.path.isfile(path):
                os.unlink(path)


# ── Tests d'intégration ────────────────────────────────────────────

class TestIntegration(unittest.TestCase):

    def test_full_diagnostic_suite(self):
        """Génère les 4 graphes en cascade + un rapport Word complet."""
        from system_diagnostics import SystemDiagnostics
        from report_generator import WordReportGenerator

        class _FakePump:
            parsed = {"label": "P-1", "flow_lps": 10.0, "pump_head_m": 30.0,
                      "npsh_required_m": 3.0, "npsh_available_m": 5.0,
                      "pressure_discharge_bar": 8.0, "pressure_suction_bar": 1.5}
            curve_points = [(0, 50), (40, 10)]
            def interpolate_head(self, q): return -2.0*q + 50.0

        d = SystemDiagnostics(
            pump_parsers=[_FakePump()],
            transient_status={"success": True, "max_pressure_bar": 12.0,
                              "min_pressure_bar": 2.0, "max_gas_volume_l": 100.0},
        )
        checks = d.run_checks()
        summary = d.get_summary()

        paths = []
        try:
            paths.append(build_kpi_donut(summary))
            paths.append(build_category_stack(checks))
            paths.append(build_compliance_bars(
                {"success": True, "max_pressure_bar": 12.0,
                 "min_pressure_bar": 2.0, "max_gas_volume_l": 100.0},
                16.0, 0.5, 200.0,
            ))
            paths.append(build_profile_chart(
                [{"pk_m": 0, "z_m": 10}, {"pk_m": 100, "z_m": 12}],
                pipe_dn_mm=250.0,
            ))

            # Tous les PNGs sont valides
            for p in paths:
                self.assertIsNotNone(p)
                self.assertTrue(_is_valid_png(p))

            # Le rapport Word complet est généré
            gen = WordReportGenerator()
            doc = gen.generate(
                metadata={"nom_projet": "Test", "ingenieur": "MK", "date": "06/06/2026"},
                steady=None,
                transient={"success": True, "max_pressure_bar": 12.0,
                           "min_pressure_bar": 2.0, "max_gas_volume_l": 100.0},
                pn_label="PN 16", pn_value=16.0,
                pmin_label="Pmin", pmin_value=0.5,
                chart_png_path=None,
                diagnostics_checks=checks, diagnostics_summary=summary,
                diag_kpi_chart_path=paths[0],
                diag_category_chart_path=paths[1],
                diag_compliance_chart_path=paths[2],
                diag_profile_chart_path=paths[3],
            )
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_out:
                out = tmp_out.name
            doc.save(out)
            try:
                self.assertGreater(os.path.getsize(out), 50000)  # au moins 50 KB avec graphes
            finally:
                if os.path.isfile(out):
                    os.unlink(out)
        finally:
            for p in paths:
                if p and os.path.isfile(p):
                    os.unlink(p)


class TestImageQuality(unittest.TestCase):

    def test_all_charts_have_reasonable_size(self):
        """Les PNG générés doivent faire au moins 5 KB (qualité imprimable)."""
        checks = _make_checks()
        summary = _make_summary(checks)
        profile = [{"pk_m": 0, "z_m": 10}, {"pk_m": 100, "z_m": 12}]
        ts = {"success": True, "max_pressure_bar": 12.0,
              "min_pressure_bar": 2.0, "max_gas_volume_l": 100.0}

        paths = []
        try:
            paths.append(("kpi",       build_kpi_donut(summary)))
            paths.append(("category",  build_category_stack(checks)))
            paths.append(("compliance", build_compliance_bars(ts, 16.0, 0.5, 200.0)))
            paths.append(("profile",   build_profile_chart(profile)))

            for label, p in paths:
                self.assertIsNotNone(p, f"{label} returned None")
                self.assertTrue(_is_valid_png(p), f"{label} is not a valid PNG")
                size = os.path.getsize(p)
                self.assertGreater(size, 5000,
                    f"{label} PNG is too small ({size} bytes)")
        finally:
            for _, p in paths:
                if p and os.path.isfile(p):
                    os.unlink(p)


if __name__ == "__main__":
    unittest.main()
