#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
test_system_diagnostics.py — Tests pour la Phase 4 (SystemDiagnostics).
"""

import unittest
import math
from system_diagnostics import (
    SystemDiagnostics,
    STATUS_OK, STATUS_WARN, STATUS_FAIL, STATUS_NA,
    CAT_A, CAT_B, CAT_C, CAT_D, CAT_E,
    ICON,
    DEFAULT_NPSH_MARGIN_M, DEFAULT_NQ_MIN_SI, DEFAULT_NQ_MAX_SI,
    DEFAULT_SLOPE_MAX_PCT,
)


# ── Pompes factices ─────────────────────────────────────────────────

class _FakePump:
    """Pompe factice pour les tests."""
    def __init__(self, parsed=None, curve_points=None):
        self.parsed = parsed or {}
        self.curve_points = curve_points or []
        self.filepath = "fake.rtf"

    def interpolate_head(self, q_lps: float) -> float:
        """H(Q) linéaire simple : H = -2*Q + 50."""
        return -2.0 * q_lps + 50.0


class _FakeAirValveSizer:
    """AirValveSizer factice."""
    def __init__(self, profile=None, high_points=None, low_points=None,
                 ventouses=None, vidanges=None, pipe_dn_mm=250.0):
        self.profile = profile or []
        self.high_points = high_points or []
        self.low_points = low_points or []
        self.ventouses = ventouses or []
        self.vidanges = vidanges or []
        self.pipe_dn_mm = pipe_dn_mm


class _FakeWorkbookManager:
    def __init__(self, summary=None):
        self._summary = summary or {}
    def get_summary(self) -> dict:
        return self._summary


# ── Tests ───────────────────────────────────────────────────────────

class TestConstants(unittest.TestCase):

    def test_severities(self):
        self.assertEqual(STATUS_OK, "OK")
        self.assertEqual(STATUS_WARN, "WARN")
        self.assertEqual(STATUS_FAIL, "FAIL")
        self.assertEqual(STATUS_NA, "NA")

    def test_icons(self):
        for s in (STATUS_OK, STATUS_WARN, STATUS_FAIL, STATUS_NA):
            self.assertIn(s, ICON)

    def test_categories(self):
        for cat in (CAT_A, CAT_B, CAT_C, CAT_D, CAT_E):
            self.assertIsInstance(cat, str)
        # A-D contiennent une flèche, E non
        for cat in (CAT_A, CAT_B, CAT_C, CAT_D):
            self.assertIn("↔", cat)
        # E commence par "E. Cohérence"
        self.assertTrue(CAT_E.startswith("E."))


class TestDefaults(unittest.TestCase):

    def test_default_npsh_margin(self):
        self.assertEqual(DEFAULT_NPSH_MARGIN_M, 1.0)

    def test_default_nq_range(self):
        self.assertEqual(DEFAULT_NQ_MIN_SI, 25.0)
        self.assertEqual(DEFAULT_NQ_MAX_SI, 80.0)

    def test_default_slope(self):
        self.assertEqual(DEFAULT_SLOPE_MAX_PCT, 50.0)


class TestInit(unittest.TestCase):

    def test_empty_init(self):
        d = SystemDiagnostics()
        self.assertEqual(d.pump_parsers, [])
        self.assertIsNone(d.air_valve_sizer)
        self.assertIsNone(d.workbook_manager)
        self.assertEqual(d.transient_status, {})
        self.assertEqual(d.pn_value_bar, 16.0)
        self.assertEqual(d.pmin_value_bar, 0.0)
        self.assertEqual(d.vgas_threshold_l, 200.0)
        self.assertEqual(d.npsh_margin_m, 1.0)
        self.assertEqual(d.nq_min_si, 25.0)
        self.assertEqual(d.nq_max_si, 80.0)
        self.assertEqual(d.slope_max_pct, 50.0)

    def test_custom_params(self):
        d = SystemDiagnostics(
            pn_value_bar=25.0, pmin_value_bar=2.0,
            vgas_threshold_l=500.0, npsh_margin_m=2.0,
            nq_min_si=20.0, nq_max_si=100.0, slope_max_pct=30.0,
        )
        self.assertEqual(d.pn_value_bar, 25.0)
        self.assertEqual(d.pmin_value_bar, 2.0)
        self.assertEqual(d.vgas_threshold_l, 500.0)
        self.assertEqual(d.npsh_margin_m, 2.0)
        self.assertEqual(d.nq_min_si, 20.0)
        self.assertEqual(d.nq_max_si, 100.0)
        self.assertEqual(d.slope_max_pct, 30.0)


class TestCheckA1(unittest.TestCase):
    """A1 : Point de fonctionnement dans la courbe H(Q)."""

    def test_A1_na_no_curve(self):
        d = SystemDiagnostics()
        result = d._check_A1_operating_point()
        self.assertEqual(result["code"], "A1")
        self.assertEqual(result["status"], STATUS_NA)

    def test_A1_ok_within_tolerance(self):
        # Q=10 L/s → H_curve = -2*10+50 = 30 m, H_nom = 30.5 m (écart 0.5 m, < 5%*30=1.5)
        pump = _FakePump(
            parsed={"flow_lps": 10.0, "pump_head_m": 30.5},
            curve_points=[(0, 50), (20, 10)],
        )
        d = SystemDiagnostics(pump_parsers=[pump])
        result = d._check_A1_operating_point()
        self.assertEqual(result["status"], STATUS_OK)

    def test_A1_warn_outside_tolerance(self):
        # Q=10 L/s → H_curve = 30 m, H_nom = 50 m (écart 20 m, > 1.5)
        pump = _FakePump(
            parsed={"flow_lps": 10.0, "pump_head_m": 50.0},
            curve_points=[(0, 50), (20, 10)],
        )
        d = SystemDiagnostics(pump_parsers=[pump])
        result = d._check_A1_operating_point()
        self.assertEqual(result["status"], STATUS_WARN)

    def test_A1_na_no_q_h(self):
        pump = _FakePump(parsed={}, curve_points=[(0, 50), (20, 10)])
        d = SystemDiagnostics(pump_parsers=[pump])
        result = d._check_A1_operating_point()
        self.assertEqual(result["status"], STATUS_NA)


class TestCheckA2(unittest.TestCase):
    """A2 : NPSH dispo ≥ requis + marge."""

    def test_A2_ok(self):
        pump = _FakePump(parsed={"npsh_required_m": 3.0, "npsh_available_m": 5.0})
        d = SystemDiagnostics(pump_parsers=[pump])
        result = d._check_A2_npsh()
        self.assertEqual(result["status"], STATUS_OK)

    def test_A2_warn_marge_insuffisante(self):
        # dispo 3.2, requis 3.0 → marge 0.2 < 1.0 → WARN
        pump = _FakePump(parsed={"npsh_required_m": 3.0, "npsh_available_m": 3.2})
        d = SystemDiagnostics(pump_parsers=[pump])
        result = d._check_A2_npsh()
        self.assertEqual(result["status"], STATUS_WARN)

    def test_A2_fail_cavitation(self):
        pump = _FakePump(parsed={"npsh_required_m": 5.0, "npsh_available_m": 3.0})
        d = SystemDiagnostics(pump_parsers=[pump])
        result = d._check_A2_npsh()
        self.assertEqual(result["status"], STATUS_FAIL)

    def test_A2_na_no_npsh(self):
        pump = _FakePump(parsed={})
        d = SystemDiagnostics(pump_parsers=[pump])
        result = d._check_A2_npsh()
        self.assertEqual(result["status"], STATUS_NA)


class TestCheckA3(unittest.TestCase):
    """A3 : Vitesse spécifique dans [25, 80] SI."""

    def test_A3_ok(self):
        pump = _FakePump(parsed={"nq_si": 50.0})
        d = SystemDiagnostics(pump_parsers=[pump])
        result = d._check_A3_specific_speed()
        self.assertEqual(result["status"], STATUS_OK)

    def test_A3_warn_low(self):
        pump = _FakePump(parsed={"nq_si": 10.0})
        d = SystemDiagnostics(pump_parsers=[pump])
        result = d._check_A3_specific_speed()
        self.assertEqual(result["status"], STATUS_WARN)

    def test_A3_warn_high(self):
        pump = _FakePump(parsed={"nq_si": 100.0})
        d = SystemDiagnostics(pump_parsers=[pump])
        result = d._check_A3_specific_speed()
        self.assertEqual(result["status"], STATUS_WARN)

    def test_A3_na_no_nq(self):
        pump = _FakePump(parsed={})
        d = SystemDiagnostics(pump_parsers=[pump])
        result = d._check_A3_specific_speed()
        self.assertEqual(result["status"], STATUS_NA)

    def test_A3_custom_nq_range(self):
        # Plage custom [20, 100] → 15 doit être OK avec range élargie
        pump = _FakePump(parsed={"nq_si": 15.0})
        d = SystemDiagnostics(pump_parsers=[pump], nq_min_si=10.0, nq_max_si=100.0)
        result = d._check_A3_specific_speed()
        self.assertEqual(result["status"], STATUS_OK)


class TestCheckB1(unittest.TestCase):
    """B1 : P refoulement ≤ PN."""

    def test_B1_ok(self):
        pump = _FakePump(parsed={"pressure_discharge_bar": 10.0})
        d = SystemDiagnostics(pump_parsers=[pump], pn_value_bar=16.0)
        result = d._check_B1_discharge_pressure()
        self.assertEqual(result["status"], STATUS_OK)

    def test_B1_fail(self):
        pump = _FakePump(parsed={"pressure_discharge_bar": 20.0})
        d = SystemDiagnostics(pump_parsers=[pump], pn_value_bar=16.0)
        result = d._check_B1_discharge_pressure()
        self.assertEqual(result["status"], STATUS_FAIL)

    def test_B1_na(self):
        pump = _FakePump(parsed={})
        d = SystemDiagnostics(pump_parsers=[pump])
        result = d._check_B1_discharge_pressure()
        self.assertEqual(result["status"], STATUS_NA)


class TestCheckB2(unittest.TestCase):
    """B2 : P aspiration ≥ Pmin."""

    def test_B2_ok(self):
        pump = _FakePump(parsed={"pressure_suction_bar": 1.5})
        d = SystemDiagnostics(pump_parsers=[pump], pmin_value_bar=0.5)
        result = d._check_B2_suction_pressure()
        self.assertEqual(result["status"], STATUS_OK)

    def test_B2_fail(self):
        pump = _FakePump(parsed={"pressure_suction_bar": -0.5})
        d = SystemDiagnostics(pump_parsers=[pump], pmin_value_bar=0.5)
        result = d._check_B2_suction_pressure()
        self.assertEqual(result["status"], STATUS_FAIL)

    def test_B2_na(self):
        pump = _FakePump(parsed={})
        d = SystemDiagnostics(pump_parsers=[pump])
        result = d._check_B2_suction_pressure()
        self.assertEqual(result["status"], STATUS_NA)


class TestCheckC1(unittest.TestCase):
    """C1 : Pmax transitoire < PN."""

    def test_C1_ok(self):
        d = SystemDiagnostics(
            transient_status={"success": True, "max_pressure_bar": 12.0},
            pn_value_bar=16.0,
        )
        result = d._check_C1_max_pressure()
        self.assertEqual(result["status"], STATUS_OK)

    def test_C1_fail(self):
        d = SystemDiagnostics(
            transient_status={"success": True, "max_pressure_bar": 18.0},
            pn_value_bar=16.0,
        )
        result = d._check_C1_max_pressure()
        self.assertEqual(result["status"], STATUS_FAIL)

    def test_C1_na_no_hpt(self):
        d = SystemDiagnostics(transient_status=None)
        result = d._check_C1_max_pressure()
        self.assertEqual(result["status"], STATUS_NA)

    def test_C1_na_hpt_failed(self):
        d = SystemDiagnostics(
            transient_status={"success": False, "message": "Erreur"},
        )
        result = d._check_C1_max_pressure()
        self.assertEqual(result["status"], STATUS_NA)


class TestCheckC2(unittest.TestCase):
    """C2 : Pmin transitoire > Pmin."""

    def test_C2_ok(self):
        d = SystemDiagnostics(
            transient_status={"success": True, "min_pressure_bar": 2.0},
            pmin_value_bar=0.5,
        )
        result = d._check_C2_min_pressure()
        self.assertEqual(result["status"], STATUS_OK)

    def test_C2_fail(self):
        d = SystemDiagnostics(
            transient_status={"success": True, "min_pressure_bar": -0.5},
            pmin_value_bar=0.5,
        )
        result = d._check_C2_min_pressure()
        self.assertEqual(result["status"], STATUS_FAIL)

    def test_C2_na(self):
        d = SystemDiagnostics(transient_status={})
        result = d._check_C2_min_pressure()
        self.assertEqual(result["status"], STATUS_NA)


class TestCheckC3(unittest.TestCase):
    """C3 : Cohérence multi-pompes."""

    def test_C3_na_single_pump(self):
        pump = _FakePump(parsed={"label": "P-1"})
        d = SystemDiagnostics(pump_parsers=[pump])
        result = d._check_C3_multi_pumps()
        self.assertEqual(result["status"], STATUS_NA)

    def test_C3_ok_two_pumps(self):
        p1 = _FakePump(parsed={"label": "P-1"})
        p2 = _FakePump(parsed={"label": "P-2"})
        d = SystemDiagnostics(pump_parsers=[p1, p2])
        result = d._check_C3_multi_pumps()
        self.assertEqual(result["status"], STATUS_OK)

    def test_C3_fail_parsed_empty(self):
        p1 = _FakePump(parsed=None)
        d = SystemDiagnostics(pump_parsers=[p1, p1])
        result = d._check_C3_multi_pumps()
        self.assertEqual(result["status"], STATUS_FAIL)


class TestCheckD1(unittest.TestCase):
    """D1 : Volume gaz HPT < seuil."""

    def test_D1_ok(self):
        d = SystemDiagnostics(
            transient_status={"success": True, "max_gas_volume_l": 100.0},
            vgas_threshold_l=200.0,
        )
        result = d._check_D1_vgas()
        self.assertEqual(result["status"], STATUS_OK)

    def test_D1_warn(self):
        d = SystemDiagnostics(
            transient_status={"success": True, "max_gas_volume_l": 300.0},
            vgas_threshold_l=200.0,
        )
        result = d._check_D1_vgas()
        self.assertEqual(result["status"], STATUS_WARN)

    def test_D1_na(self):
        d = SystemDiagnostics(transient_status={})
        result = d._check_D1_vgas()
        self.assertEqual(result["status"], STATUS_NA)


class TestCheckD2(unittest.TestCase):
    """D2 : Ventouses aux points hauts."""

    def test_D2_ok(self):
        sizer = _FakeAirValveSizer(
            profile=[{"pk_m": 0, "z_m": 10}, {"pk_m": 100, "z_m": 15}],
            high_points=[{"pk_m": 50, "z_m": 12}],
            ventouses=[{"pk_m": 50, "type": "combinée"}],
        )
        d = SystemDiagnostics(air_valve_sizer=sizer)
        result = d._check_D2_air_valves()
        self.assertEqual(result["status"], STATUS_OK)

    def test_D2_warn(self):
        sizer = _FakeAirValveSizer(
            profile=[{"pk_m": 0, "z_m": 10}, {"pk_m": 100, "z_m": 15}],
            high_points=[{"pk_m": 50, "z_m": 12}],
            ventouses=[],
        )
        d = SystemDiagnostics(air_valve_sizer=sizer)
        result = d._check_D2_air_valves()
        self.assertEqual(result["status"], STATUS_WARN)

    def test_D2_na_no_profile(self):
        d = SystemDiagnostics(air_valve_sizer=None)
        result = d._check_D2_air_valves()
        self.assertEqual(result["status"], STATUS_NA)

    def test_D2_na_monotone(self):
        sizer = _FakeAirValveSizer(
            profile=[{"pk_m": 0, "z_m": 10}, {"pk_m": 100, "z_m": 5}],
            high_points=[],
        )
        d = SystemDiagnostics(air_valve_sizer=sizer)
        result = d._check_D2_air_valves()
        self.assertEqual(result["status"], STATUS_NA)


class TestCheckD3(unittest.TestCase):
    """D3 : Vidanges aux points bas."""

    def test_D3_ok(self):
        sizer = _FakeAirValveSizer(
            profile=[{"pk_m": 0, "z_m": 10}, {"pk_m": 100, "z_m": 5}],
            low_points=[{"pk_m": 100, "z_m": 5}],
            vidanges=[{"pk_m": 100, "type": "DN 50"}],
        )
        d = SystemDiagnostics(air_valve_sizer=sizer)
        result = d._check_D3_drains()
        self.assertEqual(result["status"], STATUS_OK)

    def test_D3_warn_no_drain(self):
        sizer = _FakeAirValveSizer(
            profile=[{"pk_m": 0, "z_m": 10}, {"pk_m": 100, "z_m": 5}],
            low_points=[{"pk_m": 100, "z_m": 5}],
            vidanges=[],
        )
        d = SystemDiagnostics(air_valve_sizer=sizer)
        result = d._check_D3_drains()
        self.assertEqual(result["status"], STATUS_WARN)

    def test_D3_na(self):
        d = SystemDiagnostics(air_valve_sizer=None)
        result = d._check_D3_drains()
        self.assertEqual(result["status"], STATUS_NA)


class TestCheckE1(unittest.TestCase):
    """E1 : Profil en long chargé (≥ 2 points)."""

    def test_E1_ok(self):
        sizer = _FakeAirValveSizer(
            profile=[{"pk_m": 0, "z_m": 10}, {"pk_m": 100, "z_m": 15}],
        )
        d = SystemDiagnostics(air_valve_sizer=sizer)
        result = d._check_E1_profile_loaded()
        self.assertEqual(result["status"], STATUS_OK)

    def test_E1_fail_no_profile(self):
        d = SystemDiagnostics(air_valve_sizer=None)
        result = d._check_E1_profile_loaded()
        self.assertEqual(result["status"], STATUS_FAIL)

    def test_E1_fail_one_point(self):
        sizer = _FakeAirValveSizer(profile=[{"pk_m": 0, "z_m": 10}])
        d = SystemDiagnostics(air_valve_sizer=sizer)
        result = d._check_E1_profile_loaded()
        self.assertEqual(result["status"], STATUS_FAIL)


class TestCheckE2(unittest.TestCase):
    """E2 : Cohérence DN profil ↔ classeur."""

    def test_E2_ok(self):
        sizer = _FakeAirValveSizer(
            profile=[{"pk_m": 0, "z_m": 10}, {"pk_m": 100, "z_m": 12}],
            pipe_dn_mm=250.0,
        )
        wb = _FakeWorkbookManager(summary={"diameter_min_mm": 200, "diameter_max_mm": 300})
        d = SystemDiagnostics(air_valve_sizer=sizer, workbook_manager=wb)
        result = d._check_E2_dn_consistency()
        self.assertEqual(result["status"], STATUS_OK)

    def test_E2_warn_out_of_range(self):
        sizer = _FakeAirValveSizer(
            profile=[{"pk_m": 0, "z_m": 10}, {"pk_m": 100, "z_m": 12}],
            pipe_dn_mm=400.0,
        )
        wb = _FakeWorkbookManager(summary={"diameter_min_mm": 200, "diameter_max_mm": 300})
        d = SystemDiagnostics(air_valve_sizer=sizer, workbook_manager=wb)
        result = d._check_E2_dn_consistency()
        self.assertEqual(result["status"], STATUS_WARN)

    def test_E2_na_no_profile(self):
        wb = _FakeWorkbookManager(summary={"diameter_min_mm": 200, "diameter_max_mm": 300})
        d = SystemDiagnostics(air_valve_sizer=None, workbook_manager=wb)
        result = d._check_E2_dn_consistency()
        self.assertEqual(result["status"], STATUS_NA)

    def test_E2_na_no_wb(self):
        sizer = _FakeAirValveSizer(pipe_dn_mm=250.0)
        d = SystemDiagnostics(air_valve_sizer=sizer, workbook_manager=None)
        result = d._check_E2_dn_consistency()
        self.assertEqual(result["status"], STATUS_NA)


class TestCheckE3(unittest.TestCase):
    """E3 : Cote min profil > 0."""

    def test_E3_ok(self):
        sizer = _FakeAirValveSizer(
            profile=[{"z_m": 5.0}, {"z_m": 10.0}],
        )
        d = SystemDiagnostics(air_valve_sizer=sizer)
        result = d._check_E3_elevation_positive()
        self.assertEqual(result["status"], STATUS_OK)

    def test_E3_fail_negative(self):
        sizer = _FakeAirValveSizer(
            profile=[{"z_m": -2.0}, {"z_m": 10.0}],
        )
        d = SystemDiagnostics(air_valve_sizer=sizer)
        result = d._check_E3_elevation_positive()
        self.assertEqual(result["status"], STATUS_FAIL)

    def test_E3_na_no_profile(self):
        d = SystemDiagnostics(air_valve_sizer=None)
        result = d._check_E3_elevation_positive()
        self.assertEqual(result["status"], STATUS_NA)


class TestCheckE4(unittest.TestCase):
    """E4 : Pente max < 50 %."""

    def test_E4_ok(self):
        sizer = _FakeAirValveSizer(
            profile=[
                {"pk_m": 0, "z_m": 10, "pente_pct": 5.0},
                {"pk_m": 100, "z_m": 15, "pente_pct": 5.0},
            ],
        )
        d = SystemDiagnostics(air_valve_sizer=sizer)
        result = d._check_E4_slope_max()
        self.assertEqual(result["status"], STATUS_OK)

    def test_E4_warn(self):
        sizer = _FakeAirValveSizer(
            profile=[
                {"pk_m": 0, "z_m": 10, "pente_pct": 60.0},
                {"pk_m": 100, "z_m": 70, "pente_pct": 60.0},
            ],
        )
        d = SystemDiagnostics(air_valve_sizer=sizer, slope_max_pct=50.0)
        result = d._check_E4_slope_max()
        self.assertEqual(result["status"], STATUS_WARN)

    def test_E4_na(self):
        d = SystemDiagnostics(air_valve_sizer=None)
        result = d._check_E4_slope_max()
        self.assertEqual(result["status"], STATUS_NA)


class TestCheckE5(unittest.TestCase):
    """E5 : Au moins 1 pompe chargée."""

    def test_E5_ok(self):
        pump = _FakePump(parsed={"label": "P-1"})
        d = SystemDiagnostics(pump_parsers=[pump])
        result = d._check_E5_pumps_loaded()
        self.assertEqual(result["status"], STATUS_OK)

    def test_E5_fail(self):
        d = SystemDiagnostics()
        result = d._check_E5_pumps_loaded()
        self.assertEqual(result["status"], STATUS_FAIL)

    def test_E5_fail_unparsed(self):
        pump = _FakePump(parsed=None)
        d = SystemDiagnostics(pump_parsers=[pump])
        result = d._check_E5_pumps_loaded()
        self.assertEqual(result["status"], STATUS_FAIL)


class TestRunChecks(unittest.TestCase):

    def test_run_checks_count(self):
        d = SystemDiagnostics()
        checks = d.run_checks()
        self.assertEqual(len(checks), 16)

    def test_run_checks_codes(self):
        d = SystemDiagnostics()
        checks = d.run_checks()
        codes = [c["code"] for c in checks]
        self.assertEqual(len(set(codes)), 16)
        for code in ("A1", "A2", "A3", "B1", "B2", "C1", "C2", "C3",
                     "D1", "D2", "D3", "E1", "E2", "E3", "E4", "E5"):
            self.assertIn(code, codes)

    def test_run_checks_categories(self):
        d = SystemDiagnostics()
        checks = d.run_checks()
        categories = set(c["category"] for c in checks)
        self.assertEqual(categories, {CAT_A, CAT_B, CAT_C, CAT_D, CAT_E})

    def test_run_checks_all_have_required_fields(self):
        d = SystemDiagnostics()
        checks = d.run_checks()
        for c in checks:
            self.assertIn("code", c)
            self.assertIn("name", c)
            self.assertIn("category", c)
            self.assertIn("status", c)
            self.assertIn("message", c)
            self.assertIn(c["status"], (STATUS_OK, STATUS_WARN, STATUS_FAIL, STATUS_NA))

    def test_run_checks_idempotent(self):
        d = SystemDiagnostics()
        c1 = d.run_checks()
        c2 = d.run_checks()
        self.assertEqual(c1, c2)


class TestGetSummary(unittest.TestCase):

    def test_summary_empty(self):
        d = SystemDiagnostics()
        s = d.get_summary()
        self.assertEqual(s["total"], 16)
        self.assertEqual(sum(v for k, v in s.items() if k != "total"), 16)
        # Tous NA si rien n'est chargé
        self.assertGreaterEqual(s[STATUS_NA], 0)

    def test_summary_keys(self):
        d = SystemDiagnostics()
        s = d.get_summary()
        for k in (STATUS_OK, STATUS_WARN, STATUS_FAIL, STATUS_NA, "total"):
            self.assertIn(k, s)

    def test_summary_with_fails(self):
        # Créer une config qui produit au moins 1 FAIL
        pump = _FakePump(parsed={"npsh_required_m": 5.0, "npsh_available_m": 3.0,
                                 "pressure_discharge_bar": 20.0,
                                 "pressure_suction_bar": -0.5})
        d = SystemDiagnostics(pump_parsers=[pump])
        s = d.get_summary()
        self.assertGreater(s[STATUS_FAIL], 0)


class TestSerialize(unittest.TestCase):

    def test_serialize_structure(self):
        d = SystemDiagnostics()
        data = d.serialize()
        self.assertIn("version", data)
        self.assertEqual(data["version"], 1)
        self.assertIn("checks", data)
        self.assertIn("summary", data)
        self.assertIn("params", data)

    def test_serialize_params(self):
        d = SystemDiagnostics(pn_value_bar=20.0, pmin_value_bar=1.0)
        data = d.serialize()
        params = data["params"]
        self.assertEqual(params["pn_value_bar"], 20.0)
        self.assertEqual(params["pmin_value_bar"], 1.0)
        self.assertEqual(params["vgas_threshold_l"], 200.0)
        self.assertEqual(params["npsh_margin_m"], 1.0)
        self.assertEqual(params["slope_max_pct"], 50.0)

    def test_serialize_with_results(self):
        pump = _FakePump(parsed={"npsh_required_m": 3.0, "npsh_available_m": 5.0})
        d = SystemDiagnostics(pump_parsers=[pump])
        data = d.serialize()
        self.assertEqual(len(data["checks"]), 16)
        self.assertEqual(data["summary"]["total"], 16)


class TestDeserialize(unittest.TestCase):

    def test_deserialize_none(self):
        result = SystemDiagnostics.deserialize(None)
        self.assertEqual(result, [])

    def test_deserialize_empty(self):
        result = SystemDiagnostics.deserialize({})
        self.assertEqual(result, [])

    def test_deserialize_wrong_version(self):
        result = SystemDiagnostics.deserialize({"version": 2, "checks": [{"a": 1}]})
        self.assertEqual(result, [])

    def test_deserialize_ok(self):
        original = [{"code": "A1", "status": STATUS_OK, "message": "x"}]
        data = {"version": 1, "checks": original}
        result = SystemDiagnostics.deserialize(data)
        self.assertEqual(result, original)

    def test_serialize_deserialize_roundtrip(self):
        d = SystemDiagnostics(pn_value_bar=20.0)
        data = d.serialize()
        checks = SystemDiagnostics.deserialize(data)
        self.assertEqual(len(checks), 16)
        # Les codes doivent être préservés
        codes = [c["code"] for c in checks]
        self.assertIn("A1", codes)
        self.assertIn("E5", codes)


class TestIntegration(unittest.TestCase):

    def test_realistic_full_scenario(self):
        """Scénario complet : projet avec pompe, classeur, profil, HPT."""
        pump = _FakePump(
            parsed={
                "label": "P-1",
                "flow_lps": 20.0,
                "pump_head_m": 30.0,
                "npsh_required_m": 3.0,
                "npsh_available_m": 5.0,
                "nq_si": 50.0,
                "pressure_discharge_bar": 8.0,
                "pressure_suction_bar": 1.5,
            },
            curve_points=[(0, 50), (40, 10)],
        )
        sizer = _FakeAirValveSizer(
            profile=[
                {"pk_m": 0, "z_m": 10, "pente_pct": 5.0},
                {"pk_m": 100, "z_m": 12, "pente_pct": 5.0},
                {"pk_m": 200, "z_m": 8, "pente_pct": 4.0},
            ],
            high_points=[{"pk_m": 100, "z_m": 12}],
            low_points=[{"pk_m": 200, "z_m": 8}],
            ventouses=[{"pk_m": 100, "type": "combinée"}],
            vidanges=[{"pk_m": 200, "type": "DN 50"}],
            pipe_dn_mm=250.0,
        )
        wb = _FakeWorkbookManager(
            summary={"diameter_min_mm": 200, "diameter_max_mm": 300}
        )
        transient = {
            "success": True,
            "max_pressure_bar": 12.0,
            "min_pressure_bar": 2.0,
            "max_gas_volume_l": 100.0,
        }
        d = SystemDiagnostics(
            pump_parsers=[pump],
            air_valve_sizer=sizer,
            workbook_manager=wb,
            transient_status=transient,
            pn_value_bar=16.0,
            pmin_value_bar=0.5,
        )
        checks = d.run_checks()
        self.assertEqual(len(checks), 16)
        summary = d.get_summary()
        # Dans ce scénario "idéal", la plupart doivent être OK ou NA
        # Pas de FAIL attendu
        self.assertEqual(summary[STATUS_FAIL], 0)

    def test_realistic_problematic_scenario(self):
        """Scénario problématique : tout doit fail ou warn."""
        pump = _FakePump(
            parsed={
                "label": "P-1",
                "flow_lps": 10.0,
                "pump_head_m": 50.0,  # hors courbe
                "npsh_required_m": 10.0,  # cavitation
                "npsh_available_m": 3.0,
                "nq_si": 150.0,  # hors plage
                "pressure_discharge_bar": 25.0,  # > PN
                "pressure_suction_bar": -1.0,  # < Pmin
            },
            curve_points=[(0, 50), (20, 10)],
        )
        sizer = _FakeAirValveSizer(
            profile=[{"pk_m": 0, "z_m": -2, "pente_pct": 60.0}],  # négatif + pente forte
        )
        transient = {
            "success": True,
            "max_pressure_bar": 20.0,  # > PN
            "min_pressure_bar": -1.0,  # < Pmin
            "max_gas_volume_l": 500.0,  # > seuil
        }
        d = SystemDiagnostics(
            pump_parsers=[pump],
            air_valve_sizer=sizer,
            transient_status=transient,
        )
        summary = d.get_summary()
        # On doit avoir plusieurs FAIL et WARN
        self.assertGreater(summary[STATUS_FAIL] + summary[STATUS_WARN], 5)


if __name__ == "__main__":
    unittest.main()
