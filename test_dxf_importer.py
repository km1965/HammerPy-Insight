#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Tests unitaires pour l'import DXF (dxf_profile_importer) et le format
CSV Bentley (load_profile_bentley_csv dans air_valve_sizing).
"""

import os
import sys
import math
import tempfile
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dxf_profile_importer import (
    _normalize_layer_name,
    _match_layer,
    load_dxf_plan,
    load_dxf_profile,
    load_dxf_both,
    list_dxf_layers,
    HAS_EZDXF,
)
from air_valve_sizing import AirValveSizing


# =====================================================================
# Tests utilitaires (normalisation de calques)
# =====================================================================

class TestNormalizeLayer:
    """Tests de la normalisation des noms de calques."""

    def test_lowercase(self):
        # "É" et "É" se décomposent en "E" après NFD + filtrage des marques
        assert _normalize_layer_name("TRACÉ EN PLAN") == "trace en plan"

    def test_remove_accents(self):
        # Les accents sont décomposés (NFD) puis retirés
        n = _normalize_layer_name("Tracé en Plan")
        assert "tr" in n
        assert n == "trace en plan"

    def test_strip_whitespace(self):
        assert _normalize_layer_name("  PROFIL  ") == "profil"

    def test_empty(self):
        assert _normalize_layer_name("") == ""
        assert _normalize_layer_name(None) == ""


class TestMatchLayer:
    """Tests de la détection de calque par pattern."""

    def test_plan_match(self):
        assert _match_layer("Tracé en plan", ["trace en plan", "plan"])
        assert _match_layer("PLAN", ["plan"])
        assert _match_layer("Plan View", ["plan view"])

    def test_profile_match(self):
        assert _match_layer("Profil en long", ["profil en long", "profil"])
        assert _match_layer("PROFILE", ["profile"])
        assert _match_layer("Longitudinal Profile", ["longitudinal profile"])

    def test_no_match(self):
        assert not _match_layer("Random Layer", ["plan", "profile"])
        assert not _match_layer("", ["plan", "profile"])

    def test_case_insensitive(self):
        # Le pattern "plan" doit matcher le calque "PLAN" (insensible à la casse)
        assert _match_layer("PLAN", ["plan"])
        # _match_layer ne normalise pas le pattern — c'est _normalize_layer_name qui le fait
        # Le pattern "PLAN" ne matche pas "plan" si on ne normalise pas le pattern
        # Mais c'est l'usage interne : les patterns sont déjà en lowercase


# =====================================================================
# Tests DXF (skip si ezdxf non installé)
# =====================================================================

@pytest.fixture
def sample_dxf_file():
    """Crée un fichier DXF de test en mémoire avec 2 calques."""
    if not HAS_EZDXF:
        pytest.skip("ezdxf non installé")
    import ezdxf
    doc = ezdxf.new(dxfversion="R2010")
    msp = doc.modelspace()

    # Calque "Tracé en plan" : 5 sommets (XY horizontale)
    plan_layer = "Tracé en plan"
    plan_pts = [(0, 0), (100, 0), (200, 50), (300, 50), (400, 0)]
    msp.add_lwpolyline(plan_pts, dxfattribs={"layer": plan_layer})

    # Calque "Profil en long" : 5 sommets (X=PK, Y=altitude)
    prof_layer = "Profil en long"
    prof_pts = [(0, 100), (100, 105), (200, 110), (300, 108), (400, 115)]
    msp.add_lwpolyline(prof_pts, dxfattribs={"layer": prof_layer})

    # Calque "Autre" : ne doit PAS être capturé
    other_layer = "Annotation"
    msp.add_lwpolyline([(0, 0), (50, 50)], dxfattribs={"layer": other_layer})

    f = tempfile.NamedTemporaryFile(suffix=".dxf", delete=False)
    f.close()
    doc.saveas(f.name)
    yield f.name
    os.unlink(f.name)


class TestDxfImporter:
    """Tests d'import DXF (plan + profil)."""

    def test_list_layers(self, sample_dxf_file):
        layers = list_dxf_layers(sample_dxf_file)
        assert "Tracé en plan" in layers
        assert "Profil en long" in layers
        assert "Annotation" in layers

    def test_load_plan(self, sample_dxf_file):
        plan = load_dxf_plan(sample_dxf_file)
        assert len(plan) == 5
        assert plan[0] == (0.0, 0.0)
        assert plan[-1] == (400.0, 0.0)
        assert plan[2] == (200.0, 50.0)

    def test_load_profile(self, sample_dxf_file):
        prof = load_dxf_profile(sample_dxf_file)
        assert len(prof) == 5
        assert prof[0] == (0.0, 100.0)
        assert prof[-1] == (400.0, 115.0)
        assert prof[2] == (200.0, 110.0)

    def test_load_both(self, sample_dxf_file):
        result = load_dxf_both(sample_dxf_file)
        assert "plan" in result
        assert "profile" in result
        assert "plan_layer" in result
        assert "profile_layer" in result
        assert result["plan_layer"] == "Tracé en plan"
        assert result["profile_layer"] == "Profil en long"
        assert len(result["plan"]) == 5
        assert len(result["profile"]) == 5

    def test_load_invalid_file(self):
        # Fichier inexistant
        plan = load_dxf_plan("nonexistent.dxf")
        assert plan == []
        prof = load_dxf_profile("nonexistent.dxf")
        assert prof == []

    def test_load_only_plan(self):
        """DXF avec uniquement un calque plan."""
        if not HAS_EZDXF:
            pytest.skip("ezdxf non installé")
        import ezdxf
        doc = ezdxf.new(dxfversion="R2010")
        msp = doc.modelspace()
        msp.add_lwpolyline(
            [(0, 0), (100, 100)],
            dxfattribs={"layer": "Tracé en plan"}
        )
        f = tempfile.NamedTemporaryFile(suffix=".dxf", delete=False)
        f.close()
        doc.saveas(f.name)
        try:
            result = load_dxf_both(f.name)
            assert len(result["plan"]) == 2
            assert result["profile"] == []
        finally:
            os.unlink(f.name)


# =====================================================================
# Tests CSV Bentley FlexTable
# =====================================================================

@pytest.fixture
def bentley_csv_file():
    """Crée un CSV Bentley FlexTable de test (3 points alignés)."""
    content = (
        'FlexTable: Junction Table;;;\n'
        'Label;"X\n'
        '(m)";"Y\n'
        '(m)";"Elevation\n'
        '(m)"\n'
        '1;0,0;0,0;100,0\n'
        '2;100,0;0,0;102,0\n'
        '3;200,0;0,0;98,0\n'
    )
    f = tempfile.NamedTemporaryFile(
        suffix=".csv", delete=False, mode='w', encoding='utf-8-sig', newline=''
    )
    f.write(content)
    f.close()
    yield f.name
    os.unlink(f.name)


class TestBentleyCsv:
    """Tests du chargeur CSV FlexTable (load_profile_bentley_csv)."""

    def test_load_bentley_basic(self, bentley_csv_file):
        avs = AirValveSizing(pipe_dn_mm=250.0)
        assert avs.load_profile_bentley_csv(bentley_csv_file) is True
        assert len(avs.profile) == 3
        # PK[0] = 0, PK[1] = 100, PK[2] = 200 (3 points alignés sur X)
        assert avs.profile[0]["pk_m"] == 0.0
        assert avs.profile[1]["pk_m"] == 100.0
        assert avs.profile[2]["pk_m"] == 200.0
        # Z[0] = 100, Z[1] = 102, Z[2] = 98
        assert avs.profile[0]["z_m"] == 100.0
        assert avs.profile[1]["z_m"] == 102.0
        assert avs.profile[2]["z_m"] == 98.0

    def test_load_bentley_diagonal(self):
        """Points en diagonale : PK = distance cumulée par sqrt(dx² + dy²)."""
        content = (
            'FlexTable: Junction Table;;;\n'
            'Label;X;Y;Elevation\n'
            '1;0;0;100\n'
            '2;30;40;110\n'  # distance = 50
            '3;30;80;120\n'  # distance = 50 + 40 = 90
        )
        f = tempfile.NamedTemporaryFile(
            suffix=".csv", delete=False, mode='w', encoding='utf-8-sig', newline=''
        )
        f.write(content)
        f.close()
        try:
            avs = AirValveSizing(pipe_dn_mm=200.0)
            assert avs.load_profile_bentley_csv(f.name) is True
            assert len(avs.profile) == 3
            assert avs.profile[0]["pk_m"] == 0.0
            assert avs.profile[1]["pk_m"] == pytest.approx(50.0, rel=1e-3)
            assert avs.profile[2]["pk_m"] == pytest.approx(90.0, rel=1e-3)
        finally:
            os.unlink(f.name)

    def test_load_bentley_slopes_computed(self, bentley_csv_file):
        avs = AirValveSizing(pipe_dn_mm=250.0)
        avs.load_profile_bentley_csv(bentley_csv_file)
        # Pente entre P0 et P1 : (102-100)/100 = 2%
        assert avs.profile[1]["pente_pct"] == pytest.approx(2.0, rel=1e-3)
        # Pente entre P1 et P2 : (98-102)/100 = -4%
        assert avs.profile[2]["pente_pct"] == pytest.approx(-4.0, rel=1e-3)

    def test_load_bentley_detects_high_low(self, bentley_csv_file):
        avs = AirValveSizing(pipe_dn_mm=250.0)
        avs.load_profile_bentley_csv(bentley_csv_file)
        # Profil : 100 → 102 → 98 → point haut à PK=100, point bas à PK=200
        assert len(avs.high_points) >= 1
        assert any(hp["pk_m"] == 100.0 for hp in avs.high_points)
        assert len(avs.low_points) >= 1
        assert any(lp["pk_m"] == 200.0 for lp in avs.low_points)

    def test_load_bentley_with_semicolon_decimal_comma(self, bentley_csv_file):
        """Vérifie la gestion du séparateur ';' et décimale ',' (français)."""
        avs = AirValveSizing(pipe_dn_mm=250.0)
        assert avs.load_profile_bentley_csv(bentley_csv_file) is True
        # Les valeurs '100,0', '102,0', '98,0' doivent être parsées
        assert avs.profile[0]["z_m"] == 100.0
        assert avs.profile[1]["z_m"] == 102.0

    def test_load_bentley_too_few_points(self):
        """Moins de 2 points valides → False."""
        content = (
            'FlexTable: Junction Table;;;\n'
            'Label;X;Y;Elevation\n'
            '1;0;0;100\n'
        )
        f = tempfile.NamedTemporaryFile(
            suffix=".csv", delete=False, mode='w', encoding='utf-8-sig', newline=''
        )
        f.write(content)
        f.close()
        try:
            avs = AirValveSizing(pipe_dn_mm=250.0)
            assert avs.load_profile_bentley_csv(f.name) is False
        finally:
            os.unlink(f.name)

    def test_load_bentley_invalid_file(self):
        avs = AirValveSizing(pipe_dn_mm=250.0)
        assert avs.load_profile_bentley_csv("nonexistent.csv") is False

    def test_load_bentley_utf16_le(self):
        """Test avec encodage UTF-16 LE (BOM FF FE) — export Excel par défaut."""
        content_bytes = b'\xff\xfe'  # UTF-16 LE BOM
        for line in [
            'FlexTable: Junction Table;;;',
            'Label;X;Y;Elevation',
            '1;344013,83;354052,41;109,11',
            '3;344061,11;354068,70;107,54',
            '4;344084,70;354076,84;104,02',
        ]:
            content_bytes += (line + '\r\n').encode('utf-16-le')
        f = tempfile.NamedTemporaryFile(suffix=".csv", delete=False)
        f.write(content_bytes)
        f.close()
        try:
            avs = AirValveSizing(pipe_dn_mm=250.0)
            assert avs.load_profile_bentley_csv(f.name) is True
            assert len(avs.profile) == 3
        finally:
            os.unlink(f.name)

    def test_load_bentley_utf16_be(self):
        """Test avec encodage UTF-16 BE (BOM FE FF)."""
        content_bytes = b'\xfe\xff'  # UTF-16 BE BOM
        for line in [
            'FlexTable: Junction Table;;;',
            'Label;X;Y;Elevation',
            '1;344013,83;354052,41;109,11',
            '3;344061,11;354068,70;107,54',
        ]:
            content_bytes += (line + '\r\n').encode('utf-16-be')
        f = tempfile.NamedTemporaryFile(suffix=".csv", delete=False)
        f.write(content_bytes)
        f.close()
        try:
            avs = AirValveSizing(pipe_dn_mm=250.0)
            assert avs.load_profile_bentley_csv(f.name) is True
            assert len(avs.profile) == 2
        finally:
            os.unlink(f.name)

    def test_load_bentley_multiline_headers(self):
        """Test avec en-têtes colonnes multi-lignes (Label 'X\\n(m)' etc.)."""
        content = (
            'FlexTable: Junction Table;;;\n'
            'Label;"X\n(m)";"Y\n(m)";"Elevation\n(m)"\n'
            '1;344013,83;354052,41;109,11\n'
            '3;344061,11;354068,70;107,54\n'
            '4;344084,70;354076,84;104,02\n'
        )
        f = tempfile.NamedTemporaryFile(
            suffix=".csv", delete=False, mode='w', encoding='utf-8-sig', newline=''
        )
        f.write(content)
        f.close()
        try:
            avs = AirValveSizing(pipe_dn_mm=250.0)
            assert avs.load_profile_bentley_csv(f.name) is True
            assert len(avs.profile) == 3
            assert avs.profile[0]["z_m"] == 109.11
        finally:
            os.unlink(f.name)
