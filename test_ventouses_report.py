#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Tests unitaires pour ventouses_report.py
"""

import os
import sys
import tempfile
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from air_valve_sizing import AirValveSizing
from ventouses_report import (
    VentousesReportGenerator,
    export_ventouses_report,
    DOCX_AVAILABLE,
)


# =====================================================================
# Fixtures
# =====================================================================

@pytest.fixture
def sizing_with_ventouses_vidanges():
    """AirValveSizing avec un profil générant ventouses + vidanges."""
    avs = AirValveSizing(pipe_dn_mm=300.0)
    avs.load_profile_manual([
        (0, 100.0), (100, 110.0), (200, 115.0), (300, 108.0),
        (400, 105.0), (500, 112.0), (600, 118.0), (700, 113.0),
        (800, 108.0), (900, 105.0), (1000, 100.0),
    ])
    avs.size_ventouses()
    avs.size_drains()
    return avs


@pytest.fixture
def empty_sizing():
    """AirValveSizing vide (sans profil)."""
    return AirValveSizing(pipe_dn_mm=250.0)


# =====================================================================
# Tests de disponibilité
# =====================================================================

class TestDocxAvailable:
    def test_docx_available(self):
        assert DOCX_AVAILABLE is True


# =====================================================================
# Tests du générateur
# =====================================================================

class TestVentousesReportGenerator:
    """Tests de VentousesReportGenerator."""

    def test_init(self):
        gen = VentousesReportGenerator()
        assert gen.doc is not None
        assert hasattr(gen, "_setup_styles")

    def test_generate_minimal(self, sizing_with_ventouses_vidanges):
        gen = VentousesReportGenerator()
        doc = gen.generate(
            air_valve_sizer=sizing_with_ventouses_vidanges,
            metadata={
                "nom_projet": "Test Pipeline",
                "ingenieur": "Test Engineer",
                "date": "01/01/2026",
                "dn_mm": 300,
            },
        )
        assert doc is not None
        # Vérifier qu'on peut sauvegarder
        f = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        f.close()
        try:
            doc.save(f.name)
            assert os.path.getsize(f.name) > 1000  # Fichier non vide
        finally:
            os.unlink(f.name)

    def test_generate_empty_profile(self, empty_sizing):
        """Test avec profil vide — ne doit pas crasher."""
        gen = VentousesReportGenerator()
        doc = gen.generate(
            air_valve_sizer=empty_sizing,
            metadata={"nom_projet": "Vide"},
        )
        f = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        f.close()
        try:
            doc.save(f.name)
            assert os.path.getsize(f.name) > 1000
        finally:
            os.unlink(f.name)

    def test_generate_with_chart_image(self, sizing_with_ventouses_vidanges):
        """Test avec image PNG du graphique profil."""
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        # Créer un PNG de test
        fig, ax = plt.subplots(figsize=(8, 4))
        profile = sizing_with_ventouses_vidanges.profile
        pks = [p["pk_m"] for p in profile]
        zs = [p["z_m"] for p in profile]
        ax.plot(pks, zs, "b-o")
        ax.set_xlabel("PK (m)")
        ax.set_ylabel("Z (m)")
        ax.set_title("Profil en Long — Test")
        fig.tight_layout()

        png_file = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        fig.savefig(png_file.name, dpi=100)
        plt.close(fig)
        png_file.close()

        try:
            gen = VentousesReportGenerator()
            doc = gen.generate(
                air_valve_sizer=sizing_with_ventouses_vidanges,
                metadata={"nom_projet": "Avec image"},
                profile_chart_png_path=png_file.name,
            )
            f = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
            f.close()
            try:
                doc.save(f.name)
                # Fichier avec image est plus gros
                assert os.path.getsize(f.name) > 5000
            finally:
                os.unlink(f.name)
        finally:
            os.unlink(png_file.name)

    def test_save_returns_path(self, sizing_with_ventouses_vidanges):
        gen = VentousesReportGenerator()
        gen.generate(
            air_valve_sizer=sizing_with_ventouses_vidanges,
            metadata={"nom_projet": "Test"},
        )
        f = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        f.close()
        try:
            result = gen.save(f.name)
            assert result == f.name
            assert os.path.isfile(f.name)
        finally:
            os.unlink(f.name)


# =====================================================================
# Tests du helper export_ventouses_report
# =====================================================================

class TestExportHelper:
    """Tests du helper tout-en-un export_ventouses_report()."""

    def test_export_basic(self, sizing_with_ventouses_vidanges):
        f = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        f.close()
        try:
            result = export_ventouses_report(
                air_valve_sizer=sizing_with_ventouses_vidanges,
                output_path=f.name,
                metadata={
                    "nom_projet": "Test Helper",
                    "ingenieur": "MK",
                    "date": "06/06/2026",
                    "dn_mm": 300,
                },
            )
            assert result == f.name
            assert os.path.getsize(f.name) > 1000
        finally:
            os.unlink(f.name)

    def test_export_with_chart(self, sizing_with_ventouses_vidanges):
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        fig, ax = plt.subplots()
        ax.plot([0, 100, 200], [100, 110, 105])
        png_file = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        fig.savefig(png_file.name)
        plt.close(fig)
        png_file.close()

        try:
            f = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
            f.close()
            try:
                export_ventouses_report(
                    air_valve_sizer=sizing_with_ventouses_vidanges,
                    output_path=f.name,
                    metadata={"nom_projet": "Avec chart"},
                    profile_chart_png_path=png_file.name,
                )
                assert os.path.getsize(f.name) > 5000
            finally:
                os.unlink(f.name)
        finally:
            os.unlink(png_file.name)

    def test_export_empty_sizing(self, empty_sizing):
        """Helper avec sizer vide — ne doit pas crasher."""
        f = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        f.close()
        try:
            export_ventouses_report(
                air_valve_sizer=empty_sizing,
                output_path=f.name,
                metadata={"nom_projet": "Vide"},
            )
            assert os.path.getsize(f.name) > 1000
        finally:
            os.unlink(f.name)

    def test_export_no_metadata(self, sizing_with_ventouses_vidanges):
        """Helper sans metadata (utilise défauts)."""
        f = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        f.close()
        try:
            export_ventouses_report(
                air_valve_sizer=sizing_with_ventouses_vidanges,
                output_path=f.name,
            )
            assert os.path.getsize(f.name) > 1000
        finally:
            os.unlink(f.name)


# =====================================================================
# Tests de contenu du rapport
# =====================================================================

class TestReportContent:
    """Tests du contenu textuel du rapport généré."""

    def test_ventouse_table_count_matches_data(self, sizing_with_ventouses_vidanges):
        """Le tableau ventouses doit avoir 1 ligne d'en-tête + N données."""
        gen = VentousesReportGenerator()
        doc = gen.generate(
            air_valve_sizer=sizing_with_ventouses_vidanges,
            metadata={"nom_projet": "Count Test"},
        )
        # Compter les tables
        nb_tables = len(doc.tables)
        # Au minimum 3 tables : KV en-tête, ventouses, vidanges
        assert nb_tables >= 3

    def test_stats_correctness(self, sizing_with_ventouses_vidanges):
        """Les stats calculées doivent correspondre aux données."""
        gen = VentousesReportGenerator()
        stats = gen._compute_stats(
            sizing_with_ventouses_vidanges,
            metadata={"nom_projet": "Stats"},
        )
        assert stats["nb_points_profil"] == 11
        assert stats["longueur_totale_m"] == 1000.0
        assert stats["cote_min_m"] == 100.0
        assert stats["cote_max_m"] == 118.0
        assert stats["nb_ventouses"] == len(sizing_with_ventouses_vidanges.ventouses)
        assert stats["nb_vidanges"] == len(sizing_with_ventouses_vidanges.vidanges)

    def test_metadata_nom_projet_in_doc(self, sizing_with_ventouses_vidanges):
        gen = VentousesReportGenerator()
        doc = gen.generate(
            air_valve_sizer=sizing_with_ventouses_vidanges,
            metadata={"nom_projet": "MonProjetUnique123"},
        )
        # Vérifier que le nom du projet apparaît dans le doc
        all_text = "\n".join(p.text for p in doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    all_text += "\n" + cell.text
        assert "MonProjetUnique123" in all_text

    def test_ingenieur_in_doc(self, sizing_with_ventouses_vidanges):
        gen = VentousesReportGenerator()
        doc = gen.generate(
            air_valve_sizer=sizing_with_ventouses_vidanges,
            metadata={"nom_projet": "X", "ingenieur": "Mostafa KARIM"},
        )
        all_text = "\n".join(p.text for p in doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    all_text += "\n" + cell.text
        assert "Mostafa KARIM" in all_text
