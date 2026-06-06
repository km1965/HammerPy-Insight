#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Tests unitaires pour column_mapper.py et column_mapper_dialog.py.
"""

import os
import sys
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from column_mapper import (
    ColumnMapper,
    get_mapper,
    reset_mapper,
    file_hash_from_path,
    file_hash_from_columns,
    FILE_TYPE_PROFILE_CSV,
    FILE_TYPE_PROFILE_BENTLEY,
    FILE_TYPE_PIPES,
)
from column_mapper_dialog import SENTINEL_SKIP, SENTINEL_CANCEL


# =====================================================================
# Tests ColumnMapper (sans UI)
# =====================================================================

class TestColumnMapper:
    """Tests du mapper de colonnes."""

    def setup_method(self):
        """Reset le mapper avant chaque test."""
        reset_mapper()

    def test_initialization(self):
        m = ColumnMapper()
        assert m.mappings == {}
        stats = m.get_stats()
        assert stats["asked"] == 0
        assert stats["auto_applied"] == 0
        assert stats["skipped"] == 0

    def test_request_without_callback(self):
        """Sans callback UI, retourne None."""
        m = ColumnMapper()
        result = m.request_mapping("PK (m)", ["Distance", "Z"], FILE_TYPE_PROFILE_CSV, "abc")
        assert result is None

    def test_request_with_callback(self):
        """Avec callback UI, demande et mémorise."""
        m = ColumnMapper()
        m.set_ui_callback(lambda **kw: "Distance")
        result = m.request_mapping("PK (m)", ["Distance", "Z"], FILE_TYPE_PROFILE_CSV, "abc")
        assert result == "Distance"
        # Vérifier que c'est mémorisé
        assert ("profile_csv", "abc", "distance") in m.mappings

    def test_request_skip(self):
        """Si l'utilisateur choisit Skip, retourne None et incrémente skipped."""
        m = ColumnMapper()
        m.set_ui_callback(lambda **kw: SENTINEL_SKIP)
        result = m.request_mapping("PK (m)", ["Distance", "Z"], FILE_TYPE_PROFILE_CSV, "abc")
        assert result is None
        assert m.get_stats()["skipped"] == 1

    def test_request_cancel(self):
        """Si l'utilisateur annule, retourne None."""
        m = ColumnMapper()
        m.set_ui_callback(lambda **kw: SENTINEL_CANCEL)
        result = m.request_mapping("PK (m)", ["Distance", "Z"], FILE_TYPE_PROFILE_CSV, "abc")
        assert result is None

    def test_auto_learning(self):
        """2ème appel = auto-apprentissage depuis cache."""
        m = ColumnMapper()
        m.set_ui_callback(lambda **kw: "Distance")
        # 1er appel : demande
        r1 = m.request_mapping("PK (m)", ["Distance", "Z"], FILE_TYPE_PROFILE_CSV, "abc")
        assert r1 == "Distance"
        assert m.get_stats()["asked"] == 1
        # 2ème appel : auto-applied depuis cache
        r2 = m.request_mapping("PK (m)", ["Distance", "Z"], FILE_TYPE_PROFILE_CSV, "abc")
        assert r2 == "Distance"
        assert m.get_stats()["auto_applied"] == 1
        assert m.get_stats()["asked"] == 1  # Pas redemandé

    def test_auto_learning_different_files(self):
        """Mappings différents pour fichiers différents (hash différent)."""
        m = ColumnMapper()
        m.set_ui_callback(lambda **kw: "Distance" if kw.get("file_hash") == "abc" else "Z")
        m.request_mapping("PK (m)", ["Distance", "Z"], FILE_TYPE_PROFILE_CSV, "abc")
        m.request_mapping("PK (m)", ["Distance", "Z"], FILE_TYPE_PROFILE_CSV, "xyz")
        # Vérifier que les 2 mappings sont stockés
        assert ("profile_csv", "abc", "distance") in m.mappings
        assert ("profile_csv", "xyz", "z") in m.mappings

    def test_learn_mapping_explicit(self):
        """Apprentissage explicite via learn_mapping()."""
        m = ColumnMapper()
        m.learn_mapping(FILE_TYPE_PROFILE_CSV, "Distance", "PK (m)", "abc")
        # Auto-application
        result = m.request_mapping("PK (m)", ["Distance", "Z"], FILE_TYPE_PROFILE_CSV, "abc")
        assert result == "Distance"
        assert m.get_stats()["auto_applied"] == 1

    def test_auto_apply(self):
        """Renommage auto de plusieurs colonnes d'un coup."""
        m = ColumnMapper()
        m.learn_mapping(FILE_TYPE_PROFILE_CSV, "Distance", "PK (m)", "abc")
        m.learn_mapping(FILE_TYPE_PROFILE_CSV, "Z", "Elevation", "abc")
        result = m.auto_apply(["Distance", "Z"], FILE_TYPE_PROFILE_CSV, "abc")
        assert result == {"Distance": "PK (m)", "Z": "Elevation"}

    def test_auto_apply_empty(self):
        """Aucun mapping connu = résultat vide."""
        m = ColumnMapper()
        result = m.auto_apply(["Distance", "Z"], FILE_TYPE_PROFILE_CSV, "abc")
        assert result == {}

    def test_auto_apply_different_file(self):
        """Mapping ne s'applique pas à un autre fichier (hash différent)."""
        m = ColumnMapper()
        m.learn_mapping(FILE_TYPE_PROFILE_CSV, "Distance", "PK (m)", "abc")
        result = m.auto_apply(["Distance", "Z"], FILE_TYPE_PROFILE_CSV, "xyz")
        assert result == {}

    def test_auto_apply_missing_column(self):
        """Si la colonne n'est plus dans le fichier, on l'ignore."""
        m = ColumnMapper()
        m.learn_mapping(FILE_TYPE_PROFILE_CSV, "Distance", "PK (m)", "abc")
        m.learn_mapping(FILE_TYPE_PROFILE_CSV, "OldName", "Elevation", "abc")
        result = m.auto_apply(["Distance", "Z"], FILE_TYPE_PROFILE_CSV, "abc")
        assert result == {"Distance": "PK (m)"}

    def test_clear_mappings_all(self):
        """Effacer tous les mappings."""
        m = ColumnMapper()
        m.learn_mapping(FILE_TYPE_PROFILE_CSV, "Distance", "PK (m)", "abc")
        m.learn_mapping(FILE_TYPE_PIPES, "Mat", "Material", "xyz")
        m.clear_mappings()
        assert m.mappings == {}

    def test_clear_mappings_by_type(self):
        """Effacer mappings par type de fichier."""
        m = ColumnMapper()
        m.learn_mapping(FILE_TYPE_PROFILE_CSV, "Distance", "PK (m)", "abc")
        m.learn_mapping(FILE_TYPE_PIPES, "Mat", "Material", "xyz")
        m.clear_mappings(file_type=FILE_TYPE_PROFILE_CSV)
        assert ("profile_csv", "abc", "distance") not in m.mappings
        assert ("pipes", "xyz", "mat") in m.mappings

    def test_clear_mappings_by_file(self):
        """Effacer mappings par fichier (hash)."""
        m = ColumnMapper()
        m.learn_mapping(FILE_TYPE_PROFILE_CSV, "Distance", "PK (m)", "abc")
        m.learn_mapping(FILE_TYPE_PROFILE_CSV, "Z", "Elevation", "xyz")
        m.clear_mappings(file_hash="abc")
        assert ("profile_csv", "abc", "distance") not in m.mappings
        assert ("profile_csv", "xyz", "z") in m.mappings

    def test_serialize_deserialize(self):
        """Sérialisation / désérialisation."""
        m = ColumnMapper()
        m.learn_mapping(FILE_TYPE_PROFILE_CSV, "Distance", "PK (m)", "abc")
        m.learn_mapping(FILE_TYPE_PIPES, "Mat", "Material", "xyz")
        data = m.serialize()
        assert data["version"] == 1
        assert len(data["mappings"]) == 2

        m2 = ColumnMapper()
        m2.deserialize(data)
        assert m2.mappings == m.mappings

    def test_serialize_deserialize_empty(self):
        """Sérialisation d'un mapper vide."""
        m = ColumnMapper()
        data = m.serialize()
        m2 = ColumnMapper()
        m2.deserialize(data)
        assert m2.mappings == {}

    def test_deserialize_invalid_version(self):
        """Désérialisation avec mauvaise version = ignorée."""
        m = ColumnMapper()
        m.deserialize({"version": 999, "mappings": []})
        assert m.mappings == {}

    def test_deserialize_none(self):
        """Désérialisation avec None = no-op."""
        m = ColumnMapper()
        m.deserialize(None)
        assert m.mappings == {}

    def test_get_mappings_for_file(self):
        """Récupérer les mappings d'un fichier spécifique."""
        m = ColumnMapper()
        m.learn_mapping(FILE_TYPE_PROFILE_CSV, "Distance", "PK (m)", "abc")
        m.learn_mapping(FILE_TYPE_PROFILE_CSV, "Z", "Elevation", "abc")
        m.learn_mapping(FILE_TYPE_PIPES, "Mat", "Material", "abc")
        result = m.get_mappings_for_file(FILE_TYPE_PROFILE_CSV, "abc")
        assert len(result) == 2
        assert {"original": "distance", "mapped": "PK (m)"} in result
        assert {"original": "z", "mapped": "Elevation"} in result


# =====================================================================
# Tests du singleton global
# =====================================================================

class TestGlobalMapper:
    """Tests du singleton global."""

    def test_singleton(self):
        m1 = get_mapper()
        m2 = get_mapper()
        assert m1 is m2

    def test_reset(self):
        m1 = get_mapper()
        m1.learn_mapping(FILE_TYPE_PROFILE_CSV, "Dist", "PK", "abc")
        reset_mapper()
        m2 = get_mapper()
        assert m1 is not m2
        assert m2.mappings == {}


# =====================================================================
# Tests des utilitaires de hash
# =====================================================================

class TestFileHash:
    """Tests des fonctions de hash."""

    def test_hash_from_path(self):
        import tempfile
        content = b"col1,col2,col3\n1,2,3\n4,5,6\n"
        f = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
        f.write(content)
        f.close()
        try:
            h = file_hash_from_path(f.name)
            assert len(h) == 16
            # Même contenu = même hash
            f2 = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
            f2.write(content)
            f2.close()
            try:
                h2 = file_hash_from_path(f2.name)
                assert h == h2
            finally:
                os.unlink(f2.name)
        finally:
            os.unlink(f.name)

    def test_hash_from_columns(self):
        cols = ["Distance", "Z", "Label"]
        h = file_hash_from_columns(cols)
        assert len(h) == 16
        # Même colonnes = même hash
        h2 = file_hash_from_columns(cols)
        assert h == h2
        # Colonnes différentes = hash différent
        h3 = file_hash_from_columns(["X", "Y", "Z"])
        assert h != h3
