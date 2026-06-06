#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
column_mapper.py — Auto-apprentissage et demande de mapping de colonnes.

Permet à l'utilisateur de mapper interactivement les colonnes
non reconnues d'un fichier CSV/XLSX vers les noms attendus.

Fonctionnalités :
- Auto-apprentissage : les mappings sont mémorisés
- Persistance : stocké dans .hpi + cache global
- Mode batch : applique tous les mappings connus d'un coup
- Mode interactif : demande à l'utilisateur pour chaque colonne

Usage :
    from column_mapper import get_mapper, set_ui_callback

    mapper = get_mapper()
    mapper.set_ui_callback(my_callback_function)  # optionnel
    mapped_name = mapper.request_mapping(
        unknown_col="PK (m)",
        available_cols=df.columns.tolist(),
        file_type="profile_csv",
        file_hash="abc123"
    )
"""

import hashlib
import json


# Types de fichiers reconnus
FILE_TYPE_PROFILE_CSV = "profile_csv"
FILE_TYPE_PROFILE_BENTLEY = "profile_bentley"
FILE_TYPE_PIPES = "pipes"
FILE_TYPE_NODES = "nodes"
FILE_TYPE_PUMPS = "pumps"
FILE_TYPE_RESERVOIRS = "reservoirs"
FILE_TYPE_HPT = "hpt"
FILE_TYPE_AIR_VALVES = "air_valves"


class ColumnMapper:
    """
    Gestionnaire de mapping de colonnes avec auto-apprentissage.

    Mappings stockés en mémoire + sérialisables pour .hpi.
    """

    def __init__(self):
        # {(file_type, file_hash, original_col): mapped_col}
        self.mappings: dict[tuple[str, str, str], str] = {}
        # Compteur d'appels pour statistiques
        self._stats = {"asked": 0, "auto_applied": 0, "skipped": 0}

    # ------------------------------------------------------------------
    # API publique
    # ------------------------------------------------------------------

    def set_ui_callback(self, callback):
        """
        Définit le callback UI pour demander interactivement un mapping.
        callback(unknown_col, available_cols, file_type, file_hash) -> str | None
        Retourne le nom de colonne mappé, ou None pour ignorer.
        """
        self._ui_callback = callback

    def request_mapping(
        self,
        unknown_col: str,
        available_cols: list[str],
        file_type: str,
        file_hash: str = "",
    ) -> str | None:
        """
        Demande un mapping pour une colonne non reconnue.
        Cherche d'abord en cache (auto-apprentissage).
        Si pas trouvé et qu'un callback UI est défini, le demande.

        Sémantique :
        - unknown_col = nom STANDARD attendu (ex: "PK (m)")
        - Le user choisit une colonne PRÉSENTE dans available_cols
        - Le mapping mémorisé : (file_type, file_hash, chosen_col.lower()) → unknown_col
        - Le caller peut alors renommer la colonne du fichier via auto_apply()
        """
        key = (file_type, file_hash, unknown_col.lower())

        # 1) Auto-application depuis cache
        #    On cherche si une colonne du fichier est mappée vers unknown_col
        for (ft, fh, orig_lower), mapped_target in self.mappings.items():
            if ft == file_type and fh == file_hash and mapped_target == unknown_col:
                # orig_lower = nom dans le fichier
                for col in available_cols:
                    if col.lower() == orig_lower:
                        self._stats["auto_applied"] += 1
                        return col  # Retourne le nom dans le fichier

        # 2) Demander à l'utilisateur via callback UI
        if not hasattr(self, "_ui_callback") or self._ui_callback is None:
            return None

        self._stats["asked"] += 1
        chosen = self._ui_callback(
            unknown_col=unknown_col,
            available_cols=available_cols,
            file_type=file_type,
            file_hash=file_hash,
        )

        if chosen is None or chosen == "" or chosen == "__SKIP__" or chosen == "__CANCEL__":
            self._stats["skipped"] += 1
            return None

        # Mémoriser : (file_type, file_hash, chosen.lower()) → unknown_col
        cache_key = (file_type, file_hash, chosen.lower())
        self.mappings[cache_key] = unknown_col
        return chosen

    def learn_mapping(
        self,
        file_type: str,
        original_col: str,
        mapped_col: str,
        file_hash: str = "",
    ):
        """Mémorise explicitement un mapping (sans passer par l'UI)."""
        key = (file_type, file_hash.lower(), original_col.lower())
        self.mappings[key] = mapped_col

    def auto_apply(
        self,
        available_cols: list[str],
        file_type: str,
        file_hash: str = "",
    ) -> dict[str, str]:
        """
        Applique automatiquement tous les mappings connus pour ce fichier.
        Retourne {original_col_in_file: expected_col_name} pour les colonnes
        que l'utilisateur a déjà mappées.

        Sémantique :
        - original = le nom de colonne PRÉSENT dans le fichier (ex: "Distance")
        - mapped   = le nom STANDARD attendu (ex: "PK (m)")
        - Le caller peut alors renommer : df.rename(columns=result)
        """
        result = {}
        for (ft, fh, orig_lower), mapped in self.mappings.items():
            if ft == file_type and fh == file_hash:
                # Chercher la colonne 'orig_lower' (le nom dans le fichier)
                # dans available_cols
                for col in available_cols:
                    if col.lower() == orig_lower:
                        result[col] = mapped
                        break
        return result

    def get_mappings_for_file(self, file_type: str, file_hash: str = "") -> list[dict]:
        """Retourne tous les mappings connus pour un type de fichier / hash."""
        result = []
        for (ft, fh, orig_lower), mapped in self.mappings.items():
            if ft == file_type and fh == file_hash:
                result.append({
                    "original": orig_lower,
                    "mapped": mapped,
                })
        return result

    def clear_mappings(self, file_type: str | None = None, file_hash: str | None = None):
        """Efface les mappings (par type et/ou hash, ou tout)."""
        if file_type is None and file_hash is None:
            self.mappings.clear()
        else:
            keys_to_del = [
                k for k in self.mappings
                if (file_type is None or k[0] == file_type)
                and (file_hash is None or k[1] == file_hash)
            ]
            for k in keys_to_del:
                del self.mappings[k]

    def get_stats(self) -> dict:
        return dict(self._stats)

    # ------------------------------------------------------------------
    # Sérialisation
    # ------------------------------------------------------------------

    def serialize(self) -> dict:
        """Sérialise les mappings pour stockage .hpi."""
        return {
            "version": 1,
            "mappings": [
                {
                    "file_type": k[0],
                    "file_hash": k[1],
                    "original": k[2],
                    "mapped": v,
                }
                for k, v in self.mappings.items()
            ]
        }

    def deserialize(self, data: dict | None):
        """Charge les mappings depuis .hpi."""
        if not data or not isinstance(data, dict):
            return
        if data.get("version") != 1:
            return
        for entry in data.get("mappings", []):
            key = (
                entry.get("file_type", ""),
                entry.get("file_hash", ""),
                entry.get("original", "").lower(),
            )
            self.mappings[key] = entry.get("mapped", "")


# ------------------------------------------------------------------
# Singleton global
# ------------------------------------------------------------------

_global_mapper: ColumnMapper | None = None


def get_mapper() -> ColumnMapper:
    """Retourne l'instance globale du ColumnMapper."""
    global _global_mapper
    if _global_mapper is None:
        _global_mapper = ColumnMapper()
    return _global_mapper


def reset_mapper():
    """Reset le mapper global (pour les tests)."""
    global _global_mapper
    _global_mapper = ColumnMapper()


def file_hash_from_path(filepath: str) -> str:
    """Calcule un hash court du fichier pour identifier le mapping."""
    try:
        with open(filepath, 'rb') as f:
            content = f.read(8192)  # Premiers 8 Ko suffisent
        return hashlib.md5(content).hexdigest()[:16]
    except Exception:
        return ""


def file_hash_from_columns(columns: list[str]) -> str:
    """Hash basé sur les noms de colonnes (pour les DataFrames)."""
    joined = "|".join(columns)
    return hashlib.md5(joined.encode('utf-8')).hexdigest()[:16]
