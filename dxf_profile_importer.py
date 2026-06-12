#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
dxf_profile_importer.py — Import de profil en long et tracé en plan depuis DXF.

Format attendu : 1 fichier DXF contenant 2 calques (LWPOLYLINE ou POLYLINE) :
  - "Tracé en plan" (vue XY horizontale)
  - "Profil en long" (X = distance cumulée, Y = altitude)

Les noms de calques sont auto-détectés (case-insensitive, accents flexibles).
Supporte également les noms propres de Bentley HAMMER :
  - Plan : C300, C400, C500, Polyline canalisation, P0, PMP, R, Bab ballon anti belier, Bloc, etc.
  - Profil : INITIALPRESSUREPROFILE, GRID, VOLUMEGRAPH, AXISLABELS, etc.
"""

import math
import re
import unicodedata

try:
    import ezdxf
    from ezdxf import entities as _ezdxf_entities
    HAS_EZDXF = True
except ImportError:
    HAS_EZDXF = False
    ezdxf = None
    _ezdxf_entities = None


# Types DXF supportés (LWPOLYLINE + POLYLINE)
_POLYLINE_TYPES = "LWPOLYLINE POLYLINE"

# Noms de calques reconnus (normalisés sans accents, lowercase, sans espaces multiples)
# ── Patterns Hammer (noms propres au logiciel Bentley HAMMER) ──
_PLAN_PATTERNS = [
    # Calques standards (français/anglais)
    "trace en plan", "trace", "plan", "plan view", "plan_view",
    "view plan", "view_plan", "alignment", "alignement",
    "trace plan", "xymap",
    # Calques Hammer (noms propres)
    "c300", "c400", "c500", "polyline canalisation",
    "p0", "pmp", "bab ballon anti belier", "reservoirs",
    "bloc", "bloc piont", "bloc pompes", "bloc reservoirs",
]
_PROFILE_PATTERNS = [
    # Calques standards (français/anglais)
    "profil en long", "profil", "profile", "longitudinal profile",
    "longitudinal_profile", "long profile", "long_profile",
    "lp", "elevation", "altitude", "profil longitudinal",
    # Calques Hammer (noms propres)
    "axislabels", "grid", "initialpressureprofile",
    "maximumpressureprofile", "minimumpressureprofile",
    "profiletitle", "referencelinetext",
    "vapourpressureprofile", "volumegraph",
]


def _normalize_layer_name(name: str) -> str:
    """Normalise un nom de calque : minuscules, sans accents, espaces simples."""
    if not name:
        return ""
    n = unicodedata.normalize("NFD", name)
    n = "".join(c for c in n if unicodedata.category(c) != "Mn")
    n = n.lower().strip()
    n = re.sub(r"\s+", " ", n)
    return n


def _match_layer(layer_name: str, patterns: list[str]) -> bool:
    """Vérifie si un nom de calque correspond à un pattern (substring).
    Les patterns et le nom sont normalisés (lowercase, sans accents)."""
    norm = _normalize_layer_name(layer_name)
    for pat in patterns:
        norm_pat = _normalize_layer_name(pat)
        if norm_pat in norm:
            return True
    return False


def _get_vertices(entity) -> list[tuple[float, float]]:
    """Extrait les sommets (X, Y) d'une LWPOLYLINE ou POLYLINE."""
    try:
        if entity.dxftype() == "LWPOLYLINE":
            return [(float(vx), float(vy)) for vx, vy in entity.vertices()]
        elif entity.dxftype() == "POLYLINE":
            return [(float(v.dxf.location.x), float(v.dxf.location.y))
                    for v in entity.vertices]
    except Exception:
        return []


def list_dxf_layers(filepath: str) -> list[str]:
    """
    Liste tous les calques contenant des LWPOLYLINE ou POLYLINE dans un DXF.

    Args:
        filepath: Chemin vers le fichier .dxf

    Returns:
        Liste des noms de calques trouvés (peut être vide si erreur)
    """
    if not HAS_EZDXF:
        return []
    try:
        doc = ezdxf.readfile(filepath)
        msp = doc.modelspace()
        layers = set()
        for entity in msp.query(_POLYLINE_TYPES):
            layer = entity.dxf.layer
            if layer:
                layers.add(layer)
        return sorted(layers)
    except Exception:
        return []


def _extract_polylines(filepath: str, target_layers: list[str]) -> list:
    """
    Extrait les LWPOLYLINE / POLYLINE d'un DXF dont le calque matche.

    Returns:
        Liste d'entités (peut être vide)
    """
    if not HAS_EZDXF:
        return []
    try:
        doc = ezdxf.readfile(filepath)
        msp = doc.modelspace()
        matched = []
        for entity in msp.query(_POLYLINE_TYPES):
            layer = entity.dxf.layer or ""
            if _match_layer(layer, target_layers):
                matched.append(entity)
        return matched
    except Exception:
        return []


def _points_from_entities(entities: list) -> list[tuple[float, float]]:
    """Extrait les points de la polyligne la plus longue d'une liste."""
    if not entities:
        return []
    best = max(entities, key=lambda e: len(_get_vertices(e)))
    return _get_vertices(best)


def load_dxf_plan(filepath: str) -> list[tuple[float, float]]:
    """
    Charge le tracé en plan depuis un DXF (calque "Tracé en plan").
    Retourne une liste de points (X, Y) en mètres.
    Si plusieurs polylignes sont trouvées, retourne la plus longue.

    Args:
        filepath: Chemin vers le fichier .dxf

    Returns:
        Liste de tuples (X_m, Y_m), ou [] si erreur
    """
    entities = _extract_polylines(filepath, _PLAN_PATTERNS)
    return _points_from_entities(entities)


def load_dxf_profile(filepath: str) -> list[tuple[float, float]]:
    """
    Charge le profil en long depuis un DXF (calque "Profil en long").
    Retourne une liste de points (distance, élévation) en mètres.
    X du DXF = distance cumulée, Y du DXF = altitude.

    Si plusieurs polylignes sont trouvées, retourne la plus longue.

    Args:
        filepath: Chemin vers le fichier .dxf

    Returns:
        Liste de tuples (pk_m, z_m), ou [] si erreur
    """
    entities = _extract_polylines(filepath, _PROFILE_PATTERNS)
    return _points_from_entities(entities)


def load_dxf_both(filepath: str) -> dict:
    """
    Charge tracé en plan ET profil en long depuis un seul DXF.
    Détecte automatiquement les 2 calques.
    Supporte LWPOLYLINE et POLYLINE.

    Returns:
        dict avec clés 'plan' (list[(x,y)]), 'profile' (list[(pk,z)]),
              'plan_layer' (str ou None), 'profile_layer' (str ou None)
    """
    result = {
        "plan": [],
        "profile": [],
        "plan_layer": None,
        "profile_layer": None,
    }
    if not HAS_EZDXF:
        return result
    try:
        doc = ezdxf.readfile(filepath)
        msp = doc.modelspace()
        all_polys = list(msp.query(_POLYLINE_TYPES))

        plan_entities = [e for e in all_polys
                         if _match_layer(e.dxf.layer or "", _PLAN_PATTERNS)]
        if plan_entities:
            best = max(plan_entities, key=lambda e: len(_get_vertices(e)))
            result["plan"] = _get_vertices(best)
            result["plan_layer"] = best.dxf.layer

        prof_entities = [e for e in all_polys
                         if _match_layer(e.dxf.layer or "", _PROFILE_PATTERNS)]
        if prof_entities:
            best = max(prof_entities, key=lambda e: len(_get_vertices(e)))
            result["profile"] = _get_vertices(best)
            result["profile_layer"] = best.dxf.layer
    except Exception:
        pass
    return result
