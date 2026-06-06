#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
dxf_profile_importer.py — Import de profil en long et tracé en plan depuis DXF.

Format attendu : 1 fichier DXF contenant 2 calques (LWPOLYLINE) :
  - "Tracé en plan" (vue XY horizontale)
  - "Profil en long" (X = distance cumulée, Y = altitude)

Les noms de calques sont auto-détectés (case-insensitive, accents flexibles).
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


# Noms de calques reconnus (normalisés sans accents, lowercase, sans espaces multiples)
_PLAN_PATTERNS = [
    "tracé en plan", "tracé en plan", "trace en plan",
    "tracé", "trace", "plan", "plan view", "plan_view",
    "view plan", "view_plan", "alignment", "alignement",
    "tracé plan", "trace plan", "xymap",
]
_PROFILE_PATTERNS = [
    "profil en long", "profil", "profile", "longitudinal profile",
    "longitudinal_profile", "long profile", "long_profile",
    "lp", "elevation", "altitude", "profil longitudinal",
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


def list_dxf_layers(filepath: str) -> list[str]:
    """
    Liste tous les calques contenant des LWPOLYLINE dans un fichier DXF.

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
        for entity in msp.query("LWPOLYLINE"):
            layer = entity.dxf.layer
            if layer:
                layers.add(layer)
        return sorted(layers)
    except Exception:
        return []


def _extract_lwpolylines(filepath: str, target_layers: list[str]) -> list:
    """
    Extrait les LWPOLYLINE d'un DXF dont le calque matche les patterns cibles.

    Returns:
        Liste d'objets LWPOLYLINE (peut être vide)
    """
    if not HAS_EZDXF:
        return []
    try:
        doc = ezdxf.readfile(filepath)
        msp = doc.modelspace()
        matched = []
        for entity in msp.query("LWPOLYLINE"):
            layer = entity.dxf.layer or ""
            if _match_layer(layer, target_layers):
                matched.append(entity)
        return matched
    except Exception:
        return []


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
    entities = _extract_lwpolylines(filepath, _PLAN_PATTERNS)
    if not entities:
        return []
    # Prendre la polyligne avec le plus de sommets
    best = max(entities, key=lambda e: len(list(e.vertices())))
    points = []
    for vx, vy in best.vertices():
        points.append((float(vx), float(vy)))
    return points


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
    entities = _extract_lwpolylines(filepath, _PROFILE_PATTERNS)
    if not entities:
        return []
    best = max(entities, key=lambda e: len(list(e.vertices())))
    points = []
    for vx, vy in best.vertices():
        points.append((float(vx), float(vy)))
    return points


def load_dxf_both(filepath: str) -> dict:
    """
    Charge tracé en plan ET profil en long depuis un seul DXF.
    Détecte automatiquement les 2 calques.

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
        # Plan
        plan_entities = [e for e in msp.query("LWPOLYLINE")
                         if _match_layer(e.dxf.layer or "", _PLAN_PATTERNS)]
        if plan_entities:
            best = max(plan_entities, key=lambda e: len(list(e.vertices())))
            result["plan"] = [(float(x), float(y)) for x, y in best.vertices()]
            result["plan_layer"] = best.dxf.layer
        # Profil
        prof_entities = [e for e in msp.query("LWPOLYLINE")
                         if _match_layer(e.dxf.layer or "", _PROFILE_PATTERNS)]
        if prof_entities:
            best = max(prof_entities, key=lambda e: len(list(e.vertices())))
            result["profile"] = [(float(x), float(y)) for x, y in best.vertices()]
            result["profile_layer"] = best.dxf.layer
    except Exception:
        pass
    return result
