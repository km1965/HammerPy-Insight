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

import io
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
    """Extrait les sommets (X, Y) d'une LWPOLYLINE ou POLYLINE de façon robuste.
    Gère à la fois les LWPOLYLINE et les POLYLINE 2D/3D complexes.
    """
    try:
        t = entity.dxftype()
        if t == "LWPOLYLINE":
            pts = entity.get_points()
            return [(float(px), float(py)) for px, py, *rest in pts]
        elif t == "POLYLINE":
            return [(float(v.dxf.location.x), float(v.dxf.location.y)) for v in entity.vertices]
    except Exception:
        pass
    return []


def _distance(p1: tuple[float, float], p2: tuple[float, float]) -> float:
    """Calcule la distance euclidienne entre deux points 2D."""
    return math.sqrt((p1[0] - p2[0]) ** 2 + (p1[1] - p2[1]) ** 2)


def assemble_segments(segments: list[list[tuple[float, float]]], epsilon: float = 0.1) -> list[tuple[float, float]]:
    """
    Assemble une liste de segments de points non ordonnés en un ou plusieurs chemins continus.
    Chaque segment est une liste/tuple de points (typiquement 2 points d'un tronçon).
    """
    if not segments:
        return []
    # Copier et filtrer les segments valides
    available = [list(seg) for seg in segments if len(seg) >= 2]
    if not available:
        return []

    paths = []
    while available:
        curr_path = available.pop(0)
        extended = True
        while extended:
            extended = False
            p_first = curr_path[0]
            p_last = curr_path[-1]
            
            for idx, seg in enumerate(available):
                # Connexions à la fin du chemin actuel
                if _distance(p_last, seg[0]) < epsilon:
                    curr_path.extend(seg[1:])
                    available.pop(idx)
                    extended = True
                    break
                elif _distance(p_last, seg[-1]) < epsilon:
                    curr_path.extend(reversed(seg[:-1]))
                    available.pop(idx)
                    extended = True
                    break
                # Connexions au début du chemin actuel
                elif _distance(p_first, seg[-1]) < epsilon:
                    curr_path = seg[:-1] + curr_path
                    available.pop(idx)
                    extended = True
                    break
                elif _distance(p_first, seg[0]) < epsilon:
                    curr_path = list(reversed(seg[1:])) + curr_path
                    available.pop(idx)
                    extended = True
                    break
        paths.append(curr_path)

    if not paths:
        return []
    # Prendre le chemin avec le plus grand nombre de points
    return max(paths, key=len)


def _read_dxf_file_robustly(filepath: str):
    """
    Lit un fichier DXF en gérant l'encodage et en convertissant en mémoire
    les virgules décimales mal formées (spécifiques aux exports régionaux FR) en points.
    Retourne un objet Drawing ezdxf.
    """
    encodings = ['utf-8', 'cp1252', 'latin-1', 'utf-16']
    content = None
    for enc in encodings:
        try:
            with open(filepath, 'r', encoding=enc, errors='replace') as f:
                content = f.read()
            break
        except Exception:
            continue
            
    if content is None:
        raise ValueError(f"Impossible de lire le fichier DXF : {filepath}")

    # Remplacer les virgules par des points sur les lignes contenant uniquement un float
    comma_decimal_re = re.compile(r'^\s*(-?\d+),(\d+)\s*$')
    cleaned_lines = []
    for line in content.splitlines():
        match = comma_decimal_re.match(line)
        if match:
            cleaned_lines.append(f"{match.group(1)}.{match.group(2)}")
        else:
            cleaned_lines.append(line)
            
    stream = io.StringIO("\n".join(cleaned_lines))
    return ezdxf.read(stream)


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
        doc = _read_dxf_file_robustly(filepath)
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
        doc = _read_dxf_file_robustly(filepath)
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
    """Extrait les points des entités polylignes en les groupant par calque
    pour éviter de mélanger les tracés, puis assemble les segments du calque le plus riche.
    """
    if not entities:
        return []
    # Grouper par calque
    by_layer = {}
    for e in entities:
        l = e.dxf.layer or ""
        if l not in by_layer:
            by_layer[l] = []
        by_layer[l].append(e)
        
    best_path = []
    for layer, ents in by_layer.items():
        segments = [pts for e in ents if (pts := _get_vertices(e))]
        assembled = assemble_segments(segments, epsilon=0.1)
        if len(assembled) > len(best_path):
            best_path = assembled
    return best_path


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
        doc = _read_dxf_file_robustly(filepath)
        msp = doc.modelspace()
        all_polys = list(msp.query(_POLYLINE_TYPES))

        # Regrouper les polylignes par calque
        polys_by_layer = {}
        for e in all_polys:
            l = e.dxf.layer
            if l:
                if l not in polys_by_layer:
                    polys_by_layer[l] = []
                polys_by_layer[l].append(e)

        # Chercher le meilleur calque plan
        best_plan_pts = []
        best_plan_layer = None
        for layer, entities in polys_by_layer.items():
            if _match_layer(layer, _PLAN_PATTERNS):
                segments = [pts for e in entities if (pts := _get_vertices(e))]
                assembled = assemble_segments(segments, epsilon=0.1)
                if len(assembled) > len(best_plan_pts):
                    best_plan_pts = assembled
                    best_plan_layer = layer
        result["plan"] = best_plan_pts
        result["plan_layer"] = best_plan_layer

        # Chercher le meilleur calque profil
        best_prof_pts = []
        best_prof_layer = None
        for layer, entities in polys_by_layer.items():
            if _match_layer(layer, _PROFILE_PATTERNS):
                segments = [pts for e in entities if (pts := _get_vertices(e))]
                assembled = assemble_segments(segments, epsilon=0.1)
                if len(assembled) > len(best_prof_pts):
                    best_prof_pts = assembled
                    best_prof_layer = layer
        result["profile"] = best_prof_pts
        result["profile_layer"] = best_prof_layer
    except Exception:
        pass
    return result
