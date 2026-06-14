#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test script to reproduce DXF import issues caused by non‑standard layer names.
The script creates a temporary DXF file with layer names that differ from the
patterns expected by ``dxf_profile_importer`` (e.g. "Plan", "Profil", accents,
mixed case) and then attempts to load them using the existing importer logic.
"""

import os
import sys
import tempfile

# Ensure the project root is on PYTHONPATH
ROOT = os.path.abspath(os.path.dirname(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

# Dependency on ezdxf – safely import only if available
try:
    import ezdxf
    HAS_EZDXF = True
except ImportError:          # pragma: no cover
    HAS_EZDXF = False
    ezdxf = None

# Import the importer module (the same file we are testing)
from dxf_profile_importer import (
    load_dxf_plan,
    load_dxf_profile,
    load_dxf_both,
    list_dxf_layers,
    HAS_EZDXF,
    _normalize_layer_name,
    _match_layer,
)

# ----------------------------------------------------------------------
# Helper: create a DXF with *non‑standard* layer names
# ----------------------------------------------------------------------
def create_dxf_with_custom_layers(filepath: str):
    """
    Generates a DXF containing:
      - Layer "Plan"      – horizontal polyline (simple rectangle)
      - Layer "Profil"    – vertical profile polyline (X = distance, Y = altitude)
    The layer names are deliberately different from the canonical
    “Tracé en plan” / “Profil en long” patterns used by the importer.
    """
    if not HAS_EZDXF:
        raise RuntimeError("ezdxf not installed – aborting DXF creation")

    doc = ezdxf.new(dxfversion="R2010")
    msp = doc.modelspace()

    # ---- Layer "Plan" -------------------------------------------------
    plan_layer = "Plan"  # <-- hull of the issue: differs from expected names
    plan_coords = [
        (0, 0), (200, 0), (200, 50), (0, 50), (0, 0)
    ]  # simple rectangle
    msp.add_lwpolyline(plan_coords, dxfattribs={"layer": plan_layer})

    # ---- Layer "Profil" -----------------------------------------------
    prof_layer = "Profil"  # <-- differs from “Profil en long”, “longitudinal profile”, etc.
    # X = cumulative distance, Y = altitude (e.g. 0, 10; 50, 12; 100, 10; 150, 8; 200, 5)
    prof_coords = [
        (0, 0), (100, 15), (200, 12), (300, 10), (400, 8)
    ]
    msp.add_lwpolyline(prof_coords, dxfattribs={"layer": prof_layer})

    # Save to the requested temporary file
    doc.saveas(filepath)
    doc.close()


# ----------------------------------------------------------------------
# Test scenario
# ----------------------------------------------------------------------
def main():
    # 1️⃣ Create a temporary DXF with custom layer names
    with tempfile.NamedTemporaryFile(suffix=".dxf", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        create_dxf_with_custom_layers(tmp_path)
        print(f"\n✅ DXF temporaire créé : {tmp_path}\n")

        # 2️⃣ List detected layers – useful for debugging
        layers = list_dxf_layers(tmp_path)
        print("🔎 Calques détectés :", layers)

        # 3️⃣ Try the three importer entry‑points
        print("\n📂 Chargement « plan » …")
        plan_pts = load_dxf_plan(tmp_path)
        print(f"   Points récupérés : {len(plan_pts)} … {plan_pts[:3]}")

        print("\n📈 Chargement « profil » …")
        prof_pts = load_dxf_profile(tmp_path)
        print(f"   Points récupérés : {len(prof_pts)} … {prof_pts[:3]}")

        print("\n🔀 Chargement complet (plan + profil) …")
        result = load_dxf_both(tmp_path)
        print(
            f"   plan_layer   : {result.get('plan_layer')!r}\n"
            f"   profile_layer: {result.get('profile_layer')!r}\n"
            f"   points plan  : {len(result.get('plan', []))}\n"
            f"   points profil: {len(result.get('profile', []))}"
        )

        # 4️⃣ Diagnostic prints for the pattern matcher
        print("\n🧩 Pattern‑matching diagnostic :")
        for pat in ["Plan", "Profil"]:
            matched = _match_layer(pat, ["trace en plan", "plan", "profil en long", "profil"])
            print(f"   « {pat} » vs. patterns → {matched}")

        # 5️⃣ Normalisation illustration
        print("\n⚙️  Normalisation des noms de calques :")
        for original in ["Plan", "Profil", "Tracé en plan", "Profil en long"]:
            print(f"   {original!r} → {_normalize_layer_name(original)!r}")

    finally:
        # Clean up the temporary file
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
            print(f"\n🗑️  Fichier temporaire supprimé : {tmp_path}")


if __name__ == "__main__":
    main()