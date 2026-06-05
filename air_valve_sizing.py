#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
air_valve_sizing.py — Dimensionnement des ventouses et profil en long.
Contient AirValveSizing pour le calcul des points hauts/bas,
le pré-dimensionnement des ventouses et la localisation des vidanges.
"""

import csv
import os


# =====================================================================
# CONSTANTES MÉTIER
# =====================================================================

# Types de ventouses
VENTOUSE_SIMPLE = "Anti-vide simple"
VENTOUSE_COMBINEE = "Combinée (admission + dégazage)"
VENTOUSE_GRANDE_ORIFICE = "Grande orifice (admission rapide)"

# Règles de dimensionnement DN ventouse
RATIO_DN_SIMPLE = 12          # DN ≥ DN_conduite / 12
RATIO_DN_GRANDE_ORIFICE = 8   # DN ≥ DN_conduite / 8
RATIO_DN_COMBINEE = 10        # DN ≥ DN_conduite / 10

# Règles de dimensionnement DN vidange
RATIO_DN_VIDANGE = 10         # DN ≥ DN_conduite / 10

# Distances minimales
DISTANCE_MIN_VENTOUSES_M = 50     # Distance min entre deux ventouses
DISTANCE_MIN_VIDANGE_M = 50       # Distance min entre vidange et ventouse
DISTANCE_MAX_VIDANGES_M = 500     # Distance max entre deux vidanges

# Seuil de pente horizontale (%)
PENTE_HORIZONTALE_THRESHOLD = 0.5


class AirValveSizing:
    """
    Dimensionnement des ventouses et localisation des vidanges
    sur le profil en long d'une conduite.
    """

    def __init__(self, pipe_dn_mm: float = 250.0):
        """
        Args:
            pipe_dn_mm: Diamètre nominal de la conduite (mm).
        """
        self.pipe_dn_mm = pipe_dn_mm
        self.profile: list[dict] = []       # [{pk_m, z_m, pente_pct}]
        self.high_points: list[dict] = []   # Points hauts (ventouses)
        self.low_points: list[dict] = []    # Points bas (vidanges)
        self.ventouses: list[dict] = []     # Recommandations ventouses
        self.vidanges: list[dict] = []      # Recommandations vidanges

    # ------------------------------------------------------------------
    # Import profil en long
    # ------------------------------------------------------------------

    def load_profile_csv(self, filepath: str) -> bool:
        """
        Charge un profil en long depuis un CSV (3 colonnes : pk, z, [pente]).
        Retourne True si OK, False sinon.
        """
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                sample = f.read(4096)
                f.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=',;\t')
                except csv.Error:
                    dialect = csv.excel
                reader = csv.reader(f, dialect)
                rows = list(reader)

            # Détecter si la première ligne est un en-tête
            start = 0
            if rows and rows[0]:
                try:
                    float(str(rows[0][0]).replace(',', '.').replace('\xa0', ''))
                except ValueError:
                    start = 1

            self.profile = []
            for row in rows[start:]:
                if len(row) < 2:
                    continue
                pk = self._parse_val(row[0])
                z = self._parse_val(row[1])
                if pk is None or z is None:
                    continue
                pente = self._parse_val(row[2]) if len(row) > 2 else None
                self.profile.append({"pk_m": pk, "z_m": z, "pente_pct": pente})

            if len(self.profile) < 2:
                return False

            # Calculer les pentes si non fournies
            self._compute_slopes()
            self._find_high_low_points()
            return True

        except Exception:
            return False

    def load_profile_manual(self, points: list[tuple[float, float]]):
        """
        Charge un profil en long depuis une liste de tuples (pk_m, z_m).
        """
        self.profile = [{"pk_m": pk, "z_m": z, "pente_pct": None}
                        for pk, z in sorted(points, key=lambda p: p[0])]
        self._compute_slopes()
        self._find_high_low_points()

    # ------------------------------------------------------------------
    # Calcul des pentes
    # ------------------------------------------------------------------

    def _compute_slopes(self):
        """Calcule la pente locale (%) entre chaque point consecutive."""
        for i in range(len(self.profile)):
            if i == 0:
                self.profile[i]["pente_pct"] = 0.0
            else:
                dz = self.profile[i]["z_m"] - self.profile[i - 1]["z_m"]
                dpk = self.profile[i]["pk_m"] - self.profile[i - 1]["pk_m"]
                if dpk > 0:
                    self.profile[i]["pente_pct"] = round((dz / dpk) * 100, 2)
                else:
                    self.profile[i]["pente_pct"] = 0.0

    # ------------------------------------------------------------------
    # Détection points hauts / bas
    # ------------------------------------------------------------------

    def _find_high_low_points(self):
        """
        Identifie les points hauts (maxima locaux) et bas (minima locaux)
        du profil en long.
        """
        self.high_points = []
        self.low_points = []

        if len(self.profile) < 3:
            # Pas assez de points pour un extremum local — prendre min/max globaux
            if self.profile:
                max_pt = max(self.profile, key=lambda p: p["z_m"])
                min_pt = min(self.profile, key=lambda p: p["z_m"])
                self.high_points.append(max_pt)
                self.low_points.append(min_pt)
            return

        for i in range(1, len(self.profile) - 1):
            prev_z = self.profile[i - 1]["z_m"]
            curr_z = self.profile[i]["z_m"]
            next_z = self.profile[i + 1]["z_m"]

            if curr_z > prev_z and curr_z > next_z:
                self.high_points.append(self.profile[i])
            elif curr_z < prev_z and curr_z < next_z:
                self.low_points.append(self.profile[i])

        # Inclure les extrémités si elles sont des extrema
        if self.profile:
            first = self.profile[0]
            last = self.profile[-1]
            if len(self.profile) > 1:
                if first["z_m"] > self.profile[1]["z_m"]:
                    self.high_points.insert(0, first)
                elif first["z_m"] < self.profile[1]["z_m"]:
                    self.low_points.insert(0, first)

                if last["z_m"] > self.profile[-2]["z_m"]:
                    self.high_points.append(last)
                elif last["z_m"] < self.profile[-2]["z_m"]:
                    self.low_points.append(last)

    # ------------------------------------------------------------------
    # Pré-dimensionnement ventouses
    # ------------------------------------------------------------------

    def size_ventouses(self):
        """
        Pré-dimensionne les ventouses aux points hauts du profil.
        Règles :
          - Ventouse simple (anti-vide) : DN ≥ DN/12
          - Grande orifice (admission rapide) : DN ≥ DN/8
          - Ventouse combinée : DN ≥ DN/10
        """
        self.ventouses = []
        dn_pipe = self.pipe_dn_mm

        for pt in self.high_points:
            # Déterminer le type selon la pente locale
            pente = abs(pt.get("pente_pct") or 0)

            if pente < PENTE_HORIZONTALE_THRESHOLD:
                # Tronçon horizontal → ventouse combinée (admission + dégazage)
                vent_type = VENTOUSE_COMBINEE
                dn = self._calc_dn(dn_pipe, RATIO_DN_COMBINEE)
            elif pente > 3.0:
                # Forte pente → ventouse simple (anti-vide)
                vent_type = VENTOUSE_SIMPLE
                dn = self._calc_dn(dn_pipe, RATIO_DN_SIMPLE)
            else:
                # Pente modérée → ventouse à grande orifice
                vent_type = VENTOUSE_GRANDE_ORIFICE
                dn = self._calc_dn(dn_pipe, RATIO_DN_GRANDE_ORIFICE)

            self.ventouses.append({
                "pk_m": pt["pk_m"],
                "z_m": pt["z_m"],
                "pente_pct": pt.get("pente_pct", 0),
                "type": vent_type,
                "dn_mm": dn,
            })

    def _calc_dn(self, dn_pipe: float, ratio: int) -> int:
        """Calcule le DN recommandé en fonction du DN conduite et du ratio."""
        dn = max(25, int(dn_pipe / ratio))
        # Arrondir au DN standard supérieur le plus proche
        standard_dns = [25, 32, 40, 50, 65, 80, 100, 125, 150, 200, 250, 300]
        for std in standard_dns:
            if std >= dn:
                return std
        return dn

    # ------------------------------------------------------------------
    # Localisation des vidanges
    # ------------------------------------------------------------------

    def size_drains(self):
        """
        Localise les vidanges aux points bas entre deux ventouses.
        Algorithme :
          1. Pour chaque segment entre 2 ventouses consécutives
          2. Trouver le point bas (minimum local d'altitude)
          3. Si distance > 50m aux ventouses → placer une vidange
          4. Sinon → recommander vidange combinée avec ventouse proche
        """
        self.vidanges = []

        if not self.ventouses or not self.low_points:
            return

        # Trier les ventouses par PK
        sorted_vents = sorted(self.ventouses, key=lambda v: v["pk_m"])

        for i in range(len(sorted_vents) - 1):
            v_left = sorted_vents[i]
            v_right = sorted_vents[i + 1]

            # Trouver les points bas dans ce segment
            segment_lows = [
                lp for lp in self.low_points
                if v_left["pk_m"] < lp["pk_m"] < v_right["pk_m"]
            ]

            if not segment_lows:
                continue

            # Prendre le point bas le plus bas dans le segment
            lowest = min(segment_lows, key=lambda p: p["z_m"])

            dist_left = lowest["pk_m"] - v_left["pk_m"]
            dist_right = v_right["pk_m"] - lowest["pk_m"]

            if dist_left >= DISTANCE_MIN_VIDANGE_M and dist_right >= DISTANCE_MIN_VIDANGE_M:
                vid_type = "Vidange à bride"
                dn = self._calc_dn(self.pipe_dn_mm, RATIO_DN_VIDANGE)
                self.vidanges.append({
                    "pk_m": lowest["pk_m"],
                    "z_m": lowest["z_m"],
                    "type": vid_type,
                    "dn_mm": dn,
                    "left_ventouse_pk": v_left["pk_m"],
                    "right_ventouse_pk": v_right["pk_m"],
                    "distance_to_left_m": round(dist_left, 1),
                    "distance_to_right_m": round(dist_right, 1),
                })
            else:
                # Trop près d'une ventouse → vidange combinée
                nearest_v = v_left if dist_left < dist_right else v_right
                vid_type = "Vidange combinée avec ventouse"
                dn = self._calc_dn(self.pipe_dn_mm, RATIO_DN_VIDANGE)
                self.vidanges.append({
                    "pk_m": lowest["pk_m"],
                    "z_m": lowest["z_m"],
                    "type": vid_type,
                    "dn_mm": dn,
                    "left_ventouse_pk": v_left["pk_m"],
                    "right_ventouse_pk": v_right["pk_m"],
                    "distance_to_left_m": round(dist_left, 1),
                    "distance_to_right_m": round(dist_right, 1),
                })

    # ------------------------------------------------------------------
    # Export CSV
    # ------------------------------------------------------------------

    def export_csv(self, filepath: str):
        """Exporte les recommandations ventouses + vidanges en CSV."""
        with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(["Type", "PK (m)", "Côte (m)", "Type设备", "DN (mm)",
                             "Dist. ventouse G (m)", "Dist. ventouse D (m)"])
            for v in self.ventouses:
                writer.writerow([
                    "Ventouse", f"{v['pk_m']:.1f}", f"{v['z_m']:.2f}",
                    v["type"], v["dn_mm"], "", ""
                ])
            for d in self.vidanges:
                writer.writerow([
                    "Vidange", f"{d['pk_m']:.1f}", f"{d['z_m']:.2f}",
                    d["type"], d["dn_mm"],
                    f"{d['distance_to_left_m']:.1f}",
                    f"{d['distance_to_right_m']:.1f}"
                ])

    # ------------------------------------------------------------------
    # Utilitaires
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_val(s: str) -> float | None:
        """Parse une valeur numérique (gère virgule FR, espaces insécables)."""
        if s is None:
            return None
        s = str(s).strip().replace('\xa0', '').replace(' ', '')
        if not s or s.lower() in ('nan', 'none', 'n/a'):
            return None
        if ',' in s:
            s = s.replace(',', '.')
        try:
            return float(s)
        except (ValueError, TypeError):
            return None
