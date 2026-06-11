#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
air_valve_sizing.py — Dimensionnement des ventouses et profil en long.
Contient AirValveSizing pour le calcul des points hauts/bas,
le pré-dimensionnement des ventouses et la localisation des vidanges.
"""

import csv
import hashlib
import math
import os


def _file_hash(filepath: str) -> str:
    """Calcule un hash court (16 caractères) d'un fichier pour le mapping."""
    try:
        with open(filepath, 'rb') as f:
            content = f.read(8192)
        return hashlib.md5(content).hexdigest()[:16]
    except Exception:
        return ""


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
        self._column_mapper = None          # Optionnel — défini via set_column_mapper

    def set_column_mapper(self, mapper):
        """
        Attache un ColumnMapper pour le mapping interactif des colonnes
        inconnues. Si non attaché, fallback sur positions ou erreur.
        """
        self._column_mapper = mapper

    # ------------------------------------------------------------------
    # Helpers de lecture CSV / XLSX
    # ------------------------------------------------------------------

    def _read_csv_rows(self, filepath: str) -> list[list[str]] | None:
        """Lit un fichier CSV avec détection d'encodage et de séparateur."""
        encodings = ['utf-8-sig', 'utf-16', 'utf-16-le', 'utf-8', 'cp1252', 'latin-1']
        for enc in encodings:
            try:
                with open(filepath, 'r', encoding=enc) as f:
                    sample = f.read(4096)
                    f.seek(0)
                    try:
                        dialect = csv.Sniffer().sniff(sample, delimiters=',;\t')
                    except csv.Error:
                        dialect = csv.excel
                    reader = csv.reader(f, dialect)
                    rows = list(reader)
                return rows
            except (UnicodeDecodeError, UnicodeError):
                continue
        return None

    def _read_xlsx_rows(self, filepath: str) -> list[list[str]] | None:
        """Lit un fichier XLSX/XLS et retourne les lignes (comme csv.reader)."""
        try:
            import pandas as pd
            df = pd.read_excel(filepath, engine='openpyxl')
            rows = [list(df.columns)]
            for _, row in df.iterrows():
                rows.append([str(v) if pd.notna(v) else '' for v in row])
            return rows
        except Exception:
            return None

    # ------------------------------------------------------------------
    # Import profil en long
    # ------------------------------------------------------------------

    def load_profile_csv(self, filepath: str) -> bool:
        """
        Charge un profil en long depuis un CSV ou XLSX (3 colonnes : pk, z, [pente]).
        Retourne True si OK, False sinon.

        Modes supportés :
        - Avec en-tête (pk, z) ou variantes : utilise les noms de colonnes
        - Sans en-tête (juste des nombres) : utilise les positions [0]=PK, [1]=Z
        - Mapping interactif : si en-tête présent mais colonnes non reconnues,
          appelle column_mapper (si fourni) pour demander à l'utilisateur
        """
        try:
            ext = os.path.splitext(filepath)[1].lower()
            if ext in ('.xlsx', '.xls'):
                rows = self._read_xlsx_rows(filepath)
            else:
                rows = self._read_csv_rows(filepath)

            if rows is None:
                return False

            # Détecter si la première ligne est un en-tête
            start = 0
            has_header = False
            if rows and rows[0]:
                try:
                    float(str(rows[0][0]).replace(',', '.').replace('\xa0', ''))
                except ValueError:
                    has_header = True
                    start = 1

            # Si en-tête présent, essayer de mapper les colonnes par nom
            pk_idx, z_idx, slope_idx = 0, 1, 2
            if has_header and rows[0]:
                header = [str(c).strip() for c in rows[0]]
                pk_idx, z_idx, slope_idx = self._resolve_profile_columns(
                    header, filepath
                )
                # Si -1 = non trouvé et ignoré, fallback sur positions
                if pk_idx < 0 or z_idx < 0:
                    pk_idx, z_idx, slope_idx = 0, 1, 2

            self.profile = []
            for row in rows[start:]:
                if len(row) < 2:
                    continue
                pk = self._parse_val(row[pk_idx]) if pk_idx < len(row) else None
                z = self._parse_val(row[z_idx]) if z_idx < len(row) else None
                if pk is None or z is None:
                    continue
                pente = self._parse_val(row[slope_idx]) if 0 <= slope_idx < len(row) else None
                self.profile.append({"pk_m": pk, "z_m": z, "pente_pct": pente})

            if len(self.profile) < 2:
                return False

            # Calculer les pentes si non fournies
            self._compute_slopes()
            self._find_high_low_points()
            return True

        except Exception:
            return False

    def _resolve_profile_columns(
        self,
        header: list[str],
        filepath: str,
    ) -> tuple[int, int, int]:
        """
        Résout les indices de colonnes PK, Z, Pente à partir des noms.
        Utilise le ColumnMapper si fourni en attribut (_column_mapper).

        Returns:
            (pk_idx, z_idx, slope_idx) ou (-1, -1, -1) si ignoré
        """
        # Synonymes connus
        PK_SYNONYMS = ["pk", "distance", "abscisse", "absc", "point kilométrique", "linear", "station"]
        Z_SYNONYMS = ["z", "cote", "côte", "altitude", "elevation", "élévation", "tn", "terrain", "level"]
        SLOPE_SYNONYMS = ["pente", "slope", "gradient"]

        def find_idx(synonyms: list[str]) -> int:
            for i, col in enumerate(header):
                col_lower = col.lower()
                if any(syn in col_lower for syn in synonyms):
                    return i
            return -1

        pk_idx = find_idx(PK_SYNONYMS)
        z_idx = find_idx(Z_SYNONYMS)
        slope_idx = find_idx(SLOPE_SYNONYMS)

        # Si PK et Z sont trouvés, on a tout ce qu'il faut
        if pk_idx >= 0 and z_idx >= 0:
            return pk_idx, z_idx, slope_idx

        # Sinon, demander via ColumnMapper (si attaché)
        mapper = getattr(self, "_column_mapper", None)
        if mapper is None:
            # Pas de mapper = fallback sur positions [0]=PK, [1]=Z
            return 0, 1, 2

        # Demander pour PK
        if pk_idx < 0:
            chosen = mapper.request_mapping(
                unknown_col="PK (m)",
                available_cols=header,
                file_type="profile_csv",
                file_hash=_file_hash(filepath),
            )
            if chosen is None:
                return -1, -1, -1  # User a annulé
            pk_idx = header.index(chosen)

        # Demander pour Z
        if z_idx < 0:
            chosen = mapper.request_mapping(
                unknown_col="Z (m)",
                available_cols=header,
                file_type="profile_csv",
                file_hash=_file_hash(filepath),
            )
            if chosen is None:
                return -1, -1, -1
            z_idx = header.index(chosen)

        return pk_idx, z_idx, slope_idx

    def load_profile_manual(self, points: list[tuple[float, float]]):
        """
        Charge un profil en long depuis une liste de tuples (pk_m, z_m).
        """
        self.profile = [{"pk_m": pk, "z_m": z, "pente_pct": None}
                        for pk, z in sorted(points, key=lambda p: p[0])]
        self._compute_slopes()
        self._find_high_low_points()

    def load_profile_bentley_csv(self, filepath: str) -> bool:
        """
        Charge un profil en long depuis un CSV ou XLSX exporté de Bentley HAMMER
        (FlexTable: Junction Table) avec colonnes Label, X, Y, Elevation.

        Format attendu (CSV) :
          Ligne 1 : "FlexTable: Junction Table;;;"
          Ligne 2 : Label;X (m);Y (m);Elevation (m)
          Lignes suivantes : données

        Format XLSX :
          Feuille unique avec colonnes Label, X (m), Y (m), Elevation (m)

        La distance cumulée (PK) est calculée par cumul de distances
        successives entre points (X,Y), en respectant l'ordre du fichier
        (amont → aval, tel qu'organisé par l'utilisateur).

        Args:
            filepath: Chemin vers le fichier CSV ou XLSX

        Returns:
            True si OK, False sinon
        """
        try:
            ext = os.path.splitext(filepath)[1].lower()
            if ext in ('.xlsx', '.xls'):
                rows = self._read_xlsx_rows(filepath)
            else:
                rows = self._read_csv_rows(filepath)

            if rows is None:
                return False

            if not rows:
                return False

            # Sauter la 1ère ligne (en-tête FlexTable: Junction Table;;;)
            start = 0
            if rows[0] and any("flextable" in str(c).lower() for c in rows[0]):
                start = 1

            # Sauter la ligne d'en-tête des colonnes (Label, X, Y, Elevation)
            # On la saute si elle contient des mots-clés connus
            if start < len(rows):
                first_data_row = rows[start]
                if first_data_row and any(
                    kw in str(first_data_row[0]).lower()
                    for kw in ("label", "id", "junction", "node")
                ):
                    start += 1

            # Parser les points (X, Y, Z) — sans tri, ordre du fichier
            raw_points = []
            for row in rows[start:]:
                if len(row) < 4:
                    continue
                x = self._parse_val(row[1])
                y = self._parse_val(row[2])
                z = self._parse_val(row[3])
                if x is None or y is None or z is None:
                    continue
                raw_points.append((x, y, z))

            if len(raw_points) < 2:
                return False

            # Calcul de la distance cumulée par cumul de distances successives
            self.profile = []
            cum_dist = 0.0
            prev_x, prev_y = raw_points[0][0], raw_points[0][1]
            self.profile.append({
                "pk_m": 0.0,
                "z_m": raw_points[0][2],
                "pente_pct": None,
            })
            for i in range(1, len(raw_points)):
                x, y, z = raw_points[i]
                seg = math.sqrt((x - prev_x) ** 2 + (y - prev_y) ** 2)
                cum_dist += seg
                self.profile.append({
                    "pk_m": round(cum_dist, 3),
                    "z_m": z,
                    "pente_pct": None,
                })
                prev_x, prev_y = x, y

            self._compute_slopes()
            self._find_high_low_points()
            return True

        except Exception:
            return False

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
    # Export DXF
    # ------------------------------------------------------------------

    def export_dxf(self, filepath: str) -> tuple[bool, str]:
        """
        Exporte le profil en long + ventouses + vidanges au format DXF.
        Calques créés :
          - "Profil en long" : polyligne du profil (PK, Z)
          - "Ventouses"      : cercles aux points hauts + texte du type
          - "Vidanges"       : cercles aux points bas + texte du type

        Args:
            filepath: Chemin de sortie (.dxf)

        Returns:
            (True, "") si réussi, (False, raison) sinon
        """
        if not self.profile or len(self.profile) < 2:
            return False, "Profil insuffisant (minimum 2 points requis)."
        try:
            import ezdxf
        except ImportError:
            return False, "La librairie ezdxf n'est pas installée."
        try:
            doc = ezdxf.new("R2010")
            msp = doc.modelspace()

            # -- Profil en long (LWPOLYLINE) --
            points = [(p["pk_m"], p["z_m"], 0.0) for p in self.profile]
            msp.add_lwpolyline(points, dxfattribs={"layer": "Profil en long", "color": 5})

            # -- Ventouses (cercles + texte) --
            for v in self.ventouses:
                x, y = v["pk_m"], v["z_m"]
                msp.add_circle((x, y, 0.0), radius=1.5,
                               dxfattribs={"layer": "Ventouses", "color": 3})
                msp.add_text(v["type"], dxfattribs={
                    "layer": "Ventouses", "color": 3, "height": 2.0
                }).set_pos((x, y + 2.5), align="CENTER")

            # -- Vidanges (cercles + texte) --
            for d in self.vidanges:
                x, y = d["pk_m"], d["z_m"]
                msp.add_circle((x, y, 0.0), radius=1.5,
                               dxfattribs={"layer": "Vidanges", "color": 1})
                msp.add_text(d["type"], dxfattribs={
                    "layer": "Vidanges", "color": 1, "height": 2.0
                }).set_pos((x, y + 2.5), align="CENTER")

            doc.saveas(filepath)
            return True, ""
        except Exception as exc:
            return False, f"Erreur lors de la création du DXF : {exc}"

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
