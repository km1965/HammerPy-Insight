#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
system_diagnostics.py — Vérifications croisées Pompe ↔ Réseau ↔ HPT ↔ Ventouses.

Phase 4 de HammerPy Insight.

16 checks en 5 catégories (A-E) :
  A. Pompe ↔ Réseau (3)
  B. Pompe ↔ HPT (2)
  C. Réseau ↔ HPT (3)
  D. HPT ↔ Ventouses / Vidanges (3)
  E. Cohérence globale (5)

Sévérités : OK / WARN / FAIL / NA
"""

# ── Sévérités ──────────────────────────────────────────────────────
STATUS_OK   = "OK"
STATUS_WARN = "WARN"
STATUS_FAIL = "FAIL"
STATUS_NA   = "NA"

# ── Catégories ─────────────────────────────────────────────────────
CAT_A = "A. Pompe ↔ Réseau"
CAT_B = "B. Pompe ↔ HPT"
CAT_C = "C. Réseau ↔ HPT"
CAT_D = "D. HPT ↔ Ventouses / Vidanges"
CAT_E = "E. Cohérence globale"

# ── Icônes (cohérent avec report_generator.py) ─────────────────────
ICON = {
    STATUS_OK:   "✔",
    STATUS_WARN: "⚠",
    STATUS_FAIL: "✘",
    STATUS_NA:   "—",
}

# ── Marges / seuils par défaut ─────────────────────────────────────
DEFAULT_NPSH_MARGIN_M = 1.0          # Marge mini NPSH dispo > requis
DEFAULT_NQ_MIN_SI     = 25.0
DEFAULT_NQ_MAX_SI     = 80.0
DEFAULT_SLOPE_MAX_PCT = 50.0         # Pente max au-delà = warning


class SystemDiagnostics:
    """
    Exécute les 16 vérifications croisées et retourne une liste
    de résultats structurés.
    """

    def __init__(
        self,
        pump_parsers: list | None = None,
        air_valve_sizer=None,
        workbook_manager=None,
        transient_status: dict | None = None,
        pn_value_bar: float = 16.0,
        pmin_value_bar: float = 0.0,
        vgas_threshold_l: float = 200.0,
        npsh_margin_m: float = DEFAULT_NPSH_MARGIN_M,
        nq_min_si: float = DEFAULT_NQ_MIN_SI,
        nq_max_si: float = DEFAULT_NQ_MAX_SI,
        slope_max_pct: float = DEFAULT_SLOPE_MAX_PCT,
    ):
        self.pump_parsers = pump_parsers or []
        self.air_valve_sizer = air_valve_sizer
        self.workbook_manager = workbook_manager
        self.transient_status = transient_status or {}
        self.pn_value_bar = float(pn_value_bar)
        self.pmin_value_bar = float(pmin_value_bar)
        self.vgas_threshold_l = float(vgas_threshold_l)
        self.npsh_margin_m = float(npsh_margin_m)
        self.nq_min_si = float(nq_min_si)
        self.nq_max_si = float(nq_max_si)
        self.slope_max_pct = float(slope_max_pct)

    # ────────────────────────────────────────────────────────────
    # Helpers
    # ────────────────────────────────────────────────────────────

    def _check(
        self,
        code: str,
        name: str,
        category: str,
        status: str,
        message: str,
        value=None,
        threshold=None,
    ) -> dict:
        return {
            "code": code,
            "name": name,
            "category": category,
            "status": status,
            "message": message,
            "value": value,
            "threshold": threshold,
        }

    def _first_pump_parsed(self) -> dict | None:
        """Retourne le `parsed` de la première pompe chargée (ou None)."""
        for p in self.pump_parsers:
            if getattr(p, "parsed", None):
                return p.parsed
        return None

    def _any_curve_points(self) -> bool:
        for p in self.pump_parsers:
            if len(getattr(p, "curve_points", [])) >= 2:
                return True
        return False

    def _first_pump_with_curve(self):
        for p in self.pump_parsers:
            if len(getattr(p, "curve_points", [])) >= 2:
                return p
        return None

    def _pump_head_at_nominal(self, pump) -> float | None:
        """H(Q_nominal) interpolé via la courbe H(Q)."""
        try:
            if not getattr(pump, "curve_points", None):
                return None
            q_nominal = pump.parsed.get("flow_lps")
            if q_nominal is None:
                return None
            return pump.interpolate_head(float(q_nominal))
        except Exception:
            return None

    # ────────────────────────────────────────────────────────────
    # A. Pompe ↔ Réseau (3 checks)
    # ────────────────────────────────────────────────────────────

    def _check_A1_operating_point(self) -> dict:
        """Le point de fonctionnement nominal est-il dans la courbe H(Q) ?"""
        pump = self._first_pump_with_curve()
        if not pump:
            return self._check(
                "A1", "Point de fonctionnement dans la courbe H(Q)",
                CAT_A, STATUS_NA,
                "Aucune courbe H(Q) saisie — ajoutez des points H(Q) dans l'onglet Rapport Technique (Pompe)",
            )
        q_nom = pump.parsed.get("flow_lps")
        h_nom = pump.parsed.get("pump_head_m")
        h_curve = self._pump_head_at_nominal(pump)
        if q_nom is None or h_nom is None or h_curve is None:
            return self._check(
                "A1", "Point de fonctionnement dans la courbe H(Q)",
                CAT_A, STATUS_NA,
                "Données Q/H nominales manquantes — vérifiez le rapport pompe chargé",
            )
        ecart = abs(h_curve - h_nom)
        seuil = max(0.05 * h_nom, 1.0)  # 5% ou 1 m
        if ecart <= seuil:
            return self._check(
                "A1", "Point de fonctionnement dans la courbe H(Q)",
                CAT_A, STATUS_OK,
                f"H(Q_nom)={h_curve:.2f} m ≈ H_nom={h_nom:.2f} m (écart {ecart:.2f} m)",
                value=f"Q={q_nom:.1f} L/s, H={h_nom:.2f} m",
                threshold=f"écart ≤ {seuil:.1f} m",
            )
        return self._check(
            "A1", "Point de fonctionnement dans la courbe H(Q)",
            CAT_A, STATUS_WARN,
            f"Écart H(Q_nom)={h_curve:.2f} m vs H_nom={h_nom:.2f} m = {ecart:.2f} m",
            value=f"Q={q_nom:.1f} L/s",
            threshold=f"écart ≤ {seuil:.1f} m",
        )

    def _check_A2_npsh(self) -> dict:
        """NPSH disponible ≥ NPSH requis + marge ?"""
        if not self._first_pump_parsed():
            return self._check(
                "A2", "NPSH disponible ≥ requis + marge",
                CAT_A, STATUS_NA,
                "Aucune pompe chargée — chargez un rapport pompe via l'onglet Rapport Technique",
            )
        # Chercher npsh_req et npsh_avl dans tous les parseurs (définition + instance)
        npsh_req = next(
            (p.parsed.get("npsh_required_m") for p in self.pump_parsers
             if p.parsed.get("npsh_required_m") is not None), None)
        npsh_avl = next(
            (p.parsed.get("npsh_available_m") for p in self.pump_parsers
             if p.parsed.get("npsh_available_m") is not None), None)
        if npsh_req is None or npsh_avl is None:
            return self._check(
                "A2", "NPSH disponible ≥ requis + marge",
                CAT_A, STATUS_NA,
                "NPSH requis non défini dans le rapport Bentley — complétez le modèle ou saisissez NPSH manuellement",
            )
        marge = npsh_avl - npsh_req
        if marge >= self.npsh_margin_m:
            return self._check(
                "A2", "NPSH disponible ≥ requis + marge",
                CAT_A, STATUS_OK,
                f"Marge = {marge:.2f} m (dispo {npsh_avl:.2f} m, requis {npsh_req:.2f} m)",
                value=f"{marge:.2f} m",
                threshold=f"≥ {self.npsh_margin_m:.1f} m",
            )
        if marge >= 0:
            return self._check(
                "A2", "NPSH disponible ≥ requis + marge",
                CAT_A, STATUS_WARN,
                f"Marge insuffisante : {marge:.2f} m (seuil {self.npsh_margin_m:.1f} m)",
                value=f"{marge:.2f} m",
                threshold=f"≥ {self.npsh_margin_m:.1f} m",
            )
        return self._check(
            "A2", "NPSH disponible ≥ requis + marge",
            CAT_A, STATUS_FAIL,
            f"Cavitation : NPSH dispo {npsh_avl:.2f} m < requis {npsh_req:.2f} m (déficit {-marge:.2f} m)",
            value=f"{marge:.2f} m",
            threshold=f"≥ {self.npsh_margin_m:.1f} m",
        )

    def _check_A3_specific_speed(self) -> dict:
        """Vitesse spécifique pompe (Nq) dans la plage usuelle [25-80] SI ?"""
        if not self._first_pump_parsed():
            return self._check(
                "A3", "Vitesse spécifique (Nq) dans plage usuelle",
                CAT_A, STATUS_NA,
                "Aucune pompe chargée — chargez un rapport pompe via l'onglet Rapport Technique",
            )
        nq = next(
            (p.parsed.get("nq_si") for p in self.pump_parsers
             if p.parsed.get("nq_si") is not None),
            None
        )
        if nq is None:
            return self._check(
                "A3", "Vitesse spécifique (Nq) dans plage usuelle",
                CAT_A, STATUS_NA,
                "Nq absent du rapport — importez la définition pompe (contient 'Specific Speed') ou saisissez Nq manuellement",
            )
        if self.nq_min_si <= nq <= self.nq_max_si:
            return self._check(
                "A3", "Vitesse spécifique (Nq) dans plage usuelle",
                CAT_A, STATUS_OK,
                f"Nq = {nq:.1f} SI ∈ [{self.nq_min_si:.0f}, {self.nq_max_si:.0f}]",
                value=f"{nq:.1f}",
                threshold=f"[{self.nq_min_si:.0f}, {self.nq_max_si:.0f}]",
            )
        return self._check(
            "A3", "Vitesse spécifique (Nq) dans plage usuelle",
            CAT_A, STATUS_WARN,
            f"Nq = {nq:.1f} SI hors plage [{self.nq_min_si:.0f}, {self.nq_max_si:.0f}]",
            value=f"{nq:.1f}",
            threshold=f"[{self.nq_min_si:.0f}, {self.nq_max_si:.0f}]",
        )

    # ────────────────────────────────────────────────────────────
    # B. Pompe ↔ HPT (2 checks)
    # ────────────────────────────────────────────────────────────

    def _check_B1_discharge_pressure(self) -> dict:
        """Pression refoulement pompe ≤ PN ?"""
        pump = self._first_pump_parsed()
        if not pump:
            return self._check(
                "B1", "Pression refoulement ≤ PN",
                CAT_B, STATUS_NA,
                "Aucune pompe chargée",
            )
        p_discharge = pump.get("pressure_discharge_bar")
        if p_discharge is None:
            return self._check(
                "B1", "Pression refoulement ≤ PN",
                CAT_B, STATUS_NA,
                "Pression de refoulement manquante dans le rapport pompe",
            )
        if p_discharge <= self.pn_value_bar:
            return self._check(
                "B1", "Pression refoulement ≤ PN",
                CAT_B, STATUS_OK,
                f"P_refoulement = {p_discharge:.2f} bar ≤ PN = {self.pn_value_bar:.2f} bar",
                value=f"{p_discharge:.2f} bar",
                threshold=f"≤ {self.pn_value_bar:.2f} bar",
            )
        return self._check(
            "B1", "Pression refoulement ≤ PN",
            CAT_B, STATUS_FAIL,
            f"P_refoulement {p_discharge:.2f} bar > PN {self.pn_value_bar:.2f} bar "
            f"(dépassement {p_discharge - self.pn_value_bar:.2f} bar)",
            value=f"{p_discharge:.2f} bar",
            threshold=f"≤ {self.pn_value_bar:.2f} bar",
        )

    def _check_B2_suction_pressure(self) -> dict:
        """Pression aspiration pompe ≥ Pmin ?"""
        pump = self._first_pump_parsed()
        if not pump:
            return self._check(
                "B2", "Pression aspiration ≥ Pmin",
                CAT_B, STATUS_NA,
                "Aucune pompe chargée",
            )
        p_suction = pump.get("pressure_suction_bar")
        if p_suction is None:
            return self._check(
                "B2", "Pression aspiration ≥ Pmin",
                CAT_B, STATUS_NA,
                "Pression d'aspiration manquante dans le rapport pompe",
            )
        if p_suction >= self.pmin_value_bar:
            return self._check(
                "B2", "Pression aspiration ≥ Pmin",
                CAT_B, STATUS_OK,
                f"P_aspiration = {p_suction:.2f} bar ≥ Pmin = {self.pmin_value_bar:.2f} bar",
                value=f"{p_suction:.2f} bar",
                threshold=f"≥ {self.pmin_value_bar:.2f} bar",
            )
        return self._check(
            "B2", "Pression aspiration ≥ Pmin",
            CAT_B, STATUS_FAIL,
            f"P_aspiration {p_suction:.2f} bar < Pmin {self.pmin_value_bar:.2f} bar "
            f"(déficit {self.pmin_value_bar - p_suction:.2f} bar)",
            value=f"{p_suction:.2f} bar",
            threshold=f"≥ {self.pmin_value_bar:.2f} bar",
        )

    # ────────────────────────────────────────────────────────────
    # C. Réseau ↔ HPT (3 checks)
    # ────────────────────────────────────────────────────────────

    def _check_C1_max_pressure(self) -> dict:
        """Pmax transitoire < PN ?"""
        if not self.transient_status or not self.transient_status.get("success"):
            return self._check(
                "C1", "Pmax transitoire < PN",
                CAT_C, STATUS_NA,
                "Analyse transitoire non disponible ou en échec",
            )
        pmax = self.transient_status.get("max_pressure_bar")
        if pmax is None:
            return self._check(
                "C1", "Pmax transitoire < PN",
                CAT_C, STATUS_NA,
                "Pmax transitoire non extrait du fichier HPT",
            )
        if pmax < self.pn_value_bar:
            return self._check(
                "C1", "Pmax transitoire < PN",
                CAT_C, STATUS_OK,
                f"Pmax = {pmax:.2f} bar < PN = {self.pn_value_bar:.2f} bar "
                f"(marge {self.pn_value_bar - pmax:.2f} bar)",
                value=f"{pmax:.2f} bar",
                threshold=f"< {self.pn_value_bar:.2f} bar",
            )
        return self._check(
            "C1", "Pmax transitoire < PN",
            CAT_C, STATUS_FAIL,
            f"Pmax {pmax:.2f} bar ≥ PN {self.pn_value_bar:.2f} bar "
            f"(dépassement {pmax - self.pn_value_bar:.2f} bar)",
            value=f"{pmax:.2f} bar",
            threshold=f"< {self.pn_value_bar:.2f} bar",
        )

    def _check_C2_min_pressure(self) -> dict:
        """Pmin transitoire > Pmin ?"""
        if not self.transient_status or not self.transient_status.get("success"):
            return self._check(
                "C2", "Pmin transitoire > Pmin admissible",
                CAT_C, STATUS_NA,
                "Analyse transitoire non disponible ou en échec",
            )
        pmin_t = self.transient_status.get("min_pressure_bar")
        if pmin_t is None:
            return self._check(
                "C2", "Pmin transitoire > Pmin admissible",
                CAT_C, STATUS_NA,
                "Pmin transitoire non extrait du fichier HPT",
            )
        if pmin_t > self.pmin_value_bar:
            return self._check(
                "C2", "Pmin transitoire > Pmin admissible",
                CAT_C, STATUS_OK,
                f"Pmin = {pmin_t:.2f} bar > Pmin admissible = {self.pmin_value_bar:.2f} bar "
                f"(marge {pmin_t - self.pmin_value_bar:.2f} bar)",
                value=f"{pmin_t:.2f} bar",
                threshold=f"> {self.pmin_value_bar:.2f} bar",
            )
        return self._check(
            "C2", "Pmin transitoire > Pmin admissible",
            CAT_C, STATUS_FAIL,
            f"Pmin {pmin_t:.2f} bar ≤ Pmin admissible {self.pmin_value_bar:.2f} bar",
            value=f"{pmin_t:.2f} bar",
            threshold=f"> {self.pmin_value_bar:.2f} bar",
        )

    def _check_C3_multi_pumps(self) -> dict:
        """Cohérence multi-pompes : si plusieurs, vérifier H/Q."""
        n_pumps = len(self.pump_parsers)
        if n_pumps <= 1:
            return self._check(
                "C3", "Cohérence multi-pompes",
                CAT_C, STATUS_NA,
                "Une seule pompe chargée (check non applicable)",
            )
        # Vérifier qu'aucune pompe n'a un statut incohérent
        labels = [p.parsed.get("label", "?") for p in self.pump_parsers
                  if getattr(p, "parsed", None)]
        if not labels:
            return self._check(
                "C3", "Cohérence multi-pompes",
                CAT_C, STATUS_FAIL,
                f"{n_pumps} pompe(s) chargée(s) mais aucune avec données parsées",
            )
        return self._check(
            "C3", "Cohérence multi-pompes",
            CAT_C, STATUS_OK,
            f"{n_pumps} pompe(s) chargée(s) : {', '.join(labels)}",
            value=f"{n_pumps} pompes",
            threshold="labels distincts",
        )

    # ────────────────────────────────────────────────────────────
    # D. HPT ↔ Ventouses / Vidanges (3 checks)
    # ────────────────────────────────────────────────────────────

    def _check_D1_vgas(self) -> dict:
        """Volume gaz HPT max < seuil utilisateur ?"""
        if not self.transient_status or not self.transient_status.get("success"):
            return self._check(
                "D1", "Volume gaz HPT max < seuil",
                CAT_D, STATUS_NA,
                "Analyse transitoire non disponible",
            )
        vgas = self.transient_status.get("max_gas_volume_l")
        if vgas is None:
            return self._check(
                "D1", "Volume gaz HPT max < seuil",
                CAT_D, STATUS_NA,
                "Volume de gaz max non extrait du fichier HPT",
            )
        if vgas <= self.vgas_threshold_l:
            return self._check(
                "D1", "Volume gaz HPT max < seuil",
                CAT_D, STATUS_OK,
                f"V_gaz max = {vgas:.1f} L ≤ seuil {self.vgas_threshold_l:.0f} L",
                value=f"{vgas:.1f} L",
                threshold=f"≤ {self.vgas_threshold_l:.0f} L",
            )
        return self._check(
            "D1", "Volume gaz HPT max < seuil",
            CAT_D, STATUS_WARN,
            f"V_gaz max {vgas:.1f} L > seuil {self.vgas_threshold_l:.0f} L — "
            f"vérifier l'HPT ou augmenter sa taille",
            value=f"{vgas:.1f} L",
            threshold=f"≤ {self.vgas_threshold_l:.0f} L",
        )

    def _check_D2_air_valves(self) -> dict:
        """Ventouses présentes si profil non monotone ?"""
        if not self.air_valve_sizer or not getattr(self.air_valve_sizer, "profile", None):
            return self._check(
                "D2", "Ventouses aux points hauts",
                CAT_D, STATUS_NA,
                "Aucun profil en long chargé",
            )
        profile = self.air_valve_sizer.profile
        n_highs = len(getattr(self.air_valve_sizer, "high_points", []))
        n_vents = len(getattr(self.air_valve_sizer, "ventouses", []))
        # Si profil monotone (0 points hauts) → check NA
        if n_highs == 0:
            return self._check(
                "D2", "Ventouses aux points hauts",
                CAT_D, STATUS_NA,
                "Profil monotone (aucun point haut) — ventouses non requises",
            )
        if n_vents >= n_highs:
            return self._check(
                "D2", "Ventouses aux points hauts",
                CAT_D, STATUS_OK,
                f"{n_vents} ventouse(s) pour {n_highs} point(s) haut(s)",
                value=f"{n_vents}",
                threshold=f"≥ {n_highs}",
            )
        return self._check(
            "D2", "Ventouses aux points hauts",
            CAT_D, STATUS_WARN,
            f"{n_vents} ventouse(s) seulement pour {n_highs} point(s) haut(s)",
            value=f"{n_vents}",
            threshold=f"≥ {n_highs}",
        )

    def _check_D3_drains(self) -> dict:
        """Vidanges présentes si profil non monotone ?"""
        if not self.air_valve_sizer or not getattr(self.air_valve_sizer, "profile", None):
            return self._check(
                "D3", "Vidanges aux points bas",
                CAT_D, STATUS_NA,
                "Aucun profil en long chargé",
            )
        profile = self.air_valve_sizer.profile
        n_lows = len(getattr(self.air_valve_sizer, "low_points", []))
        n_drains = len(getattr(self.air_valve_sizer, "vidanges", []))
        if n_lows == 0:
            return self._check(
                "D3", "Vidanges aux points bas",
                CAT_D, STATUS_NA,
                "Profil monotone (aucun point bas) — vidanges non requises",
            )
        if n_drains >= 1:
            return self._check(
                "D3", "Vidanges aux points bas",
                CAT_D, STATUS_OK,
                f"{n_drains} vidange(s) localisée(s) entre {n_lows} point(s) bas",
                value=f"{n_drains}",
                threshold=f"≥ 1",
            )
        return self._check(
            "D3", "Vidanges aux points bas",
            CAT_D, STATUS_WARN,
            f"Aucune vidange localisée pour {n_lows} point(s) bas — "
            f"vérifier les distances inter-ventouses",
            value=f"{n_drains}",
            threshold=f"≥ 1",
        )

    # ────────────────────────────────────────────────────────────
    # E. Cohérence globale (5 checks)
    # ────────────────────────────────────────────────────────────

    def _check_E1_profile_loaded(self) -> dict:
        """Profil en long chargé (≥ 2 points) ?"""
        if not self.air_valve_sizer or not getattr(self.air_valve_sizer, "profile", None):
            return self._check(
                "E1", "Profil en long chargé",
                CAT_E, STATUS_FAIL,
                "Aucun profil en long — chargez un profil avant diagnostic complet",
            )
        n = len(self.air_valve_sizer.profile)
        if n < 2:
            return self._check(
                "E1", "Profil en long chargé",
                CAT_E, STATUS_FAIL,
                f"Profil avec seulement {n} point(s) — au moins 2 requis",
            )
        return self._check(
            "E1", "Profil en long chargé",
            CAT_E, STATUS_OK,
            f"Profil chargé avec {n} points",
            value=f"{n} points",
            threshold="≥ 2",
        )

    def _check_E2_dn_consistency(self) -> dict:
        """DN profil = DN pipes sheet ?"""
        if not self.air_valve_sizer or not getattr(self.air_valve_sizer, "profile", None):
            return self._check(
                "E2", "Cohérence DN profil ↔ classeur",
                CAT_E, STATUS_NA,
                "Aucun profil en long chargé",
            )
        if not self.workbook_manager:
            return self._check(
                "E2", "Cohérence DN profil ↔ classeur",
                CAT_E, STATUS_NA,
                "Aucun classeur HAMMER chargé",
            )
        try:
            summary = self.workbook_manager.get_summary()
        except Exception:
            return self._check(
                "E2", "Cohérence DN profil ↔ classeur",
                CAT_E, STATUS_NA,
                "Résumé classeur indisponible",
            )
        # Récupérer DN min/max du classeur
        dn_min = summary.get("diameter_min_mm")
        dn_max = summary.get("diameter_max_mm")
        dn_profile = self.air_valve_sizer.pipe_dn_mm
        if dn_min is None or dn_max is None:
            return self._check(
                "E2", "Cohérence DN profil ↔ classeur",
                CAT_E, STATUS_NA,
                "DN non disponible dans le classeur HAMMER",
            )
        if dn_min <= dn_profile <= dn_max:
            return self._check(
                "E2", "Cohérence DN profil ↔ classeur",
                CAT_E, STATUS_OK,
                f"DN profil {dn_profile:.0f} mm ∈ [{dn_min:.0f}, {dn_max:.0f}] mm (classeur)",
                value=f"{dn_profile:.0f} mm",
                threshold=f"[{dn_min:.0f}, {dn_max:.0f}] mm",
            )
        return self._check(
            "E2", "Cohérence DN profil ↔ classeur",
            CAT_E, STATUS_WARN,
            f"DN profil {dn_profile:.0f} mm hors plage classeur [{dn_min:.0f}, {dn_max:.0f}] mm",
            value=f"{dn_profile:.0f} mm",
            threshold=f"[{dn_min:.0f}, {dn_max:.0f}] mm",
        )

    def _check_E3_elevation_positive(self) -> dict:
        """Cote min profil > 0 ?"""
        if not self.air_valve_sizer or not getattr(self.air_valve_sizer, "profile", None):
            return self._check(
                "E3", "Cote min profil > 0",
                CAT_E, STATUS_NA,
                "Aucun profil en long chargé",
            )
        z_min = min((p["z_m"] for p in self.air_valve_sizer.profile), default=None)
        if z_min is None:
            return self._check(
                "E3", "Cote min profil > 0",
                CAT_E, STATUS_NA,
                "Cote min non calculable",
            )
        if z_min > 0:
            return self._check(
                "E3", "Cote min profil > 0",
                CAT_E, STATUS_OK,
                f"Z_min = {z_min:.2f} m > 0 (physiquement cohérent)",
                value=f"{z_min:.2f} m",
                threshold="> 0 m",
            )
        return self._check(
            "E3", "Cote min profil > 0",
            CAT_E, STATUS_FAIL,
            f"Z_min = {z_min:.2f} m ≤ 0 (incohérence géométrique)",
            value=f"{z_min:.2f} m",
            threshold="> 0 m",
        )

    def _check_E4_slope_max(self) -> dict:
        """Pente max < 50 % ?"""
        if not self.air_valve_sizer or not getattr(self.air_valve_sizer, "profile", None):
            return self._check(
                "E4", "Pente max < seuil",
                CAT_E, STATUS_NA,
                "Aucun profil en long chargé",
            )
        slopes = [abs(p.get("pente_pct") or 0)
                  for p in self.air_valve_sizer.profile]
        s_max = max(slopes) if slopes else 0.0
        if s_max <= self.slope_max_pct:
            return self._check(
                "E4", "Pente max < seuil",
                CAT_E, STATUS_OK,
                f"Pente max = {s_max:.2f} % ≤ {self.slope_max_pct:.1f} %",
                value=f"{s_max:.2f} %",
                threshold=f"≤ {self.slope_max_pct:.1f} %",
            )
        return self._check(
            "E4", "Pente max < seuil",
            CAT_E, STATUS_WARN,
            f"Pente max = {s_max:.2f} % > {self.slope_max_pct:.1f} % (point singulier ?)",
            value=f"{s_max:.2f} %",
            threshold=f"≤ {self.slope_max_pct:.1f} %",
        )

    def _check_E5_pumps_loaded(self) -> dict:
        """Au moins 1 pompe chargée ?"""
        n = sum(1 for p in self.pump_parsers if getattr(p, "parsed", None))
        if n == 0:
            return self._check(
                "E5", "Au moins 1 pompe chargée",
                CAT_E, STATUS_FAIL,
                "Aucune pompe chargée — vérifiez les rapports RTF pompe",
            )
        return self._check(
            "E5", "Au moins 1 pompe chargée",
            CAT_E, STATUS_OK,
            f"{n} pompe(s) chargée(s) et parsée(s)",
            value=f"{n}",
            threshold="≥ 1",
        )

    # ────────────────────────────────────────────────────────────
    # Exécution globale
    # ────────────────────────────────────────────────────────────

    def run_checks(self) -> list[dict]:
        """Exécute les 16 vérifications et retourne la liste de résultats.
        Chaque appel de vérification est protégé pour éviter qu'une exception
        non prévue n'interrompe l'ensemble du diagnostic. En cas d'erreur
        inattendue, le check correspondant renvoie un statut FAILURE avec le
        message d'exception, garantissant que le processus complet continue.
        """
        checks = []
        for method in [
            self._check_A1_operating_point,
            self._check_A2_npsh,
            self._check_A3_specific_speed,
            self._check_B1_discharge_pressure,
            self._check_B2_suction_pressure,
            self._check_C1_max_pressure,
            self._check_C2_min_pressure,
            self._check_C3_multi_pumps,
            self._check_D1_vgas,
            self._check_D2_air_valves,
            self._check_D3_drains,
            self._check_E1_profile_loaded,
            self._check_E2_dn_consistency,
            self._check_E3_elevation_positive,
            self._check_E4_slope_max,
            self._check_E5_pumps_loaded,
        ]:
            try:
                checks.append(method())
            except Exception as exc:
                # Retourner un résultat générique d'échec pour le check
                checks.append({
                    "code": getattr(method, "__name__", "unknown"),
                    "name": method.__doc__.split("\n")[0] if method.__doc__ else "",
                    "category": "UNKNOWN",
                    "status": STATUS_FAIL,
                    "message": f"Exception inattendue : {exc}",
                    "value": None,
                    "threshold": None,
                })
        return checks

    def get_summary(self) -> dict:
        """Compteurs OK / WARN / FAIL / NA."""
        checks = self.run_checks()
        summary = {STATUS_OK: 0, STATUS_WARN: 0, STATUS_FAIL: 0, STATUS_NA: 0}
        for c in checks:
            s = c.get("status", STATUS_NA)
            summary[s] = summary.get(s, 0) + 1
        summary["total"] = len(checks)
        return summary

    # ────────────────────────────────────────────────────────────
    # Sérialisation
    # ────────────────────────────────────────────────────────────

    def serialize(self) -> dict:
        """Sérialise les checks pour .hpi."""
        checks = self.run_checks()
        return {
            "version": 1,
            "checks": checks,
            "summary": self.get_summary(),
            "params": {
                "pn_value_bar": self.pn_value_bar,
                "pmin_value_bar": self.pmin_value_bar,
                "vgas_threshold_l": self.vgas_threshold_l,
                "npsh_margin_m": self.npsh_margin_m,
                "nq_min_si": self.nq_min_si,
                "nq_max_si": self.nq_max_si,
                "slope_max_pct": self.slope_max_pct,
            },
        }

    @staticmethod
    def deserialize(data: dict | None) -> list[dict]:
        """Désérialise une liste de checks depuis .hpi."""
        if not data or not isinstance(data, dict):
            return []
        if data.get("version") != 1:
            return []
        return data.get("checks", [])
