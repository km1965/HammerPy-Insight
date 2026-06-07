#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
diagnostics_charts.py — Génération de graphiques Matplotlib imprimables
(palette claire) pour la section 6 du rapport Word Phase 4.

Quatre fonctions publiques :
  - build_kpi_donut      : diagramme en anneau OK/WARN/FAIL/NA
  - build_category_stack : barres horizontales empilées par catégorie A-E
  - build_compliance_bars: 3 sous-graphes Pmax/Pmin/Vgaz vs seuils
  - build_profile_chart  : profil en long + ventouses + vidanges (palette claire)

Chaque fonction :
  - prend les données nécessaires en paramètres
  - retourne un chemin PNG temporaire (str) ou None si données insuffisantes
  - utilise une palette imprimable (fond blanc, texte noir) pour le rapport Word
  - écrit le PNG via tempfile.NamedTemporaryFile(delete=False) pour que
    le fichier survive après fermeture

L'appelant (main.py) doit nettoyer les PNG via os.unlink() dans finally:.
"""

import os
import tempfile
from collections import defaultdict

from matplotlib.figure import Figure
import matplotlib.patches as mpatches

# Palette imprimable
COLOR_OK   = "#1a7a1a"
COLOR_WARN = "#c08000"
COLOR_FAIL = "#c00000"
COLOR_NA   = "#888888"
COLOR_TXT  = "#202020"
COLOR_BG   = "#ffffff"
COLOR_AX   = "#fafafa"
COLOR_GRID = "#cccccc"
COLOR_LINE = "#1f538d"
COLOR_PROFILE_FILL = "#4cc9f0"
COLOR_VENTOUSE = "#06d6a0"
COLOR_VIDANGE = "#ef476f"

# Status → numéro (pour heatmap optionnelle)
STATUS_NUM = {"OK": 4, "WARN": 3, "FAIL": 1, "NA": 2}
STATUS_COLOR = {"OK": COLOR_OK, "WARN": COLOR_WARN, "FAIL": COLOR_FAIL, "NA": COLOR_NA}


def _save_fig_to_tmp(fig: Figure, dpi: int = 150) -> str:
    """Sauvegarde une Figure en PNG temporaire, retourne le chemin."""
    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    fig.savefig(tmp.name, dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    tmp.close()
    return tmp.name


def _apply_print_style(fig: Figure, ax) -> None:
    """Applique un style imprimable à une figure/axe."""
    fig.patch.set_facecolor(COLOR_BG)
    ax.set_facecolor(COLOR_AX)
    ax.tick_params(colors=COLOR_TXT, labelsize=8)
    for spine in ax.spines.values():
        spine.set_color("#888888")
    ax.title.set_color(COLOR_TXT)
    ax.xaxis.label.set_color(COLOR_TXT)
    ax.yaxis.label.set_color(COLOR_TXT)
    ax.grid(True, ls=":", color=COLOR_GRID, alpha=0.6)


# ════════════════════════════════════════════════════════════════════
# Chart 1 — Donut OK/WARN/FAIL/NA
# ════════════════════════════════════════════════════════════════════

def build_kpi_donut(summary: dict | None) -> str | None:
    """
    Diagramme en anneau (donut) avec les compteurs OK / WARN / FAIL / NA.
    Retourne le chemin d'un PNG temporaire ou None si pas de données.
    """
    if not summary:
        return None
    labels = ["OK", "WARN", "FAIL", "NA"]
    values = [summary.get(k, 0) for k in labels]
    if sum(values) == 0:
        return None

    colors = [COLOR_OK, COLOR_WARN, COLOR_FAIL, COLOR_NA]
    fig = Figure(figsize=(6, 4.5), dpi=150)
    fig.patch.set_facecolor(COLOR_BG)
    ax = fig.add_subplot(111)
    ax.set_facecolor(COLOR_BG)

    wedges, _texts, autotexts = ax.pie(
        values,
        labels=[f"{l} : {v}" for l, v in zip(labels, values)],
        colors=colors,
        autopct="%1.0f%%",
        startangle=90,
        wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2),
        textprops=dict(fontsize=11, color=COLOR_TXT),
    )
    for at in autotexts:
        at.set_color("white")
        at.set_fontsize(10)
        at.set_fontweight("bold")
    ax.set_title("Synthèse des 16 vérifications croisées",
                 fontsize=12, fontweight="bold", pad=14, color=COLOR_TXT)
    fig.tight_layout()
    return _save_fig_to_tmp(fig)


# ════════════════════════════════════════════════════════════════════
# Chart 2 — Barres horizontales empilées par catégorie (A → E)
# ════════════════════════════════════════════════════════════════════

def build_category_stack(checks: list[dict] | None) -> str | None:
    """
    Barres horizontales empilées : une barre par catégorie (A → E),
    décomposée en OK / WARN / FAIL / NA. Retourne PNG temp ou None.
    """
    if not checks:
        return None

    by_cat = defaultdict(lambda: {"OK": 0, "WARN": 0, "FAIL": 0, "NA": 0})
    cat_order = []
    for c in checks:
        cat = c.get("category", "—")
        status = c.get("status", "NA")
        if cat not in by_cat:
            cat_order.append(cat)
        # Statuts inconnus sont comptés comme N/A (robustesse)
        if status not in by_cat[cat]:
            status = "NA"
        by_cat[cat][status] += 1

    if not cat_order:
        return None

    ok   = [by_cat[c]["OK"]   for c in cat_order]
    warn = [by_cat[c]["WARN"] for c in cat_order]
    fail = [by_cat[c]["FAIL"] for c in cat_order]
    na   = [by_cat[c]["NA"]   for c in cat_order]

    fig = Figure(figsize=(9, 4), dpi=150)
    fig.patch.set_facecolor(COLOR_BG)
    ax = fig.add_subplot(111)

    y = list(range(len(cat_order)))
    # Étiquettes courtes (retire le préfixe lettre + point)
    short_labels = []
    for cat in cat_order:
        if ". " in cat:
            short_labels.append(cat.split(". ", 1)[1])
        else:
            short_labels.append(cat)

    left_w = [a for a in ok]
    left_f = [a + b for a, b in zip(ok, warn)]
    left_n = [a + b + c for a, b, c in zip(ok, warn, fail)]

    ax.barh(y, ok,   color=COLOR_OK,   label="✔ OK",   edgecolor="white")
    ax.barh(y, warn, left=left_w, color=COLOR_WARN, label="⚠ WARN", edgecolor="white")
    ax.barh(y, fail, left=left_f, color=COLOR_FAIL, label="✘ FAIL", edgecolor="white")
    ax.barh(y, na,   left=left_n, color=COLOR_NA,   label="— N/A",  edgecolor="white")

    # Valeurs dans les segments
    for i, (a, b, c, d) in enumerate(zip(ok, warn, fail, na)):
        if a > 0:
            ax.text(a / 2, i, str(a), va="center", ha="center",
                    color="white", fontsize=9, fontweight="bold")
        if b > 0:
            ax.text(left_w[i] + b / 2, i, str(b), va="center", ha="center",
                    color="white", fontsize=9, fontweight="bold")
        if c > 0:
            ax.text(left_f[i] + c / 2, i, str(c), va="center", ha="center",
                    color="white", fontsize=9, fontweight="bold")
        if d > 0:
            ax.text(left_n[i] + d / 2, i, str(d), va="center", ha="center",
                    color="white", fontsize=9, fontweight="bold")

    ax.set_yticks(y)
    ax.set_yticklabels(short_labels, fontsize=9, color=COLOR_TXT)
    ax.set_xlabel("Nombre de vérifications", color=COLOR_TXT, fontsize=9)
    ax.set_xlim(0, max(sum(x) for x in zip(ok, warn, fail, na)) + 0.5)
    ax.set_title("Résultats par catégorie (5 domaines A → E)",
                 fontsize=12, fontweight="bold", pad=10, color=COLOR_TXT)
    ax.legend(loc="lower right", fontsize=9, frameon=True, edgecolor=COLOR_GRID)
    _apply_print_style(fig, ax)
    ax.set_axisbelow(True)
    fig.tight_layout()
    return _save_fig_to_tmp(fig)


# ════════════════════════════════════════════════════════════════════
# Chart 3 — Conformité Pmax / Pmin / Vgaz vs seuils
# ════════════════════════════════════════════════════════════════════

def build_compliance_bars(transient_status: dict | None,
                          pn_value: float | None,
                          pmin_value: float | None,
                          vgas_threshold: float | None) -> str | None:
    """
    3 sous-graphes en barres : Pmax vs PN, Pmin vs Pmin admissible, Vgaz vs seuil.
    Couleurs conditionnelles (rouge si hors limite, vert si OK).
    """
    if not transient_status or not transient_status.get("success"):
        return None
    if pn_value is None or pmin_value is None or vgas_threshold is None:
        return None

    pmax = transient_status.get("max_pressure_bar")
    pmin = transient_status.get("min_pressure_bar")
    vgas = transient_status.get("max_gas_volume_l")

    if pmax is None and pmin is None and vgas is None:
        return None

    fig = Figure(figsize=(9, 3.2), dpi=150)
    fig.patch.set_facecolor(COLOR_BG)

    # ── Subplot 1 : Pmax vs PN ──────────────────────────────────────
    if pmax is not None:
        ax1 = fig.add_subplot(1, 3, 1)
        ax1.set_facecolor(COLOR_BG)
        ok = pmax <= pn_value
        color_pmax = COLOR_OK if ok else COLOR_FAIL
        ax1.bar(["Pmax\ntransitoire"], [pmax], color=color_pmax, edgecolor="white", width=0.5)
        ax1.axhline(pn_value, color=COLOR_FAIL, ls="--", lw=1.5,
                    label=f"PN = {pn_value:.1f} bar")
        ax1.set_ylim(0, max(pmax, pn_value) * 1.15)
        ax1.set_ylabel("Pression (bar)", color=COLOR_TXT, fontsize=9)
        ax1.set_title("Surpression vs PN", fontsize=10, fontweight="bold", color=COLOR_TXT)
        ax1.legend(fontsize=8, loc="upper right")
        for sp in ax1.spines.values():
            sp.set_color("#888888")
        ax1.tick_params(colors=COLOR_TXT, labelsize=8)
        ax1.grid(True, axis="y", ls=":", color=COLOR_GRID, alpha=0.6)
        # Valeur au-dessus de la barre
        ax1.text(0, pmax, f"{pmax:.2f}", ha="center", va="bottom",
                 fontsize=9, fontweight="bold", color=COLOR_TXT)

    # ── Subplot 2 : Pmin vs Pmin admissible ─────────────────────────
    if pmin is not None:
        ax2 = fig.add_subplot(1, 3, 2)
        ax2.set_facecolor(COLOR_BG)
        ok = pmin > pmin_value
        color_pmin = COLOR_OK if ok else COLOR_FAIL
        bar_vals = max(pmin, 0.0)
        ax2.bar(["Pmin\ntransitoire"], [bar_vals], color=color_pmin, edgecolor="white", width=0.5)
        ax2.axhline(pmin_value, color=COLOR_FAIL, ls="--", lw=1.5,
                    label=f"Pmin adm. = {pmin_value:.2f} bar")
        ymin = min(pmin, pmin_value, 0) - 0.5
        ymax = max(bar_vals * 1.15, 2.0)
        ax2.set_ylim(ymin, ymax)
        ax2.set_ylabel("Pression (bar)", color=COLOR_TXT, fontsize=9)
        ax2.set_title("Dépression vs Pmin admissible", fontsize=10,
                      fontweight="bold", color=COLOR_TXT)
        ax2.legend(fontsize=8, loc="upper right")
        for sp in ax2.spines.values():
            sp.set_color("#888888")
        ax2.tick_params(colors=COLOR_TXT, labelsize=8)
        ax2.grid(True, axis="y", ls=":", color=COLOR_GRID, alpha=0.6)
        ax2.text(0, pmin, f"{pmin:.2f}", ha="center", va="top" if pmin < 0 else "bottom",
                 fontsize=9, fontweight="bold", color=COLOR_TXT)

    # ── Subplot 3 : Vgaz vs seuil ───────────────────────────────────
    if vgas is not None:
        ax3 = fig.add_subplot(1, 3, 3)
        ax3.set_facecolor(COLOR_BG)
        ok = vgas <= vgas_threshold
        color_vgas = COLOR_OK if ok else COLOR_WARN
        ax3.bar(["Vgaz\nmax"], [vgas], color=color_vgas, edgecolor="white", width=0.5)
        ax3.axhline(vgas_threshold, color=COLOR_FAIL, ls="--", lw=1.5,
                    label=f"Seuil = {vgas_threshold:.0f} L")
        ax3.set_ylim(0, max(vgas, vgas_threshold) * 1.15)
        ax3.set_ylabel("Volume (L)", color=COLOR_TXT, fontsize=9)
        ax3.set_title("Volume gaz HPT vs seuil", fontsize=10,
                      fontweight="bold", color=COLOR_TXT)
        ax3.legend(fontsize=8, loc="upper right")
        for sp in ax3.spines.values():
            sp.set_color("#888888")
        ax3.tick_params(colors=COLOR_TXT, labelsize=8)
        ax3.grid(True, axis="y", ls=":", color=COLOR_GRID, alpha=0.6)
        ax3.text(0, vgas, f"{vgas:.1f}", ha="center", va="bottom",
                 fontsize=9, fontweight="bold", color=COLOR_TXT)

    fig.suptitle("Conformité des paramètres transitoires (HPT)",
                 fontsize=12, fontweight="bold", y=1.02, color=COLOR_TXT)
    fig.tight_layout()
    return _save_fig_to_tmp(fig)


# ════════════════════════════════════════════════════════════════════
# Chart 4 — Profil en long + ventouses + vidanges (palette claire)
# ════════════════════════════════════════════════════════════════════

def build_profile_chart(profile: list[dict] | None,
                        ventouses: list[dict] | None = None,
                        vidanges: list[dict] | None = None,
                        pipe_dn_mm: float | None = None) -> str | None:
    """
    Profil en long de la conduite avec marqueurs ventouses (▲ vert) et
    vidanges (▼ rouge). Palette imprimable (fond blanc). Retourne PNG temp
    ou None si pas de profil.
    """
    if not profile or len(profile) < 2:
        return None

    pks = [p.get("pk_m", 0) for p in profile]
    zs  = [p.get("z_m", 0) for p in profile]
    ventouses = ventouses or []
    vidanges  = vidanges  or []

    fig = Figure(figsize=(9, 3.8), dpi=150)
    fig.patch.set_facecolor(COLOR_BG)
    ax = fig.add_subplot(111)
    ax.set_facecolor(COLOR_AX)

    # Profil + remplissage
    ax.plot(pks, zs, color=COLOR_LINE, lw=1.5, marker="o",
            markersize=3, markerfacecolor=COLOR_LINE,
            markeredgecolor="white", label="Profil TN")
    if zs:
        z_min = min(zs)
        ax.fill_between(pks, zs, z_min - 5,
                        color=COLOR_PROFILE_FILL, alpha=0.10)

    # Ventouses (▲ vert)
    for v in ventouses:
        ax.plot(v.get("pk_m", 0), v.get("z_m", 0), marker="^",
                color=COLOR_VENTOUSE, markersize=10,
                markeredgecolor="white", markeredgewidth=0.8)

    # Vidanges (▼ rouge)
    for d in vidanges:
        ax.plot(d.get("pk_m", 0), d.get("z_m", 0), marker="v",
                color=COLOR_VIDANGE, markersize=10,
                markeredgecolor="white", markeredgewidth=0.8)

    # Légende custom
    legend_items = [
        mpatches.Patch(color=COLOR_LINE,            label="Profil TN"),
        mpatches.Patch(color=COLOR_VENTOUSE,         label=f"Ventouses ({len(ventouses)})"),
        mpatches.Patch(color=COLOR_VIDANGE,          label=f"Vidanges ({len(vidanges)})"),
    ]
    ax.legend(handles=legend_items, loc="lower right", fontsize=9,
              frameon=True, edgecolor=COLOR_GRID)

    title = "Profil en long & protections"
    if pipe_dn_mm is not None and pipe_dn_mm > 0:
        title += f" (DN {pipe_dn_mm:.0f} mm)"
    ax.set_title(title, fontsize=11, fontweight="bold", color=COLOR_TXT, pad=10)
    ax.set_xlabel("PK (m)", color=COLOR_TXT, fontsize=9)
    ax.set_ylabel("Côte Z (m)", color=COLOR_TXT, fontsize=9)
    _apply_print_style(fig, ax)
    fig.tight_layout()
    return _save_fig_to_tmp(fig)
