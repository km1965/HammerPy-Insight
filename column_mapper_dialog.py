#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
column_mapper_dialog.py — Boîte de dialogue modale pour mapping de colonnes.

Présente une fenêtre CTkToplevel avec :
- Label explicatif (quelle colonne est manquante)
- Dropdown (CTkComboBox) listant les colonnes disponibles
- Boutons OK / Skip / Cancel
"""

import customtkinter as ctk


SENTINEL_SKIP = "__SKIP__"
SENTINEL_CANCEL = "__CANCEL__"


def ask_column_mapping(
    parent,
    unknown_col: str,
    available_cols: list[str],
    file_type: str = "",
    file_hash: str = "",
) -> str | None:
    """
    Affiche une boîte de dialogue modale pour demander un mapping.

    Args:
        parent: Fenêtre parente (CTk)
        unknown_col: Nom de la colonne attendue (non trouvée)
        available_cols: Liste des colonnes réellement présentes dans le fichier
        file_type: Type de fichier (info affichée)
        file_hash: Hash du fichier (info affichée)

    Returns:
        Nom de la colonne choisie, SENTINEL_SKIP, SENTINEL_CANCEL, ou None
    """
    result = {"value": SENTINEL_CANCEL}

    dialog = ctk.CTkToplevel(parent)
    dialog.title("Mapping de colonne")
    dialog.geometry("520x340")
    dialog.transient(parent)
    dialog.grab_set()
    dialog.resizable(False, False)

    # Centrer
    dialog.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() // 2) - 260
    y = parent.winfo_rooty() + (parent.winfo_height() // 2) - 170
    dialog.geometry(f"+{x}+{y}")

    # ── En-tête ─────────────────────────────────────────────────
    ctk.CTkLabel(
        dialog,
        text="⚠️ Colonne non reconnue",
        font=ctk.CTkFont(size=16, weight="bold"),
        text_color="#f59e0b"
    ).pack(pady=(20, 8), padx=20, anchor="w")

    ctk.CTkLabel(
        dialog,
        text=f"Le programme attend la colonne :",
        font=ctk.CTkFont(size=12),
    ).pack(pady=(0, 4), padx=20, anchor="w")

    ctk.CTkLabel(
        dialog,
        text=f"  « {unknown_col} »",
        font=ctk.CTkFont(size=13, weight="bold"),
        text_color="#4cc9f0",
    ).pack(pady=(0, 12), padx=20, anchor="w")

    if file_type:
        type_label = {
            "profile_csv": "Profil CSV libre",
            "profile_bentley": "Profil CSV Bentley FlexTable",
            "pipes": "Pipes (classeur HAMMER)",
            "nodes": "Nœuds (classeur HAMMER)",
            "pumps": "Pompes (classeur HAMMER)",
            "reservoirs": "Réservoirs (classeur HAMMER)",
            "hpt": "HPT (classeur HAMMER)",
            "air_valves": "Ventouses (classeur HAMMER)",
        }.get(file_type, file_type)
        ctk.CTkLabel(
            dialog,
            text=f"Type de fichier : {type_label}",
            font=ctk.CTkFont(size=10),
            text_color="gray",
        ).pack(pady=(0, 4), padx=20, anchor="w")

    # ── Dropdown ────────────────────────────────────────────────
    ctk.CTkLabel(
        dialog,
        text="Quelle colonne du fichier lui correspond ?",
        font=ctk.CTkFont(size=12),
    ).pack(pady=(8, 4), padx=20, anchor="w")

    combo_var = ctk.StringVar(value=available_cols[0] if available_cols else "")
    combo = ctk.CTkComboBox(
        dialog,
        values=available_cols,
        variable=combo_var,
        width=460,
        height=32,
        font=ctk.CTkFont(size=12),
    )
    combo.pack(pady=(0, 8), padx=20)

    # ── Case à cocher : mémoriser ───────────────────────────────
    remember_var = ctk.BooleanVar(value=True)
    ctk.CTkCheckBox(
        dialog,
        text="Mémoriser ce mapping (jamais redemandé)",
        variable=remember_var,
        font=ctk.CTkFont(size=11),
    ).pack(pady=(4, 12), padx=20, anchor="w")

    # ── Boutons ─────────────────────────────────────────────────
    btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    btn_frame.pack(pady=(8, 16), padx=20, fill="x")

    def on_ok():
        val = combo_var.get().strip()
        if val and val != "":
            result["value"] = val
            if not remember_var.get():
                # L'utilisateur ne veut pas mémoriser → on retourne le nom
                # mais le mapper ne le stockera pas (signal via SENTINEL_SKIP)
                result["no_remember"] = True
            dialog.destroy()

    def on_skip():
        result["value"] = SENTINEL_SKIP
        dialog.destroy()

    def on_cancel():
        result["value"] = SENTINEL_CANCEL
        dialog.destroy()

    ctk.CTkButton(
        btn_frame,
        text="✓ OK",
        command=on_ok,
        fg_color="#2d6a4f",
        hover_color="#1b4332",
        width=120,
    ).pack(side="left", padx=(0, 8))

    ctk.CTkButton(
        btn_frame,
        text="⏭ Ignorer",
        command=on_skip,
        fg_color="transparent",
        border_width=1,
        text_color=("gray10", "gray90"),
        width=120,
    ).pack(side="left", padx=8)

    ctk.CTkButton(
        btn_frame,
        text="✕ Annuler",
        command=on_cancel,
        fg_color="#9b2226",
        hover_color="#6a1010",
        width=120,
    ).pack(side="right")

    # Touche Echap = Annuler, Entrée = OK
    dialog.bind("<Escape>", lambda e: on_cancel())
    dialog.bind("<Return>", lambda e: on_ok())

    # Attendre la fermeture
    dialog.wait_window()

    if result.get("no_remember"):
        # Le mapping est retourné mais le caller ne doit pas le stocker
        return ("__NO_REMEMBER__", result["value"])
    return result["value"]
