#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
model_builder_dialog.py — Dialogue interactif pour le Model Builder.
"""

import os
import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd

from model_builder import (
    ModelBuilder, ModelData,
    PIPE_COLUMNS, NODE_COLUMNS,
    PIPE_MANDATORY, NODE_MANDATORY,
    FILE_TYPE_MODEL_PIPES, FILE_TYPE_MODEL_NODES,
)
from column_mapper import get_mapper, file_hash_from_path
from column_mapper_dialog import ask_column_mapping


class ModelBuilderDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Model Builder — Import r\u00e9seau")
        self.geometry("680x600")
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)

        self.builder = ModelBuilder()
        self.filepath = ""
        self.sheet_names = []
        self.pipe_df = None
        self.node_df = None
        self.result = None

        self._build_ui()

        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - 340
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - 300
        self.geometry(f"+{x}+{y}")

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=0)
        self.grid_rowconfigure(3, weight=1)
        self.grid_rowconfigure(4, weight=0)

        # ── Titre ──
        ctk.CTkLabel(
            self, text="\U0001f9f1  Model Builder — Import r\u00e9seau",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).grid(row=0, column=0, padx=20, pady=(18, 8), sticky="w")

        ctk.CTkLabel(
            self, text="1. S\u00e9lectionnez un classeur Excel\n"
                       "2. Choisissez la feuille Canalisations et la feuille N\u0153uds\n"
                       "3. Mappez les colonnes pour chaque feuille\n"
                       "4. Cliquez sur \"Construire le mod\u00e8le\"",
            font=ctk.CTkFont(size=11), text_color="gray",
            justify="left",
        ).grid(row=1, column=0, padx=20, pady=(0, 12), sticky="w")

        # ── Fichier ──
        file_frame = ctk.CTkFrame(self)
        file_frame.grid(row=2, column=0, padx=20, pady=4, sticky="ew")
        file_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            file_frame, text="\U0001f4c2  Choisir classeur", width=140,
            command=self._select_file,
        ).grid(row=0, column=0, padx=(10, 8), pady=10, sticky="w")

        self.lbl_file = ctk.CTkLabel(
            file_frame, text="Aucun fichier s\u00e9lectionn\u00e9",
            text_color="gray", anchor="w",
        )
        self.lbl_file.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="ew")

        # ── Feuilles ──
        sheet_frame = ctk.CTkFrame(self)
        sheet_frame.grid(row=3, column=0, padx=20, pady=4, sticky="nsew")
        sheet_frame.grid_columnconfigure(1, weight=1)
        sheet_frame.grid_columnconfigure(3, weight=1)
        sheet_frame.grid_rowconfigure(4, weight=1)

        ctk.CTkLabel(
            sheet_frame, text="\U0001f4cb  S\u00e9lection des feuilles",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=0, columnspan=4, padx=12, pady=(12, 8), sticky="w")

        # Pipes
        ctk.CTkLabel(
            sheet_frame, text="Feuille Canalisations :",
            font=ctk.CTkFont(size=12),
        ).grid(row=1, column=0, padx=12, pady=6, sticky="w")

        self.sheet_pipe_var = ctk.StringVar(value="")
        self.opt_sheet_pipe = ctk.CTkOptionMenu(
            sheet_frame, variable=self.sheet_pipe_var,
            values=[], width=200, state="disabled",
        )
        self.opt_sheet_pipe.grid(row=1, column=1, padx=(0, 8), pady=6, sticky="w")

        self.btn_map_pipes = ctk.CTkButton(
            sheet_frame, text="\u2699  Mapper colonnes", width=140,
            state="disabled", command=self._map_pipe_columns,
        )
        self.btn_map_pipes.grid(row=1, column=2, padx=(0, 8), pady=6, sticky="w")

        self.lbl_pipe_status = ctk.CTkLabel(
            sheet_frame, text="", text_color="gray", font=ctk.CTkFont(size=10),
        )
        self.lbl_pipe_status.grid(row=1, column=3, padx=(0, 12), pady=6, sticky="w")

        # Nodes
        ctk.CTkLabel(
            sheet_frame, text="Feuille N\u0153uds :",
            font=ctk.CTkFont(size=12),
        ).grid(row=2, column=0, padx=12, pady=6, sticky="w")

        self.sheet_node_var = ctk.StringVar(value="")
        self.opt_sheet_node = ctk.CTkOptionMenu(
            sheet_frame, variable=self.sheet_node_var,
            values=[], width=200, state="disabled",
        )
        self.opt_sheet_node.grid(row=2, column=1, padx=(0, 8), pady=6, sticky="w")

        self.btn_map_nodes = ctk.CTkButton(
            sheet_frame, text="\u2699  Mapper colonnes", width=140,
            state="disabled", command=self._map_node_columns,
        )
        self.btn_map_nodes.grid(row=2, column=2, padx=(0, 8), pady=6, sticky="w")

        self.lbl_node_status = ctk.CTkLabel(
            sheet_frame, text="", text_color="gray", font=ctk.CTkFont(size=10),
        )
        self.lbl_node_status.grid(row=2, column=3, padx=(0, 12), pady=6, sticky="w")

        # Résumé
        self.txt_summary = ctk.CTkTextbox(
            sheet_frame, font=ctk.CTkFont(family="Consolas", size=11),
            state="disabled", height=120,
        )
        self.txt_summary.grid(row=4, column=0, columnspan=4, padx=12, pady=(8, 12), sticky="nsew")

        # ── Boutons ──
        btn_frame = ctk.CTkFrame(self)
        btn_frame.grid(row=4, column=0, padx=20, pady=(4, 16), sticky="ew")
        btn_frame.grid_columnconfigure(0, weight=1)

        self.btn_build = ctk.CTkButton(
            btn_frame, text="\U0001f3d7  Construire le mod\u00e8le",
            fg_color="#2d6a4f", hover_color="#1b4332",
            font=ctk.CTkFont(weight="bold"),
            state="disabled", command=self._build_model,
        )
        self.btn_build.grid(row=0, column=1, padx=6)

        self.btn_cancel = ctk.CTkButton(
            btn_frame, text="Fermer",
            fg_color="transparent", border_width=1,
            text_color=("gray10", "gray90"),
            command=self.destroy,
        )
        self.btn_cancel.grid(row=0, column=2, padx=6)

        self._pipe_map = {}
        self._node_map = {}

    # ── Actions ──────────────────────────────────────────────────

    def _select_file(self):
        filepath = filedialog.askopenfilename(
            title="S\u00e9lectionner le classeur Excel",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Tous", "*.*")]
        )
        if not filepath:
            return
        self.filepath = filepath
        self.lbl_file.configure(
            text=os.path.basename(filepath),
            text_color=("gray20", "gray80"),
        )
        self.sheet_names = self.builder.open_workbook(filepath)
        if not self.sheet_names:
            messagebox.showerror(
                "Erreur",
                "Impossible d'ouvrir le classeur. V\u00e9rifiez le format.",
            )
            return
        self.opt_sheet_pipe.configure(values=self.sheet_names, state="normal")
        self.opt_sheet_node.configure(values=self.sheet_names, state="normal")
        self.sheet_pipe_var.set(self.sheet_names[0] if self.sheet_names else "")
        self.sheet_node_var.set(self.sheet_names[1] if len(self.sheet_names) > 1 else (self.sheet_names[0] if self.sheet_names else ""))
        self._reset_mappings()

    def _reset_mappings(self):
        self._pipe_map = {}
        self._node_map = {}
        self.lbl_pipe_status.configure(text="")
        self.lbl_node_status.configure(text="")
        self.btn_map_pipes.configure(state="disabled")
        self.btn_map_nodes.configure(state="disabled")
        self.btn_build.configure(state="disabled")
        self.txt_summary.configure(state="normal")
        self.txt_summary.delete("1.0", "end")
        self.txt_summary.configure(state="disabled")

    def _load_pipe_sheet(self):
        sheet = self.sheet_pipe_var.get()
        if not sheet:
            return None
        df = self.builder.load_sheet(self.filepath, sheet)
        if df is None or len(df) == 0:
            messagebox.showerror("Erreur", f"Impossible de charger la feuille '{sheet}'.")
            return None
        return df

    def _load_node_sheet(self):
        sheet = self.sheet_node_var.get()
        if not sheet:
            return None
        df = self.builder.load_sheet(self.filepath, sheet)
        if df is None or len(df) == 0:
            messagebox.showerror("Erreur", f"Impossible de charger la feuille '{sheet}'.")
            return None
        return df

    def _map_columns_for(self, standard_cols, mandatory_keys, df, file_type):
        sheet_cols = df.columns.tolist()
        file_hash = file_hash_from_path(self.filepath)
        mapper = get_mapper()
        mapper.set_ui_callback(
            lambda unknown_col, available_cols, ft, fh:
                ask_column_mapping(self, unknown_col, available_cols, ft, fh)
        )

        result_map = {}
        for key, display_name in standard_cols.items():
            chosen = mapper.request_mapping(
                unknown_col=display_name,
                available_cols=sheet_cols,
                file_type=file_type,
                file_hash=file_hash,
            )
            if chosen and isinstance(chosen, str) and chosen not in ("__SKIP__", "__CANCEL__"):
                result_map[key] = chosen
            elif key in mandatory_keys:
                # Obligatoire mais pas mappé → essayer auto-détection
                found = [c for c in sheet_cols if key.lower() in c.lower()]
                if found:
                    result_map[key] = found[0]

        return result_map

    def _map_pipe_columns(self):
        df = self._load_pipe_sheet()
        if df is None:
            return
        self._pipe_map = self._map_columns_for(
            PIPE_COLUMNS, PIPE_MANDATORY, df, FILE_TYPE_MODEL_PIPES,
        )
        ok = all(k in self._pipe_map for k in PIPE_MANDATORY)
        self.lbl_pipe_status.configure(
            text=f"\u2705 {len(self._pipe_map)}/{len(PIPE_COLUMNS)} colonnes" if ok
                 else f"\u26a0  Colonnes obligatoires manquantes",
            text_color="#33a02c" if ok else "orange",
        )

    def _map_node_columns(self):
        df = self._load_node_sheet()
        if df is None:
            return
        self._node_map = self._map_columns_for(
            NODE_COLUMNS, NODE_MANDATORY, df, FILE_TYPE_MODEL_NODES,
        )
        ok = all(k in self._node_map for k in NODE_MANDATORY)
        self.lbl_node_status.configure(
            text=f"\u2705 {len(self._node_map)}/{len(NODE_COLUMNS)} colonnes" if ok
                 else f"\u26a0  Colonnes obligatoires manquantes",
            text_color="#33a02c" if ok else "orange",
        )
        self.btn_build.configure(state="normal" if ok else "disabled")

    def _build_model(self):
        pipe_df = self._load_pipe_sheet()
        node_df = self._load_node_sheet()
        if pipe_df is None or node_df is None:
            return

        self.builder.data = self.builder.build_model(
            pipe_df, node_df, self._pipe_map, self._node_map,
        )
        data = self.builder.data

        self.txt_summary.configure(state="normal")
        self.txt_summary.delete("1.0", "end")

        if data.errors:
            for e in data.errors:
                self.txt_summary.insert("end", f"\u274c  {e}\n")
            self.txt_summary.configure(state="disabled")
            messagebox.showerror("Erreur", "\n".join(data.errors[:3]))
            return

        summary = data.get_summary()
        lines = [
            f"\u2705  Mod\u00e8le construit avec succ\u00e8s !\n",
            f"    Canalisations : {summary['pipes_count']}",
            f"    N\u0153uds :         {summary['nodes_count']}",
            f"    Longueur totale : {summary['total_length_m']} m",
        ]
        if summary['diameters']:
            dn_str = ", ".join(f"DN{int(d)}" for d in summary['diameters'])
            lines.append(f"    Diam\u00e8tres :     {dn_str}")
        if summary['materials']:
            lines.append(f"    Mat\u00e9riaux :      {', '.join(summary['materials'])}")
        if summary['unlinked_nodes']:
            ul = summary['unlinked_nodes'][:10]
            lines.append(f"\n\u26a0  N\u0153uds r\u00e9f\u00e9renc\u00e9s sans position : {', '.join(ul)}")
            if len(summary['unlinked_nodes']) > 10:
                lines[-1] += f" ... (+{len(summary['unlinked_nodes']) - 10})"
        if data.warnings:
            lines.append(f"\n\u26a0  {len(data.warnings)} avertissement(s)")

        self.txt_summary.insert("end", "\n".join(lines))
        self.txt_summary.configure(state="disabled")

        self.btn_build.configure(
            text="\u2705  Mod\u00e8le construit", fg_color="#33a02c",
            state="disabled",
        )
        self.result = data
        messagebox.showinfo(
            "Mod\u00e8le construit",
            f"{summary['pipes_count']} canalisations, {summary['nodes_count']} n\u0153uds\n"
            f"Longueur totale : {summary['total_length_m']} m",
        )
