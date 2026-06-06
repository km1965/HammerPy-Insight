# HammerPy Insight ⚡  — v3.0

> Application de bureau Python pour l'interprétation automatisée des résultats transitoires exportés depuis **Bentley HAMMER** (Coup de Bélier / Water Hammer Analysis).

---

## 🎯 Objectif

Permettre aux ingénieurs hydrauliques de :
1. Charger rapidement les résultats de simulation HAMMER (CSV/Excel **et classeur Flex Tables .xlsx/.xls**)
2. Configurer les paramètres réglementaires de l'étude (Classe PN, pression min de sécurité)
3. Visualiser interactivement les courbes enveloppes de pressions et volumes de gaz
4. Consulter le **résumé du modèle hydraulique** (6 feuilles : Pipes, Nœuds, Pompes, Réservoirs, HPT, Ventouses)
5. **Analyser les courbes H(Q) de pompes** (multi-pompes en parallèle ou continu)
6. **Dimensionner les ventaises et vidanges** sur le profil en long de la conduite
7. Générer en un clic une **note technique structurée au format Word (.docx)** prête à intégrer au dossier d'étude

---

## 🚀 Fonctionnalités

### Interface Utilisateur
- Design moderne **CustomTkinter** — thème **sombre par défaut** (basculable vers Clair/Système)
- **Icône personnalisée** `hammerpy_icon.ico` (éclair bleu/jaune, 6 tailles)
- **Sidebar de configuration** : Classe PN (PN6 → PN40) et pression minimale admissible
- **Sidebar — Unités d'affichage** : sélecteur d'unités pour le **débit** (m³/h ↔ L/s) et le **volume** (L ↔ m³), avec **seuil HPT éditable** (suit l'unité de volume choisie)
- **Barre de menu supérieure** : boutons **Ouvrir · Enregistrer · Enregistrer sous · Quitter** avec indicateur d'état (modifications non sauvegardées)
- Fenêtre responsive (1200×750 min)

### 💾 Gestion de projet (`.hpi` v3.0)
- **Sauvegarde complète** de la session dans un fichier unique `.hpi` (JSON lisible) :
  - Métadonnées (nom de projet, ingénieur, dates)
  - Configuration (PN, pression min admissible, unités, seuil HPT)
  - Données station + données HPT **embarquées**
  - **6 feuilles du classeur HAMMER** (Pipes, Nœuds, Réservoirs, Pompes, HPT, Ventouses) en `orient='records'` optimisé
  - **Batterie de pompes** : liste de rapports détaillés parsés + mode Continu/Parallèle
  - **Profil en long + ventaises + vidanges** (Phase 3)
  - **Diagnostic système** (Phase 4) : résultats 16 checks + résumé
  - Texte complet de la prévisualisation du rapport
  - Les graphiques Matplotlib sont **automatiquement régénérés** à l'ouverture
- **Rétrocompatibilité** : les `.hpi` v2.x restent lisibles par v3.0
- **Détection des modifications** : indicateur visuel orange dès qu'un changement est effectué
- **Protection contre la perte de données** : confirmation avant d'écraser un projet modifié (Ouverture / Quitter)

### 🔄 Système d'unités
- **Choix des unités d'affichage** (sidebar) — s'applique instantanément aux KPI, au graphique (axes + seuil), au rapport texte et à l'export Word
  - **Débit** : m³/h ou L/s
  - **Volume** : L ou m³
- **Auto-détection de l'unité des fichiers CSV source** : scan multi-feuilles Excel avec priorité à la feuille "Pompes"
- **Override manuel** via menu déroulant "Unité :" à côté de chaque bouton d'import
- **Seuil HPT éditable** dans la sidebar (suit l'unité de volume sélectionnée, ex. `200` en L ou `0.200` en m³)
- **Stockage canonique interne** (L, m³/h) → conversions sans perte de précision

### Onglet 1 — Régime Permanent
- Champs de saisie : **Nom du projet** et **Ingénieur responsable**
- Importation du fichier station (CSV/Excel multi-feuilles)
- Extraction automatique du **Débit nominal Q (m³/h)** et de la **HMT (mCE)** (priorité feuille "Pompes")
- Affichage des colonnes détectées pour diagnostic

### Onglet 2 — Analyse Transitoire
- Importation du fichier d'enveloppe HPT Bentley HAMMER
- **Indicateurs KPI colorés** (Pression Min, Pression Max, Volume Gaz)
- **Section « Modèle HAMMER »** :
  - Bouton **Charger classeur HAMMER** (.xlsx/.xls) — lit le fichier Flex Tables complet
  - **6 compteurs** de feuilles : Pipes, Nœuds, Pompes, Réservoirs, HPT, Ventouses
  - **4 mini-stats** : P max, P min, Q max pompe, Matériaux détectés
  - Bouton **Réinitialiser** pour vider le classeur chargé
  - Validation stricte : refuse si Pipes, Nœuds ou Pompes manquants
- **Section « Batterie de Pompes » (multi-pompes)** :
  - Sélecteur de **mode** : Continu ou Parallèle
  - **Liste des pompes** avec boutons +/retirer
  - KPI de la pompe sélectionnée (Label, Q, HMT, NPSH, nb points)
  - **Saisie manuelle des points de courbe** : champs Q (L/s) + H (m) + boutons Ajouter / Effacer
  - **Graphique Matplotlib H(Q)** : courbe interpolée (polyfit numpy) + courbe combinée en mode parallèle
  - Adaptation automatique au thème clair/sombre
- **Graphique HPT double-subplot interactif** :
  - Courbes Pression Min/Max le long de la conduite
  - **Seuils critiques tracés dynamiquement** (ligne PN, ligne P min sécurité)
  - **Zones de dépassement surlignées** automatiquement
  - Volume de Gaz au HPT avec seuil 200 L
  - Barre d'outils Matplotlib intégrée (zoom, pan, sauvegarde PNG)
- Adaptation instantanée des couleurs au changement de thème

### Onglet 3 — Rapport Technique
- **Prévisualisation éditable** de la note complète
- Diagnostic de sécurité automatique (surpression, dépression, volume HPT)
- **Section « Modèle Hydraulique »** dans le rapport Word :
  - Tableau récapitulatif (6 composants du réseau)
  - Matériaux et diamètres des conduites
  - Pressions transitoires du modèle avec alertes conformité
- **Section « Batterie de Pompes »** dans le rapport Word :
  - Tableau récapitulatif par pompe (ID, label, débit, HMT, NPSH dispo.)
  - Alerte NPSH si dispo. < requis
- **Section « Profil en Long »** dans le rapport Word :
  - Table ventaises (vert) + Table vidanges (rouge)
- **Export .txt** (note brute)
- **Export Word (.docx)** professionnel structuré

### Onglet 4 — Ventaises & Vidanges (Phase 3)
- **Imports multi-format** :
  - **CSV Libre** (PK, Z) — format simple
  - **CSV Bentley FlexTable** (Label, X, Y, Elevation) — export HAMMER
    - Distance cumulée calculée par cumul de distances successives (ordre amont→aval)
  - **DXF AutoCAD** — 1 fichier avec 2 calques LWPOLYLINE :
    - Calque `Tracé en plan` (vue XY horizontale)
    - Calque `Profil en long` (X = PK, Y = altitude)
    - Auto-détection accents/casse (`Tracé en plan`, `Profil en long`, `Plan`, `Profile`…)
  - **Profil exemple** démo intégré
- **Encodages supportés** : UTF-8, UTF-8-sig, UTF-16 LE, UTF-16 BE, CP1252, Latin-1
- **Mapping interactif des colonnes** (Phase 3.6) :
  - Si une colonne attendue (PK, Z) n'est pas trouvée dans le CSV,
    une boîte de dialogue modale demande à l'utilisateur quelle colonne utiliser
  - **Auto-apprentissage** : les mappings sont mémorisés (jamais redemandés)
  - Persistance dans `.hpi` (rétrocompatible v3.0)
- **Saisie du DN conduite** (mm)
- **Calcul automatique** des ventaises et vidanges :
  - Détection des points hauts/bas du profil
  - Pré-dimensionnement ventaises (anti-vide, combinée, grande orifice)
  - Localisation vidanges aux points bas entre 2 ventaises
- **2 graphiques Matplotlib** :
  - Profil en long (PK × Z) avec marqueurs ▲ ventaises / ▼ vidanges
  - Tracé en plan (X × Y) si DXF chargé
- **Tableaux récapitulatifs** : PK, côte, type, DN, distances
- **Export CSV** des recommandations
- **Export Rapport Word (.docx)** dédié :
  - En-tête projet (nom, ingénieur, date, DN)
  - Image du profil en long avec marqueurs
  - Tableau ventouses + synthèse (DN min/max, types)
  - Tableau vidanges + distances aux ventouses
  - Section méthodologie & hypothèses

### Onglet 5 — Système & Diagnostics (Phase 4)
- **16 vérifications croisées** en 5 catégories (A-E) :
  - **A. Pompe ↔ Réseau** (3) : point de fonctionnement dans la courbe H(Q), NPSH, vitesse spécifique
  - **B. Pompe ↔ HPT** (2) : pression refoulement/aspiration vs PN/Pmin
  - **C. Réseau ↔ HPT** (3) : Pmax/Pmin transitoires, cohérence multi-pompes
  - **D. HPT ↔ Ventouses / Vidanges** (3) : volume gaz HPT, ventouses aux points hauts, vidanges aux points bas
  - **E. Cohérence globale** (5) : profil chargé, DN cohérent, cote min positive, pente max, ≥1 pompe
- **Sévérités** : ✔ OK / ⚠ WARN / ✘ FAIL / — N/A
- **Bandeau KPI** : 4 compteurs (OK / WARN / FAIL / NA) en temps réel
- **Tableau plat triable** (TTK Treeview) : code, catégorie, statut, nom, détail
- **Bouton "Lancer le diagnostic"** : exécute les 16 checks en cascade
- **Persistance `.hpi`** : résultats sauvegardés, rechargés à l'ouverture
- **Intégration rapport Word** : section 6 « Diagnostic Système » avec 6.1 Synthèse + 6.2 Détail par catégorie

---

## 🔧 Parser — Formats Supportés

### Fichiers CSV/Excel (import individuel)

| Format | Séparateurs | Décimales | Encodages |
|--------|-------------|-----------|-----------|
| `.csv` | `,`  `;`  `Tab` | `.` ou `,` | UTF-8, Latin-1, CP1252 |
| `.xlsx` | — | — | — |
| `.xls` | — | — | — |

### Classeur HAMMER — Flex Tables (import complet)

| Feuille | Contenu | Obligatoire |
|---------|---------|-------------|
| **Pipes** | Conduites : longueurs, diamètres, matériaux, HW C, V/P transitoires | ✅ Oui |
| **Noeuds** | Nœuds : élévation, pressions, HGL, coordonnées X/Y | ✅ Oui |
| **Pumps** | Pompes : débit, HMT, pressions transitoires | ✅ Oui |
| **Reservoirs** | Réservoirs à niveau variable | ⬜ Optionnel |
| **HPT** | Réservoirs anti-bélier (volumes, pressions gaz) | ⬜ Optionnel |
| **Air valves** | Ventouses : type, diamètres orifices, volumes air | ⬜ Optionnel |

### Rapport Pompe Détaillé (import pour courbe H(Q))

| Format | Contenu |
|--------|---------|
| `.rtf` | Rapport HAMMER « Pump Detailed Report » — extraction de 13 champs (ID, label, débit, HMT, NPSH, etc.) |
| `.txt` | Format texte brut du même rapport |

**Champs extraits du rapport pompe :**
- `pump_id`, `label`, `downstream_pipe`, `speed_factor`, `status_initial`
- `flow_lps`, `pump_head_m`, `pressure_suction_bar`, `pressure_discharge_bar`
- `npsh_available_m`, `npsh_required_m`, `controlled`, `hydraulic_grade_*`

### Profil en Long (import pour ventaises/vidanges)

| Format | Contenu |
|--------|---------|
| `.csv` | 3 colonnes : PK (m), Z (m), [pente %] — séparateur auto-détecté |

---

## 🛠️ Installation

```bash
# 1. Créer et activer l'environnement virtuel
python -m venv env
env\Scripts\activate   # Windows

# 2. Installer les dépendances
pip install -r requirements.txt
```

**Dépendances principales :**
`customtkinter` · `pandas` · `openpyxl` · `matplotlib` · `numpy` · `python-docx` · `pillow`

---

## ▶️ Lancement

```bash
python main.py
```

---

## 📁 Fichiers du Projet

| Fichier | Description |
|---------|-------------|
| `main.py` | Script principal — GUI + entry point (~2600 lignes) |
| `utils.py` | Fonctions utilitaires : parse_number, constantes PN/Unités |
| `data_parser.py` | Parser CSV/Excel multi-feuilles (station + HPT) |
| `workbook.py` | WorkbookManager — chargement classeur HAMMER |
| `pump_parser.py` | PumpReportParser — extraction données pompe depuis RTF |
| `report_generator.py` | WordReportGenerator — génération rapport Word |
| `air_valve_sizing.py` | AirValveSizing — dimensionnement ventaises/vidanges |
| `hammerpy_icon.ico` | Icône de l'application |
| `test_workbook_parser.py` | Tests unitaires (65 tests) |
| `requirements.txt` | Dépendances Python |
| `ROADMAP.md` | Feuille de route v3.0 |

---

## 🗂️ Architecture du Code

```
main.py                          # GUI + entry point
├── HammerPyApp (ctk.CTk)        # Interface graphique principale
│   ├── _create_top_bar()        # Barre de menu Ouvrir/Enregistrer/Quitter
│   ├── _create_sidebar()        # Configuration PN, thème, aide
│   ├── Onglet 1                 # Régime permanent
│   ├── Onglet 2                 # Analyse transitoire + Matplotlib + Multi-pompes
│   ├── Onglet 3                 # Rapport et exports
│   ├── Onglet 4                 # Ventaises & Vidanges (Phase 3)
│   └── Onglet 5                 # Système & Diagnostics (Phase 4)
│
├── utils.py                     # Fonctions utilitaires partagées
│   ├── parse_number()           # Parsing numérique HAMMER (nbsp, virgule FR)
│   ├── find_col_in_df()         # Recherche de colonnes par mots-clés
│   ├── PN_CLASSES, PMIN_OPTIONS # Constantes de pression
│   └── FLOW_UNITS, VOLUME_UNITS # Constantes d'unités
│
├── data_parser.py               # Parser CSV/Excel robuste
│   ├── HammerDataParser         # Parsing station + HPT
│   ├── _load_file()             # Chargement universel (encodages, séparateurs)
│   ├── parse_station_file()     # Régime permanent multi-feuilles (Q, HMT)
│   └── parse_hpt_file()         # Transitoire (Pmin, Pmax, Vol.Gaz)
│
├── workbook.py                  # WorkbookManager — Classeur HAMMER
│   ├── load()                   # Détection format, chargement 6 feuilles
│   ├── validate()               # Vérifie les 3 feuilles obligatoires
│   └── get_summary()            # Résumé : counts, Pmax/Pmin, matériaux
│
├── pump_parser.py               # Parser rapport pompe HAMMER
│   ├── PumpReportParser         # Extraction 13 champs depuis RTF/TXT
│   ├── _strip_rtf()             # Suppression balises RTF, images
│   ├── add_curve_point()        # Ajout point Q/H avec tri automatique
│   └── get_summary()            # Résumé pompe (label, Q, H, NPSH)
│
├── report_generator.py          # Génération rapport Word (.docx)
│   ├── WordReportGenerator      # Rapport complet multi-sections
│   └── generate()               # Sections: Modèle + Pompe + Ventaises/vidanges
│
├── dxf_profile_importer.py       # Import profil + tracé depuis DXF (ezdxf)
│   ├── load_dxf_both()          # Charge plan + profil d'un DXF
│   ├── load_dxf_plan()          # Extrait polyligne calque "Tracé en plan"
│   ├── load_dxf_profile()       # Extrait polyligne calque "Profil en long"
│   └── list_dxf_layers()        # Liste calques LWPOLYLINE disponibles
│
├── column_mapper.py               # Mapping interactif colonnes CSV/XLSX
│   ├── ColumnMapper             # Auto-apprentissage + UI callback
│   ├── request_mapping()        # Demande utilisateur via modale
│   ├── auto_apply()             # Renommage auto de plusieurs colonnes
│   ├── learn_mapping()          # Apprentissage explicite
│   └── serialize()/deserialize() # Persistance dans .hpi
│
├── column_mapper_dialog.py        # Boîte de dialogue modale (CTk)
│   └── ask_column_mapping()     # Dropdown + OK/Skip/Cancel
│
├── ventouses_report.py            # Rapport Word (.docx) dédié ventouses
│   ├── VentousesReportGenerator  # Sections : en-tête, profil, ventouses, vidanges
│   └── export_ventouses_report() # Helper tout-en-un (sizer + metadata + PNG)
│
└── air_valve_sizing.py          # Dimensionnement ventaises/vidanges
    ├── AirValveSizing           # Calcul points hauts/bas + sizing
    ├── load_profile_csv()       # Import profil en long (CSV libre)
    ├── load_profile_bentley_csv() # Import CSV FlexTable Bentley (X,Y,Z)
    ├── load_profile_manual()    # Saisie manuelle points (PK, Z)
    ├── size_ventaises()         # Pré-dimensionnement ventaises
    ├── size_drains()            # Localisation vidanges
    └── export_csv()             # Export recommandations
```

---

## 🧪 Tests

```bash
# Lancer tous les tests
python -m pytest test_workbook_parser.py -v

# Résultat attendu : 65 passed
```

**Couverture des tests :**
- 19 tests `_parse_number()` : natif, None/NaN, nbsp, virgule FR, négatifs
- 4 tests `_find_col_in_df()` : match exact, partiel, casse
- 17 tests intégration workbook : chargement, validation, résumé
- 3 tests erreurs : extension invalide, fichier absent, feuille manquante
- 4 tests rétrocompatibilité : CSV HPT + station lisibles
- 12 tests `PumpReportParser` : chargement RTF réel, strip RTF, courbe points, interpolation, résumé
- 7 tests `AirValveSizing` : profil, points hauts/bas, sizing ventaises/vidanges, DN, export CSV
- 14 tests DXF & CSV Bentley : normalisation calques, extraction LWPOLYLINE, parsing FlexTable, distance cumulée, pentes, détection points hauts/bas, encodages UTF-16
- 24 tests ColumnMapper : auto-apprentissage, cache, UI callback, skip/cancel, sérialisation, hash fichier
- 14 tests ventouses_report : génération, sauvegarde, intégration image PNG, contenu, stats

---

*Document mis à jour — HammerPy Insight v3.0 Phase 3 + Phase 3.5 Imports multi-format + Phase 3.6 Mapping interactif + Rapport Ventouses + Phase 4 SystemDiagnostics — Juin 2026*
