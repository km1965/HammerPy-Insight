# HammerPy Insight ⚡  — v3.0

> Application de bureau Python pour l'interprétation automatisée des résultats transitoires exportés depuis **Bentley HAMMER** (Coup de Bélier / Water Hammer Analysis).

---

## 🎯 Objectif

Permettre aux ingénieurs hydrauliques de :
1. Charger rapidement les résultats de simulation HAMMER (CSV/Excel **et classeur Flex Tables .xlsx/.xls**)
2. Configurer les paramètres réglementaires de l'étude (Classe PN, pression min de sécurité)
3. Visualiser interactivement les courbes enveloppes de pressions et volumes de gaz
4. Consulter le **résumé du modèle hydraulique** (6 feuilles : Pipes, Nœuds, Pompes, Réservoirs, HPT, Ventouses)
5. Générer en un clic une **note technique structurée au format Word (.docx)** prête à intégrer au dossier d'étude

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
  - Texte complet de la prévisualisation du rapport
  - Le graphique Matplotlib est **automatiquement régénéré** à l'ouverture
- **Rétrocompatibilité** : les `.hpi` v2.x restent lisibles par v3.0
- **Détection des modifications** : indicateur visuel orange dès qu'un changement est effectué
- **Protection contre la perte de données** : confirmation avant d'écraser un projet modifié (Ouverture / Quitter)

### 🔄 Système d'unités
- **Choix des unités d'affichage** (sidebar) — s'applique instantanément aux KPI, au graphique (axes + seuil), au rapport texte et à l'export Word
  - **Débit** : m³/h ou L/s
  - **Volume** : L ou m³
- **Auto-détection de l'unité des fichiers CSV source** : scan des noms de colonnes (ex. `Flow (m3/h)`, `Volume of Gas (L)`)
- **Override manuel** via menu déroulant "Unité :" à côté de chaque bouton d'import
- **Seuil HPT éditable** dans la sidebar (suit l'unité de volume sélectionnée, ex. `200` en L ou `0.200` en m³)
- **Stockage canonique interne** (L, m³/h) → conversions sans perte de précision

### Onglet 1 — Régime Permanent
- Champs de saisie : **Nom du projet** et **Ingénieur responsable**
- Importation du fichier station (CSV/Excel)
- Extraction automatique du **Débit nominal Q (m³/h)** et de la **HMT (mCE)**
- Affichage des colonnes détectées pour diagnostic

### Onglet 2 — Analyse Transitoire
- Importation du fichier d'enveloppe HPT Bentley HAMMER
- **Indicateurs KPI colorés** (Pression Min, Pression Max, Volume Gaz)
- **Nouveau : Section « Modèle HAMMER »** :
  - Bouton **Charger classeur HAMMER** (.xlsx/.xls) — lit le fichier Flex Tables complet
  - **6 compteurs** de feuilles : Pipes, Nœuds, Pompes, Réservoirs, HPT, Ventouses
  - **4 mini-stats** : P max, P min, Q max pompe, Matériaux détectés
  - Bouton **Réinitialiser** pour vider le classeur chargé
  - Validation stricte : refuse si Pipes, Nœuds ou Pompes manquants
- **Graphique double-subplot interactif Matplotlib** :
  - Courbes Pression Min/Max le long de la conduite
  - **Seuils critiques tracés dynamiquement** (ligne PN, ligne P min sécurité)
  - **Zones de dépassement surlignées** automatiquement
  - Volume de Gaz au HPT avec seuil 200 L
  - Barre d'outils Matplotlib intégrée (zoom, pan, sauvegarde PNG)
- Adaptation instantanée des couleurs au changement de thème

### Onglet 3 — Rapport Technique
- **Prévisualisation éditable** de la note complète
- Diagnostic de sécurité automatique (surpression, dépression, volume HPT)
- **Nouvelle section « Modèle Hydraulique »** dans le rapport Word :
  - Tableau récapitulatif (6 composants du réseau)
  - Matériaux et diamètres des conduites
  - Pressions transitoires du modèle avec alertes conformité
- **Export .txt** (note brute)
- **Export Word (.docx)** professionnel structuré :
  - En-tête centré
  - Tableau de métadonnées du projet
  - Section Modèle Hydraulique (nouveau)
  - Tableau de résultats avec codes couleurs vert/rouge (OK / Dépassement)
  - Interprétation et recommandations par section (Surpression, Dépression, Volume HPT)
  - **Graphique Matplotlib intégré automatiquement** dans le document

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

**Colonnes HAMMER reconnues automatiquement** (anglais et français) :
- `Pressure (Minimum)` / `Pression Min` / `P Min`
- `Pressure (Maximum)` / `Pression Max` / `P Max`
- `Volume of Gas (Maximum)` / `Volume Gaz` / `Air Volume`
- `Distance` / `Abscisse` / `Chainage` / `Station`
- `Start Node` / `Stop Node` / `Diameter` / `Material`

**Gestion des formats numériques HAMMER** :
- Espace insécable `\xa0` comme séparateur de milliers (`1\xa0029,00` → 1029.00)
- Virgule décimale française (`12,5` → 12.5)

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
`customtkinter` · `pandas` · `openpyxl` · `matplotlib` · `python-docx` · `pillow`

---

## ▶️ Lancement

```bash
python main.py
```

---

## 📁 Fichiers du Projet

| Fichier | Description |
|---------|-------------|
| `main.py` | Script principal — application complète |
| `hammerpy_icon.ico` | Icône de l'application (éclair HammerPy, 6 tailles) |
| `Flex Tables.xlsx` | Classeur HAMMER d'exemple (6 feuilles) |
| `station_steady_state_test.csv` | Débit nominal et HMT d'une station AEP fictive |
| `hpt_transient_test.csv` | Enveloppes de pression et volume de gaz |
| `test_workbook_parser.py` | Tests unitaires (48 tests) |
| `requirements.txt` | Dépendances Python |
| `ROADMAP.md` | Feuille de route v3.0 (Pompes & Ventouses) |

---

## 🗂️ Architecture du Code

```
main.py
├── _parse_number()                # Helper : parsing numérique HAMMER (nbsp, virgule FR)
├── _find_col_in_df()              # Helper : recherche de colonnes par mots-clés
│
├── WorkbookManager                # Chargement classeur HAMMER (.xlsx/.xls)
│   ├── load()                     # Détection format, chargement 6 feuilles
│   ├── validate()                 # Vérifie les 3 feuilles obligatoires
│   ├── get_summary()              # Résumé : counts, Pmax/Pmin, matériaux, diamètres
│   └── get_sheet()                # Accès à une feuille par nom canonique
│
├── HammerDataParser               # Logique de parsing (CSV/Excel, robuste et flexible)
│   ├── _load_file()               # Chargement universel avec gestion encodages
│   ├── _find_col()                # Recherche flexible de colonnes par mots-clés
│   ├── parse_station_file()       # Régime permanent (Q, HMT)
│   └── parse_hpt_file()           # Transitoire (Pmin, Pmax, Vol.Gaz)
│
├── WordReportGenerator            # Génération du rapport Word (.docx)
│   └── generate()                 # Rapport complet avec section Modèle Hydraulique
│
└── HammerPyApp (ctk.CTk)          # Interface graphique principale
    ├── _create_top_bar()          # Barre de menu Ouvrir/Enregistrer/Enregistrer sous/Quitter
    ├── _create_sidebar()          # Configuration PN, thème, aide
    ├── Onglet 1                   # Régime permanent
    ├── Onglet 2                   # Analyse transitoire + Matplotlib + Modèle HAMMER
    │   ├── _import_workbook()     # Chargement classeur Flex Tables
    │   ├── _reset_workbook()      # Réinitialisation du classeur
    │   └── _update_chart()        # Tracé dynamique avec seuils configurables
    ├── Onglet 3                   # Rapport et exports
    ├── Gestion de projet          # Sauvegarde/chargement .hpi v3.0 (JSON)
    │   ├── _open_project()
    │   ├── _save_project_as()
    │   ├── _save_project_quick()
    │   ├── _load_project()        # Restauration complète + régénération graphique
    │   └── _quit_app()            # Confirmation si modifications non sauvegardées
    ├── _export_txt()              # Export note brute
    └── _export_word()             # Export Word professionnel avec section Modèle
```

---

## 🧪 Tests

```bash
# Lancer tous les tests
python -m pytest test_workbook_parser.py -v

# Résultat attendu : 48 passed
```

**Couverture des tests :**
- 19 tests `_parse_number()` : natif, None/NaN, nbsp, virgule FR, négatifs
- 4 tests `_find_col_in_df()` : match exact, partiel, casse
- 17 tests intégration workbook : chargement, validation, résumé
- 3 tests erreurs : extension invalide, fichier absent, feuille manquante
- 4 tests rétrocompatibilité : CSV HPT + station lisibles

---

*Document mis à jour automatiquement — HammerPy Insight v3.0 — Juin 2026*
