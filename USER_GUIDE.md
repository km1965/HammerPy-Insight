# Guide Utilisateur — HammerPy Insight

## 1. Introduction

HammerPy Insight est un outil d'automatisation et d'interprétation des résultats transitoires exportés depuis Bentley HAMMER. Il permet de :

- Charger et visualiser les enveloppes transitoires (Pmin, Pmax, volume gaz)
- Analyser les régimes permanents des stations hydrauliques
- Importer et exploiter les rapports détaillés des pompes
- Dimensionner les ventouses et vidanges sur le profil en long
- Générer un rapport Word complet avec diagnostics croisés

---

## 2. Interface Principale

### 2.1 Barre latérale (Paramètres)

- **Projet / Ingénieur** : informations d'entête du rapport
- **Classe PN** : pression nominale de la conduite (PN 6 à PN 40)
- **P min admissible** : seuil de dépression (0 à -1.0 bar)
- **Unités** : choix L/s ou m³/h pour les débits, L ou m³ pour les volumes

### 2.2 Onglets

| Onglet | Fonction |
|--------|----------|
| **Transitoire** | Graphique enveloppe HPT + KPIs + import fichier transitoire |
| **Station** | Régime permanent (débit, HMT) |
| **Pompes** | Import rapports RTF, courbes H(Q), données Nq/NPSH |
| **Ventouses** | Profil en long, dimensionnement ventouses/vidanges, export DXF |
| **Classeur** | Chargement classeur HAMMER FlexTables |
| **Diagnostic** | 16 vérifications croisées A–E |
| **Rapport** | Génération du rapport Word final |

---

## 3. Import des Données

### 3.1 Fichier Transitoire HPT

Charge l'enveloppe des résultats transitoires (Pmin, Pmax, Volume gaz).

**Formats supportés** : CSV, XLSX

**Colonnes attendues** (détection automatique) :
- `Volume of gas (L)`, `Volume gaz`, `Gas Volume`...
- `Pressure (Minimum) (bar)`, `Pmin`, `Pression min`...
- `Pressure (Maximum) (bar)`, `Pmax`, `Pression max`...

**Template** : générez `template_hpt_enveloppe.xlsx` via le bouton `📄 Générer templates XLSX`.

### 3.2 Régime Permanent (Station)

Charge les données de débit et hauteur de la station.

**Colonnes attendues** : `Flow (m3/h)` ou `Débit` + `HMT (m)` ou `Head`.

### 3.3 Rapports Pompe (RTF)

Importez les rapports détaillés pompe exportés depuis Bentley HAMMER (format `.rtf` ou `.txt`). L'application extrait automatiquement :

- Débit nominal (Flow)
- Hauteur manométrique (Head)
- NPSH requis et disponible
- Vitesse spécifique Nq
- Pressions aspiration/refoulement

### 3.4 Données Pompe (CSV/XLSX)

Enrichit les données des pompes avec Nq, NPSH, et points de courbe H(Q).

**Colonnes attendues** : `pump_id`, `nq_si`, `npsh_required_m`, `npsh_available_m`

**Pour la courbe H(Q)** : `pump_id`, `flow_lps`, `head_m` (plusieurs lignes par pompe).

### 3.5 Profil en Long (Ventouses)

**Formats supportés** : CSV, XLSX, DXF

- CSV/XLSX : colonnes `PK (m)`, `Z (m)`, `Pente (%)` (optionnel)
- DXF : calque "Profil en long" (LWPOLYLINE ou POLYLINE)
- Bentley CSV : format FlexTable Junction Table

### 3.6 Model Builder (Import réseau libre)

Importe un classeur Excel quelconque avec mapping interactif des colonnes
(type constructeur de modèle HAMMER).

**Étapes** :
1. Cliquez sur `🧱 Model Builder (Excel libre)` dans la section Classeur HAMMER
2. Sélectionnez le fichier Excel (`.xlsx` / `.xls`)
3. Choisissez la **feuille Canalisations** et la **feuille Nœuds**
4. Pour chaque feuille, cliquez sur `⚙ Mapper colonnes` :
   - **Canalisations** (5 obligatoires) : Label, Nœud départ, Nœud arrivée, Longueur, Diamètre
   - **Canalisations** (optionnel) : Matériau, Rugosité (C), Célérité (m/s)
   - **Nœuds** (3 obligatoires) : Label, X (m), Y (m)
   - **Nœuds** (optionnel) : Z (m)
5. Cliquez `🏗 Construire le modèle`

Le modèle construit :
- Positionne les nœuds par leurs coordonnées XY
- Relie les canalisations aux nœuds (départ/arrivée)
- Détecte les nœuds sans position (non reliés)
- Calcule la longueur totale, les diamètres et matériaux

### 3.7 Classeur HAMMER

Charge un classeur FlexTables HAMMER (.xlsx) avec 6 feuilles possibles :

| Feuille | Contenu |
|---------|---------|
| `Pipes` (Conduites) | Diamètres, matériaux, longueurs, pressions transitoires |
| `Nodes` (Nœuds) | Élévations, pressions transitoires |
| `Pumps` (Pompes) | Débits, hauteurs |
| `Reservoirs` | Élévations |
| `HPT` | Volumes de gaz |
| `Air Valves` | Volumes d'air |

---

## 4. Dimensionnement Ventouses & Vidanges

### 4.1 Chargement du profil

1. Importez un profil via CSV/XLSX (bouton "Profil") ou DXF (bouton "Importer DXF")
2. Le profil s'affiche dans le graphique supérieur

### 4.2 Calcul

1. Saisissez le **DN conduite** (mm)
2. Cliquez **"Calculer ventouses + vidanges"**
3. L'application détecte automatiquement les points hauts (ventouses) et points bas (vidanges)

### 4.3 Règles de dimensionnement

| Type de ventouse | Condition | DN min |
|-----------------|-----------|--------|
| Simple (anti-vide) | Pente > 3% | DN_conduite / 12 |
| Combinée | Pente < 0.5% | DN_conduite / 10 |
| Grande orifice | 0.5% ≤ pente ≤ 3% | DN_conduite / 8 |

**Vidanges** : placées aux points bas entre deux ventouses, distance max 500 m.

### 4.4 Export DXF

Exporte le profil + ventouses + vidanges au format DXF (3 calques : Profil en long, Ventouses, Vidanges).

---

## 5. Diagnostic Système (16 vérifications)

| Code | Vérification | Catégorie |
|------|-------------|-----------|
| A1 | Point de fonctionnement dans la courbe H(Q) | Pompe ↔ Réseau |
| A2 | NPSH requis vs disponible | Pompe ↔ Réseau |
| A3 | Vitesse spécifique Nq | Pompe ↔ Réseau |
| B1 | Pression refoulement vs PN | Pompe ↔ HPT |
| B2 | Pression aspiration vs Pmin | Pompe ↔ HPT |
| C1 | Pmax transitoire vs PN | Réseau ↔ HPT |
| C2 | Pmin transitoire vs Pmin admissible | Réseau ↔ HPT |
| C3 | Cohérence multi-pompes | Réseau ↔ HPT |
| D1 | Volume gaz HPT vs seuil | HPT ↔ Ventouses |
| D2 | Présence ventouses | HPT ↔ Ventouses |
| D3 | Présence vidanges | HPT ↔ Ventouses |
| E1–E5 | Cohérence globale | Synthèse |

---

## 6. Génération du Rapport Word

Le rapport complet inclut :

1. **Entête** : projet, ingénieur, date
2. **Modèle hydraulique** : composition du réseau + métré canalisations (DN/matériau)
3. **Données pompes** : caractéristiques nominales + courbes H(Q)
4. **Enveloppe transitoire** : KPIs (Pmin, Pmax, Vgas, PN requis) + graphique
5. **Profil en long** : ventouses et vidanges
6. **Diagnostic système** : 16 vérifications avec synthèse et détails
7. **Annexe** : graphiques des vérifications

---

## 7. Formats de Fichiers

| Extension | Usage | Lecture | Écriture |
|-----------|-------|---------|----------|
| `.csv` | Données tabulaires | ✅ | ✅ (export vidanges) |
| `.xlsx` | Classeur Excel | ✅ | ✅ (templates) |
| `.xls` | Classeur Excel ancien | ✅ | ❌ |
| `.rtf` | Rapports pompe HAMMER | ✅ | ❌ |
| `.txt` | Rapports pompe (texte) | ✅ | ❌ |
| `.dxf` | Profil/tracé en plan AutoCAD | ✅ | ✅ (export ventouses) |
| `.hpi` | Projet HammerPy Insight | ✅ | ✅ |
| `.docx` | Rapport Word | ❌ | ✅ |

---

## 8. Raccourcis et Astuces

- **Générer des templates** : bouton `📄 Générer templates XLSX` dans l'onglet Pompe
- **Importer des données pompe** : bouton vert `📋 Importer données pompe (XLSX/CSV)`
- **Recalculer les diagnostics** : rechargez les données ou cliquez sur l'onglet Diagnostic
- **Sauvegarde** : le projet `.hpi` conserve toutes les données (pompes, profil, transitoire, mappings)
- **CRLF** : les fichiers CSV générés sous Excel (Windows) utilisent CP1252 — l'application le détecte automatiquement

---

## 9. Dépendances

- Python ≥ 3.10
- `customtkinter` — Interface graphique
- `pandas` — Analyse de données
- `openpyxl` — Lecture/écriture XLSX
- `ezdxf` — Import/export DXF
- `matplotlib` — Graphiques
- `numpy` — Calculs numériques
- `python-docx` — Génération Word
