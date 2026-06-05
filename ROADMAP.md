# 🗺️ Feuille de Route — HammerPy Insight v3.0
## *« Pompes & Ventouses : du réservoir HPT au système complet »*

> **Statut** : 🟢 Phase 1 terminée — Parser classeur HAMMER + UI + .hpi v3.0 (juin 2026)
> **Phase 2** : 🟢 Terminée — Parser rapport pompe RTF + UI courbe H(Q) graphique (juin 2026)
> **Phase 3** : 🟢 Terminée — AirValveSizing + profil en long + ventaises + vidanges + UI + sérialisation + Word (juin 2026)
> **Compatibilité ascendante** : `.hpi` v2.x → v3.0 (migration automatique) ✅

---

## 🎯 1. Vision & Objectifs

### 1.1 Contexte actuel (v3.0 — Phase 3 terminée ✅)
HammerPy Insight v3 sait :
- ✅ Charger et tracer les **enveloppes de pression** (Pmin, Pmax) issues de Bentley HAMMER
- ✅ Vérifier la conformité vis-à-vis de la **classe PN** et de la **pression min admissible**
- ✅ Afficher le **volume de gaz HPT** et son seuil (200 L par défaut)
- ✅ Générer une **note technique Word** avec tableau de résultats
- ✅ **Charger le classeur HAMMER complet** (.xlsx/.xls) — 6 feuilles Flex Tables
- ✅ **Afficher le résumé du modèle** : counts, Pmax/Pmin, matériaux, diamètres, Qmax pompe
- ✅ **Sauvegarder les 6 feuilles** dans le .hpi v3.0 (optimisé orient='records')
- ✅ **Section « Modèle Hydraulique »** dans le rapport Word
- ✅ **Parser le rapport pompe détaillé** HAMMER (.rtf / .txt) — 13 champs extraits
- ✅ **Multi-pompes** : batterie de pompes avec mode Continu/Parallèle + courbe combinée
- ✅ **Saisie manuelle des points de courbe H(Q)** avec graphique Matplotlib interpolé
- ✅ **UI Courbe H(Q) Pompe** : KPI, saisie Q/H, courbe interpolée + point nominal
- ✅ **Section « Batterie de Pompes »** dans le rapport Word avec alerte NPSH
- ✅ **Profil en long** : import CSV ou profil exemple, détection points hauts/bas
- ✅ **AirValveSizing** : pré-dimensionnement ventaises (anti-vide, combinée, grande orifice)
- ✅ **DrainValveSizing** : localisation vidanges aux points bas entre ventaises
- ✅ **Onglet UI « Ventaises & Vidanges »** : graphique + tableaux + export CSV
- ✅ **Sérialisation Phase 3** : ventaises/vidanges dans .hpi (rétrocompatible v3.0)
- ✅ **Section « Profil en Long »** dans le rapport Word (tableaux ventaises + vidanges)
- ✅ **65 tests unitaires** validés

### 1.2 Phase 4 — Module Système (prévue)
- ❌ **SystemDiagnostics** : vérifications croisées (pompe ↔ réseau ↔ HPT ↔ vidanges)
- ❌ **Indicateurs d'adéquation** : ✔ cohérent / ⚠ à vérifier / ✘ non conforme
- ❌ **Section « Système complet »** dans le rapport Word

| Objectif | Bénéfice utilisateur |
|---|---|
| **Pré-dimensionner les ventouses** le long du profil en long | Recommandations automatiques (DN, type, position) |
| **Pré-dimensionner les vidanges** aux points bas | Emplacement optimal entre deux ventouses, DN recommandé |
| Calculer le **NPSH disponible** vs NPSH requis (approfondi) | Détection précoce des risques de cavitation |
| **Modéliser le profil en long** de la conduite | Visualisation des pentes, points critiques |
| Vérifier l'**adéquation pompe × réseau** | Superposition graphique HMT pompe vs pertes de charge réseau |
| Étendre la **note technique Word** avec ces analyses | Dossier d'étude complet en un clic |

---

## 📦 2. Périmètre fonctionnel

### 2.1 ✅ Phase 1 — Parser classeur HAMMER (Terminée — Juin 2026)

#### A. WorkbookManager (nouveau ✅)
- Chargement du **classeur HAMMER complet** (.xlsx/.xls) avec les 6 feuilles Flex Tables
- **Détection par alias multilingue** : pipes/conduites, noeuds/nodes, pumps/pompe, etc.
- **Validation stricte** : refuse si Pipes, Noeuds ou Pompes manquants (feuilles obligatoires)
- **Helper `_parse_number()`** : gère l'espace insécable `\xa0` (séparateur milliers) et la virgule décimale française
- **Résumé automatique** : counts, Pmax/Pmin, matériaux, diamètres, Qmax pompe
- **6 compteurs + 4 mini-stats** dans l'UI (section « Modèle HAMMER »)

#### B. Format `.hpi` v3.0 (nouveau ✅)
- Les 6 feuilles du classeur sont embarquées dans le fichier `.hpi` (optimisé `orient='records'`)
- Nettoyage des types numpy → types natifs Python (JSON-sérialisable)
- Rétrocompatibilité : les `.hpi` v2.x restent lisibles

#### C. Rapport Word v3.0 (mis à jour ✅)
- Nouvelle section **« 2. Modèle Hydraulique »** avec tableau récapitulatif
- Matériaux et diamètres des conduites
- Pressions transitoires du modèle avec alertes conformité

### 2.2 ✅ Phase 2 — Module Pompes (Terminée — Juin 2026)

#### A. PumpReportParser (nouveau ✅)
- Extraction **13 champs** depuis le rapport détaillé HAMMER (.rtf / .txt)
- Deux phases de parsing : ID/Label (section General) + données opérées (section Pump Data)
- `_strip_rtf()` : suppression balises RTF, images (shppict/pict), fonttbl, contrôles
- Gestion des nombres de positionnement RTF (ex: -229, -432)
- **Filtrage `_is_positioning_number()`** : négatifs entiers > 20

#### B. UI Courbe H(Q) Pompe (nouveau ✅)
- **5 KPI** : Label, Q nominal (L/s), HMT pompe (m), NPSH disponible (m), Nombre de points courbe
- **Saisie manuelle** : champs Q (L/s) + H (m) + boutons Ajouter / Effacer tout
- **Liste des points** : zone texte avec tableau formaté (#, Q, H)
- **Graphique Matplotlib** : courbe interpolée (numpy polyfit, deg ≤ 3), points saisis (pastilles orange), point nominal (losange rouge)
- Adaptation automatique au thème clair/sombre

#### C. Intégration sérialisation (mise à jour ✅)
- Section `"pump"` dans le `.hpi` v3.0 : `filepath`, `parsed` dict, `curve_points` list
- Rétrocompatibilité v2.x préservée

#### D. Rapport Word — Section Pompe (mise à jour ✅)
- Nouvelle section **« 2b. Données Pompe »** avec tableau récapitulatif 10 lignes
- Alerte NPSH si disponible < requis
- Couleur thème violet (#8b5cf6)

#### E. Tests (58 tests validés ✅)
- 12 tests `PumpReportParser` : chargement RTF réel, strip RTF, courbe points, interpolation, résumé

### 2.3 📋 Phase 3 — Module Ventouses (Prévue)

#### A. Profil en long
- **Saisie du profil en long** de la conduite : tableau `distance (m) | côte TN (m) | pente locale (%)`
  - Import CSV (3 colonnes) ou saisie manuelle
- **Calcul des points hauts/bas** automatiques
- **Pré-dimensionnement** par règle métier :
  - **Ventouse simple** (anti-vide) : DN ≥ DN_conduite / 12, à chaque point haut
  - **Ventouse combinée** (admission + dégazage) : recommandée sur les longs tronçons horizontaux
  - **Ventouse à grande orifice** (admission d'air rapide) : DN ≥ DN_conduite / 8
- **Tableau de localisation** : PK, côte, type recommandé, DN suggéré
- **Chart** : profil en long de la conduite avec marqueurs des ventouses

#### B. Vidanges (drainage)
- **Localisation automatique** des vidanges aux points bas du profil en long
- **Règle métier** : une vidange au point bas entre deux ventouses consécutives
- **Algorithme de positionnement** :
  1. Identifier les ventouses (points hauts)
  2. Pour chaque segment entre 2 ventouses, trouver le point bas (minimum local d'altitude)
  3. Si le point bas existe et est suffisamment éloigné des ventouses (> 50m), placer une vidange
  4. Sinon, recommander une vidange combinée avec la ventouse la plus proche
- **Dimensionnement** : DN ≥ DN_conduite / 10, vanne d'arrêt + bride de vidange
- **Paramètres** : longueur max entre 2 vidanges (défaut : 500m), pente min pour drainage
- **Tableau de localisation** : PK, côte, type vidange, DN, distance aux ventouses voisines
- **Chart** : profil en long avec marqueurs ventouses (▲) et vidanges (▼)

#### C. Module Système complet
- **Vue d'ensemble** récapitulant :
  - Pompe sélectionnée + point de fonctionnement
  - Réseau : DN, longueur, rugosité, pertes
  - Protection : réservoir HPT (déjà en v2), ventouses recommandées
- **Indicateurs d'adéquation** : ✔ cohérent / ⚠ à vérifier / ✘ non conforme

### 2.4 🚫 Hors périmètre v3.0 (différé en v3.1+)
- Calcul réseau maillé (loi des mailles) — v3.1
- Courbes de pompe à vitesse variable (loi d'affinité) — v3.1
- Simulation transitoire intégrée (MOC) sans HAMMER — v4.0
- Import direct des fichiers HAMMER `.csv` de pompe (`.pump`) — v3.1
- Intégration de catalogues constructeur (Grundfos, KSB, etc.) — v3.1
- Multi-pompes en parallèle / série — v3.2

---

## 🏗️ 3. Architecture technique

### 3.1 Vue d'ensemble (couches applicatives)

```
┌─────────────────────────────────────────────────────────┐
│  UI (CustomTkinter)                                     │
│  ┌─────────┬─────────┬─────────┬─────────┬─────────┐   │
│  │ Station │ Transi- │ Rapport │  Pompe  │Ventouses│   │
│  │ (v2)    │ toire   │ (v2)    │ (✅ v3) │+Vidanges│   │
│  │         │ (v2)    │         │         │(📋 v3)  │   │
│  └─────────┴─────────┴─────────┴─────────┴─────────┘   │
│                        │                                │
│  ┌─────────────────────▼─────────────────────────┐     │
│  │  Métier (Logique)                              │     │
│  │  • HammerDataParser (v2)                       │     │
│  │  • PumpReportParser     ✅ NEW (Phase 2)       │     │
│  │  • AirValveSizing         ← NEW (Phase 3)      │     │
│  │  • DrainValveSizing       ← NEW (Phase 3)      │     │
│  │  • SystemDiagnostics      ← NEW (Phase 3)      │     │
│  │  • WordReportGenerator (étendu)                │     │
│  └────────────────────┬──────────────────────────┘     │
│                       │                                │
│  ┌────────────────────▼──────────────────────────┐     │
│  │  Persistance                                    │     │
│  │  • .hpi v3.0 (JSON) — extension de v2.0        │     │
│  │  • Champs backward-compatible                  │     │
│  └────────────────────────────────────────────────┘     │
└─────────────────────────────────────────────────────────┘
```

### 3.2 Classes / modules

| Classe | Responsabilité | Statut |
|---|---|---|
| `PumpReportParser` | Extraction données pompe depuis rapport RTF/TXT | ✅ Terminé |
| `AirValveSizing` | Règles de dimensionnement ventouses, détection points hauts | 📋 Phase 3 |
| `DrainValveSizing` | Localisation vidanges aux points bas entre ventouses | 📋 Phase 3 |
| `SystemDiagnostics` | Vérifications croisées (pompe ↔ réseau ↔ HPT ↔ vidanges) | 📋 Phase 3 |

> **Décision** : tout reste dans `main.py` pour la v3.0 (cohérence avec v2.x monolithique). Extraction en modules séparés si la base dépasse ~3500 lignes.

### 3.3 Format `.hpi` v3.0 (extension de v2.0)

```json
{
  "version": "3.0",
  "metadata": { ... },
  "config": {
    "pn_label": "PN 16", "pn_value": 16.0,
    "pmin_label": "...", "pmin_value": -0.1,
    "flow_unit": "m³/h", "volume_unit": "L", "volume_threshold_l": 200.0
  },
  "station": { ... },
  "hpt":     { ... },
  "report_text": "...",

  "pump": {                                       ✅ NEW (Phase 2)
    "filepath": "Pump detailed report.rtf",
    "parsed": {
      "pump_id": "122",
      "label": "PMP-2",
      "downstream_pipe": "P-4",
      "flow_lps": 100.0,
      "pump_head_m": 110.80,
      "npsh_available_m": 15.35,
      "npsh_required_m": null,
      "pressure_suction_bar": 0.54,
      "pressure_discharge_bar": 11.38,
      "speed_factor": 1.0,
      "status_initial": "On",
      "controlled": false
    },
    "curve_points": [
      {"flow_lps": 0.0,   "head_m": 200.0},
      {"flow_lps": 75.0,  "head_m": 150.0},
      {"flow_lps": 150.0, "head_m": 100.0}
    ]
  },

  "air_valves": {                                  ← NEW (Phase 3)
    "profile": [
      [0,    125.5, 0.0],
      [250,  138.2, 5.1],
      [500,  142.0, 1.5]
    ],
    "pipe_dn_mm": 250,
    "recommendations": []
  },

  "drain_valves": {                                ← NEW (Phase 3)
    "recommendations": [
      {
        "pk_m": 375.0,
        "z_m": 130.5,
        "type": "Vidange à bride",
        "dn_mm": 25,
        "left_ventouse_pk": 250.0,
        "right_ventouse_pk": 500.0,
        "distance_to_left_m": 125.0,
        "distance_to_right_m": 125.0
      }
    ]
  },

  "system_diagnostics": []                         ← NEW (Phase 3)
}
```

**Rétrocompatibilité** : un `.hpi` v2.0 est lisible par v3.0. Les sections `pump` et `air_valves` sont simplement absentes → l'onglet Affiche un état vide + bouton "Charger".

---

## 🖥️ 4. Architecture UI (CustomTkinter)

### 4.1 Organisation des onglets

```
[ Onglet 1 — Régime Permanent ]     [ Onglet 2 — Analyse Transitoire ]
                                    [ Onglet 3 — Rapport Technique ]
```

### 4.2 Onglet 2 — Analyse Transitoire (disposition actuelle)

```
┌──────────────────────────────────────────────────────────────┐
│  Analyse des Pressions & Volumes Transitoires                │
├──────────────────────────────────────────────────────────────┤
│  [Importer données HPT (.csv/.xlsx)]  Aucun fichier         │
│                                                              │
│  ┌── KPI ──────────────────────────────────────────────┐    │
│  │ Pression Min    │ Pression Max    │ Volume Gaz Max   │    │
│  │ -0.15 bar       │ 14.20 bar       │ 185.3 L          │    │
│  └─────────────────────────────────────────────────────┘    │
│                                                              │
│  ┌── MODÈLE HAMMER ────────────────────────────────────┐    │
│  │ [Charger classeur (.xlsx)]  Flex Tables.xlsx         │    │
│  │ Pipes: 12 │ Nœuds: 8 │ Pompes: 2 │ Réservoirs: 1   │    │
│  │ P max: 14.2 bar │ P min: -0.15 bar │ Matériaux: PE  │    │
│  └─────────────────────────────────────────────────────┘    │
│                                                              │
│  ┌── COURBE H(Q) POMPE ────────────────────────────────┐    │
│  │ [Charger rapport pompe (.rtf)]  Pump report.rtf      │    │
│  │ Pompe: PMP-2 │ Q nom: 100.0 L/s │ HMT: 110.8 m     │    │
│  │ NPSH dispo: 15.3 m │ Pts courbe: 3                  │    │
│  │                                                       │    │
│  │ Q (L/s): [75]  H (m): [150]  [+ Ajouter] [Effacer] │    │
│  │  #     Q (L/s)       H (m)                           │    │
│  │  ──────────────────────────────                      │    │
│  │  1       0.0         200.0                           │    │
│  │  2      75.0         150.0                           │    │
│  │  3     150.0         100.0                           │    │
│  └─────────────────────────────────────────────────────┘    │
│                                                              │
│  ┌── Graphique H(Q) ───────────────────────────────────┐    │
│  │  H (m)                                                │    │
│  │  200│●                                                 │    │
│  │     │ ╲                                               │    │
│  │  150│  ●─ ─ ─ ─ ─ ★ (nominal)                        │    │
│  │     │      ╲                                          │    │
│  │  100│       ●                                         │    │
│  │     └──────────────────────▶ Q (L/s)                  │    │
│  │       0    50   100   150                             │    │
│  └──────────────────────────────────────────────────────┘    │
│                                                              │
│  ┌── Graphique HPT ────────────────────────────────────┐    │
│  │  Pression Min/Max + Volume Gaz                       │    │
│  └──────────────────────────────────────────────────────┘    │
└──────────────────────────────────────────────────────────────┘
```

### 4.3 Rapport (Onglet 3) — Sections Word

- **§2. Modèle Hydraulique** : tableau récap 6 composants, matériaux, pressions
- **§2b. Données Pompe** : tableau récap 10 lignes, alerte NPSH
- **§3. Résultats Transitoires** : Pmin/Pmax/Vol.Gaz
- **§4. Profil en Long** : tableau ventouses + vidanges, graphique profil
- **§5. Interprétation & Recommandations**

### 4.4 Onglet 4 — Ventouses & Vidanges (Phase 3)

```
┌──────────────────────────────────────────────────────────────┐
│  💨  Profil en Long & Dimensionnement Ventouses + Vidanges   │
├──────────────────────────────────────────────────────────────┤
│  [Importer profil (.csv)]   Profil en long — 12 points        │
│                                                               │
│  ┌───── Profil en long (côte TN) ────────────────────────┐   │
│  │   Z (m)                                                │   │
│  │  145│      ●─●                    ▲ V2 (ventouse)      │   │
│  │     │   ●─╱    ╲─●                                     │   │
│  │  135│  ╱          ╲                                     │   │
│  │     │ ╱            ●─●                                  │   │
│  │  125│●          ▼ D1          ●─●                   ●   │   │
│  │     │        (vidange)                                 │   │
│  │     └────────────────────────────────────────▶ PK (m)   │   │
│  │       0      250     500     750    1000               │   │
│  │              ▲ V1              ▲ V3   ▲ V4             │   │
│  └─────────────────────────────────────────────────────────┘  │
│                                                               │
│  ┌───── Ventouses ───────────────────────────────────────┐   │
│  │  PK (m) │  Côte (m) │ Type              │ DN (mm)  │  │  │
│  │  ───────┼───────────┼───────────────────┼──────────│  │  │
│  │   250   │  138.2    │ Combinée GO       │   80     │  │  │
│  │   500   │  142.0    │ Anti-vide simple  │   50     │  │  │
│  └────────────────────────────────────────────────────────┘  │
│                                                               │
│  ┌───── Vidanges ────────────────────────────────────────┐   │
│  │  PK (m) │  Côte (m) │ Type         │ DN  │ Dist.V1/V2│  │
│  │  ───────┼───────────┼──────────────┼─────┼───────────│  │
│  │   375   │  130.5    │ À bride      │  25 │ 125/125 m │  │
│  └────────────────────────────────────────────────────────┘  │
│                                                               │
│  [Exporter liste ventouses + vidanges (.csv)]                 │
└──────────────────────────────────────────────────────────────┘
```

---

## 📅 5. Phasage

| Phase | Livrables | Statut |
|---|---|---|
| **P1 — Parser classeur HAMMER** | WorkbookManager + 6 feuilles + .hpi v3.0 + Rapport Word | ✅ Terminé |
| **P2 — Module Pompe** | PumpReportParser + UI courbe H(Q) + graphique + 58 tests | ✅ Terminé |
| **P3 — Module Ventouses & Vidanges** | AirValveSizing + DrainValveSizing + profil en long + UI + tests + sérialisation + Word | ✅ Terminé |
| **P4 — Module Système** | SystemDiagnostics + Rapport Word complet | 📋 À démarrer |
| **P5 — Documentation** | README, CHANGELOG, guide utilisateur v3.0 | 📋 |

---

## 🧪 6. Stratégie de tests

### 6.1 Tests unitaires (65 tests ✅)
- 19 tests `_parse_number()` : natif, None/NaN, nbsp, virgule FR, négatifs
- 4 tests `_find_col_in_df()` : match exact, partiel, casse
- 17 tests intégration workbook : chargement, validation, résumé
- 3 tests erreurs : extension invalide, fichier absent, feuille manquante
- 4 tests rétrocompatibilité : CSV HPT + station lisibles
- 12 tests `PumpReportParser` : RTF réel, strip RTF, courbe points, interpolation, résumé
- 7 tests `AirValveSizing` : profil, points hauts/bas, sizing, DN, export CSV

### 6.2 Tests Phase 3 (prévu)
- `AirValveSizing` : détection points hauts, dimensionnement ventouses (8+ cas)
- `DrainValveSizing` : localisation vidanges aux points bas, distance aux ventouses (6+ cas)
- `SystemDiagnostics` : chaque check avec cas OK / WARN / FAIL / NA
- Migration v2.x → v3.0 : lecture de 3 fichiers de test legacy

### 6.3 Tests visuels (manuels)
- Courbe H(Q) interpolée lisible et correcte
- Point nominal bien positionné sur la courbe
- Adaptation thème clair/sombre vérifiée

---

## ⚠️ 7. Risques & questions ouvertes

### 7.1 Risques techniques
| Risque | Probabilité | Impact | Mitigation |
|---|---|---|---|
| Interpolation polynomiale instable avec 2 points | Faible | Moyen | Degré limité à min(nb-1, 3) |
| Format RTF variable entre versions HAMMER | Moyen | Moyen | `_strip_rtf()` robuste + tests avec fichier réel |
| Performance UI avec >50 points de courbe | Faible | Faible | Tri automatique par Q croissant |

### 7.2 Questions ouvertes (pour Phase 3)
1. **Interpolation avancée** : spline cubique pour courbes à >4 points ?
2. **Auto-parser points courbe** : extraction automatique depuis le RTF (Shutoff/Design/Max Op)
3. **Ventouses sur sections en charge vs refoulement** : différenciation ?
4. **Multilangue** : i18n FR/EN dès v3.0 ? → Recommandation : NON

---

## 📊 8. Métriques de succès

| Métrique | Cible | Actuel |
|---|---|---|
| Couverture de tests unitaires | > 85% | 65 tests ✅ |
| Bugs de régression | < 5 | 0 ✅ |
| Temps de chargement `.hpi` v3.0 (< 5 MB) | < 1 s | ✅ |
| Temps d'export Word complet (< 15 pages) | < 5 s | ✅ |

---

## 📎 9. Annexes

### A. Arborescence actuelle
```
HammerPy Insight/
├── main.py                          # GUI + entry point (~2600 lignes)
├── utils.py                         # Fonctions utilitaires, constantes
├── data_parser.py                   # Parser CSV/Excel multi-feuilles
├── workbook.py                      # WorkbookManager — classeur HAMMER
├── pump_parser.py                   # PumpReportParser — rapport pompe RTF
├── report_generator.py              # WordReportGenerator — rapport Word
├── air_valve_sizing.py              # AirValveSizing — ventaises/vidanges
├── hammerpy_icon.ico                # Icône multi-tailles (370 Ko)
├── test_workbook_parser.py          # 65 tests unitaires
├── requirements.txt                 # Dépendances
├── README.md                        # Documentation v3.0 Phase 3
├── ROADMAP.md                       # Ce document
├── Flex Tables.xlsx                 # Classeur HAMMER exemple (exclu du dépôt)
├── station_steady_state_test.csv    # CSV test station
├── hpt_transient_test.csv           # CSV test HPT
└── .gitignore                       # Exclut données perso + Flex Tables
```

### B. Dépendances
- **Aucune dépendance externe obligatoire** ajoutée en Phase 2
- `numpy` : utilisé pour l'interpolation polyfit (déjà présent via matplotlib)

### C. Glossaire
- **HPT** : Hydropneumatic Tank (réservoir anti-bélier)
- **NPSH** : Net Positive Suction Head (hauteur de charge nette à l'aspiration)
- **MOC** : Method of Characteristics (méthode des caractéristiques)
- **PN** : Pression Nominale (classe de conduite)
- **DN** : Diamètre Nominal (conduite ou appareil)
- **η** : Rendement (de la pompe)
- **RTF** : Rich Text Format (format du rapport pompe HAMMER)

---

*Document rédigé le 4 juin 2026 — HammerPy Insight v3.0 Phase 2 — Roadmap*
*Phase 2 terminée le 5 juin 2026 — Parser rapport pompe + UI courbe H(Q)*
*Phase 3 terminée le 5 juin 2026 — AirValveSizing + profil en long + sérialisation + Word*
*Prochaine revue : démarrage Phase 4 (module Système)*
