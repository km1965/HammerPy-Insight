# 🗺️ Feuille de Route — HammerPy Insight v3.0
## *« Pompes & Ventouses : du réservoir HPT au système complet »*

> **Statut** : 🟢 Phase 1 terminée — Parser classeur HAMMER + UI + .hpi v3.0 (juin 2026)
> **Phase 2** : 📋 En cours — Courbe H(Q) pompe + NPSH
> **Phase 3** : 📋 Prévue — Ventouses + profil en long
> **Compatibilité ascendante** : `.hpi` v2.x → v3.0 (migration automatique) ✅

---

## 🎯 1. Vision & Objectifs

### 1.1 Contexte actuel (v3.0 — Phase 1 terminée ✅)
HammerPy Insight v3 sait :
- ✅ Charger et tracer les **enveloppes de pression** (Pmin, Pmax) issues de Bentley HAMMER
- ✅ Vérifier la conformité vis-à-vis de la **classe PN** et de la **pression min admissible**
- ✅ Afficher le **volume de gaz HPT** et son seuil (200 L par défaut)
- ✅ Générer une **note technique Word** avec tableau de résultats
- ✅ **Charger le classeur HAMMER complet** (.xlsx/.xls) — 6 feuilles Flex Tables
- ✅ **Afficher le résumé du modèle** : counts, Pmax/Pmin, matériaux, diamètres, Qmax pompe
- ✅ **Sauvegarder les 6 feuilles** dans le .hpi v3.0 (optimisé orient='records')
- ✅ **Section « Modèle Hydraulique »** dans le rapport Word

### 1.2 Phase 2 — Pompes (en cours)
La v3.0 Phase 1 charge les données transitoires des pompes, mais reste muette sur :
- ❌ Le **comportement de la pompe** (point de fonctionnement, risque de cavitation, NPSH disponible)
- ❌ L'**adéquation pompe × réseau** (interaction courbe pompe H(Q) × courbe réseau)

### 1.3 Phase 3 — Ventouses (prévue)
- ❌ La **localisation et le dimensionnement des ventouses** sur la conduite

| Objectif | Bénéfice utilisateur |
|---|---|
| Intégrer la **courbe caractéristique de pompe** H(Q), η(Q), P(Q) | Valider le point de fonctionnement et le rendement |
| Calculer le **NPSH disponible** vs NPSH requis | Détection précoce des risques de cavitation |
| **Pré-dimensionner les ventouses** le long du profil en long | Recommandations automatiques (DN, type, position) |
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

### 2.2 📋 Phase 2 — Module Pompes (En cours)

#### A. Courbe caractéristique H(Q)
- Saisie / import d'une **courbe caractéristique** de pompe (H(Q), η(Q), P(Q)) :
  - Saisie manuelle (table éditable)
  - Import CSV (3 colonnes : Q, H, η)
  - **Auto-détection** du format (m³/h ou L/s, mCE ou m, kW ou CV)
- **Paramètres** : vitesse de rotation N (tr/min), diamètre de roue, NPSH requis (NPSH₃%)
- **Calculs** :
  - **Point de fonctionnement** (intersection courbe pompe × courbe réseau) — résolution itérative
  - **Rendement** au point de fonctionnement
  - **Puissance absorbée** P = ρgQH/η
  - **NPSH disponible** = (P_atm − P_vapeur) / (ρg) ± ΔH_statique − pertes
- **Diagnostic** : concordance Q,HMT avec les valeurs extraites du CSV station
- **Chart** : courbe H(Q) superposée à la courbe réseau, marqueur du point de fonctionnement

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

#### C. Module Système complet
- **Vue d'ensemble** récapitulant :
  - Pompe sélectionnée + point de fonctionnement
  - Réseau : DN, longueur, rugosité, pertes
  - Protection : réservoir HPT (déjà en v2), ventouses recommandées
- **Indicateurs d'adéquation** : ✔ cohérent / ⚠ à vérifier / ✘ non conforme

### 2.2 🚫 Hors périmètre v3.0 (différé en v3.1+)
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
│  │ (v2)    │ toire   │ (v2)    │ (NEW v3)│(NEW v3) │   │
│  │         │ (v2)    │         │         │         │   │
│  └─────────┴─────────┴─────────┴─────────┴─────────┘   │
│                        │                                │
│  ┌─────────────────────▼─────────────────────────┐     │
│  │  Métier (Logique)                              │     │
│  │  • HammerDataParser (v2)                       │     │
│  │  • PumpCurveCalculator    ← NEW                │     │
│  │  • AirValveSizing         ← NEW                │     │
│  │  • SystemDiagnostics      ← NEW                │     │
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

### 3.2 Nouvelles classes / modules

| Classe | Responsabilité | Fichier |
|---|---|---|
| `PumpCurveCalculator` | Calcul point de fonctionnement, NPSH, puissance, rendement | `main.py` (ou `pump_engine.py` si volume) |
| `AirValveSizing` | Règles de dimensionnement, détection points hauts/bas | `main.py` |
| `PumpProfileData` (dataclass) | Modèle de données pompe | `main.py` |
| `AirValveData` (dataclass) | Modèle de données ventouse | `main.py` |
| `SystemDiagnostics` | Vérifications croisées (pompe ↔ réseau ↔ HPT) | `main.py` |

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

  "pump": {                                       ← NEW
    "name": "Grundfos CR 64-3",
    "speed_rpm": 2900,
    "curve_data": [                              ← Q (m³/h), H (m), η (%), P (kW)
      [0.0,    72.0, 0.0,  3.5],
      [20.0,   70.5, 35.0, 4.2],
      [40.0,   67.0, 58.0, 4.9],
      [60.0,   61.5, 68.0, 5.5],
      [80.0,   54.0, 64.0, 6.1],
      [100.0,  44.5, 52.0, 6.4]
    ],
    "npsh_required_m": 3.2,
    "npsh_available_m": 6.8,                     ← calculé
    "operating_point": {                          ← calculé
      "flow_m3h": 64.5,
      "head_m": 60.2,
      "efficiency_pct": 67.8,
      "power_kw": 5.1
    },
    "source_unit_flow": "m³/h",
    "source_unit_head": "mCE"
  },

  "air_valves": {                                  ← NEW
    "profile": [                                  ← PK (m), Z (m), slope (%)
      [0,    125.5, 0.0],
      [250,  138.2, 5.1],
      [500,  142.0, 1.5],
      [750,  135.8, -2.5],
      [1000, 128.0, -3.1]
    ],
    "pipe_dn_mm": 250,
    "recommendations": [                           ← calculé
      {"pk_m": 500,  "z_m": 142.0, "type": "Combinée grande orifice", "dn_mm": 80, "reason": "Point haut"},
      {"pk_m": 1000, "z_m": 128.0, "type": "Anti-vide simple",       "dn_mm": 50, "reason": "Extrémité aval"}
    ]
  },

  "system_diagnostics": [                          ← NEW (résultats agrégés)
    {"check": "Adéquation pompe × Q nominal",     "status": "OK",   "value": "..."},
    {"check": "NPSH disponible > NPSH requis",    "status": "OK",   "value": "..."},
    {"check": "Ventouses aux points hauts",        "status": "WARN", "value": "..."},
    ...
  ]
}
```

**Rétrocompatibilité** : un `.hpi` v2.0 est lisible par v3.0. Les sections `pump`, `air_valves`, `system_diagnostics` sont simplement absentes → onglet "Pompe" et "Ventouses" affichent un état vide + bouton "Démarrer l'analyse".

---

## 🖥️ 4. Architecture UI (CustomTkinter)

### 4.1 Nouvelle organisation des onglets

```
[ Onglet 1 — Régime Permanent ] (v2)    [ Onglet 2 — Pompe ]        (NEW v3)
[ Onglet 3 — Analyse Transitoire ] (v2)  [ Onglet 4 — Ventouses ]    (NEW v3)
[ Onglet 5 — Rapport Technique ] (v2 + nouvelles sections)
```

### 4.2 Onglet 2 (Pompe) — Maquette

```
┌──────────────────────────────────────────────────────────────┐
│  ⚙️  Caractéristique & Point de Fonctionnement de la Pompe    │
├──────────────────────────────────────────────────────────────┤
│  [Importer courbe (.csv)]   Aucun fichier       [Unité: Auto▾]│
│                                                               │
│  ┌───── Courbe caractéristique ─────┐  ┌─── KPI ──────────┐  │
│  │  H (m)                            │  │ Q fonctionn.     │  │
│  │  ▲  ╲                              │  │ 64.5 m³/h        │  │
│  │  │   ╲___                          │  │                  │  │
│  │  │       ╲___                      │  │ HMT au point     │  │
│  │  │           ╲___  ★ Point fonct.  │  │ 60.2 mCE         │  │
│  │  └──────────────────▶ Q (m³/h)     │  │                  │  │
│  │   + courbe réseau                  │  │ Rendement η      │  │
│  │   + ligne de pompe                 │  │ 67.8 %           │  │
│  └────────────────────────────────────┘  │                  │  │
│                                          │ Puissance P      │  │
│  ┌───── NPSH ────────────────────────┐  │ 5.1 kW           │  │
│  │ NPSH requis  : 3.20 m              │  └──────────────────┘  │
│  │ NPSH dispo   : 6.80 m              │                         │
│  │ Marge        : +3.60 m ✔           │  [✔ Marge satisfaisante]│
│  └────────────────────────────────────┘                         │
│                                                               │
│  Paramètres : [Nom: _________] [N (tr/min): 2900] [NPSH req: 3.2]│
└──────────────────────────────────────────────────────────────┘
```

### 4.3 Onglet 4 (Ventouses) — Maquette

```
┌──────────────────────────────────────────────────────────────┐
│  💨  Profil en Long & Dimensionnement des Ventouses           │
├──────────────────────────────────────────────────────────────┤
│  [Importer profil (.csv)]   Profil en long — 12 points        │
│                                                               │
│  ┌───── Profil en long (côte TN) ────────────────────────┐   │
│  │   Z (m)                                                │   │
│  │  145│      ●─●                                          │   │
│  │     │   ●─╱    ╲─●                                      │   │
│  │  135│  ╱          ╲                                      │   │
│  │     │ ╱            ●─●                                   │   │
│  │  125│●                                                ●│   │
│  │     └────────────────────────────────────────▶ PK (m)   │   │
│  │       0      250     500     750    1000                 │   │
│  │              ▼ V1    ▼ V2   ▼ V3   ▼ V4  (ventouses)   │   │
│  └─────────────────────────────────────────────────────────┘   │
│                                                               │
│  ┌───── Recommandations ─────────────────────────────────┐  │
│  │  PK (m) │  Côte (m) │ Type              │ DN (mm) │ R │ │
│  │  ───────┼───────────┼───────────────────┼─────────┼─── │ │
│  │   500   │  142.0    │ Combinée GO       │   80    │PH │ │
│  │   750   │  135.8    │ Anti-vide simple  │   50    │PD │ │
│  │  1000   │  128.0    │ Anti-vide simple  │   50    │EX │ │
│  └────────────────────────────────────────────────────────┘  │
│                                                               │
│  [Exporter liste ventouses (.csv)]                            │
└──────────────────────────────────────────────────────────────┘
```

### 4.4 Sidebar — Nouvelles sections

```
ACTIF                                          ← NOUVEAU
─────────────────
Pompe installée  : [✔]                         ← v3 (bascule onglet 2)
Ventouses        : [✔]                         ← v3 (bascule onglet 4)
```

### 4.5 Rapport (Onglet 5) — Nouvelles sections Word

Insertion automatique de **3 nouveaux chapitres** dans `WordReportGenerator.generate()` :

- **§5  Caractéristique de la Pompe & Point de Fonctionnement**
  - Tableau : Nom, N, NPSH req, Q, HMT, η, P
  - Image : graphique H(Q) avec point de fonctionnement
  - Diagnostic : marges, adéquation

- **§6  Ventousage & Protection Anti-Vide**
  - Tableau : PK, Côte, Type, DN
  - Image : profil en long avec emplacements
  - Diagnostic : couverture, conformités

- **§7  Synthèse Système & Recommandations**
  - Tableau de diagnostic global
  - Conclusions & actions prioritaires

---

## ⚙️ 5. Spécifications fonctionnelles détaillées

### 5.1 `PumpCurveCalculator`

**Méthodes** :
```python
class PumpCurveCalculator:
    def __init__(self, curve_data: list[tuple[float, float, float, float]],
                 npsh_required_m: float):
        # curve_data : [(Q_m3h, H_m, eta_pct, P_kW), ...]
        ...

    def find_operating_point(self, hmt_target_m: float, q_target_m3h: float) -> dict:
        """Interpolation linéaire de la courbe pour trouver (Q, H) au point voulu."""
        ...

    def compute_npsh_available(self, p_atm_m: float, p_vapor_m: float,
                                h_static_m: float, losses_m: float) -> float:
        """NPSHd = (Patm - Pvapeur) / (ρg) + ΔH_static - pertes"""
        ...

    def compute_power(self, q_m3h: float, h_m: float, eta_pct: float) -> float:
        """P_kW = ρ * g * Q * H / (1000 * η)"""
        ...

    def get_curve_data_for_display(self) -> dict:
        """Renvoie les arrays pour matplotlib (Q_array, H_array, eta_array, P_array)."""
        ...
```

**Validations** :
- Au moins 3 points dans la courbe
- Q ≥ 0, H ≥ 0, 0 ≤ η ≤ 100, P ≥ 0
- Avertissement si la courbe n'est pas monotone décroissante

### 5.2 `AirValveSizing`

**Méthodes** :
```python
class AirValveSizing:
    def __init__(self, profile: list[tuple[float, float, float]],
                 pipe_dn_mm: int):
        # profile : [(PK_m, Z_m, slope_pct), ...]
        ...

    def find_critical_points(self) -> list[dict]:
        """Détecte les points hauts (maxima locaux Z) et points bas (minima).
        Retourne : [{pk, z, type: 'high'|'low'|'end'}, ...]"""
        ...

    def recommend_valves(self) -> list[dict]:
        """Pour chaque point critique, recommande type + DN.
        Règles :
          - Point haut : ventouse combinée grande orifice, DN ≥ DN/8
          - Point bas  : pas de ventouse (sauf vanne de vidange)
          - Extrémités : ventouse anti-vide simple, DN ≥ DN/12
          - Long tronçon horizontal (> 500m sans singularité) : ventouse intermédiaire
        """
        ...

    def get_profile_for_display(self) -> dict:
        """Renvoie arrays pour matplotlib + marqueurs ventouses."""
        ...
```

### 5.3 `SystemDiagnostics`

**Méthodes** :
```python
class SystemDiagnostics:
    def __init__(self, pump, air_valves, hpt_result, station_result, config):
        ...

    def run_all_checks(self) -> list[dict]:
        """Exécute toutes les vérifications croisées et retourne un rapport.
        Format : [{check, status, value, message}, ...]
        Status ∈ {'OK', 'WARN', 'FAIL', 'NA'}
        """
        checks = [
            self.check_pump_matches_flow(),
            self.check_npsh_margin(),
            self.check_pump_power_vs_electrical(),
            self.check_valves_at_high_points(),
            self.check_hpt_volume_under_threshold(),
            self.check_pressure_envelopes(),
        ]
        return checks

    def check_npsh_margin(self) -> dict:
        """Marge ≥ +1 m → OK ; 0 à 1 m → WARN ; < 0 → FAIL"""
        ...

    def check_valves_at_high_points(self) -> dict:
        """Toutes les points hauts doivent être couverts → sinon WARN"""
        ...
```

---

## 🔄 6. Migration `.hpi` v2.x → v3.0

### 6.1 Règle de lecture
```python
def _load_project_v2_or_v3(self, filepath: str):
    payload = json.load(open(filepath))
    version = payload.get("version", "1.0")

    if version.startswith("1.") or version.startswith("2."):
        # Migration transparente : ajout des sections vides
        payload.setdefault("pump", None)
        payload.setdefault("air_valves", None)
        payload.setdefault("system_diagnostics", None)
        payload["version"] = "3.0"   # Promotion à la lecture
        self._migrated_from = version

    # Suite : chargement normal
    ...
```

### 6.2 Règle d'écriture
- Toujours écrire `version: "3.0"` pour les nouveaux projets
- Conserver l'avertissement à l'utilisateur si le projet a été migré depuis v2.x :
  > *"Ce projet a été créé avec la v2.x. Certaines sections (Pompe, Ventouses) sont vides et n'ont pas encore été remplies."*

### 6.3 Format legacy `.hpi` v2.x
- Rester lisible par v2.x ? → **NON** (cohérence interne)
- Documenter la non-réversibilité dans le rapport d'export et la doc

---

## 🧪 7. Stratégie de tests

### 7.1 Tests unitaires (couverture > 85%)
- `PumpCurveCalculator` : interpolation, NPSH, puissance (10+ cas)
- `AirValveSizing` : détection points critiques, dimensionnement (8+ cas)
- `SystemDiagnostics` : chaque check avec cas OK / WARN / FAIL / NA
- Migration v2.x → v3.0 : lecture de 3 fichiers de test legacy

### 7.2 Tests d'intégration
- Création projet v3 → ajout pompe → ajout ventouses → sauvegarde → rechargement
- Vérification que la régénération du graphique fonctionne avec sections vides
- Export Word v3 : présence des nouvelles sections

### 7.3 Tests visuels (manuels)
- Maquettes des onglets 2 & 4 validées par 2 ingénieurs hydrauliques
- Vérification de la lisibilité des couleurs KPI (vert/orange/rouge)
- Impression PDF de la note technique complète (15-20 pages)

### 7.4 Tests de régression
- Tous les tests v2.x doivent toujours passer
- Projets `.hpi` de test (3 fichiers legacy) se chargent sans erreur

---

## 📅 8. Phasage proposé

| Phase | Livrables | Durée estimée | Statut |
|---|---|---|---|
| **P1 — Spec & maquettes** | Ce document + wireframes Figma/PNG | 2 semaines | ✅ Fait |
| **P2 — Module Pompe** | `PumpCurveCalculator` + Onglet 2 + tests unitaires | 3 semaines | 🔲 À démarrer |
| **P3 — Module Ventouses** | `AirValveSizing` + Onglet 4 + tests | 2 semaines | 🔲 |
| **P4 — Module Système** | `SystemDiagnostics` + Onglet 5 (Rapport) + intégration Word | 3 semaines | 🔲 |
| **P5 — Migration & UI** | Migration v2.x, sidebar ACTIF, polish | 2 semaines | 🔲 |
| **P6 — Tests & QA** | Tests intégration + visuels + régression | 2 semaines | 🔲 |
| **P7 — Documentation** | README, CHANGELOG, guide utilisateur v3.0 | 1 semaine | 🔲 |
| **P8 — Release** | Tag v3.0.0, distribution, communication | 1 semaine | 🔲 |
| **TOTAL** | | **~16 semaines** (4 mois) | |

---

## ⚠️ 9. Risques & questions ouvertes

### 9.1 Risques techniques
| Risque | Probabilité | Impact | Mitigation |
|---|---|---|---|
| Volume excessif du `.hpi` avec courbes de pompe longues | Faible | Moyen | Compression gzip optionnelle à l'écriture |
| Calcul NPSH avec données incomplètes | Moyen | Élevé | Valeurs par défaut documentées + bandeau d'avertissement |
| Incohérence entre `pump` et `hpt` (deux analyses du même débit) | Moyen | Élevé | `SystemDiagnostics.check_pump_matches_flow()` |
| Performance UI avec 100+ points de profil en long | Faible | Faible | Sous-échantillonnage pour affichage, données brutes conservées |
| Régression visuelle v2.x après refactor layout | Moyen | Moyen | Tests visuels comparatifs + beta-testeurs |

### 9.2 Questions ouvertes (à arbitrer avant P2)

1. **Formats de fichiers pompe acceptés** : en plus de CSV, faut-il supporter Excel (.xlsx) ?  
   *Recommandation : OUI, par cohérence avec les autres onglets.*

2. **Pompe à vitesse variable** : faut-il une courbe `H(Q, N)` ou juste la courbe à N constant ?  
   *Recommandation : N constant en v3.0, vitesse variable en v3.1.*

3. **Calcul de pertes de charge réseau** : faut-il les demander ou les extraire de HAMMER ?  
   *Recommandation : les deux — saisie manuelle OU extraction depuis fichier HPT (colonne "Headloss").*

4. **Ventouses sur les sections en charge vs en refoulement** : différenciation ?  
   *Recommandation : NON en v3.0 — juste le type et DN, pas le côté de la conduite.*

5. **Multilangue** : faut-il préparer l'i18n (FR/EN) dès la v3.0 ?  
   *Recommandation : NON — garder le français comme v2.x, i18n en v4.0.*

6. **Base de données de ventouses pré-dimensionnées** : inclure quelques modèles courants (GA, ARI, Venting et al.) ?  
   *Recommandation : NON en v3.0 (scope creep). Documentation des règles métier en v3.1.*

---

## 📊 10. Métriques de succès

| Métrique | Cible |
|---|---|
| Couverture de tests unitaires | > 85% |
| Bugs de régression détectés en P6 | < 5 |
| Temps de chargement d'un `.hpi` v3.0 (5 MB) | < 1 s |
| Temps d'export Word complet (15 pages) | < 5 s |
| Satisfaction beta-testeurs (échelle 1-5) | ≥ 4.2 |
| Adoption (téléchargements v3.0 / v2.x à 3 mois) | ≥ 60% |

---

## 📎 11. Annexes

### A. Arborescence cible (post v3.0)
```
HammerPy Insight/
├── main.py                          # ~3500 lignes (avec v3)
├── requirements.txt                 # inchangé
├── README.md                        # maj v3.0
├── ROADMAP.md                       # ce document
├── CHANGELOG.md                     # NOUVEAU (historique versions)
├── test_units.py                    # tests v2 (conservé)
├── test_v3_pump.py                  # NOUVEAU
├── test_v3_valves.py                # NOUVEAU
├── test_v3_integration.py           # NOUVEAU
├── test_legacy_migration.py         # NOUVEAU
├── data/                            # NOUVEAU (optionnel)
│   ├── sample_pump_curve.csv
│   ├── sample_profile.csv
│   └── legacy_v2_project.hpi        # pour test migration
└── docs/                            # NOUVEAU
    ├── user_guide_v3.md
    ├── technical_spec_v3.md
    └── wireframes/
        ├── tab2_pump.png
        └── tab4_valves.png
```

### B. Nouvelles dépendances (toutes optionnelles)
- **Aucune dépendance externe obligatoire** pour la v3.0 core
- `scipy` (interpolation cubique optionnelle) — *à confirmer en P2*
- `reportlab` (déjà présent) — pour la mise en page avancée du profil en long

### C. Glossaire
- **HPT** : Hydropneumatic Tank (réservoir anti-bélier)
- **NPSH** : Net Positive Suction Head (hauteur de charge nette à l'aspiration)
- **MOC** : Method of Characteristics (méthode des caractéristiques)
- **PN** : Pression Nominale (classe de conduite)
- **DN** : Diamètre Nominal (conduite ou appareil)
- **η** : Rendement (de la pompe)
- **V1, V2, V3** : repères de ventouses dans les plans

---

*Document rédigé le 4 juin 2026 — HammerPy Insight v2.1 → v3.0 — Roadmap*
*Prochaine revue : fin de phase P2 (module Pompe)*
