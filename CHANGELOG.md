# Changelog — HammerPy Insight

## v3.1 (Juin 2026 — En cours)

### Nouvelles fonctionnalités
- **Model Builder** : import de classeur Excel libre avec mapping interactif des colonnes
  - Sélection de la feuille Canalisations (label, nœud départ, nœud arrivée, longueur, Ø, matériau, rugosité)
  - Sélection de la feuille Nœuds/Jonctions (label, X, Y, Z)
  - Mapping interactif via le ColumnMapper existant (auto-apprentissage)
  - Validation de la connectivité (nœuds sans position détectés)
  - Résumé du modèle après construction
- PN Requis Fascicule 74 : 4e KPI dans le strip transitoire (vert/orange/rouge)
- Export des 4 graphiques au format PNG/PDF/SVG (boutons 💾 Export)
- USER_GUIDE.md et CHANGELOG.md créés

---

## v3.0 (Juin 2026)

### Nouvelles fonctionnalités
- Phase 4 — **SystemDiagnostics** (16 vérifications croisées A–E) avec onglet UI dédié et section Word
- 4 graphiques imprimables dans la section 6 du rapport Word (diagnostics)
- **Templates XLSX** pour tous les imports : HPT enveloppe, régime permanent, données pompe, courbe H(Q), profil ventouse, profil Bentley, classeur FlexTables (6 feuilles)
- **Export DXF** du profil en long avec ventouses et vidanges (3 calques : Profil en long, Ventouses, Vidanges)
- **Import XLSX** des données pompe + courbe H(Q) multipoints
- **Pmin admissible** : option -0.5 bar ajoutée
- **Métré des canalisations** par DN et matériau dans le rapport Word

### Corrections
- Support POLYLINE (HAMMER) dans l'import DXF (ne cherchait que LWPOLYLINE)
- `set_pos` → `set_placement` dans l'export DXF (API ezdxf)
- Couleur hex `#1a1a2e` correctement passée à Tkinter (toolbar pump curve)
- Typo `ventaises` → `ventouses` dans les rapports, UI et documentation
- Import courbe H(Q) : `flow_lps`/`head_m` stockés dans `curve_points` au lieu de `parsed`
- Message d'erreur DXF liste tous les calques disponibles
- `openpyxl` importé globalement (évite freeze UI + KeyboardInterrupt)
- Noms de colonnes normalisés : `flow_lps`, `head_m`, `debit_lps`, `hmt_m`, `hauteur_m`

### Technique
- Python 3.10+, CustomTkinter, pandas, openpyxl, ezdxf, matplotlib, numpy
- 235 tests unitaires (pytest)
- ROADMAP.md mis à jour (Phase 4.2)

---

## v2.0 (Mai 2026)

### Nouvelles fonctionnalités
- Phase 3 — **AirValveSizing** : dimensionnement ventouses et vidanges
- Profil en long (CSV, XLSX, Bentley FlexTable) avec détection points hauts/bas
- **Rapport Word** (.docx) dédié Ventouses/Vidanges avec profil en long
- **Mapping interactif** des colonnes CSV avec auto-apprentissage
- **Import DXF** du tracé en plan et profil en long
- Sauvegarde/chargement projet (.hpi) avec sérialisation complète
- Section « Profil en Long » dans le rapport Word (tableaux ventouses + vidanges)

### Corrections
- Détection robuste des colonnes (français/anglais, accents, séparateurs)
- Encodage automatique UTF-8 / CP1252 / Latin-1 pour les CSV

---

## v1.0 (Avril 2026)

### Première version
- Parsing des fichiers CSV/Excel HAMMER (enveloppe transitoire, régime permanent station)
- Visualisation graphique interactive avec seuils critiques (PN, Pmin)
- Export de rapport Word (.docx) avec tableau de bord + graphique intégré
- Import et parsing des rapports pompe Bentley HAMMER (.rtf/.txt)
- Gestion multi-pompe (batterie) avec courbes H(Q)
- Import classeur HAMMER Flex Tables (.xlsx)
