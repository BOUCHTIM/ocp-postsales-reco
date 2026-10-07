# Livrables, structure du dossier, contrôle qualité

Tout est produit par `scripts/pipeline.py` (une commande) ; les templates lisent `ref.json` (make_ref.py) et `_summary.json` (expert.py). Un livrable ne contient aucun chiffre ni nom tapé à la main : pour changer un texte, changer la donnée ou le template, puis relancer.

## Dossier livré (un seul, sur le Desktop ou là où le demandeur l'attend)

```
<Nom mission>/
├── 01 - Livrables/
│   ├── Rapprochement … .xlsx            (reconcile.py)
│   ├── Analyse experte … .xlsx          (expert.py) + _summary.json
│   ├── Presentation … .pptx + .pdf      (templates/deck.js, 12 slides)
│   ├── Rapport detaille … .docx + .pdf  (templates/report.js, ~16 pages)
│   └── Note d'expertise … .docx + .pdf  (templates/note.js, ~6 pages)
├── 02 - Sources recues/                 fichiers d'origine, non modifiés, renommés si le nom est illisible (ex. « \.xlsx »)
├── 03 - Apercus/                        PNG de chaque slide (qa_render.py ou pdftoppm)
├── 04 - Outil rejouable/                scripts + templates + run.txt (commande de rejeu) + modele_factures_fret.xlsx
├── LISEZ-MOI.txt                        contenu + chiffres clés
└── Mail de reponse.txt                  brouillon
```

## Classeur de rapprochement (6 onglets)
Synthèse (formules SUMIFS/COUNTIFS sur Rapprochement : par année, type, statut, statut tracker, méthode, région ; contrôle des écarts ; côté SAP non suivi) · Rapprochement (1 ligne par clé) · SAP non suivi · Détail tracker (toutes lignes, périmètre) · Base SAP 2024+ · Notes & hypothèses.

## Classeur d'analyse experte (10 onglets)
Tableau de bord (formules) · Chaîne par voyage (①②③④ + typologie + exposé + alertes) · Typologie (valeurs, contrôle des formules) · Délais & encours vieilli · Par client · Par région · Référentiel navires (graphies, IMO, absents de SAP) · Réémissions SAP · Contrôles (16 tests OK/KO) · Échantillon audit (≈18 dossiers, pièces à réunir, colonnes à remplir en jaune) · Plan d'action chiffré · Hypothèses & paramètres.

## Présentation (12 slides, pptxgenjs, formes natives pour les graphiques)
1 Titre · 2 L'enjeu (4 maillons + risque) · 3 Demande et données · 4 Méthode · 5 Chiffres clés (4 tuiles) · 6 Par année (barres dessinées + tableau) · 7 Entonnoir · 8 Non refacturés (T1 vs T2) · 9 Anomalies · 10 Limite majeure (maillon ①) · 11 Plan d'action · 12 Guide Excel.
Titres 24 pt sur une ligne, corps ≥ 11 pt, pas de barres décoratives.

## Rapport détaillé (docx-js)
Résumé · 1 Objet · 2 Données · 3 Méthodologie (normalisation, clé, niveaux, montants/statuts, périmètre, outils & QA) · 4 Résultats (global, année, type, statut, région) · 5 Clés sans document (répartition, interprétation export) · 6 Anomalies (écarts, surplus SAP, saisie) · 7 Refacturé non encaissé · 8 SAP non suivi · 9 Limites · 10 Recommandations · 11 Guide Excel · Annexe A (liste PAID/PENDING sans doc) · Annexe B (hypothèses).
Listes numérotées : une référence de numérotation par liste, sinon Word numérote en continu.

## Note d'expertise (docx-js, 6 pages)
1 Synthèse exécutive (tableau 5 scénarios + 3 constats) · 2 Chaîne de valeur et typologies · 3 Délais et encours vieilli · 4 Qualité des données et référentiel · 5 Échantillon d'audit · 6 Plan d'action chiffré · 7 Outil rejouable · 8 Ce qu'il manque.

## Mail de réponse
Deux registres selon le demandeur : formel (méthode, tableau par année, points d'attention a–e, pièces jointes) ou amical (5 puces de résultats, 2 demandes). Toujours : la demande d'export SAP complet et la demande des factures de fret.

## Contrôle qualité avant livraison (ne pas sauter)
1. Chiffres : tout livrable lit `ref.json` / `_summary.json` ; aucun chiffre tapé à la main. Recalculer un total en pandas et le comparer aux formules (`formulas` : `ExcelModel().loads(p).finish().calculate()`).
2. XML : `validate.py` des skills pptx/docx (ou `python-pptx`/`python-docx` qui ouvrent sans erreur).
3. Texte : grep `nan|undefined|NaN|null|_x00` dans pptx/docx.
4. Visuel : `qa_render.py pptx` puis lire chaque PNG ; `qa_render.py pdf` sur chaque PDF. Chercher : titres sur deux lignes qui chevauchent, dates coupées (« 23/12/202 5 »), en-têtes coupés (« Clé s »), listes numérotées continues, graphiques vides.
5. Cohérence : les mêmes chiffres dans Excel, slides, rapport, note, mail.

## Classeurs des modules métier (présents selon les fichiers déposés ; voir `modules.md`)
- `Order-to-cash … .xlsx` : Synthèse (formules) · Cargaisons (1 ligne par livraison, statut O2C, délais, alertes, remplissage conditionnel) · Par mois · Par client.
- `Balance agee … .xlsx` : Synthèse (tranches en formules, part > 90 j, DSO) · Postes ouverts (tranche, retard, relance) · Par client (client × tranche, crédits non imputés, net, % > 90 j) · Relances · DSO.
- `Suivi LC … .xlsx` : Synthèse · Registre (statut calculé, alerte, délais) · Alertes · Par banque · Par client.
- `Reclamations clients … .xlsx` : Synthèse · Registre (état, délai, ancienneté, hors délai) · Hors délai · Par type · Par client · Par gestionnaire · Par mois.
- `Laytime … .xlsx` : Synthèse · Escales (laytime autorisé, temps compté, demurrage / despatch calculés, écart vs payé / réclamé, contrôle) · Données manquantes.
- `Tableau de bord … .xlsx` + `Tableau de bord.md` (racine du dossier) : KPI par domaine, alertes consolidées, évolution si plusieurs mois (`--history`).
Chaque module écrit un `_summary.json` ; le tableau de bord ne lit que ces résumés. Les livrables Office (présentation, rapport, note) restent propres à la chaîne réclamations ; pour les modules, le tableau de bord Markdown sert de texte de synthèse à reprendre dans le mail.
