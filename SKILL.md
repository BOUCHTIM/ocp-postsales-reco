---
name: ocp-postsales-reco
description: Customer-service / post-sales toolkit for a commodity exporter (OCP and similar) — reconcile claim trackers (demurrage, despatch, discounts, price adjustments, LC charges, dead freight, claims) with SAP debit/credit notes on the vessel + BL-date key, classify leakage and cash direction, measure delays; plus order-to-cash per cargo (orders, deliveries, invoices, payments), aged balance / DSO / dunning list, letter-of-credit tracking, customer-claims register with team workload, laytime recalculation, and a monthly management dashboard consolidating everything. Profiles any unknown Excel/CSV, reconciles any two tables, produces all deliverables in one command. Use for any OCP / SAP post-sales, surestaries, demurrage, despatch, laytime, remises, débit/crédit notes, DN/CN, factures, balance âgée, DSO, LC, réclamations clients, tableau de bord request — even from bare Excel files with a short French instruction.
---

# Post-sales / Customer Service : rapprochement, suivi et expertise

Métier couvert : celui d'un **Customer Service Director en commodity trading**. Après une vente export, l'argent circule entre vendeur, armateur, banques et acheteur. Le travail : suivre chaque cargaison de la commande à l'encaissement, chaque réclamation de son fait générateur à son règlement, chaque crédit documentaire jusqu'au paiement, et dire à la direction combien est récupérable, combien est dû, où sont les retards et qui les porte. Tous les modules partent de fichiers Excel / CSV (exports SAP ou registres tenus par l'équipe), ne les modifient jamais, et produisent des classeurs à synthèse en formules plus un résumé JSON consolidé dans un tableau de bord.

Lire `references/modules.md` (entrées, règles, limites de chaque module), `references/familles.md` (familles de réclamations), `references/typologies.md` avant d'interpréter un chiffre, `references/colonnes.md` avant de lire un fichier, `references/pieges.md` avant de conclure, `references/livrables.md` pour la forme. `GUIDE-UTILISATEUR.md` est le mode d'emploi destiné à l'utilisateur.

## 1. Fichier inconnu ? Profiler d'abord
```bash
python3 scripts/profile.py "<fichier.xlsx|csv>" profil.md --json profil.json
```
Rôle probable de chaque colonne, qualité (dates / nombres en texte, espaces, plages), clés candidates, famille probable. Lire le profil, puis choisir : chaîne réclamations (§2), module métier (§3), deux tables (§4) ou analyse ad hoc en pandas. Dire à l'utilisateur ce que le fichier est et n'est pas avant de calculer.

## 2. Tout le dossier en une commande
```bash
python3 scripts/pipeline.py --setup                                   # une fois par machine (pip/uv + npm)
python3 scripts/pipeline.py --inputs "<dossier reçu>" --out "<dossier mission>" --mission "Post-sales 2026-09" \
    [--claims demdesp|priceadj|all] [--start 2024-01-01] [--delai-alerte 180] [--arrete 2026-09-30] [--config aliases.json]
```
`detect_inputs.py` donne un rôle à chaque onglet (export SAP, tracker + familles, factures de fret, commandes, livraisons, factures, règlements, postes ouverts, CA mensuel, registre LC, registre réclamations, escales). Puis, **selon ce qui est présent** :
- export SAP + tracker → `reconcile.py` (clé navire + BL, n° DN/CN cités, ±3 j, une ligne par clé), `make_ref.py` (ref.json : source unique des chiffres et phrases), `expert.py` (typologie T0–T5, sens créance / dette par Incoterm et type, délais, encours, référentiel navires, 16 contrôles, échantillon d'audit, plan chiffré), puis présentation / rapport / note / PDF / aperçus ;
- commandes + livraisons + factures (+ règlements) → `o2c.py` ; postes ouverts (+ CA) → `aging.py` ; registre LC → `lc.py` ; registre réclamations → `claims_register.py` ; escales → `laytime.py` ;
- toujours à la fin : `dashboard.py` consolide tous les `_summary.json` (KPI, alertes, évolution avec `--history`), puis `LISEZ-MOI.txt`, `Mail de reponse.txt`, `04 - Outil rejouable` (scripts + modèles).
Chaque étape absente ou en échec est écrite dans le journal : le dire, ne pas le cacher. Colonnes différentes → alias dans `--config` (voir `assets/config_exemple.json` : sections `tracker_columns`, `orders`, `deliveries`, `invoices`, `payments`, `open`, `sales`, `lc`, `claims`, `calls`).

## 3. Modules un par un (quand un seul sujet est demandé)
```bash
python3 scripts/o2c.py --orders VA05.xlsx --deliveries VL06.xlsx --invoices VF05.xlsx --payments FBL5N.xlsx --out o2c.xlsx --arrete 2026-09-30
python3 scripts/aging.py --open "FBL5N.xlsx:Postes ouverts" --sales ca.xlsx --out balance.xlsx --arrete 2026-09-30 --relance 30
python3 scripts/lc.py --lc registre_lc.xlsx --out lc.xlsx --arrete 2026-09-30 --alerte 15
python3 scripts/claims_register.py --claims registre.xlsx --out reclamations.xlsx --sla 30
python3 scripts/laytime.py --calls escales.xlsx --out laytime.xlsx
python3 scripts/dashboard.py --out tdb.xlsx --md tdb.md --mois 2026-09 --summary o2c_summary.json --summary balance_summary.json ... [--history <dossier des mois précédents>]
```
Modèles de fichiers : `assets/modeles/modele_{o2c,balance_agee,lc,reclamations,laytime}.xlsx` (en-têtes attendues, ligne d'exemple, légende). Les règles (statut O2C, tranches, DSO count-back, alertes LC et UCP 600 21 j, SLA réclamations, laytime = quantité ÷ cadence × 24 h, temps compté = brut − exceptions) sont dans `references/modules.md`. Ces modules ont été construits sur des structures SAP standard et validés sur un jeu synthétique (`scripts/make_synth.py`) : **au premier fichier réel, profiler, ajouter les alias, relire une ligne à la main**.

## 4. Deux tables quelconques à rapprocher
```bash
python3 scripts/generic_reconcile.py --left A.xlsx:Onglet --right B.xlsx:Onglet --out rapp.xlsx \
    --key "VESSEL=Navire" --key "B/L DATE=Date BL" --amount "FINAL AMOUNT=Montant Facturée" --abs --date-tolerance 3 --agg sum
```
Texte normalisé, numéros en entier, dates ± N jours, montants en valeur absolue, agrégation par clé ; sortie : Rapprochés, Gauche seule, Droite seule, Écarts, Synthèse en formules.

## 5. Lire comme un expert
- Onglet « Contrôles » / « Alertes » : chaque KO devient une phrase du compte rendu.
- Réclamations : scénarios toujours séparés (récupérable certain T4/T5, probable T3, à confirmer T2 — export souvent filtré —, en cours T1, risque de prescription T1b) et à part les dettes. Ne jamais additionner créances et dettes ni annoncer un total « perdu ». Sens du flux : CFR demurrage = créance, despatch = dette ; FOB inversé ; remises / claims / remboursements = dettes ; LC charges / dead freight = créances ; ajustement de prix = selon le signe du document, sinon « indéterminé ».
- Order-to-cash : une cargaison livrée non facturée est une fuite de trésorerie immédiate ; comparer facturé et quantité × prix avant d'accuser le recouvrement.
- Balance âgée : lire la part > 90 j et la liste de relance par client, pas seulement le DSO ; sans CA mensuel, le DSO est approximé (le dire).
- LC : une LC expirée sans documents présentés, ou des réserves, bloque le paiement de la cargaison : priorité absolue.
- Réclamations clients : charge par gestionnaire et délai vs SLA avant les montants ; taux d'acceptation par type dit où la qualité fuit.
- Laytime : un montant payé ou réclamé différent du calcul contractuel est un écart à qualifier, pas une erreur prouvée (SHINC/SHEX, turn time, exceptions non saisies).
- Délais et âges : la fuite principale est souvent un défaut d'émission, pas de recouvrement. Mesurer avant d'affirmer. Relire le sens métier au-delà des contrôles automatiques.

## 6. Adapter, contrôler, livrer
Les textes sont générés depuis `ref.json` / `_summary.json` / `Tableau de bord.md` ; adapter le ton sans ressaisir un chiffre. Avant de livrer : formules évaluées (`formulas` sans LibreOffice) et égales aux agrégats pandas ; ouverture sans erreur (openpyxl / python-pptx / python-docx) ; grep `nan|undefined|NaN|null|_x00` ; rendu visuel de chaque slide et page (`scripts/qa_render.py`) ; mêmes chiffres partout. Corriger à la source, relancer, re-vérifier.

Dans la réponse : tableau des chiffres clés par domaine, scénarios créances / dettes, anomalies nominatives, limites (maillon manquant, export filtré, sens indéterminés, modules validés sur synthétique seulement) et demandes au demandeur (exports SAP complets, factures de fret, contrats, registres à jour).

## Ressources
- `scripts/pipeline.py`, `detect_inputs.py`, `reconcile.py`, `make_ref.py`, `expert.py` — chaîne réclamations.
- `scripts/o2c.py`, `aging.py`, `lc.py`, `claims_register.py`, `laytime.py`, `dashboard.py`, `xlsx_common.py` — modules métier ; `make_templates.py`, `make_synth.py` — modèles et jeu d'essai.
- `scripts/profile.py`, `generic_reconcile.py` — fichier inconnu, deux tables.
- `scripts/qa_render.py`, `docx2html.py`, `pptx2html.py` — contrôle visuel et PDF sans LibreOffice.
- `assets/templates/*.js` ; `assets/modeles/*.xlsx` ; `assets/modele_factures_fret.xlsx` ; `assets/config_exemple.json`.
- `references/modules.md`, `familles.md`, `typologies.md`, `colonnes.md`, `livrables.md`, `pieges.md` ; `GUIDE-UTILISATEUR.md`.
