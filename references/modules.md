# Modules métier du Customer Service Director (commodity trading) — entrées, règles, sorties, limites

Tous les modules partagent `scripts/xlsx_common.py` : lecture tolérante (`fichier.xlsx:Onglet`, alias de colonnes `--config`, nettoyage des textes / dates / nombres en texte), classeur Arial avec synthèse en formules, résumé `_summary.json` consommé par `dashboard.py`. Les modèles de fichiers sont dans `assets/modeles/` (ligne d'exemple en bleu + légende). Sans fichier réel du client, les colonnes attendues sont une convention raisonnable : **toujours profiler le fichier réel et ajouter des alias plutôt que de forcer le fichier**.

| Module | Script | Entrées (modèle) | Ce qu'il produit | Règles / limites |
|---|---|---|---|---|
| Réclamations tracker ↔ SAP | `pipeline.py --claims` (reconcile / make_ref / expert) | export SAP, trackers, (fret) | rapprochement, typologie, sens, délais, livrables | voir `typologies.md`, `familles.md` |
| Order-to-cash par cargaison | `o2c.py` | commandes (VA05), livraisons / BL (VL06), factures (VF05), règlements (FBL5N) — `modele_o2c.xlsx` | une ligne par livraison : statut (livré non facturé / facturé non encaissé / partiel / encaissé), quantité livrée vs commandée, facturé vs quantité × prix, délais BL → facture → règlement, retard vs échéance, par mois, par client | jointure par n° de livraison, sinon n° de commande ; alerte « non facturé > N j » (`--delai-facture`, 15 j) ; montant attendu = quantité livrée × prix unitaire (si prix fourni) |
| Balance âgée / DSO / relances | `aging.py` | postes ouverts clients (FBL5N) + CA mensuel optionnel — `modele_balance_agee.xlsx` | par client × tranche (non échu, 1-30, 31-60, 61-90, 91-180, > 180 j), crédits non imputés à part, DSO count-back, liste de relance (> `--relance` j, 30 par défaut) | retard = arrêté − échéance ; DSO exact seulement avec le CA mensuel, sinon approximé sur les factures ouvertes (le dire) |
| Crédits documentaires | `lc.py` | registre LC — `modele_lc.xlsx` | alertes (LC expirée docs non présentés, BL après date limite, docs présentés après expiration, non présentés > 21 j UCP 600, réserves, expire sous N j, paiement attendu), délais BL → présentation → paiement, frais par banque / client | `--alerte` (15 j) ; « réglé » = date de paiement ou statut PAID/SETTLED/CLOSED ; réserves = nombre ou liste « ; » |
| Réclamations clients (qualité, quantité, documentation, litiges) | `claims_register.py` | registre — `modele_reclamations.xlsx` | backlog, ancienneté, délai de traitement vs SLA (`--sla`, 30 j), taux d'acceptation, par type / client / **gestionnaire (charge d'équipe)** / mois, liste hors délai | clôturé = date de clôture ou statut CLOSED/SETTLED/REJECTED/PAID |
| Recalcul laytime | `laytime.py` | escales (SOF résumé) — `modele_laytime.xlsx` | laytime autorisé, temps compté, demurrage / despatch calculés, écart vs montant payé / réclamé, données manquantes | laytime = quantité ÷ cadence × 24 h ou valeur fournie ; temps compté = brut − exceptions ; despatch = 50 % du taux demurrage si absent ; ne gère pas seul SHINC/SHEX, turn time, NOR hors heures : les traduire en « Exceptions (h) » |
| Tableau de bord direction | `dashboard.py` | les `_summary.json` des modules (+ dossier des tableaux de bord précédents pour l'évolution) | KPI par domaine, alertes consolidées, évolution mensuelle, résumé Markdown | ne recalcule rien : reflète les classeurs détaillés |
| Fichier inconnu | `profile.py` | tout xlsx / csv | profil par onglet, rôles, qualité, clés candidates, famille probable | point de départ avant toute analyse ad hoc |
| Deux tables | `generic_reconcile.py` | deux xlsx / csv | rapprochés / gauche seule / droite seule / écarts | clés texte normalisées, dates ± N j, montants en valeur absolue |

## Détection automatique (`detect_inputs.py`)
Chaque onglet reçoit un rôle (sap, tracker, fret, orders, deliveries, invoices, payments, open_items, sales, lc, claims, laytime, autre) d'après ses en-têtes et son nom ; `pipeline.py` lance chaque module dont les entrées sont présentes, puis le tableau de bord. Les quatre tables order-to-cash peuvent être quatre fichiers ou quatre onglets d'un même classeur. Vérifier la détection affichée en tête du journal : un onglet mal classé se force avec `--sap` / `--tracker` / `--fret` ou en renommant l'onglet.

## Indicateurs attendus d'un Customer Service Director et où les trouver
- Taux de refacturation et d'encaissement des réclamations, montants récupérables vs dettes → Analyse experte / Tableau de bord.
- Délai BL → facture, facture → encaissement, cargaisons livrées non facturées → Order-to-cash.
- DSO, encours > 90 j, liste de relance → Balance âgée.
- LC en risque (expiration, réserves), délais documentaires, frais bancaires → Suivi LC.
- Backlog et délai de traitement des réclamations, taux d'acceptation, charge par gestionnaire → Réclamations clients.
- Demurrage contractuel vs payé / réclamé → Laytime.
- Évolution mois par mois → Tableau de bord (`--history`).

## Ce qui reste hors champ sans données supplémentaires
Recalcul des prix (formules indexées), contrôle des lettres de crédit documentaire ligne à ligne (UCP), prévision de trésorerie, scoring crédit clients, qualité produit (laboratoire). Chacun peut devenir un module sur le même modèle : script + modèle Excel + alias + résumé JSON + ligne dans le tableau de bord.
