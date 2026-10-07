# Fichiers d'entrée : colonnes attendues, alias, pièges de lecture

Les scripts renomment automatiquement les colonnes vers les noms canoniques ci-dessous (alias insensibles à la casse, surchargeables via `--config`, voir `assets/config_exemple.json`). Vérifier l'en-tête réel avec un `pd.read_excel(..., nrows=5)` avant de lancer : un alias manquant produit une colonne vide, pas une erreur.

## Tracker post-sales (un onglet par période, préfixe « DEM »)

| Canonique | Alias courants | Usage | Pièges vus |
|---|---|---|---|
| VESSEL | Navire, Vessel name | clé | espaces finaux (« YC FORTITUDE  »), casse variable (Ability / ABILITY), un même navire sous deux graphies |
| B/L DATE | BL date, Date BL | clé | plages « 15/10/2024 - 17/10/2024 », parfois à l'envers ; lire comme intervalle |
| Claims Type | Type | DEMURRAGES ↔ ZDS, DESPATCH ↔ ZDP, « DESPATCH / DEMURRAGES » ↔ les deux | lignes mixtes à montant nul |
| FINAL AMOUNT | Montant final, Amount | montant réclamé (positif) | montant saisi en texte « 15 425,22 » |
| DN/CN Number | DN/CN, N° DN/CN | n° de documents SAP cités, séparés par « - » ou retours ligne | fautes de frappe (90027811 pour 90037811), « NOT YET STARTED » dans la cellule |
| PAYMENT STATUS | Statut | PAID / PENDING / NOT YET STARTED, texte libre | « CMO/OP-N°30/25 : PAID CMO/OP N°31/25 : PAID » → normaliser par mots-clés |
| RECEIVER | Client, Customer | client | caractères de contrôle (« ZEN\x02NOH ») : retirer `[\x00-\x1f]` |
| REGION, DELIVERY MODE, PRODUCT, INVOICE NUMBER, REMARK, Origin | | contexte | INVOICE NUMBER = factures de vente, jamais présentes dans l'export ZDS/ZDP |

Onglets « PRICE ADJ / DISCOUNTS » : remises, ajustements de prix, LC charges, claims, dead freight, trop-perçus, virements à tort. Famille `priceadj` (`--claims priceadj`) : mêmes colonnes sauf « DN/CN Number/STATUS » (numéros et statuts mélangés) et « REMARKs », souvent sans DELIVERY MODE. Ne pas les mélanger avec les surestaries dans un même livrable sans le dire (montants d'un autre ordre de grandeur).

Plusieurs fichiers tracker (ex. CTD + APAC) : le premier est maître ; vérifier si les autres sont des sous-ensembles (clé navire + BL + type + montant) avant de les fusionner.

## Export SAP ZDS / ZDP (onglet « Feuil1 » ou premier onglet)

| Canonique | Alias | Usage |
|---|---|---|
| Numéro Facture | Billing document | identifiant du document (cité par le tracker) |
| Ord.Type | Order type | ZDS = débit (demurrage), ZDP = crédit (despatch) ; codes surchargeables |
| Navire, Date BL | Vessel, BL date | clé |
| Navir IMO | IMO | référentiel navire (le tracker n'en a pas) |
| Montant Facturée | Net value | signé : ZDS +, ZDP − |
| Solde Client | Open amount | signé ; |Montant| − |Solde| = encaissé / imputé |
| Date Création Facture | Created on | date d'émission → délai BL → DN |
| Echéance, Date Reglement 1 | Due date, Clearing date | encours vieilli, délai DN → règlement |
| Report Run Date | | date d'arrêté pour tous les âges et retards |
| Entité Description, S.Grp.Description, Client Name, Devise | | filtres, régions, contrôle d'un export filtré |

Signes qu'un export est **filtré** (très fréquent) : n° DN/CN cités par le tracker absents ; régions entières vides côté SAP alors que le tracker y a des dossiers ; tableau croisé dans le fichier avec « (Plusieurs éléments) » sur l'entité ou « Fertilizer » sur la famille produit ; majorité des navires du tracker absents de SAP. Dans ce cas, classer les dossiers réclamés sans document en T2 « à confirmer », jamais en « non refacturé ».

## Factures de fret armateur (maillon ①, optionnel) — `assets/modele_factures_fret.xlsx`

Obligatoires : Navire, Date BL, Montant demurrage payé, Date paiement. Utiles : laytime, temps réel, taux USD/jour (recalcul contractuel), montant despatch reçu, devise. Une ligne par voyage. Même normalisation du navire que le tracker.
