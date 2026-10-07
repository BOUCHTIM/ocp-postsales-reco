# Familles de réclamations post-sales et comment chacune se traite

Le moteur `reconcile.py` est générique : une **famille** = des onglets du tracker (préfixe), une correspondance « type de réclamation → types de document SAP attendus », une règle de sens du flux. Tout est dans `CONFIG["claims_families"]` (surchargeable par `--config`). `--claims demdesp|priceadj|all|<famille>`.

| Famille | Onglets tracker (préfixe) | Types de réclamation vus | Document SAP attendu | Sens du flux (défaut) | Maillon ① |
|---|---|---|---|---|---|
| `demdesp` | DEM… (DEM-DESP Tracker…) | DEMURRAGES, DESPATCH, DESPATCH / DEMURRAGES | ZDS (débit) / ZDP (crédit) | selon Incoterm : CFR demurrage = créance, despatch = dette ; FOB inversé | factures de fret armateur |
| `priceadj` | PRICE ADJ…, DISCOUNT…, REMISE…, CLAIM… | PRICE ADJUST, DISCOUNT, CLAIM(S), LC CHARGES, MARGIN SHARE, DEAD FREIGHT, TROP PERCU, Virement à Tort | `*` : tout document SAP du même navire + BL ; les n° de DN/CN cités tranchent | DISCOUNT, CLAIM, MARGIN SHARE, TROP PERCU, VIREMENT → dette ; LC CHARGES, DEAD FREIGHT → créance ; PRICE ADJUST → selon le signe du document SAP, sinon « indéterminé » | accord commercial (contrat, avenant, courriel) |

Ajouter une famille : dans un JSON passé via `--config`
```json
{"claims_families": {"lc": {"label": "LC charges", "sheet_prefix": ["LC"], "map": {"LC CHARGE": ["ZL2"], "*": ["*"]}}}}
```
et, si besoin, `--sens-config '{"LC CHARGE": "creance"}'` pour `expert.py` (valeurs : incoterm, incoterm_inverse, creance, dette, signe).

## Ce qui est commun à toutes les familles
- Clé navire + date BL (+ n° DN/CN cités), typologie T0–T5, scénarios créances / dettes, délais BL → DN → règlement, encours vieilli, contrôles, échantillon d'audit, plan chiffré, livrables.
- Un export SAP limité aux ZDS/ZDP ne contient pas les documents des remises / ajustements : la famille `priceadj` sortira alors presque entièrement en « réclamé sans document » (T2). C'est un constat sur l'export, pas sur les remises : demander l'export de **tous** les types de commande (ZL2 / ZG2 débit et crédit memos, etc.).

## Pièges propres à `priceadj`
- Les onglets n'ont souvent **pas de colonne Incoterm** : le sens du flux repose sur le type de réclamation ; « PRICE ADJUST » reste indéterminé tant qu'aucun document SAP n'est rattaché. Le dire dans la note plutôt que de forcer un sens.
- Montants parfois très élevés (ajustements de prix sur une cargaison entière) : un seul dossier peut dominer tous les agrégats ; toujours montrer la liste nominative des cinq premiers.
- Colonne « DN/CN Number/STATUS » mélange numéros et statuts (« NOT YET STARTED ») ; les numéros sont extraits par regex `9\d{7}`.
- « Virement à Tort » et « TROP PERCU » sont des remboursements : dettes, pas des réclamations à recouvrer.

## Fichiers hors familles
`profile.py` décrit n'importe quel fichier (rôle probable de chaque colonne, qualité, clés candidates, famille probable). `generic_reconcile.py` rapproche deux tables sur les clés choisies (texte normalisé, dates ± N jours, montants en valeur absolue, agrégation par clé) : factures de vente ↔ SAP, relevé bancaire ↔ règlements, deux extractions. Pour un nouveau type de fichier récurrent, créer une famille ou un petit script dédié, puis le documenter ici.
