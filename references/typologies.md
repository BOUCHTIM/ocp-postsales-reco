# Chaîne de valeur, typologies de fuite, scénarios de récupération

## La chaîne en quatre maillons

| # | Maillon | Source | Ce qui peut se perdre |
|---|---|---|---|
| ① | Payé à l'armateur | Factures de fret finales (demurrage réglé, despatch perçu) | Rien n'est réclamé au client → perte sèche |
| ② | Réclamé au client | Tracker post-sales (FINAL AMOUNT, statut, n° DN/CN) | Réclamation partielle, tardive, jamais émise |
| ③ | Refacturé | SAP : ZDS = Demurrage Order → **débit note** (montant +) ; ZDP = Dispatch Order → **crédit note** (montant −) | Pas de document, document annulé non réémis, montant différent |
| ④ | Encaissé / imputé | SAP : Montant facturé − Solde client (en valeurs absolues) | Solde ouvert, relance absente, prescription |

Un montant qui ne franchit pas un maillon est une fuite. Le type de fuite dicte l'action ; ne jamais présenter un total unique « non récupéré ».

## Sens du flux : ne jamais additionner créances et dettes

| Incoterm | Demurrage | Despatch |
|---|---|---|
| CFR / CIF / DPU (vendeur paie le fret, retard au déchargement chez l'acheteur) | **Créance** : débit note ZDS à l'acheteur | **Dette** : crédit note ZDP à l'acheteur |
| FOB (acheteur affrète ; retard au port de chargement du vendeur) | **Dette** : le vendeur paie l'acheteur (ordre de paiement / swift, parfois crédit note) | **Créance** : débit note à l'acheteur |

Conséquences vues sur le cas de référence : 136 clés FOB sans aucun document dans l'export ZDS/ZDP (les paiements FOB passent par des ordres de paiement, pas par des DN) ; 0,82 MUSD de « payé sans document » étaient en réalité des sorties de trésorerie vers l'acheteur ; les « récupérables » passaient de 0,75 à 0,55 MUSD une fois les crédit notes retirées. Lire la colonne DELIVERY MODE / Mode de vente avant toute interprétation ; `expert.py` calcule « Sens du flux » et sépare « Créance exposée » et « Dette exposée ».

## Typologies (une par clé navire + date BL + type)

Ordre d'évaluation : T1/T2 (aucun doc SAP) → T3 (écart) → T4/T5 (solde) → T0.

| Code | Définition | Montant exposé | Lecture |
|---|---|---|---|
| T0 | Documents SAP trouvés, solde client ≤ tolérance | 0 | Soldé |
| T1 | Statut tracker NOT YET STARTED, aucun doc SAP, âge BL ≤ délai d'alerte | Réclamé | En cours, attendu |
| T1b | Idem mais âge BL > délai d'alerte (défaut 180 j) | Réclamé | **Risque de prescription** : plus le dossier vieillit, plus les pièces disparaissent et le client conteste |
| T2 | Statut PAID / PENDING, aucun doc SAP | Réclamé | **À confirmer** : très souvent export SAP filtré, pas fuite réelle |
| T3 | Docs trouvés, écart tracker − SAP > max(1 USD ; 0,5 %) | Solde | Écart positif (SAP < tracker) = refacturation incomplète probable |
| T4 | Docs trouvés, encaissé ≈ 0, solde > tolérance | Solde | **Récupérable certain** : relance client |
| T5 | Docs trouvés, encaissé partiel, solde > tolérance | Solde | Récupérable certain |

## Scénarios à présenter (toujours séparés)

| Scénario | Composition | Action |
|---|---|---|
| Récupérable certain | T4 + T5, **créances seulement** | Relance clients, lettrage ; vérifier les lots SAP réémis non annulés qui gonflent l'encours |
| Récupérable probable | T3, part où SAP < tracker | Compléter la refacturation ou corriger le tracker |
| À confirmer | T2 | Relancer l'export SAP complet avant toute conclusion |
| En cours | T1 | Suivre le délai d'émission |
| Risque de prescription | T1b créances | Émettre en priorité ; vérifier la clause contractuelle de délai de réclamation |
| Dettes sans document dans l'export | T2 dettes (FOB demurrage payé par OP, crédit notes absentes) | Rien à récupérer ; vérifier la comptabilisation |
| Dettes à régler / provisionner | T1, T1b, T4, T5 dettes | Émettre les crédit notes, provisionner ; jamais dans le « récupérable » |

## Indicateurs de délai (toujours calculer, c'est souvent le vrai sujet)

- BL → émission DN/CN (date de création SAP) : médiane, moyenne, P90, par année.
- DN/CN → premier règlement (Date Reglement 1).
- Âge des dossiers non démarrés depuis le BL, par tranches 0-90 / 91-180 / 181-365 / > 365 j.
- Encours vieilli des documents ouverts : retard vs échéance SAP à la date d'arrêté (Report Run Date), tranches non échu / 1-30 / 31-60 / 61-90 / 91-180 / > 180 j, par client.

Sur le cas de référence (2026) : délai médian BL → DN de 195 jours, dossiers non démarrés âgés de 344 jours en médiane, encours ouvert à 100 % au-delà de 90 jours. La fuite principale était un défaut d'émission, pas de recouvrement. Vérifier si c'est aussi le cas plutôt que de supposer.

## Montant exposé et récupérable : règles de prudence

- Exposé = réclamé si aucun document SAP, sinon solde client. Jamais la somme des deux.
- Ne pas additionner T2 au « récupérable » tant que l'export complet n'est pas reçu.
- Donner les montants par scénario, par année, par client, par région ; le total agrégé n'a pas de sens opérationnel.
- Toujours écrire la limite : si le maillon ① manque, l'indicateur « payé → récupéré » n'est pas calculé ; dire ce qui est mesuré à la place.
