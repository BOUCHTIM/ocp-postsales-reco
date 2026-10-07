# Guide utilisateur — skill « ocp-postsales-reco »

Ce skill apprend à Claude le travail d'un Customer Service en commodity trading : rapprochement des réclamations suivies (surestaries demurrage / despatch, remises, ajustements de prix, LC charges, dead freight, claims…) avec les débit / crédit notes SAP et analyse experte ; suivi order-to-cash par cargaison (commande → BL → facture → encaissement) ; balance âgée, DSO et liste de relance ; suivi des crédits documentaires (LC) ; registre des réclamations clients avec charge par gestionnaire ; recalcul des laytime ; tableau de bord mensuel de direction qui consolide le tout. Il sait aussi décrire un fichier inconnu et rapprocher deux tables quelconques. Vous n'avez rien à calculer : vous déposez vos fichiers et vous décrivez la demande.

## 1. Installer le skill

**Claude Code (ordinateur)** : copier le dossier `ocp-postsales-reco/` dans `~/.claude/skills/` (Mac / Linux) ou `%USERPROFILE%\.claude\skills\` (Windows). Au premier usage, demander à Claude : « installe les dépendances du skill ocp-postsales-reco » (il lance `scripts/pipeline.py --setup`). Prérequis : Python 3.10+ et Node.js 18+. Pour les PDF et les aperçus visuels, LibreOffice est recommandé (sinon Chrome et, sur Mac, Quick Look suffisent).

**Claude.ai (web / app)** : dans Paramètres → Capacités → Skills, importer le fichier `ocp-postsales-reco.skill`. Les dépendances sont installées automatiquement dans l'environnement de Claude.

## 2. Utiliser

Déposez dans un dossier (ou joignez à la conversation) :
- l'export SAP des commandes ZDS / ZDP (débit et crédit notes) ;
- le ou les fichiers de suivi post-sales (tracker DEM / DESP) ;
- si disponibles, les factures de fret armateur au format du modèle `assets/modele_factures_fret.xlsx`.

Puis écrivez par exemple :
> « Voici l'export SAP et les trackers post-sales. Clé navire + date BL, à partir de 2024. Fais le rapprochement complet et l'analyse experte des surestaries, nom de mission "Surestaries 2024-2026". »

Autres demandes possibles : « fais la même chose pour les remises et ajustements de prix » (famille priceadj), « tout en une fois » (famille all), « décris-moi ce fichier, je ne sais pas ce qu'il contient » (profileur), « rapproche ces deux fichiers sur le numéro de facture » (rapprochement générique).

**Autres sujets du poste** — déposez les fichiers correspondants (exports SAP ou registres tenus par l'équipe ; les modèles sont dans `assets/modeles/`), seuls ou avec les précédents, et demandez par exemple :
- « suivi order-to-cash des cargaisons du trimestre » : commandes (VA05), livraisons / BL (VL06), factures (VF05), règlements (FBL5N) ;
- « balance âgée et DSO au 30/09, liste de relance » : postes ouverts clients (FBL5N) et, si possible, le CA mensuel ;
- « où en sont les LC » : registre des crédits documentaires ;
- « état des réclamations clients et charge de l'équipe » : registre des réclamations ;
- « recalcule les laytime et compare au payé » : tableau des escales (NOR, début / fin, cadence, taux) ;
- « tableau de bord du mois » : Claude produit automatiquement `Tableau de bord.xlsx` + `.md` dès qu'un module a tourné, avec l'évolution si les mois précédents sont fournis.
Claude lance chaque module dont les fichiers sont présents ; le tableau de bord les consolide.

Claude identifie les fichiers, lance la chaîne, vérifie les résultats et vous rend un dossier :

```
<Nom de mission>/
├── 01 - Livrables/      Rapprochement.xlsx · Analyse experte.xlsx · Presentation.pptx (+pdf) · Rapport detaille.docx (+pdf) · Note d'expertise.docx (+pdf)
│                        + selon les fichiers : Order-to-cash.xlsx · Balance agee.xlsx · Suivi LC.xlsx · Reclamations clients.xlsx · Laytime.xlsx · Tableau de bord.xlsx
├── 02 - Sources recues/ vos fichiers, non modifiés
├── 03 - Apercus/        une image par slide
├── 04 - Outil rejouable/ scripts + run.txt : relancer avec de nouveaux exports
├── LISEZ-MOI.txt        chiffres clés + journal de génération
└── Mail de reponse.txt  brouillon
```

## 3. Ce que vous obtenez

- **Rapprochement** : une ligne par voyage (navire + date BL + type), documents SAP trouvés, montant refacturé, encaissé, solde, écart, alertes ; synthèse en formules.
- **Analyse experte** : typologie de fuite par voyage (soldé, non démarré récent / ancien, réclamé sans document, écart, refacturé non encaissé), sens du flux (créance ou dette selon l'Incoterm), scénarios de récupération, délais BL → DN → règlement, encours vieilli, référentiel navires, réémissions SAP, 16 contrôles automatiques, échantillon d'audit, plan d'action chiffré.
- **Présentation** 12 slides, **rapport** détaillé, **note d'expertise** 6 pages, **mail** : tous générés à partir des mêmes chiffres.
- **Order-to-cash** : une ligne par cargaison, statut (livré non facturé, facturé non encaissé, partiel, encaissé), facturé vs quantité × prix, délais BL → facture → règlement, par mois et par client.
- **Balance âgée** : encours par client et par tranche de retard, crédits non imputés, DSO, liste de relance.
- **Suivi LC** : alertes (expiration, documents non présentés, réserves, paiement attendu), délais, frais par banque.
- **Réclamations clients** : backlog, délai vs SLA, taux d'acceptation, par type / client / gestionnaire / mois.
- **Laytime** : demurrage / despatch recalculés par escale et comparés au payé / réclamé.
- **Tableau de bord** : les indicateurs de tous les modules, les alertes, l'évolution mensuelle ; résumé en Markdown prêt pour un mail.

## 4. Paramètres utiles (à dire à Claude)

| Paramètre | Défaut | Quand le changer |
|---|---|---|
| Date de début de périmètre | 2024-01-01 | « à partir de 2023 » |
| Délai d'alerte sans DN | 180 jours | aligner sur la clause contractuelle de délai de réclamation |
| Tolérance de rapprochement sur la date BL | 3 jours | fichiers avec dates BL imprécises |
| Alias de colonnes | détection automatique | si vos fichiers ont d'autres noms de colonnes (voir `assets/config_exemple.json`) |
| Famille de réclamations | demdesp | priceadj (remises / ajustements) ou all ; nouvelle famille via le même fichier de configuration |
| Date d'arrêté | aujourd'hui | balance âgée, LC, réclamations, order-to-cash : « au 30/09/2026 » |
| SLA réclamations / délai de relance / alerte LC | 30 j / 30 j / 15 j | aligner sur vos procédures |

## 5. Lire les résultats sans se tromper

- Un dossier « réclamé sans document SAP » n'est pas forcément perdu : l'export SAP est souvent filtré. Demandez l'export complet avant de conclure.
- En **FOB**, le demurrage est dû par le vendeur à l'acheteur : c'est une dette, pas une créance. Remises, claims et remboursements sont des dettes ; LC charges et dead freight des créances ; un ajustement de prix dépend du signe du document SAP. Le classeur d'analyse sépare créances et dettes ; seules les créances sont « récupérables ».
- Un export SAP limité aux surestaries (ZDS/ZDP) ne contient pas les documents des remises : demandez l'export de tous les types de commande avant de lire la famille « remises ».
- Sans les factures de fret armateur, on mesure « suivi → refacturé → encaissé », pas « payé à l'armateur → récupéré ». Le modèle de fichier permet d'ajouter ce maillon.
- Les délais (BL → émission de la DN, DN → règlement) sont souvent le vrai sujet : regardez-les avant les montants.
- Les modules order-to-cash, balance âgée, LC, réclamations clients et laytime ont été construits sur des structures SAP standard et testés sur un jeu de données d'essai, pas encore sur vos fichiers : au premier usage, faites vérifier une ligne à la main par Claude (il le propose) et donnez-lui les noms de colonnes qui diffèrent. Le laytime recalculé est un contrôle, pas un décompte contractuel : les clauses SHINC/SHEX, turn time et exceptions doivent être saisies en heures d'exception.

## 6. Relancer avec de nouveaux fichiers

Déposez les nouveaux exports dans `02 - Sources recues` et demandez à Claude de « rejouer le rapprochement », ou lancez vous-même la commande de `04 - Outil rejouable/run.txt`. Tout se recalcule ; les chiffres restent cohérents entre tous les livrables.
