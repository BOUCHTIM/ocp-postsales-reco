# Pièges rencontrés et solutions (à relire avant de conclure quoi que ce soit)

## Données
- **Un export SAP d'un seul type de commande** (ZDS/ZDP) rend toutes les autres familles « sans document » : vérifier `Ord.Type` dans le profil de l'export avant d'interpréter la famille priceadj.
- **Numéros de document en nombre flottant** (90020837.0) côté tracker : normaliser en entier avant toute jointure (fait dans reconcile.py et generic_reconcile.py).
- **Sens du flux selon l'Incoterm** : en FOB le demurrage est dû par le vendeur à l'acheteur (ordre de paiement, pas de débit note) et le despatch est une débit note. Mapper DEMURRAGES→ZDS sans regarder DELIVERY MODE transforme des dettes payées en « créances à confirmer » et gonfle le récupérable. Trouvé par une relecture indépendante, pas par les contrôles automatiques : garder une relecture humaine du sens métier.
- **Export SAP filtré** : des n° DN/CN cités par le tracker manquent, 177 navires sur 257 absents. Conséquence : « réclamé sans document » ≠ « non refacturé ». Classer T2, demander l'export complet.
- **Une ligne tracker = deux voyages** (ABILITY : 4 DN citées, 2 au 30/09, 2 au 07/10). Solution : rattacher par n° DN cité même sur une autre date BL du même navire ; signaler dans la méthode.
- **Faute de frappe sur un n° DN** (90027811 vs 90037811) : l'écart de montant disparaît avec le bon n°. Chercher le n° à ±10000 sur le même navire avant de conclure à un écart.
- **Lots SAP réémis** (annulation puis nouvelle émission) : le premier lot peut rester avec un solde ouvert → encours gonflé. Détecter : même navire/BL/type, deux dates de création, montants répétés.
- **Encaissé** : |Montant| − |Solde| en valeurs absolues, sinon les crédits ZDP compensent les débits ZDS et le taux d'encaissement est faux.
- **Plages de dates BL** « 30/09/2024 - 01/10/2024 » : intervalle, parfois inversé.
- **Caractères de contrôle** (« ZEN\x02NOH ») : openpyxl les réécrit en `_x0002_`. Nettoyer `[\x00-\x08\x0b-\x1f]` à la lecture.
- **Lignes à montant nul** (types mixtes DESPATCH / DEMURRAGES) : garder, signaler dans les contrôles.
- **Même navire, voyage différent** : ne jamais rapprocher sur le nom seul ; mentionner « navire connu SAP, autre voyage ».

## Code
- **pandas ≥ 3** : colonnes texte en dtype `str`, pas `object`. Tester `pd.api.types.is_string_dtype`. `astype(str)` conserve les NA → `fillna("")` avant.
- **JSON** : NaN n'est pas sérialisable ; nettoyer récursivement (float NaN → None, NaT → None) avant `json.dump(..., allow_nan=False)`.
- **openpyxl** : formules sans valeur en cache → `wb.calculation.fullCalcOnLoad = True` ; vérifier avec le paquet `formulas` quand LibreOffice est absent.
- **pptxgenjs** : `apply_theme.js` a besoin de `jszip` → `NODE_PATH` vers un node_modules local contenant pptxgenjs, jszip, docx. Chartes de couleurs en hex sans `#`.
- **docx-js** : une référence de numérotation par liste numérotée ; largeurs de colonnes en DXA, sommes exactes ; prévoir ≥ 1,15" pour une date « jj/mm/aaaa » à 8 pt.
- **python-pptx** : `sh.left` peut être None sur un placeholder → reprendre la position du layout.

## Contrôle visuel sans LibreOffice (macOS)
- `qlmanage -t` ne rend que la première slide/page → `qa_render.py` fait tourner chaque slide en première position et découpe les PDF (pypdf).
- Quick Look **ne rend pas les graphiques natifs** PowerPoint : dessiner barres/entonnoirs avec des formes (`addShape`) ; rendu identique partout.
- `qlmanage -p` ouvre une fenêtre « [DEBUG] » visible par l'utilisateur, avec artefacts de superposition au défilement : ne pas l'utiliser, et la fermer (`pkill -f "qlmanage -p"`) si ouverte.
- Le panneau navigateur intégré bloque `file://` et télécharge les PDF : servir le dossier avec `python3 -m http.server` pour les HTML ; pour les PDF, passer par `qa_render.py pdf`.
- PDF depuis un .docx : `docx2html.py` puis Chrome headless `--print-to-pdf` (Arial présent sur Mac). Pagination ≈ Word, pas identique ; le dire.
- Polices Cambria/Calibri absentes du Mac : l'aperçu affiche Times/Helvetica ; vérifier les débordements avec ~10 % de marge.

## Communication
- Dire explicitement ce qui est mesuré et ce qui ne l'est pas (maillon ① absent).
- Ne pas annoncer un montant « perdu » : annoncer des scénarios (certain / probable / à confirmer / en cours / à risque).
- Chaque régénération : relancer la chaîne complète (reconcile → make_ref → expert → livrables) pour que tous les chiffres restent identiques.
