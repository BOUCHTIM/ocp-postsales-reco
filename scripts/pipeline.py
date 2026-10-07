"""Chaîne complète en une commande : détection des fichiers → rapprochement → chiffres de référence → analyse experte
→ présentation, rapport, note (si Node disponible) → PDF (si LibreOffice ou Chrome) → aperçus (si LibreOffice+pdftoppm ou Quick Look macOS)
→ dossier organisé + LISEZ-MOI + brouillon de mail. Chaque étape optionnelle qui échoue est signalée, jamais silencieuse.

Usage minimal :
  python3 pipeline.py --inputs "<dossier des fichiers reçus>" --out "<dossier de sortie>" --mission "Nom de la mission"
Options : --start 2024-01-01  --delai-alerte 180  --config <alias colonnes.json>  --sap/--tracker/--fret pour forcer les fichiers
          --no-office (pas de pptx/docx)  --no-pdf  --setup (installe les dépendances Python et Node puis s'arrête)"""
import argparse, os, sys, subprocess, shutil, json, glob, platform, textwrap
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__)); SKILL = os.path.dirname(HERE); TPL = os.path.join(SKILL, "assets", "templates")
ap = argparse.ArgumentParser()
ap.add_argument("--inputs"); ap.add_argument("--out"); ap.add_argument("--mission", default="Rapprochement demurrage / despatch")
ap.add_argument("--sap"); ap.add_argument("--tracker", action="append"); ap.add_argument("--fret")
ap.add_argument("--start", default="2024-01-01"); ap.add_argument("--delai-alerte", type=int, default=180); ap.add_argument("--config")
ap.add_argument("--claims", default="demdesp", help="demdesp | priceadj | all")
ap.add_argument("--arrete", default=None, help="date d'arrêté AAAA-MM-JJ pour les modules o2c / balance / LC / réclamations (défaut : aujourd'hui)")
ap.add_argument("--no-office", action="store_true"); ap.add_argument("--no-pdf", action="store_true"); ap.add_argument("--setup", action="store_true")
a = ap.parse_args()
PY = sys.executable
LOG = []


def log(msg): print(msg); LOG.append(msg)


def run(cmd, label, must=True, env=None):
    log(f"→ {label}")
    p = subprocess.run(cmd, capture_output=True, text=True, env=env)
    tail = (p.stdout + p.stderr).strip().splitlines()[-3:]
    for t in tail: log("   " + t[:200])
    if p.returncode != 0:
        log(f"   ✗ {label} : échec" + (" (étape obligatoire)" if must else " (étape ignorée)"))
        if must: sys.exit(1)
        return False
    return True


# ---------------------------------------------------------------- outils disponibles
def which(*names):
    for n in names:
        p = shutil.which(n)
        if p: return p
    return None


def find_node():
    n = which("node")
    if n: return n
    for c in sorted(glob.glob(os.path.expanduser("~/.nvm/versions/node/*/bin/node")), reverse=True): return c
    return None


NODE = find_node(); NPM = which("npm") or (os.path.join(os.path.dirname(NODE), "npm") if NODE and os.path.exists(os.path.join(os.path.dirname(NODE), "npm")) else None)
SOFFICE = which("soffice", "libreoffice") or next((p for p in ["/Applications/LibreOffice.app/Contents/MacOS/soffice", "C:\\Program Files\\LibreOffice\\program\\soffice.exe"] if os.path.exists(p)), None)
PDFTOPPM = which("pdftoppm"); QLMANAGE = which("qlmanage")
CHROME = which("google-chrome", "chrome", "chromium") or next((p for p in ["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"] if os.path.exists(p)), None)
NODE_MODULES = next((d for d in [os.path.join(TPL, "node_modules"), os.environ.get("NODE_PATH", "")] if d and os.path.isdir(os.path.join(d, "pptxgenjs"))), None)

if a.setup:
    log("Installation des dépendances Python…")
    req = os.path.join(HERE, "requirements.txt")
    if not run([PY, "-m", "pip", "install", "-q", "-r", req], "pip install", must=False):
        uv = which("uv")
        if uv: run([uv, "pip", "install", "-q", "--python", PY, "-r", req], "uv pip install", must=True)
        else: log("Ni pip ni uv disponibles pour cet interpréteur : installer manuellement `pip install -r scripts/requirements.txt`."); sys.exit(1)
    if NPM:
        log("Installation des dépendances Node (pptxgenjs, docx, jszip) dans assets/templates…")
        run([NPM, "install", "--silent", "--prefix", TPL, "pptxgenjs", "docx", "jszip"], "npm install", must=False)
    else: log("npm introuvable : les livrables Office (pptx/docx) ne pourront pas être générés. Installer Node.js puis relancer --setup.")
    log(f"Outils : node={bool(NODE)} soffice={bool(SOFFICE)} pdftoppm={bool(PDFTOPPM)} qlmanage={bool(QLMANAGE)} chrome={bool(CHROME)}"); sys.exit(0)

if not a.inputs or not a.out: ap.error("--inputs et --out sont requis (ou --setup)")
OUT = os.path.abspath(a.out); os.makedirs(OUT, exist_ok=True)
L1 = os.path.join(OUT, "01 - Livrables"); L2 = os.path.join(OUT, "02 - Sources recues"); L3 = os.path.join(OUT, "03 - Apercus"); L4 = os.path.join(OUT, "04 - Outil rejouable"); WK = os.path.join(OUT, ".travail")
for d in (L1, L2, L3, L4, WK): os.makedirs(d, exist_ok=True)

# ---------------------------------------------------------------- 1. détection des fichiers
sys.path.insert(0, HERE); from detect_inputs import detect, find_sheet  # noqa
found = detect(a.inputs)
sap = a.sap or next((o["file"] for o in found if o["kind"] == "sap"), None)
trackers = a.tracker or [o["file"] for o in sorted([o for o in found if o["kind"] == "tracker"], key=lambda o: (o.get("role") != "tracker", -os.path.getsize(o["file"])))]
fret = a.fret or next((o["file"] for o in found if o["kind"] == "fret"), None)
for o in found: log(f"   {o['kind']:8s} {os.path.basename(o['file'])}")
MOD = {k: find_sheet(found, k) for k in ["orders", "deliveries", "invoices", "payments", "open_items", "sales", "lc", "claims", "laytime"]}
has_claims_chain = bool(sap and trackers)
if not has_claims_chain and not any(MOD.values()):
    log("✗ Aucun fichier exploitable : il faut un export SAP + un tracker (réclamations), ou des fichiers order-to-cash / balance âgée / LC / réclamations / laytime. Préciser avec --sap / --tracker."); sys.exit(1)
if not has_claims_chain: log("Pas de couple export SAP + tracker : la chaîne réclamations est ignorée ; modules détectés : " + ", ".join(k for k, v in MOD.items() if v))
for f in {*( [sap] if sap else []), *trackers, *([fret] if fret else []), *[v.rpartition(":")[0] for v in MOD.values() if v]}:
    dst = os.path.join(L2, os.path.basename(f))
    if os.path.abspath(f) != os.path.abspath(dst): shutil.copy2(f, dst)
tag = a.mission.replace("/", "-") + ("" if a.claims == "demdesp" else f" - {a.claims}")
RAP = os.path.join(L1, f"Rapprochement {tag}.xlsx"); ANA = os.path.join(L1, f"Analyse experte {tag}.xlsx"); REF = os.path.join(WK, "ref.json")
DECK = os.path.join(L1, f"Presentation {tag}.pptx"); RPT = os.path.join(L1, f"Rapport detaille {tag}.docx"); NOTE = os.path.join(L1, f"Note d'expertise {tag}.docx")

# ---------------------------------------------------------------- 2. rapprochement, référence, expertise
SUMM = ANA.replace(".xlsx", "_summary.json"); office_ok = False
if not has_claims_chain:
    REF = None
cmd = [PY, os.path.join(HERE, "reconcile.py"), "--sap", sap, "--out", RAP, "--start", a.start, "--claims", a.claims] if has_claims_chain else None
if cmd:
  for t in trackers: cmd += ["--tracker", t]
  if a.config: cmd += ["--config", a.config]
  run(cmd, "Rapprochement (reconcile.py)")
if cmd: run([PY, os.path.join(HERE, "make_ref.py"), RAP, REF, "--sap", sap, "--mission", a.mission, "--trackers", ";".join(trackers), "--start", a.start, "--delai-alerte", str(a.delai_alerte), "--famille", a.claims] + (["--fret", fret] if fret else []), "Chiffres de référence (make_ref.py)")
if cmd:
  cmd = [PY, os.path.join(HERE, "expert.py"), "--rapprochement", RAP, "--sap", sap, "--out", ANA, "--delai-alerte", str(a.delai_alerte)]
  if fret: cmd += ["--fret", fret]
  run(cmd, "Analyse experte (expert.py)")
# ---------------------------------------------------------------- 2b. modules métier (présents selon les fichiers détectés)
summaries = [SUMM] if (cmd and os.path.exists(SUMM)) else []
ARR = ["--arrete", a.arrete] if a.arrete else []
if MOD["orders"] and MOD["deliveries"] and MOD["invoices"]:
    o = os.path.join(L1, f"Order-to-cash {tag}.xlsx"); c2 = [PY, os.path.join(HERE, "o2c.py"), "--orders", MOD["orders"], "--deliveries", MOD["deliveries"], "--invoices", MOD["invoices"], "--out", o] + ARR + (["--payments", MOD["payments"]] if MOD["payments"] else []) + (["--config", a.config] if a.config else [])
    if run(c2, "Order-to-cash (o2c.py)", must=False): summaries.append(o.replace(".xlsx", "_summary.json"))
if MOD["open_items"]:
    o = os.path.join(L1, f"Balance agee {tag}.xlsx"); c2 = [PY, os.path.join(HERE, "aging.py"), "--open", MOD["open_items"], "--out", o] + ARR + (["--sales", MOD["sales"]] if MOD["sales"] else []) + (["--config", a.config] if a.config else [])
    if run(c2, "Balance âgée / DSO (aging.py)", must=False): summaries.append(o.replace(".xlsx", "_summary.json"))
if MOD["lc"]:
    o = os.path.join(L1, f"Suivi LC {tag}.xlsx"); c2 = [PY, os.path.join(HERE, "lc.py"), "--lc", MOD["lc"], "--out", o] + ARR + (["--config", a.config] if a.config else [])
    if run(c2, "Crédits documentaires (lc.py)", must=False): summaries.append(o.replace(".xlsx", "_summary.json"))
if MOD["claims"]:
    o = os.path.join(L1, f"Reclamations clients {tag}.xlsx"); c2 = [PY, os.path.join(HERE, "claims_register.py"), "--claims", MOD["claims"], "--out", o] + ARR + (["--config", a.config] if a.config else [])
    if run(c2, "Réclamations clients (claims_register.py)", must=False): summaries.append(o.replace(".xlsx", "_summary.json"))
if MOD["laytime"]:
    o = os.path.join(L1, f"Laytime {tag}.xlsx"); c2 = [PY, os.path.join(HERE, "laytime.py"), "--calls", MOD["laytime"], "--out", o] + (["--config", a.config] if a.config else [])
    if run(c2, "Recalcul laytime (laytime.py)", must=False): summaries.append(o.replace(".xlsx", "_summary.json"))
if summaries:
    c2 = [PY, os.path.join(HERE, "dashboard.py"), "--out", os.path.join(L1, f"Tableau de bord {tag}.xlsx"), "--md", os.path.join(OUT, "Tableau de bord.md"), "--mois", (a.arrete or str(date.today()))[:7]] + [x for sfile in summaries for x in ("--summary", sfile)]
    run(c2, "Tableau de bord (dashboard.py)", must=False)

# ---------------------------------------------------------------- 3. livrables Office
if not has_claims_chain: log("Livrables Office (présentation, rapport, note) : réservés à la chaîne réclamations ; voir le Tableau de bord pour les autres modules.")
elif a.no_office: log("Livrables Office désactivés (--no-office).")
elif not NODE: log("Node.js introuvable : présentation, rapport et note non générés. Installer Node puis `pipeline.py --setup`.")
elif not NODE_MODULES: log("Dépendances Node absentes (pptxgenjs/docx/jszip) : lancer `pipeline.py --setup`. Présentation, rapport et note non générés.")
else:
    env = dict(os.environ, NODE_PATH=NODE_MODULES)
    ok1 = run([NODE, os.path.join(TPL, "deck.js"), REF, DECK], "Présentation (deck.js)", must=False, env=env)
    ok2 = run([NODE, os.path.join(TPL, "report.js"), REF, RPT], "Rapport détaillé (report.js)", must=False, env=env)
    ok3 = run([NODE, os.path.join(TPL, "note.js"), REF, SUMM, NOTE], "Note d'expertise (note.js)", must=False, env=env)
    office_ok = ok1 and ok2 and ok3

# ---------------------------------------------------------------- 4. PDF et aperçus
if office_ok and not a.no_pdf:
    if SOFFICE:
        for f in (DECK, RPT, NOTE): run([SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", L1, f], f"PDF LibreOffice : {os.path.basename(f)}", must=False)
        if PDFTOPPM and os.path.exists(DECK[:-5] + ".pdf"): run([PDFTOPPM, "-jpeg", "-r", "110", DECK[:-5] + ".pdf", os.path.join(L3, "slide")], "Aperçus des slides (pdftoppm)", must=False)
    else:
        if CHROME:
            for f in (RPT, NOTE):
                html = os.path.join(WK, os.path.basename(f) + ".html")
                if run([PY, os.path.join(HERE, "docx2html.py"), f, html], f"HTML : {os.path.basename(f)}", must=False):
                    run([CHROME, "--headless=new", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={f[:-5]}.pdf", "file://" + html], f"PDF Chrome : {os.path.basename(f)}", must=False)
        else: log("Ni LibreOffice ni Chrome : pas de PDF pour le rapport et la note.")
        if QLMANAGE: run([PY, os.path.join(HERE, "qa_render.py"), "pptx", DECK, L3, "--pdf", DECK[:-5] + ".pdf"], "Aperçus des slides + PDF image (Quick Look)", must=False)
        else: log("Pas de LibreOffice ni de Quick Look : pas d'aperçu ni de PDF pour la présentation.")

# ---------------------------------------------------------------- 5. outil rejouable, LISEZ-MOI, mail
for f in ("reconcile.py", "expert.py", "make_ref.py", "detect_inputs.py", "pipeline.py", "qa_render.py", "docx2html.py", "pptx2html.py", "requirements.txt", "xlsx_common.py", "o2c.py", "aging.py", "lc.py", "claims_register.py", "laytime.py", "dashboard.py", "profile.py", "generic_reconcile.py", "make_templates.py"):
    shutil.copy2(os.path.join(HERE, f), os.path.join(L4, f))
os.makedirs(os.path.join(L4, "assets", "templates"), exist_ok=True)
for f in glob.glob(os.path.join(TPL, "*.js")): shutil.copy2(f, os.path.join(L4, "assets", "templates", os.path.basename(f)))
shutil.copy2(os.path.join(SKILL, "assets", "modele_factures_fret.xlsx"), os.path.join(L4, "modele_factures_fret.xlsx"))
os.makedirs(os.path.join(L4, "assets", "modeles"), exist_ok=True)
for f in glob.glob(os.path.join(SKILL, "assets", "modeles", "*.xlsx")): shutil.copy2(f, os.path.join(L4, "assets", "modeles", os.path.basename(f)))
ref = json.load(open(REF)) if REF and os.path.exists(REF) else {"total": {"n": 0, "trk": 0, "sap": 0, "enc": 0, "solde": 0, "nonref": 0, "nonref_amt": 0}, "dncn": {"trouves": 0, "listes": 0}, "vessels": {"absent_sap": 0, "tracker": 0}, "mission": {"arrete": "", "periode": ""}}; t = ref["total"]; summ = json.load(open(SUMM)) if os.path.exists(SUMM) else {}
sc = {s["Scénario"]: s["Montant (USD)"] for s in summ.get("scenarios", [])}
fmt = lambda v: f"{v:,.0f}".replace(",", " ")
musd = lambda v: f"{v / 1e6:.2f} MUSD".replace(".", ",")
pct = lambda x, y: f"{100 * x / y:.1f} %".replace(".", ",") if y else "–"
with open(os.path.join(L4, "run.txt"), "w") as f:
    f.write("Pour rejouer avec de nouveaux fichiers :\n\n  python3 pipeline.py --inputs \"../02 - Sources recues\" --out \"..\" --mission \"" + a.mission + "\" --start " + a.start + f" --delai-alerte {a.delai_alerte}" + (f" --arrete {a.arrete}" if a.arrete else "") + "\n\n"
            "Premier lancement sur une nouvelle machine : python3 pipeline.py --setup\nFactures de fret armateur : remplir modele_factures_fret.xlsx, le déposer dans 02 - Sources recues, relancer.\n"
            "Autres modules (order-to-cash, balance âgée / DSO, crédits documentaires, réclamations clients, laytime) : remplir les modèles de assets/modeles/ (ou déposer les exports SAP équivalents), les mettre dans 02 - Sources recues, relancer : chaque module détecté est produit et le Tableau de bord les consolide.\n")
lisez = f"""{a.mission.upper()} — DOSSIER GÉNÉRÉ LE {date.today():%d/%m/%Y}
Clé : navire + date BL · Périmètre BL ≥ {a.start} · Délai d'alerte sans DN : {a.delai_alerte} j · Arrêté SAP : {ref['mission']['arrete']}

01 - Livrables     {"Rapprochement (xlsx) · Analyse experte (xlsx) · Présentation (pptx" + ('/pdf' if office_ok else ' — non générée') + ") · Rapport détaillé (docx) · Note d'expertise (docx)" if has_claims_chain else "chaîne réclamations non lancée (pas de couple export SAP + tracker)"}{" · " + " · ".join(os.path.basename(x).replace("_summary.json", "").replace(" " + tag, "") + " (xlsx)" for x in summaries if "Analyse experte" not in x) if len(summaries) > (1 if has_claims_chain else 0) else ""}{" · Tableau de bord (xlsx + Tableau de bord.md)" if summaries else ""}
02 - Sources recues  fichiers d'origine (non modifiés)
03 - Apercus       images des slides (si générées)
04 - Outil rejouable scripts + run.txt : tout se recalcule avec de nouveaux exports

CHIFFRES CLÉS
  {t['n']} clés navire + BL · réclamé {musd(t['trk'])} · refacturé SAP {musd(t['sap'])} ({pct(t['sap'], t['trk'])}) · encaissé {musd(t['enc'])} ({pct(t['enc'], t['sap'])} du refacturé) · solde client {musd(t['solde'])}
  {t['nonref']} clés sans document SAP ({musd(t['nonref_amt'])}) · n° DN/CN cités retrouvés : {ref['dncn']['trouves']}/{ref['dncn']['listes']} · navires du tracker absents de SAP : {ref['vessels']['absent_sap']}/{ref['vessels']['tracker']}
  Scénarios (créances) : récupérable certain {musd(sc.get('Récupérable certain', 0))} · probable {musd(sc.get('Récupérable probable', 0))} · à confirmer {musd(sc.get('À confirmer (export filtré ?)', 0))} · en cours {musd(sc.get('En cours de traitement', 0))} · risque de prescription {musd(sc.get('Risque de prescription', 0))}
  Dettes : sans document dans l'export {musd(sc.get("Dettes sans document dans l'export", 0))} · à régler / provisionner {musd(sc.get('Dettes à régler ou provisionner', 0))}
  Maillon « payé à l'armateur » : {'fourni' if fret else 'ABSENT — demander les factures de fret'}

JOURNAL DE GÉNÉRATION
""" + "\n".join("  " + l for l in LOG)
open(os.path.join(OUT, "LISEZ-MOI.txt"), "w").write(lisez)
mail = f"""Objet : {a.mission} — résultats du rapprochement

Bonjour,

Voici les résultats du rapprochement surestaries sur la période {ref['mission']['periode']} (clé navire + date BL, BL à partir du {a.start}).

- {t['n']} dossiers suivis pour {musd(t['trk'])} ; {musd(t['sap'])} refacturés dans SAP ({pct(t['sap'], t['trk'])}), {musd(t['enc'])} encaissés ({pct(t['enc'], t['sap'])} du refacturé).
- {t['nonref']} dossiers sans document SAP ({musd(t['nonref_amt'])}) ; {ref['dncn']['trouves']} n° de DN/CN sur {ref['dncn']['listes']} cités sont retrouvés dans l'export : l'export semble filtré, à confirmer avant toute conclusion.
- Récupérable certain (créances refacturées non encaissées) : {musd(sc.get('Récupérable certain', 0))} ; à risque de prescription (non démarré > {a.delai_alerte} j) : {musd(sc.get('Risque de prescription', 0))}.
- Dettes envers les acheteurs (crédit notes, demurrage FOB) : {musd(sc.get('Dettes à régler ou provisionner', 0))} à régler ou provisionner.

Deux demandes pour aller au bout : l'export SAP complet (toutes entités / familles produit) et {'les factures de fret armateur (navire, date BL, montant payé) pour mesurer le maillon « payé → récupéré ».' if not fret else 'la validation des factures de fret fournies.'}

Pièces jointes : classeur de rapprochement, classeur d'analyse experte, présentation, rapport détaillé, note d'expertise.

Cordialement,
"""
open(os.path.join(OUT, "Mail de reponse.txt"), "w").write(mail)
log(f"\nTerminé. Dossier : {OUT}")
