import re
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.table import Table, TableStyleInfo

import argparse
_ap = argparse.ArgumentParser(description="Rapprochement tracker post-sales DEM/DESP <-> SAP ZDS/ZDP (clé navire + date BL)")
_ap.add_argument("--sap", required=True, help="export SAP ZDS/ZDP (onglet Feuil1)")
_ap.add_argument("--tracker", action="append", required=True, help="fichier tracker post-sales (répéter : --tracker CTD.xlsx --tracker APAC.xlsx ; le premier est le fichier maître)")
_ap.add_argument("--out", required=True, help="classeur de sortie .xlsx")
_ap.add_argument("--start", default="2024-01-01", help="date BL de début de périmètre (AAAA-MM-JJ)")
_ap.add_argument("--tolerance-jours", type=int, default=3)
_ap.add_argument("--config", default=None, help="JSON de configuration (alias de colonnes, feuilles, codes de type) ; voir assets/config_exemple.json")
_ap.add_argument("--claims", default="demdesp", help="famille(s) de réclamations : demdesp | priceadj | all | nom défini dans config claims_families")
_a = _ap.parse_args()
EXP = _a.sap
TRACKERS = _a.tracker
OUT = _a.out
START = pd.Timestamp(_a.start)
TOL_DAYS = _a.tolerance_jours
import json as _json
CONFIG = {
    "sap_sheet": "Feuil1",
    "tracker_sheet_prefix": ["DEM"],
    "claims_families": {
        "demdesp": {"label": "demurrage / despatch", "sheet_prefix": ["DEM"], "map": {"DEMURRAGES": ["ZDS"], "DEMURRAGE": ["ZDS"], "DESPATCH": ["ZDP"], "DISPATCH": ["ZDP"]}},
        "priceadj": {"label": "remises, ajustements de prix et réclamations diverses", "sheet_prefix": ["PRICE ADJ", "DISCOUNT", "REMISE", "CLAIM"], "map": {"*": ["*"]}}},
    "sap_type_debit": "ZDS", "sap_type_credit": "ZDP",
    "tracker_columns": {
        "VESSEL": ["VESSEL", "NAVIRE", "VESSEL NAME", "SHIP"], "B/L DATE": ["B/L DATE", "BL DATE", "DATE BL", "B/L", "BL"],
        "REGION": ["REGION", "RÉGION"], "DELIVERY MODE": ["DELIVERY MODE", "INCOTERM", "MODE DE VENTE"], "RECEIVER": ["RECEIVER", "CLIENT", "CUSTOMER", "BUYER"],
        "PRODUCT": ["PRODUCT", "PRODUIT"], "INVOICE NUMBER": ["INVOICE NUMBER", "INVOICE", "FACTURE VENTE", "SALES INVOICE"],
        "Claims Type": ["CLAIMS TYPE", "CLAIM TYPE", "TYPE", "TYPE DE RÉCLAMATION"], "FINAL AMOUNT": ["FINAL AMOUNT", "AMOUNT", "MONTANT", "MONTANT FINAL"],
        "DN/CN Number": ["DN/CN NUMBER", "DN/CN NUMBER/STATUS", "DN/CN", "DN CN NUMBER", "DEBIT NOTE", "N° DN/CN"], "PAYMENT STATUS": ["PAYMENT STATUS", "STATUS", "STATUT"], "REMARK": ["REMARK", "REMARKS", "COMMENTAIRE"],
        "Origin": ["ORIGIN", "ORIGINE", "ENTITÉ"]},
    "sap_columns": {
        "Numéro Facture": ["NUMÉRO FACTURE", "BILLING DOCUMENT", "N° FACTURE", "INVOICE"], "Date Création Facture": ["DATE CRÉATION FACTURE", "CREATED ON", "DATE CRÉATION"],
        "Ord.Type": ["ORD.TYPE", "ORDER TYPE", "TYPE COMMANDE"], "Client Name": ["CLIENT NAME", "CUSTOMER NAME", "NOM CLIENT"], "Navire": ["NAVIRE", "VESSEL", "SHIP"],
        "Date BL": ["DATE BL", "BL DATE", "B/L DATE"], "Montant Facturée": ["MONTANT FACTURÉE", "MONTANT FACTURE", "NET VALUE", "BILLED AMOUNT"], "Solde Client": ["SOLDE CLIENT", "OPEN AMOUNT", "SOLDE"],
        "Devise": ["DEVISE", "CURRENCY"], "Report Run Date": ["REPORT RUN DATE", "DATE EXTRACTION"], "Navir IMO": ["NAVIR IMO", "IMO"], "S.Grp.Description": ["S.GRP.DESCRIPTION", "SALES GROUP", "RÉGION"],
        "Entité Description": ["ENTITÉ DESCRIPTION", "COMPANY", "ENTITÉ"], "Echéance": ["ECHÉANCE", "ÉCHÉANCE", "DUE DATE"], "Date Reglement 1": ["DATE REGLEMENT 1", "DATE RÈGLEMENT 1", "CLEARING DATE"]}}
if _a.config:
    _user = _json.load(open(_a.config))
    for k, v in _user.items():
        if isinstance(v, dict) and isinstance(CONFIG.get(k), dict): CONFIG[k].update(v)
        else: CONFIG[k] = v


_fams = list(CONFIG["claims_families"]) if _a.claims == "all" else [f.strip() for f in _a.claims.split(",")]
for _f in _fams:
    if _f not in CONFIG["claims_families"]: raise SystemExit(f"Famille inconnue : {_f}. Disponibles : {list(CONFIG['claims_families'])}")
SHEET_PREFIXES = [p for f in _fams for p in CONFIG["claims_families"][f]["sheet_prefix"]]
CLAIMS_MAP = {}
for _f in _fams:
    for k, v in CONFIG["claims_families"][_f]["map"].items(): CLAIMS_MAP[k.upper()] = v
FAMILY_LABEL = " + ".join(CONFIG["claims_families"][f]["label"] for f in _fams)


def types_for(claim_type):
    """Types de document SAP attendus pour un type de réclamation du tracker ('*' = n'importe lequel)."""
    c = str(claim_type).upper()
    hits = [v for k, v in CLAIMS_MAP.items() if k != "*" and k in c]
    if len(hits) >= 2: return "/".join(sorted({t for h in hits for t in h}))
    if hits: return "/".join(hits[0])
    return "/".join(CLAIMS_MAP.get("*", ["*"]))


def rename_cols(df, mapping):
    """Renomme les colonnes du fichier vers les noms canoniques du script, via les alias (insensible à la casse et aux espaces)."""
    norm_cols = {re.sub(r"\s+", " ", str(c)).strip().upper(): c for c in df.columns}
    ren = {}
    for canon, aliases in mapping.items():
        if canon in df.columns: continue
        for a in aliases:
            if a.upper() in norm_cols: ren[norm_cols[a.upper()]] = canon; break
    return df.rename(columns=ren)


def norm(v):
    if pd.isna(v):
        return ""
    s = str(v).upper()
    s = re.sub(r"\bM/?V\b", " ", s)
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def parse_bl(v):
    """Return (start, end, raw_text). Handles datetime or 'dd/mm/yyyy - dd/mm/yyyy'."""
    if isinstance(v, (datetime, pd.Timestamp)):
        t = pd.Timestamp(v).normalize()
        return t, t
    s = str(v)
    ds = [pd.to_datetime(x, dayfirst=True, errors="coerce") for x in re.findall(r"\d{1,2}/\d{1,2}/\d{4}", s)]
    ds = [d for d in ds if pd.notna(d)]
    if not ds:
        t = pd.to_datetime(s, errors="coerce", dayfirst=True)
        return (t, t) if pd.notna(t) else (pd.NaT, pd.NaT)
    return min(ds), max(ds)


def parse_amt(v):
    if pd.isna(v):
        return np.nan
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(" ", "").replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return np.nan


def nums_in(v):
    return re.findall(r"\b9\d{7}\b", str(v)) if pd.notna(v) else []


def pay_norm(v):
    s = str(v).upper()
    if "NOT YET" in s:
        return "NOT YET STARTED"
    if "PAID" in s and "PENDING" not in s:
        return "PAID"
    if "PENDING" in s:
        return "PENDING"
    if "CLOSED" in s:
        return "CLOSED"
    return s.strip() or ""


# ---------------------------------------------------------------- SAP export
_xs = pd.ExcelFile(EXP)
exp = pd.read_excel(EXP, sheet_name=CONFIG["sap_sheet"] if CONFIG["sap_sheet"] in _xs.sheet_names else _xs.sheet_names[0])
exp = rename_cols(exp, CONFIG["sap_columns"])
for _c in ["Montant Facturé en MAD", "Solde Client en MAD", "Quantité Facturée", "Description Produit", "S.Off.Description", "Numéro de Commande", "Date Comptabilisation", "Ord.Type Description", "Type Commande Descp", "Numéro Comptable", "Client", "Année BL", "Mode de Vente", "Encaissé", "Montant Reglement 1", "Date Reglement 2", "Montant Reglement 2", "Date Reglement 3", "Montant Reglement 3", "Payment Terms Description", "SBU Description", "Désignation pays destination", "Nom de l'ADV", "Report Run Date", "Navir IMO", "S.Grp.Description", "Entité Description", "Echéance", "Date Reglement 1", "Devise"]:
    if _c not in exp.columns: exp[_c] = np.nan
if exp["Report Run Date"].isna().all(): exp["Report Run Date"] = pd.Timestamp.today().normalize()
if exp["Année BL"].isna().all(): exp["Année BL"] = pd.to_datetime(exp["Date BL"]).dt.year
exp["Ord.Type"] = exp["Ord.Type"].astype(str).str.strip().str.upper().replace({CONFIG["sap_type_debit"]: "ZDS", CONFIG["sap_type_credit"]: "ZDP"})
exp["Numéro Facture"] = exp["Numéro Facture"].astype(str)
exp["VK"] = exp["Navire"].map(norm)
exp["BL"] = pd.to_datetime(exp["Date BL"]).dt.normalize()
exp["Solde abs"] = exp["Solde Client"].abs()
exp["Encaissé"] = exp["Montant Facturée"].abs() - exp["Solde abs"]
sap = exp[exp["BL"] >= START].copy().reset_index(drop=True)

# ---------------------------------------------------------------- trackers
frames = []
for src, f in [("CTD" if i == 0 else f"T{i + 1}", f) for i, f in enumerate(TRACKERS)]:
    for sh in pd.ExcelFile(f).sheet_names:
        if not any(sh.upper().startswith(p.upper()) for p in SHEET_PREFIXES):
            continue
        d = rename_cols(pd.read_excel(f, sheet_name=sh), CONFIG["tracker_columns"])
        for _c in ["REGION", "DELIVERY MODE", "RECEIVER", "PRODUCT", "INVOICE NUMBER", "DN/CN Number", "PAYMENT STATUS", "REMARK", "Origin"]:
            if _c not in d.columns: d[_c] = np.nan
        d = d[[c for c in d.columns if not str(c).startswith("Unnamed")]]
        d["SOURCE"] = src
        d["ONGLET"] = sh
        frames.append(d)
tr = pd.concat(frames, ignore_index=True)
tr = tr[tr["VESSEL"].notna()].copy()
for c in tr.columns:
    if pd.api.types.is_string_dtype(tr[c]) or tr[c].dtype == object:
        tr[c] = tr[c].map(lambda v: re.sub(r"_x[0-9A-Fa-f]{4}_|[\x00-\x08\x0b-\x1f]", "", v) if isinstance(v, str) else v)
tr["VK"] = tr["VESSEL"].map(norm)
bl = tr["B/L DATE"].map(parse_bl)
tr["BL_DEB"] = [x[0] for x in bl]
tr["BL_FIN"] = [x[1] for x in bl]
tr["MONTANT"] = tr["FINAL AMOUNT"].map(parse_amt)
tr["CT"] = tr["Claims Type"].astype(str).str.upper().str.strip()
tr["TYPES_SAP"] = tr["CT"].map(types_for)
tr["PAY"] = tr["PAYMENT STATUS"].map(pay_norm)
tr["DNCN_LIST"] = tr["DN/CN Number"].map(nums_in)

# row key for APAC-vs-CTD dedupe
tr["ROWKEY"] = tr["VK"] + "|" + tr["BL_DEB"].astype(str) + "|" + tr["CT"] + "|" + tr["MONTANT"].round(2).astype(str)
ctd_keys = set(tr.loc[tr.SOURCE == "CTD", "ROWKEY"])            # CTD = fichier maître (premier --tracker)
apac_keys = set(tr.loc[tr.SOURCE != "CTD", "ROWKEY"])           # lignes présentes dans un fichier secondaire
master = pd.concat([tr[tr.SOURCE == "CTD"], tr[(tr.SOURCE != "CTD") & (~tr.ROWKEY.isin(ctd_keys))].drop_duplicates("ROWKEY")], ignore_index=True)
master["DANS_APAC"] = master["ROWKEY"].isin(apac_keys).map({True: "Oui", False: "Non"})
master["PERIMETRE"] = np.where(master["BL_DEB"].isna(), "Date BL illisible",
                        np.where(master["BL_DEB"] >= START, "2024+", "Avant 2024 (exclu)"))

# ---------------------------------------------------------------- matching (key level)
master["CLE"] = master["VK"] + " | " + master["BL_DEB"].dt.strftime("%Y-%m-%d").fillna("?") + \
    np.where(master["BL_FIN"] != master["BL_DEB"], "→" + master["BL_FIN"].dt.strftime("%Y-%m-%d").fillna(""), "") + \
    " | " + np.where(master["TYPES_SAP"] == "*", master["CT"], master["TYPES_SAP"])
sap_by_v = {v: g for v, g in sap.groupby("VK")}
sap_doc_index = sap.set_index("Numéro Facture")


def find_docs(vk, d0, d1, types):
    g = sap_by_v.get(vk)
    if g is None or pd.isna(d0):
        return sap.iloc[0:0], "Non trouvé"
    tl = types.split("/")
    gt = g if "*" in tl else g[g["Ord.Type"].isin(tl)]
    ex = gt[(gt.BL >= d0) & (gt.BL <= d1)]
    if len(ex):
        return ex, "Exact (navire + date BL)"
    near = gt[(gt.BL >= d0 - pd.Timedelta(days=TOL_DAYS)) & (gt.BL <= d1 + pd.Timedelta(days=TOL_DAYS))]
    if len(near):
        return near, f"Probable (navire + BL ±{TOL_DAYS}j)"
    other = g[(g.BL >= d0 - pd.Timedelta(days=TOL_DAYS)) & (g.BL <= d1 + pd.Timedelta(days=TOL_DAYS))]
    if len(other) and "*" not in tl:
        return sap.iloc[0:0], "Non trouvé (SAP a l'autre type ZDS/ZDP sur ce voyage)"
    if len(g):
        return sap.iloc[0:0], "Non trouvé (navire connu SAP, autre voyage)"
    return sap.iloc[0:0], "Non trouvé (navire absent SAP)"


def joinu(s, sep=" ; "):
    vals = []
    for x in s:
        if pd.isna(x):
            continue
        x = str(x).strip()
        if x and x not in vals:
            vals.append(x)
    return sep.join(vals)


scope = master[master.PERIMETRE == "2024+"].copy()
rows = []
used_docs = set()
for cle, g in scope.groupby("CLE", sort=False):
    r0 = g.iloc[0]
    docs_key, how = find_docs(r0.VK, r0.BL_DEB, r0.BL_FIN, r0.TYPES_SAP)
    listed = sorted({n for L in g.DNCN_LIST for n in L})
    listed_in_sap = [n for n in listed if n in sap_doc_index.index]
    # listed DN/CN found in SAP on the same vessel but another BL date (tracker line covering 2 voyages)
    gv = sap_by_v.get(r0.VK, sap.iloc[0:0])
    by_num = gv[gv["Numéro Facture"].isin(listed) & ~gv["Numéro Facture"].isin(docs_key["Numéro Facture"])]
    if len(by_num):
        docs_all = pd.concat([docs_key, by_num])
        other_bl = joinu(by_num["BL"].dt.strftime("%d/%m/%Y"))
        how = (f"Exact (navire + date BL) + n° DN/CN cités sur autre BL SAP ({other_bl})" if len(docs_key)
               else f"n° DN/CN cités, sur autre BL SAP ({other_bl})")
    else:
        docs_all = docs_key
    listed_in_docs = [n for n in listed if n in set(docs_all["Numéro Facture"])]
    if listed_in_docs:
        docs = docs_all[docs_all["Numéro Facture"].isin(listed_in_docs)]
        if not by_num.__len__():
            how = how.replace("Exact (navire + date BL)", "Exact (navire + date BL + n° DN/CN)")
    else:
        docs = docs_all
    extra = docs_all[~docs_all["Numéro Facture"].isin(docs["Numéro Facture"])]
    used_docs.update(docs_all["Numéro Facture"].tolist())
    sap_zds = docs.loc[docs["Montant Facturée"] > 0, "Montant Facturée"].sum()   # débits (ZDS ou tout document positif)
    sap_zdp = docs.loc[docs["Montant Facturée"] < 0, "Montant Facturée"].sum()   # crédits (ZDP ou tout document négatif)
    sap_abs = abs(sap_zds) + abs(sap_zdp)
    sap_solde = docs["Solde abs"].sum()
    sap_enc = docs["Encaissé"].sum()
    trk_amt = g.MONTANT.sum()
    ecart = trk_amt - sap_abs if len(docs) else np.nan
    if len(docs) == 0:
        statut = "Non refacturé (aucun DN/CN SAP)"
    elif abs(sap_solde) < 0.01:
        statut = "Refacturé – soldé"
    elif abs(sap_enc) < 0.01:
        statut = "Refacturé – non encaissé"
    else:
        statut = "Refacturé – partiellement encaissé"
    ecart_flag = "" if len(docs) == 0 else ("OK" if abs(ecart) <= max(1.0, 0.005 * max(trk_amt, 1)) else ("SAP < tracker" if ecart > 0 else "SAP > tracker"))
    rows.append({
        "Clé (navire | BL | type SAP)": cle,
        "Navire": joinu(g.VESSEL),
        "Date BL (début)": r0.BL_DEB.to_pydatetime(),
        "Date BL (fin)": r0.BL_FIN.to_pydatetime(),
        "Année BL": int(r0.BL_DEB.year),
        "Type réclamation": joinu(g["Claims Type"]),
        "Type SAP attendu": r0.TYPES_SAP,
        "Famille": FAMILY_LABEL,
        "Docs SAP – types": joinu(docs["Ord.Type"]),
        "Région (tracker)": joinu(g.REGION),
        "Mode livraison": joinu(g["DELIVERY MODE"]),
        "Client (tracker)": joinu(g.RECEIVER),
        "Origine": joinu(g.Origin),
        "Produit": joinu(g.PRODUCT),
        "Nb lignes tracker": len(g),
        "Source(s) tracker": joinu(g.SOURCE + " / " + g.ONGLET),
        "Présent fichier APAC": joinu(g.DANS_APAC),
        "Factures vente (tracker)": joinu(g["INVOICE NUMBER"]),
        "Montant tracker (FINAL AMOUNT)": round(trk_amt, 2),
        "DN/CN n° (tracker)": joinu(g["DN/CN Number"]),
        "Statut paiement (tracker)": joinu(g.PAY),
        "Remarque (tracker)": joinu(g.REMARK),
        "Méthode de rapprochement": how,
        "Nb docs SAP": len(docs),
        "Docs SAP (n° facture)": " ; ".join(docs["Numéro Facture"].tolist()),
        "Docs SAP – date création": joinu(docs["Date Création Facture"].dt.strftime("%d/%m/%Y")),
        "Client SAP": joinu(docs["Client Name"]),
        "Entité SAP": joinu(docs["Entité Description"]),
        "Région SAP": joinu(docs["S.Grp.Description"]),
        "SAP ZDS – Débit (demurrage)": round(sap_zds, 2),
        "SAP ZDP – Crédit (despatch)": round(sap_zdp, 2),
        "SAP montant refacturé (abs)": round(sap_abs, 2),
        "SAP encaissé / imputé (abs)": round(sap_enc, 2),
        "SAP solde client (abs)": round(sap_solde, 2),
        "Écart tracker – SAP": round(ecart, 2) if pd.notna(ecart) else None,
        "Contrôle écart": ecart_flag,
        "Autres docs SAP même voyage (nb)": len(extra),
        "Autres docs SAP même voyage (n°)": " ; ".join(extra["Numéro Facture"].tolist()),
        "Autres docs SAP même voyage (montant)": round(extra["Montant Facturée"].sum(), 2),
        "Alerte": "Docs SAP supplémentaires non cités dans le tracker (annulation / réémission ?)" if len(extra) and listed_in_docs else "",
        "DN/CN tracker listés": len(listed),
        "…dont trouvés dans SAP": len(listed_in_sap),
        "…dont dans les docs rapprochés": len(listed_in_docs),
        "Statut rapprochement": statut,
        "Devise SAP": joinu(docs["Devise"]),
    })
rap = pd.DataFrame(rows)

# ---------------------------------------------------------------- SAP side not tracked
sap["Rapproché tracker"] = sap["Numéro Facture"].isin(used_docs).map({True: "Oui", False: "Non"})
tracked_vessel_keys = set(scope.VK)
nt = sap[sap["Rapproché tracker"] == "Non"].groupby(["VK", "BL", "Ord.Type"], sort=False).agg(
    Navire=("Navire", "first"), Nb_docs=("Numéro Facture", "size"),
    Docs=("Numéro Facture", lambda s: " ; ".join(s)),
    Client=("Client Name", joinu), Entite=("Entité Description", joinu), Region=("S.Grp.Description", joinu),
    Montant=("Montant Facturée", "sum"), Encaisse=("Encaissé", "sum"), Solde=("Solde abs", "sum"),
    Creation=("Date Création Facture", "max")).reset_index()
nt["Navire dans tracker (autre voyage/type)"] = nt["VK"].isin(tracked_vessel_keys).map({True: "Oui", False: "Non"})

# ---------------------------------------------------------------- write workbook
wb = Workbook()
FONT = "Arial"
hdr_fill = PatternFill("solid", fgColor="1F3864")
hdr_font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
base_font = Font(name=FONT, size=10)
bold = Font(name=FONT, size=10, bold=True)
title_font = Font(name=FONT, size=14, bold=True, color="1F3864")
thin = Side(style="thin", color="BFBFBF")
border = Border(top=thin, bottom=thin, left=thin, right=thin)
MONEY = '#,##0.00;[Red](#,##0.00);"-"'
PCT = "0.0%"
DATE = "DD/MM/YYYY"


def write_table(ws, df, start_row=1, money_cols=(), date_cols=(), int_cols=(), widths=None, name=None):
    for j, c in enumerate(df.columns, 1):
        cell = ws.cell(row=start_row, column=j, value=c)
        cell.font = hdr_font; cell.fill = hdr_fill; cell.border = border
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for i, rec in enumerate(df.itertuples(index=False), start_row + 1):
        for j, v in enumerate(rec, 1):
            if isinstance(v, float) and np.isnan(v):
                v = None
            if isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            cell = ws.cell(row=i, column=j, value=v)
            cell.font = base_font; cell.border = border
            col = df.columns[j - 1]
            if col in money_cols:
                cell.number_format = MONEY
            elif col in date_cols:
                cell.number_format = DATE
            elif col in int_cols:
                cell.number_format = "0"
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    ws.row_dimensions[start_row].height = 42
    last = get_column_letter(len(df.columns))
    ws.auto_filter.ref = f"A{start_row}:{last}{start_row + len(df)}"
    for j, c in enumerate(df.columns, 1):
        w = (widths or {}).get(c)
        if w is None:
            mx = max([len(str(c))] + [len(str(x)) for x in df[c].head(200).tolist() if x is not None])
            w = min(max(10, mx + 2), 45)
        ws.column_dimensions[get_column_letter(j)].width = w
    return start_row + len(df)


# ---- Sheet 1: Synthèse
ws = wb.active
ws.title = "Synthèse"
R = "Rapprochement"
ws["A1"] = f"Rapprochement {FAMILY_LABEL} – Tracker post-sales vs SAP"; ws["A1"].font = title_font
ws["A2"] = f"Périmètre : lignes tracker avec date BL ≥ 01/01/2024. Clé de rapprochement : nom du navire + date BL. Export SAP daté du {exp['Report Run Date'].max():%d/%m/%Y}. Généré le {datetime.now():%d/%m/%Y}."
ws["A2"].font = Font(name=FONT, size=9, italic=True)
ws["A3"] = "Indicateur majeur : part des demurrages/despatch suivis par le post-sales effectivement refacturés (débit/crédit note SAP rattachée à la vente) et part encaissée."
ws["A3"].font = Font(name=FONT, size=9, italic=True)

# column letters in Rapprochement
col = {c: get_column_letter(i + 1) for i, c in enumerate(rap.columns)}
n = len(rap) + 1
rng = lambda c: f"{R}!${col[c]}$2:${col[c]}${n}"

def kpi_block(r0, title, crit_col=None, crit_vals=None):
    """Block: one row per crit value + total. Returns next row."""
    ws.cell(row=r0, column=1, value=title).font = bold
    heads = ["", "Nb clés navire+BL", "Nb non refacturées", "Montant tracker (USD)", "Montant refacturé SAP (USD)", "% refacturé",
             "Encaissé / imputé SAP (USD, abs)", "% encaissé (sur refacturé)", "Solde client restant SAP (USD, abs)", "Montant non refacturé (USD)"]
    for j, h in enumerate(heads, 1):
        c = ws.cell(row=r0 + 1, column=j, value=h); c.font = hdr_font; c.fill = hdr_fill; c.border = border
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r0 + 1].height = 40
    r = r0 + 2
    vals = list(crit_vals) + ["Total"]
    for v in vals:
        ws.cell(row=r, column=1, value=v).font = bold if v == "Total" else base_font
        if v == "Total":
            cnt = f'=COUNTA({rng("Clé (navire | BL | type SAP)")})'
            nonref = f'=COUNTIF({rng("Nb docs SAP")},0)'
            trk = f'=SUM({rng("Montant tracker (FINAL AMOUNT)")})'
            sapm = f'=SUM({rng("SAP montant refacturé (abs)")})'
            enc = f'=SUM({rng("SAP encaissé / imputé (abs)")})'
            sol = f'=SUM({rng("SAP solde client (abs)")})'
            nonrefamt = f'=SUMIFS({rng("Montant tracker (FINAL AMOUNT)")},{rng("Nb docs SAP")},0)'
        else:
            cr = rng(crit_col); vv = f'"{v}"' if isinstance(v, str) else v
            cnt = f'=COUNTIF({cr},{vv})'
            nonref = f'=COUNTIFS({cr},{vv},{rng("Nb docs SAP")},0)'
            trk = f'=SUMIFS({rng("Montant tracker (FINAL AMOUNT)")},{cr},{vv})'
            sapm = f'=SUMIFS({rng("SAP montant refacturé (abs)")},{cr},{vv})'
            enc = f'=SUMIFS({rng("SAP encaissé / imputé (abs)")},{cr},{vv})'
            sol = f'=SUMIFS({rng("SAP solde client (abs)")},{cr},{vv})'
            nonrefamt = f'=SUMIFS({rng("Montant tracker (FINAL AMOUNT)")},{cr},{vv},{rng("Nb docs SAP")},0)'
        cells = [cnt, nonref, trk, sapm, f"=IF(D{r}=0,0,E{r}/D{r})", enc, f"=IF(E{r}=0,0,ABS(G{r})/E{r})", sol, nonrefamt]
        for j, f in enumerate(cells, 2):
            c = ws.cell(row=r, column=j, value=f); c.border = border
            c.font = bold if v == "Total" else base_font
            c.number_format = PCT if j in (6, 8) else ("0" if j in (2, 3) else MONEY)
        r += 1
    return r + 1

r = 5
r = kpi_block(r, "1. Par année de BL", "Année BL", sorted(rap["Année BL"].unique()))
r = kpi_block(r, "2. Par type de réclamation", "Type réclamation", sorted(rap["Type réclamation"].astype(str).unique()))
ws.cell(row=r - 1, column=1, value="ZDS = Demurrage (débit note client) ; ZDP = Despatch (crédit note client).").font = Font(name=FONT, size=9, italic=True)
r += 1
r = kpi_block(r, "3. Par statut de rapprochement", "Statut rapprochement",
              ["Refacturé – soldé", "Refacturé – partiellement encaissé", "Refacturé – non encaissé", "Non refacturé (aucun DN/CN SAP)"])
r = kpi_block(r, "3b. Non refacturé – selon statut de paiement du tracker (PAID/PENDING sans doc SAP = DN/CN probablement hors périmètre de l'export)",
              "Statut paiement (tracker)", ["NOT YET STARTED", "*PAID*", "PENDING"])
r = kpi_block(r, "4. Par méthode de rapprochement", "Méthode de rapprochement", sorted(rap["Méthode de rapprochement"].unique()))
r = kpi_block(r, "5. Par région (tracker)", "Région (tracker)", sorted(rap["Région (tracker)"].unique()))

ws.cell(row=r, column=1, value="6. Contrôle des montants (clés rapprochées)").font = bold
r += 1
for lab, f in [("Clés rapprochées avec écart = OK", f'=COUNTIF({rng("Contrôle écart")},"OK")'),
               ("Clés avec SAP < tracker", f'=COUNTIF({rng("Contrôle écart")},"SAP < tracker")'),
               ("Clés avec SAP > tracker", f'=COUNTIF({rng("Contrôle écart")},"SAP > tracker")'),
               ("Somme des écarts (tracker – SAP, USD)", f'=SUM({rng("Écart tracker – SAP")})'),
               ("Clés avec docs SAP supplémentaires non cités (annulation/réémission ?)", f'=COUNTIF({rng("Alerte")},"Docs SAP*")'),
               ("Clés 'Probable (±3j)' à confirmer", f'=COUNTIF({rng("Méthode de rapprochement")},"Probable*")'),
               ("N° DN/CN listés dans le tracker", f'=SUM({rng("DN/CN tracker listés")})'),
               ("…retrouvés dans l'export SAP", f'=SUM({rng("…dont trouvés dans SAP")})')]:
    ws.cell(row=r, column=1, value=lab).font = base_font
    c = ws.cell(row=r, column=2, value=f); c.font = base_font; c.border = border
    c.number_format = MONEY if "USD" in lab else "0"
    r += 1
r += 1
ws.cell(row=r, column=1, value="7. Côté SAP : DN/CN ZDS/ZDP (BL ≥ 2024) sans ligne dans le tracker post-sales").font = bold
r += 1
NT = "SAP non suivi"
nt_cols = ["Navire", "Date BL", "Type SAP", "Nb docs SAP", "Docs SAP", "Client SAP", "Entité SAP", "Région SAP", "Montant SAP", "Encaissé / imputé (abs)", "Solde client (abs)", "Dernière création", "Navire dans tracker (autre voyage/type)"]
ntc = {c: get_column_letter(i + 1) for i, c in enumerate(nt_cols)}
nn = len(nt) + 1
for lab, f in [("Nb clés navire+BL non suivies", f"=COUNT('{NT}'!$B$2:$B${nn})"),
               ("Montant SAP (USD, signé)", f"=SUM('{NT}'!${ntc['Montant SAP']}$2:${ntc['Montant SAP']}${nn})"),
               ("…dont ZDS (demurrage)", f"=SUMIFS('{NT}'!${ntc['Montant SAP']}$2:${ntc['Montant SAP']}${nn},'{NT}'!${ntc['Type SAP']}$2:${ntc['Type SAP']}${nn},\"ZDS\")"),
               ("…dont ZDP (despatch)", f"=SUMIFS('{NT}'!${ntc['Montant SAP']}$2:${ntc['Montant SAP']}${nn},'{NT}'!${ntc['Type SAP']}$2:${ntc['Type SAP']}${nn},\"ZDP\")"),
               ("Solde client restant (USD, abs)", f"=SUM('{NT}'!${ntc['Solde client (abs)']}$2:${ntc['Solde client (abs)']}${nn})")]:
    ws.cell(row=r, column=1, value=lab).font = base_font
    c = ws.cell(row=r, column=2, value=f); c.font = base_font; c.border = border
    c.number_format = "0" if lab.startswith("Nb") else MONEY
    r += 1
ws.cell(row=r, column=1, value="NB : le tracker post-sales couvre le périmètre CTD ; l'export SAP inclut aussi Brésil / Amériques. Filtrer 'Région SAP' sur l'onglet 'SAP non suivi'.").font = Font(name=FONT, size=9, italic=True)
ws.column_dimensions["A"].width = 44
for L in "BCDEFGHIJ":
    ws.column_dimensions[L].width = 17
ws.sheet_view.showGridLines = False

# ---- Sheet 2: Rapprochement
ws2 = wb.create_sheet(R)
money = ["Montant tracker (FINAL AMOUNT)", "SAP ZDS – Débit (demurrage)", "SAP ZDP – Crédit (despatch)", "SAP montant refacturé (abs)",
         "SAP encaissé / imputé (abs)", "SAP solde client (abs)", "Écart tracker – SAP", "Autres docs SAP même voyage (montant)"]
write_table(ws2, rap, money_cols=money, date_cols=["Date BL (début)", "Date BL (fin)"],
            int_cols=["Année BL", "Nb lignes tracker", "Nb docs SAP", "Autres docs SAP même voyage (nb)", "DN/CN tracker listés", "…dont trouvés dans SAP", "…dont dans les docs rapprochés"],
            widths={"Clé (navire | BL | type SAP)": 38, "Remarque (tracker)": 50, "Docs SAP (n° facture)": 40, "DN/CN n° (tracker)": 32,
                    "Factures vente (tracker)": 32, "Produit": 28, "Client (tracker)": 30, "Client SAP": 30, "Méthode de rapprochement": 34, "Statut rapprochement": 32})
sc = col["Statut rapprochement"]; ec = col["Contrôle écart"]
last_col = get_column_letter(len(rap.columns))
ws2.conditional_formatting.add(f"A2:{last_col}{n}", FormulaRule(formula=[f'${sc}2="Non refacturé (aucun DN/CN SAP)"'], fill=PatternFill("solid", fgColor="F8CBAD")))
ws2.conditional_formatting.add(f"A2:{last_col}{n}", FormulaRule(formula=[f'${sc}2="Refacturé – non encaissé"'], fill=PatternFill("solid", fgColor="FFE699")))
ws2.conditional_formatting.add(f"A2:{last_col}{n}", FormulaRule(formula=[f'${sc}2="Refacturé – partiellement encaissé"'], fill=PatternFill("solid", fgColor="FFF2CC")))
ws2.conditional_formatting.add(f"A2:{last_col}{n}", FormulaRule(formula=[f'${sc}2="Refacturé – soldé"'], fill=PatternFill("solid", fgColor="E2EFDA")))
ws2.conditional_formatting.add(f"{ec}2:{ec}{n}", FormulaRule(formula=[f'OR(${ec}2="SAP < tracker",${ec}2="SAP > tracker")'], font=Font(name=FONT, bold=True, color="C00000")))

# ---- Sheet 3: SAP non suivi
ws3 = wb.create_sheet(NT)
nt_out = pd.DataFrame({
    "Navire": nt.Navire, "Date BL": nt.BL, "Type SAP": nt["Ord.Type"], "Nb docs SAP": nt.Nb_docs, "Docs SAP": nt.Docs,
    "Client SAP": nt.Client, "Entité SAP": nt.Entite, "Région SAP": nt.Region, "Montant SAP": nt.Montant.round(2),
    "Encaissé / imputé (abs)": nt.Encaisse.round(2), "Solde client (abs)": nt.Solde.round(2), "Dernière création": nt.Creation,
    "Navire dans tracker (autre voyage/type)": nt["Navire dans tracker (autre voyage/type)"]})
nt_out = nt_out.sort_values(["Date BL", "Navire"]).reset_index(drop=True)
write_table(ws3, nt_out, money_cols=["Montant SAP", "Encaissé / imputé (abs)", "Solde client (abs)"], date_cols=["Date BL", "Dernière création"],
            int_cols=["Nb docs SAP"], widths={"Docs SAP": 40, "Client SAP": 32})

# ---- Sheet 4: Détail tracker (all rows, incl. excluded)
ws4 = wb.create_sheet("Détail tracker")
det = master.copy()
det["Statut rapprochement"] = det["CLE"].map(rap.set_index("Clé (navire | BL | type SAP)")["Statut rapprochement"]).fillna("Hors périmètre")
det["Méthode"] = det["CLE"].map(rap.set_index("Clé (navire | BL | type SAP)")["Méthode de rapprochement"]).fillna("")
det_out = pd.DataFrame({
    "Source": det.SOURCE, "Onglet": det.ONGLET, "Présent fichier APAC": det.DANS_APAC, "Périmètre": det.PERIMETRE,
    "Clé": det.CLE, "VESSEL": det.VESSEL, "Origin": det.Origin, "B/L DATE (brut)": det["B/L DATE"].astype(str),
    "BL début": det.BL_DEB, "BL fin": det.BL_FIN, "REGION": det.REGION, "DELIVERY MODE": det["DELIVERY MODE"], "RECEIVER": det.RECEIVER,
    "PRODUCT": det.PRODUCT, "INVOICE NUMBER": det["INVOICE NUMBER"].astype(str), "Claims Type": det["Claims Type"],
    "FINAL AMOUNT (brut)": det["FINAL AMOUNT"].astype(str), "Montant (num)": det.MONTANT, "DN/CN Number": det["DN/CN Number"].astype(str),
    "PAYMENT STATUS": det["PAYMENT STATUS"], "Statut paiement (normalisé)": det.PAY, "REMARK": det.REMARK,
    "Statut rapprochement": det["Statut rapprochement"], "Méthode": det["Méthode"]})
write_table(ws4, det_out, money_cols=["Montant (num)"], date_cols=["BL début", "BL fin"], widths={"REMARK": 50, "Clé": 38})

# ---- Sheet 5: Base SAP 2024+
ws5 = wb.create_sheet("Base SAP 2024+")
keep = ["Numéro Facture", "Numéro de Commande", "Date Création Facture", "Date Comptabilisation", "Entité Description", "S.Off.Description",
        "S.Grp.Description", "Ord.Type", "Ord.Type Description", "Type Commande Descp", "Numéro Comptable", "Client", "Client Name",
        "Navire", "Date BL", "Année BL", "Mode de Vente", "Description Produit", "Quantité Facturée", "Devise", "Montant Facturée",
        "Encaissé", "Solde Client", "Solde abs", "Montant Facturé en MAD", "Echéance", "Date Reglement 1", "Montant Reglement 1",
        "Date Reglement 2", "Montant Reglement 2", "Date Reglement 3", "Montant Reglement 3", "Payment Terms Description",
        "SBU Description", "Désignation pays destination", "Nom de l'ADV", "Rapproché tracker"]
sap_out = sap[keep].sort_values(["Date BL", "Navire"]).reset_index(drop=True)
write_table(ws5, sap_out, money_cols=["Montant Facturée", "Encaissé", "Solde Client", "Solde abs", "Montant Facturé en MAD", "Montant Reglement 1", "Montant Reglement 2", "Montant Reglement 3"],
            date_cols=["Date Création Facture", "Date Comptabilisation", "Date BL", "Echéance", "Date Reglement 1", "Date Reglement 2", "Date Reglement 3"],
            int_cols=["Année BL"], widths={"Client Name": 32, "Description Produit": 30})

# ---- Sheet 6: Notes
ws6 = wb.create_sheet("Notes & hypothèses")
notes = [
    ("Sources", ""),
    ("Export SAP", f"{EXP} : {len(exp)} documents ZDS (demurrage → débit note client) et ZDP (despatch → crédit note client), Report Run Date {pd.to_datetime(exp['Report Run Date']).max():%d/%m/%Y}."),
    ("Tracker(s)", "Fichiers : " + " ; ".join(TRACKERS) + ". Le premier est le fichier maître ; les suivants sont fusionnés sans doublon (colonne 'Présent fichier APAC' = présent dans un fichier secondaire)."),
    ("Onglets PRICE ADJ / DISCOUNTS", "Non intégrés : hors demurrage/despatch (remises, ajustements de prix, LC charges…) et leurs DN/CN ne sont pas des types ZDS/ZDP."),
    ("", ""),
    ("Règles de rapprochement", ""),
    ("Clé", "Nom du navire normalisé (majuscules, ponctuation et espaces multiples retirés, 'M/V' ignoré) + date BL. Type de réclamation DEMURRAGES ↔ ZDS, DESPATCH ↔ ZDP, 'DESPATCH / DEMURRAGES' ↔ les deux."),
    ("Date BL en plage", "Quand le tracker indique une plage 'jj/mm/aaaa - jj/mm/aaaa', tout document SAP dont la date BL tombe dans la plage est rapproché (ordre inversé toléré, ex. G TAISHAN)."),
    ("Probable (±3j)", "Si aucun document exact, recherche à ±3 jours → à confirmer manuellement (indiqué dans 'Méthode de rapprochement')."),
    ("Même navire, autre voyage", "Un navire présent dans SAP à une autre date BL n'est PAS rapproché (voyages différents). Mention dans la méthode pour aide à l'analyse."),
    ("Niveau de rapprochement", "Onglet 'Rapprochement' = 1 ligne par clé navire + BL + type. Plusieurs lignes tracker partageant la clé (ex. BEATRICE 17/11/2025, 2 lignes) sont regroupées pour éviter de compter deux fois les documents SAP. Le détail ligne à ligne est dans 'Détail tracker'."),
    ("Périmètre 2024+", "Lignes tracker avec date BL ≥ 01/01/2024 (instruction). Les lignes antérieures restent visibles dans 'Détail tracker' avec Périmètre = 'Avant 2024 (exclu)'."),
    ("Montants", "Tracker : FINAL AMOUNT (positif). SAP : Montant Facturée (ZDS positif, ZDP négatif) ; 'SAP montant refacturé (abs)' = |ZDS| + |ZDP|. 'Encaissé / imputé (abs)' = |Montant Facturée| – |Solde Client| (valeurs absolues pour ne pas compenser débits ZDS et crédits ZDP). Écart = tracker – SAP abs ; 'OK' si |écart| ≤ max(1 USD ; 0,5 %)."),
    ("N° DN/CN sur un autre BL", "Si un n° de DN/CN cité par le tracker existe dans SAP sur le même navire mais une autre date BL (ex. ABILITY : 2 documents au 30/09/2024 et 2 au 07/10/2024 sur une seule ligne tracker), le document est rattaché et la méthode le signale : la ligne tracker couvre deux voyages."),
    ("Priorité aux n° DN/CN", "Si le tracker cite des n° de DN/CN retrouvés dans SAP sur le même navire/BL, seuls ces documents sont retenus pour les montants ; les autres documents SAP du même voyage (ex. deux lots créés à des dates différentes, annulation puis réémission) sont listés dans 'Autres docs SAP même voyage' avec une alerte."),
    ("Statuts", "Refacturé – soldé : docs SAP trouvés et solde client = 0. Refacturé – non encaissé : docs trouvés, rien réglé. Refacturé – partiellement encaissé : solde ≠ 0 et règlement partiel. Non refacturé : aucun document ZDS/ZDP dans l'export pour ce navire/BL/type."),
    ("", ""),
    ("Points d'attention", ""),
    ("Couverture de l'export SAP", (lambda _m: (f"{_m['n']} clés au statut PAID/PENDING citent des n° de DN/CN absents de l'export (ex. {_m['ex']}) : l'export semble filtré (entité / famille produit). Ces clés sont 'Non refacturé' faute de document, mais peuvent être hors périmètre d'extraction : vérifier le filtre de l'extraction SAP avant conclusion." if _m['n'] else "Tous les n° de DN/CN cités par le tracker et marqués payés/en cours ont été retrouvés dans l'export."))(
        {"n": int(((rap["Nb docs SAP"] == 0) & rap["Statut paiement (tracker)"].str.contains("PAID|PENDING") & (rap["DN/CN tracker listés"] > 0)).sum()),
         "ex": " ; ".join((rap.loc[(rap["Nb docs SAP"] == 0) & rap["Statut paiement (tracker)"].str.contains("PAID|PENDING") & (rap["DN/CN tracker listés"] > 0)].sort_values("Montant tracker (FINAL AMOUNT)", ascending=False)["Navire"].head(3)).tolist())})),
    ("Nature de l'export", "L'export fourni contient les débit/crédit notes émises aux clients (ZDS/ZDP), pas les paiements à l'armateur. Les montants payés à l'armateur (facture de fret finale) ne figurent pas dans ces fichiers : le rapprochement mesure tracker post-sales ↔ refacturation SAP. Pour mesurer 'payé à l'armateur ↔ refacturé', joindre l'extraction des factures de fret (même clé navire + BL, voir modèle)."),
    ("N° DN/CN non retrouvés", f"{int(rap['DN/CN tracker listés'].sum() - rap['…dont trouvés dans SAP'].sum())} n° cités sur {int(rap['DN/CN tracker listés'].sum())} sont absents de l'export. Avant de conclure, chercher les fautes de frappe (n° voisin sur le même navire) : un seul chiffre erroné crée un faux écart."),
    ("Qualité des données tracker", f"{int(master['FINAL AMOUNT'].map(lambda v: isinstance(v, str)).sum())} montant(s) saisi(s) en texte converti(s) ; {int((master['BL_FIN'] != master['BL_DEB']).sum())} date(s) BL saisie(s) en plage ; {int(master['VESSEL'].astype(str).map(lambda v: v != v.strip()).sum())} nom(s) de navire avec espaces parasites ; statuts de paiement en texte libre (normalisés dans 'Statut paiement (normalisé)')."),
    ("Formules", "L'onglet Synthèse est entièrement en formules (SUMIFS/COUNTIFS) sur l'onglet Rapprochement : il se recalcule à l'ouverture dans Excel."),
]
ws6["A1"] = "Notes, hypothèses et règles"; ws6["A1"].font = title_font
for i, (k, v) in enumerate(notes, 3):
    a = ws6.cell(row=i, column=1, value=k); a.font = bold if v == "" else base_font
    b = ws6.cell(row=i, column=2, value=v); b.font = base_font; b.alignment = Alignment(wrap_text=True, vertical="top")
ws6.column_dimensions["A"].width = 30; ws6.column_dimensions["B"].width = 140

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)

# ---- console check: replicate Synthèse with pandas
print("saved", OUT)
print("rap rows", len(rap), "| nt rows", len(nt_out), "| det rows", len(det_out), "| sap rows", len(sap_out))
chk = rap.groupby("Année BL").agg(n=("Nb docs SAP", "size"), nonref=("Nb docs SAP", lambda s: (s == 0).sum()),
                                  trk=("Montant tracker (FINAL AMOUNT)", "sum"), sap=("SAP montant refacturé (abs)", "sum"),
                                  enc=("SAP encaissé / imputé (abs)", "sum"), solde=("SAP solde client (abs)", "sum"))
chk["pct_ref"] = chk["sap"] / chk["trk"]; chk["pct_enc"] = chk["enc"].abs() / chk["sap"]
print(chk.round(2).to_string())
print(rap["Statut rapprochement"].value_counts().to_string())
print(rap["Méthode de rapprochement"].value_counts().to_string())
print(rap["Contrôle écart"].value_counts().to_string())
print("DN/CN listés", rap["DN/CN tracker listés"].sum(), "trouvés SAP", rap["…dont trouvés dans SAP"].sum(), "dans docs", rap["…dont dans les docs rapprochés"].sum())
print(nt_out.groupby("Type SAP")["Montant SAP"].agg(["size", "sum"]).round(2))
