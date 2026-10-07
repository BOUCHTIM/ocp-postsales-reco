"""Profileur générique : décrit n'importe quel fichier Excel / CSV (tous les onglets) pour décider quoi en faire.
Sortie : un rapport Markdown lisible + un JSON exploitable (dictionnaire de données, qualité, clés candidates, familles détectées).
Usage : profile.py <fichier.xlsx|csv> [<rapport.md>] [--json <sortie.json>] [--rows 2000]"""
import sys, os, re, json, argparse
import pandas as pd
import numpy as np

ap = argparse.ArgumentParser(); ap.add_argument("file"); ap.add_argument("out_md", nargs="?"); ap.add_argument("--json"); ap.add_argument("--rows", type=int, default=5000)
a = ap.parse_args()
ROLE_PATTERNS = {  # rôle probable d'une colonne d'après son nom
    "navire": r"VESSEL|NAVIRE|SHIP\b", "date_bl": r"B/?L DATE|DATE BL|BL DATE", "date": r"DATE|ÉCH[ÉE]ANCE|ECHEANCE|CREATED|DUE", "montant": r"AMOUNT|MONTANT|VALUE|SOLDE|PRICE|PRIX|TOTAL",
    "doc_sap": r"NUM[ÉE]RO FACTURE|BILLING|DN/?CN|DEBIT NOTE|CREDIT NOTE|INVOICE|FACTURE", "client": r"RECEIVER|CLIENT|CUSTOMER|BUYER|SOLD TO|SHIP TO", "type": r"TYPE|CLAIMS?|ORD\.?TYPE|PROCESS",
    "statut": r"STATUS|STATUT|PAYMENT", "region": r"REGION|RÉGION|S\.GRP|GROUP|OFFICE|PAYS|COUNTRY", "produit": r"PRODUCT|PRODUIT|MATERIAL", "imo": r"\bIMO\b", "incoterm": r"DELIVERY MODE|INCOTERM|MODE DE VENTE",
}
FAMILY_HINTS = {"demdesp": r"DEMURRAGE|DESPATCH|DEM-?DESP|ZDS|ZDP", "priceadj": r"PRICE ADJ|DISCOUNT|REMISE|LC CHARGE|DEAD FREIGHT|MARGIN SHARE|TROP PERCU", "sap_billing": r"ORD\.?TYPE|BILLING|NUMÉRO FACTURE|SOLDE CLIENT", "fret": r"ARMATEUR|SHIPOWNER|LAYTIME|FREIGHT INVOICE|FACTURE FRET"}


def role(col):
    c = re.sub(r"\s+", " ", str(col)).strip().upper()
    for k, p in ROLE_PATTERNS.items():
        if re.search(p, c): return k
    return ""


def sheets(path):
    if path.lower().endswith((".csv", ".txt")):
        yield "csv", pd.read_csv(path, sep=None, engine="python", nrows=a.rows)
    else:
        xl = pd.ExcelFile(path)
        for sh in xl.sheet_names:
            try: yield sh, pd.read_excel(path, sheet_name=sh, nrows=a.rows)
            except Exception as e: yield sh, pd.DataFrame({"_erreur": [str(e)]})


def describe(df):
    cols = []
    for c in df.columns:
        s = df[c]; nn = s.notna().sum(); info = {"colonne": str(c), "role": role(c), "non_vides": int(nn), "vides_pct": round(100 * (1 - nn / max(1, len(s))), 1), "uniques": int(s.nunique(dropna=True))}
        if pd.api.types.is_numeric_dtype(s): info.update(type="nombre", min=float(s.min()) if nn else None, max=float(s.max()) if nn else None, somme=float(s.sum()) if nn else None, negatifs=int((s < 0).sum()))
        elif pd.api.types.is_datetime64_any_dtype(s): info.update(type="date", min=str(s.min().date()) if nn else None, max=str(s.max().date()) if nn else None)
        else:
            vals = s.dropna().astype(str); info["type"] = "texte"
            info["exemples"] = vals.drop_duplicates().head(4).tolist()
            info["dates_texte"] = int(vals.str.contains(r"\d{1,2}/\d{1,2}/\d{4}").sum()); info["nombres_texte"] = int(vals.str.contains(r"^\s*-?\d[\d  ]*[,.]\d+\s*$").sum())
            info["espaces_parasites"] = int((vals != vals.str.strip()).sum()); info["plages_dates"] = int(vals.str.contains(r"\d{4}\s*-\s*\d{1,2}/").sum())
            if info["uniques"] <= 12 and nn: info["valeurs"] = {str(k): int(v) for k, v in vals.value_counts().head(12).items()}
        cols.append(info)
    return cols


def candidate_keys(cols):
    roles = {c["role"] for c in cols}
    keys = []
    if "navire" in roles and ("date_bl" in roles or "date" in roles): keys.append("navire + date BL (voyage)")
    if "doc_sap" in roles: keys.append("numéro de document (facture / DN / CN)")
    if "client" in roles and "date" in roles: keys.append("client + période")
    return keys


out = {"fichier": os.path.basename(a.file), "onglets": []}
md = [f"# Profil du fichier : {os.path.basename(a.file)}", ""]
for sh, df in sheets(a.file):
    cols = describe(df); txt = " ".join(str(c) for c in df.columns).upper() + " " + " ".join("" if pd.isna(v) else str(v) for v in df.head(50).values.ravel()).upper()
    fam = [k for k, p in FAMILY_HINTS.items() if re.search(p, txt)]
    keys = candidate_keys(cols)
    out["onglets"].append({"onglet": sh, "lignes": int(len(df)), "colonnes": len(df.columns), "familles_probables": fam, "cles_candidates": keys, "colonnes_detail": cols})
    md += [f"## Onglet « {sh} » — {len(df)} lignes × {len(df.columns)} colonnes", f"Familles probables : {', '.join(fam) or 'aucune reconnue'} · Clés candidates : {', '.join(keys) or 'à définir'}", "",
           "| Colonne | Rôle | Type | Non vides | Vides % | Uniques | Repères |", "|---|---|---|---|---|---|---|"]
    for c in cols:
        rep = []
        if c.get("type") == "nombre": rep.append(f"min {c['min']:.0f} · max {c['max']:.0f} · somme {c['somme']:,.0f}".replace(",", " ") if c["min"] is not None else "")
        if c.get("type") == "date": rep.append(f"{c['min']} → {c['max']}")
        if c.get("type") == "texte":
            if c.get("valeurs"): rep.append("valeurs : " + ", ".join(f"{k} ({v})" for k, v in list(c["valeurs"].items())[:6]))
            else: rep.append("ex. : " + " ; ".join(c["exemples"][:3]))
            for k, lab in [("dates_texte", "dates en texte"), ("nombres_texte", "nombres en texte"), ("espaces_parasites", "espaces parasites"), ("plages_dates", "plages de dates")]:
                if c.get(k): rep.append(f"{c[k]} {lab}")
        md.append(f"| {c['colonne']} | {c['role']} | {c.get('type','')} | {c['non_vides']} | {c['vides_pct']} | {c['uniques']} | {' · '.join(r for r in rep if r)} |")
    md.append("")
md += ["## Lecture", "- `role` = rôle deviné d'après le nom de colonne ; vérifier avant de s'en servir comme clé.", "- Les repères signalent les fragilités habituelles : dates ou nombres saisis en texte, espaces parasites, plages de dates.", "- Familles reconnues : demdesp (surestaries), priceadj (remises / ajustements), sap_billing (export de documents SAP), fret (factures armateur). Un onglet sans famille se traite avec generic_reconcile.py ou une analyse ad hoc."]
text = "\n".join(md)
if a.out_md: open(a.out_md, "w").write(text)
else: print(text)
if a.json: json.dump(out, open(a.json, "w"), ensure_ascii=False, indent=1, default=str)
