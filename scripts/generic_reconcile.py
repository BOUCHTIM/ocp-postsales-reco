"""Rapprochement générique de deux tables (Excel / CSV) sur des clés choisies, avec normalisation et tolérances.
Sert pour tout ce qui n'est pas couvert par reconcile.py : factures de vente ↔ SAP, tracker ↔ relevé bancaire, deux extractions SAP, etc.
Produit un classeur : Rapprochés, Gauche seule, Droite seule, Écarts, Synthèse (formules), Paramètres.

Usage :
  generic_reconcile.py --left A.xlsx[:Onglet] --right B.xlsx[:Onglet] --out rapp.xlsx \
      --key "VESSEL=Navire" --key "B/L DATE=Date BL" [--amount "FINAL AMOUNT=Montant Facturée"] \
      [--date-tolerance 3] [--abs] [--norm-text] [--tolerance 1] [--tolerance-pct 0.5] [--agg sum|first]
--key gauche=droite (répéter) ; une clé date tolère ±N jours ; --abs compare les montants en valeur absolue ;
--agg sum regroupe les lignes de même clé avant comparaison (évite les doubles comptes)."""
import argparse, re, sys
import pandas as pd, numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

ap = argparse.ArgumentParser()
ap.add_argument("--left", required=True); ap.add_argument("--right", required=True); ap.add_argument("--out", required=True)
ap.add_argument("--key", action="append", required=True); ap.add_argument("--amount", action="append", default=[])
ap.add_argument("--date-tolerance", type=int, default=0); ap.add_argument("--abs", action="store_true"); ap.add_argument("--norm-text", action="store_true", default=True)
ap.add_argument("--tolerance", type=float, default=1.0); ap.add_argument("--tolerance-pct", type=float, default=0.5); ap.add_argument("--agg", choices=["sum", "first"], default="sum")
a = ap.parse_args()


def load(spec):
    path, _, sheet = spec.partition(":")
    if path.lower().endswith(".csv"): return pd.read_csv(path, sep=None, engine="python")
    xl = pd.ExcelFile(path); return pd.read_excel(path, sheet_name=sheet if sheet in xl.sheet_names else xl.sheet_names[0])


def norm(v):
    if pd.isna(v): return ""
    if isinstance(v, (pd.Timestamp,)): return v.normalize()
    if isinstance(v, (int, float, np.integer, np.floating)) and float(v).is_integer(): return str(int(v))   # 90020837.0 → "90020837"
    s = str(v)
    if re.match(r"^\d+\.0$", s): s = s[:-2]
    if re.match(r"^\d{4}-\d{2}-\d{2}", s):
        try: return pd.Timestamp(s[:10])
        except Exception: pass
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        try: return pd.Timestamp(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except Exception: pass
    if a.norm_text: s = re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9 ]", " ", s.upper())).strip()
    return s


L, R = load(a.left), load(a.right)
pairs = [k.split("=", 1) if "=" in k else (k, k) for k in a.key]
amts = [k.split("=", 1) if "=" in k else (k, k) for k in a.amount]
for lk, rk in pairs:
    if lk not in L.columns or rk not in R.columns: sys.exit(f"Clé introuvable : {lk} / {rk}. Colonnes gauche : {list(L.columns)} ; droite : {list(R.columns)}")
for i, (lk, rk) in enumerate(pairs):
    L[f"_k{i}"] = L[lk].map(norm); R[f"_k{i}"] = R[rk].map(norm)
kcols = [f"_k{i}" for i in range(len(pairs))]
date_idx = [i for i, (lk, rk) in enumerate(pairs) if pd.api.types.is_datetime64_any_dtype(L[f"_k{i}"]) or L[f"_k{i}"].map(lambda v: isinstance(v, pd.Timestamp)).mean() > 0.5]
# agrégation par clé
if a.agg == "sum" and amts:
    La = L.groupby(kcols, as_index=False).agg({**{lk: "sum" for lk, _ in amts}, **{c: "first" for c in L.columns if c not in kcols and c not in [lk for lk, _ in amts]}})
    Ra = R.groupby(kcols, as_index=False).agg({**{rk: "sum" for _, rk in amts}, **{c: "first" for c in R.columns if c not in kcols and c not in [rk for _, rk in amts]}})
    La["_n"] = L.groupby(kcols).size().values; Ra["_n"] = R.groupby(kcols).size().values
else: La, Ra = L.copy(), R.copy(); La["_n"] = 1; Ra["_n"] = 1
# jointure exacte puis tolérance sur dates
m = La.merge(Ra, on=kcols, how="outer", indicator=True, suffixes=(" (gauche)", " (droite)"))
if a.date_tolerance and date_idx:
    lo = m[m["_merge"] == "left_only"].copy(); ro = m[m["_merge"] == "right_only"].copy(); extra = []
    nondate = [c for i, c in enumerate(kcols) if i not in date_idx]; di = kcols[date_idx[0]]
    for _, lr in lo.iterrows():
        cand = ro[(ro[nondate] == lr[nondate].values).all(axis=1)] if nondate else ro
        cand = cand[cand[di].map(lambda d: isinstance(d, pd.Timestamp) and isinstance(lr[di], pd.Timestamp) and abs((d - lr[di]).days) <= a.date_tolerance)]
        if len(cand):
            rr = cand.iloc[0]; ro = ro.drop(rr.name); row = {**{c: lr[c] for c in m.columns if c.endswith("(gauche)") or c in kcols or c == "_n (gauche)"}, **{c: rr[c] for c in m.columns if c.endswith("(droite)") or c == "_n (droite)"}}
            row["_merge"] = f"both (±{a.date_tolerance}j)"; extra.append(row)
    if extra:
        m = m[~m.index.isin(lo.index[:len(extra)])]  # retire les gauche-seules rapprochées
        m = pd.concat([m[m["_merge"] != "right_only"], pd.DataFrame(extra), ro], ignore_index=True)
m["_merge"] = m["_merge"].astype(str).replace({"both": "Rapproché", "left_only": "Gauche seule", "right_only": "Droite seule"})
for lk, rk in amts:
    lc, rc = (lk if lk in m.columns else f"{lk} (gauche)"), (rk if rk in m.columns else f"{rk} (droite)")
    lv, rv = pd.to_numeric(m[lc], errors="coerce"), pd.to_numeric(m[rc], errors="coerce")
    if a.abs: lv, rv = lv.abs(), rv.abs()
    m[f"Écart {lk}"] = (lv - rv).round(2)
    m[f"Contrôle {lk}"] = np.where(m["_merge"].str.startswith("Rapproché") is False, "", np.where((m[f"Écart {lk}"].abs() <= np.maximum(a.tolerance, a.tolerance_pct / 100 * lv.abs())), "OK", "ÉCART"))
    m.loc[~m["_merge"].str.startswith("Rapproché"), f"Contrôle {lk}"] = ""
m = m.rename(columns={"_merge": "Statut"}); m = m.drop(columns=kcols)
# écriture
wb = Workbook(); thin = Side(style="thin", color="BFBFBF"); bd = Border(top=thin, bottom=thin, left=thin, right=thin); hf = PatternFill("solid", fgColor="1F3864"); hfont = Font(name="Arial", bold=True, color="FFFFFF", size=10); bf = Font(name="Arial", size=10)


def sheet(ws, df):
    for j, c in enumerate(df.columns, 1):
        x = ws.cell(row=1, column=j, value=str(c)); x.font = hfont; x.fill = hf; x.border = bd; x.alignment = Alignment(wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = min(40, max(10, len(str(c)) + 2))
    for i, row in enumerate(df.itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            if isinstance(v, float) and np.isnan(v): v = None
            if isinstance(v, pd.Timestamp): v = v.to_pydatetime()
            if hasattr(v, "item"):
                try: v = v.item()
                except Exception: pass
            x = ws.cell(row=i, column=j, value=v); x.font = bf; x.border = bd
    ws.freeze_panes = "A2"; ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"


ws = wb.active; ws.title = "Synthèse"; ws["A1"] = f"Rapprochement générique : {a.left}  ↔  {a.right}"; ws["A1"].font = Font(name="Arial", bold=True, size=13, color="1F3864")
ws["A2"] = "Clés : " + " ; ".join(f"{l} = {r}" for l, r in pairs) + (f" · tolérance date ±{a.date_tolerance} j" if a.date_tolerance else "") + (" · montants en valeur absolue" if a.abs else ""); ws["A2"].font = Font(name="Arial", italic=True, size=9)
n = len(m) + 1
rows = [("Lignes rapprochées", f'=COUNTIF(Rapprochement!$A$2:$A${n},"Rapproché*")'), ("Gauche seule", f'=COUNTIF(Rapprochement!$A$2:$A${n},"Gauche seule")'), ("Droite seule", f'=COUNTIF(Rapprochement!$A$2:$A${n},"Droite seule")')]
cols = ["Statut"] + [c for c in m.columns if c != "Statut"]; m = m[cols]
for lk, _ in amts:
    ci = get_column_letter(cols.index(f"Contrôle {lk}") + 1)
    rows += [(f"Écarts sur {lk}", f'=COUNTIF(Rapprochement!${ci}$2:${ci}${n},"ÉCART")')]
for i, (lab, f) in enumerate(rows, 4): ws.cell(row=i, column=1, value=lab).font = bf; ws.cell(row=i, column=2, value=f).font = bf
ws.column_dimensions["A"].width = 32; ws.column_dimensions["B"].width = 14
sheet(wb.create_sheet("Rapprochement"), m)
for name, flt in [("Gauche seule", m["Statut"] == "Gauche seule"), ("Droite seule", m["Statut"] == "Droite seule")]: sheet(wb.create_sheet(name), m[flt])
if amts: sheet(wb.create_sheet("Écarts"), m[(m[[f"Contrôle {lk}" for lk, _ in amts]] == "ÉCART").any(axis=1)])
wp = wb.create_sheet("Paramètres"); [wp.append([k, str(v)]) for k, v in vars(a).items()]
wb.calculation.fullCalcOnLoad = True; wb.save(a.out)
print(f"saved {a.out} | rapprochés {int(m['Statut'].str.startswith('Rapproché').sum())} | gauche seule {int((m['Statut']=='Gauche seule').sum())} | droite seule {int((m['Statut']=='Droite seule').sum())}")
