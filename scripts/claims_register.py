"""Registre des réclamations clients (qualité, quantité, documentation, délai, litiges) : délais de traitement, taux d'acceptation, backlog,
ancienneté, par client / type / gestionnaire (charge d'équipe).
Entrée : --claims registre : N° réclamation, Client, Navire, Date BL, Produit, Type, Date réception, Montant réclamé, Montant accepté, Statut, Date clôture, Gestionnaire [, Cause, Action, Commentaire]
Sortie : classeur « Réclamations » : synthèse en formules, backlog par ancienneté, par type / client / gestionnaire, liste des dossiers hors délai.
Usage : claims_register.py --claims registre.xlsx --out reclamations.xlsx [--arrete AAAA-MM-JJ] [--sla 30] [--config alias.json]"""
import argparse, json
import pandas as pd, numpy as np
from openpyxl import Workbook
from xlsx_common import *

ap = argparse.ArgumentParser(); ap.add_argument("--claims", required=True); ap.add_argument("--out", required=True); ap.add_argument("--arrete"); ap.add_argument("--sla", type=int, default=30, help="délai cible de traitement en jours"); ap.add_argument("--config")
a = ap.parse_args()
DEF = {"claims": {"N° réclamation": ["CLAIM NO", "CLAIM ID", "N° CLAIM", "REF", "RÉFÉRENCE", "ID"], "Client": ["CUSTOMER", "RECEIVER", "CLIENT NAME", "BUYER"], "Navire": ["VESSEL", "SHIP"], "Date BL": ["B/L DATE", "BL DATE"], "Produit": ["PRODUCT", "MATERIAL"],
                  "Type": ["CLAIM TYPE", "CATEGORY", "CATÉGORIE", "NATURE", "TYPE DE RÉCLAMATION"], "Date réception": ["RECEIVED", "DATE RECEIVED", "DATE RÉCEPTION", "OPENED", "DATE OUVERTURE", "CLAIM DATE"], "Montant réclamé": ["CLAIMED AMOUNT", "AMOUNT CLAIMED", "MONTANT RÉCLAMÉ", "CLAIM AMOUNT", "AMOUNT"],
                  "Montant accepté": ["ACCEPTED AMOUNT", "SETTLED AMOUNT", "MONTANT ACCEPTÉ", "MONTANT ACCORDÉ", "AGREED"], "Statut": ["STATUS", "STATUT"], "Date clôture": ["CLOSED", "CLOSING DATE", "DATE CLÔTURE", "DATE FERMETURE", "SETTLED DATE"], "Gestionnaire": ["OWNER", "HANDLER", "RESPONSABLE", "ASSIGNED TO", "ANALYST", "CHARGÉ"],
                  "Cause": ["ROOT CAUSE", "CAUSE", "REASON"], "Action": ["ACTION", "RESOLUTION", "DÉCISION"], "Commentaire": ["COMMENT", "REMARK", "REMARKS", "COMMENTAIRE"]}}
cfg = load_config(a.config, DEF)
C = read_any(a.claims, aliases=cfg["claims"], required=["Client", "Date réception"])
ARRETE = pd.Timestamp(a.arrete) if a.arrete else pd.Timestamp.today().normalize()
for c in ["Date BL", "Date réception", "Date clôture"]:
    if c in C.columns: C[c] = C[c].map(to_date)
    else: C[c] = pd.NaT
for c in ["Montant réclamé", "Montant accepté"]:
    if c in C.columns: C[c] = C[c].map(to_num)
    else: C[c] = np.nan
for c in ["Type", "Statut", "Gestionnaire", "Produit", "Cause"]:
    if c not in C.columns: C[c] = ""
    C[c] = C[c].fillna("").astype(str).str.strip()
C["Client"] = C["Client"].astype(str).str.strip().str.upper().str.replace(r"\s+", " ", regex=True)
closed = C["Date clôture"].notna() | C["Statut"].str.upper().isin(["CLOSED", "CLÔTURÉ", "CLOTURE", "SETTLED", "REJECTED", "REJETÉ", "PAID", "RÉGLÉ"])
C["État"] = np.where(closed, "Clôturé", "Ouvert")
C["Délai de traitement (j)"] = np.where(closed, (C["Date clôture"].fillna(ARRETE) - C["Date réception"]).dt.days, np.nan)
C["Ancienneté (j)"] = np.where(~closed, (ARRETE - C["Date réception"]).dt.days, np.nan)
C["Hors délai"] = np.where((~closed) & (C["Ancienneté (j)"] > a.sla), f"Ouvert > {a.sla} j", np.where(closed & (C["Délai de traitement (j)"] > a.sla), f"Traité en > {a.sla} j", "Dans le délai"))
C["Taux d'acceptation"] = np.where(C["Montant réclamé"].fillna(0) > 0, C["Montant accepté"].fillna(0) / C["Montant réclamé"], np.nan)
C["Tranche ancienneté"] = C["Ancienneté (j)"].map(lambda d: "" if pd.isna(d) else ("0-30 j" if d <= 30 else "31-60 j" if d <= 60 else "61-90 j" if d <= 90 else "91-180 j" if d <= 180 else "> 180 j"))
C["Mois réception"] = C["Date réception"].dt.to_period("M").astype(str)
cols = [c for c in ["N° réclamation", "Client", "Navire", "Date BL", "Produit", "Type", "Cause", "Date réception", "Mois réception", "Montant réclamé", "Montant accepté", "Taux d'acceptation", "Statut", "État", "Date clôture", "Délai de traitement (j)", "Ancienneté (j)", "Tranche ancienneté", "Hors délai", "Gestionnaire", "Action", "Commentaire"] if c in C.columns]
C = C[cols].sort_values(["État", "Ancienneté (j)"], ascending=[False, False]).reset_index(drop=True)
wb = Workbook(); ws = wb.active; ws.title = "Synthèse"
ws["A1"] = "Registre des réclamations clients — délais, acceptation, backlog, charge"; ws["A1"].font = TITLE
ws["A2"] = f"Arrêté : {ARRETE:%d/%m/%Y}. Délai cible (SLA) : {a.sla} j. Formules sur l'onglet « Réclamations »."; ws["A2"].font = NOTE
cm = colmap(C); n = len(C); R_ = lambda c: rng("Réclamations", cm[c], n)
r = kpi_rows(ws, 4, [("Réclamations", f'=COUNTA({R_("Client")})', "0"), ("Ouvertes (backlog)", f'=COUNTIF({R_("État")},"Ouvert")', "0"), ("Clôturées", f'=COUNTIF({R_("État")},"Clôturé")', "0"), ("Montant réclamé", f'=SUM({R_("Montant réclamé")})', MONEY0), ("Montant accepté", f'=SUM({R_("Montant accepté")})', MONEY0),
                    ("Taux d'acceptation global", f'=IF(B{8}=0,0,B{9}/B{8})', PCT), ("Montant réclamé en cours (ouvert)", f'=SUMIFS({R_("Montant réclamé")},{R_("État")},"Ouvert")', MONEY0),
                    ("Délai médian de traitement (j, clôturées)", f'=MEDIAN({R_("Délai de traitement (j)")})', "0"), ("Ancienneté médiane du backlog (j)", f'=MEDIAN({R_("Ancienneté (j)")})', "0"),
                    (f"Ouvertes > {a.sla} j", f'=COUNTIF({R_("Hors délai")},"Ouvert > {a.sla} j")', "0"), (f"Clôturées en > {a.sla} j", f'=COUNTIF({R_("Hors délai")},"Traité en > {a.sla} j")', "0")], "1. Volumes, montants, délais")
r = kpi_rows(ws, r, [(b, f'=COUNTIFS({R_("Tranche ancienneté")},"{b}")', "0") for b in ["0-30 j", "31-60 j", "61-90 j", "91-180 j", "> 180 j"]], "2. Backlog par ancienneté")
ws.column_dimensions["A"].width = 48; ws.column_dimensions["B"].width = 18; ws.sheet_view.showGridLines = False
ws2 = wb.create_sheet("Réclamations"); write_df(ws2, C, money=["Montant réclamé", "Montant accepté"], date=["Date BL", "Date réception", "Date clôture"], pct=["Taux d'acceptation"], intc=["Délai de traitement (j)", "Ancienneté (j)"], widths={"Commentaire": 40, "Action": 30, "Hors délai": 20})
last = get_column_letter(len(C.columns)); hc = cm["Hors délai"]; ec = cm["État"]
ws2.conditional_formatting.add(f"A2:{last}{n + 1}", FormulaRule(formula=[f'LEFT(${hc}2,6)="Ouvert"'], fill=FILL_RED)); ws2.conditional_formatting.add(f"A2:{last}{n + 1}", FormulaRule(formula=[f'${ec}2="Clôturé"'], fill=FILL_GRN))


def by(col):
    g = C.groupby(col).agg(Réclamations=("Client", "size"), Ouvertes=("État", lambda s: (s == "Ouvert").sum()), Hors_délai=("Hors délai", lambda s: (s != "").sum()), Réclamé=("Montant réclamé", "sum"), Accepté=("Montant accepté", "sum"), Délai_médian=("Délai de traitement (j)", "median"), Ancienneté_médiane=("Ancienneté (j)", "median")).reset_index()
    g["Taux d'acceptation"] = (g["Accepté"] / g["Réclamé"].replace(0, np.nan)).fillna(0); return g.sort_values("Réclamé", ascending=False)


for col, name in [("Type", "Par type"), ("Client", "Par client"), ("Gestionnaire", "Par gestionnaire (charge)"), ("Mois réception", "Par mois")]:
    if C[col].astype(str).str.strip().replace("", np.nan).notna().any():
        g = by(col); write_df(wb.create_sheet(name), g if col != "Mois réception" else g.sort_values(col), money0=["Réclamé", "Accepté"], pct=["Taux d'acceptation"], intc=["Réclamations", "Ouvertes", "Hors_délai", "Délai_médian", "Ancienneté_médiane"], autofilter=False)
hd = C[C["Hors délai"] != ""]
write_df(wb.create_sheet("Hors délai"), hd, money=["Montant réclamé", "Montant accepté"], date=["Date BL", "Date réception", "Date clôture"], pct=["Taux d'acceptation"], intc=["Délai de traitement (j)", "Ancienneté (j)"])
summary = {"module": "claims", "arrete": str(ARRETE.date()), "sla": a.sla, "n": n, "ouvertes": int((C["État"] == "Ouvert").sum()), "reclame": float(C["Montant réclamé"].sum()), "accepte": float(C["Montant accepté"].sum()),
           "taux_acceptation": float(C["Montant accepté"].sum() / C["Montant réclamé"].sum()) if C["Montant réclamé"].sum() else None, "delai_median": float(C["Délai de traitement (j)"].median()) if C["Délai de traitement (j)"].notna().any() else None,
           "anciennete_mediane": float(C["Ancienneté (j)"].median()) if C["Ancienneté (j)"].notna().any() else None, "hors_delai": C["Hors délai"].value_counts().to_dict(), "par_type": by("Type").to_dict("records"), "par_gestionnaire": by("Gestionnaire").to_dict("records") if C["Gestionnaire"].str.strip().replace("", np.nan).notna().any() else [],
           "par_mois": by("Mois réception").sort_values("Mois réception").to_dict("records")}
finish(wb, a.out, summary)
print(json.dumps({k: summary[k] for k in ["n", "ouvertes", "reclame", "accepte", "taux_acceptation", "delai_median", "hors_delai"]}, ensure_ascii=False, default=str))
