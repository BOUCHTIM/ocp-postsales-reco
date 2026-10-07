"""Recalcul contractuel du demurrage / despatch par escale (laytime) et comparaison au montant payé, réclamé ou refacturé.
Entrée : --calls time sheets / statements of facts résumés : Navire, Date BL, Port, Opération (chargement/déchargement), Quantité, Cadence contractuelle (t/jour) ou Laytime autorisé (h),
         NOR tendue (date-heure), Début comptage (date-heure) [, Fin opérations (date-heure)], Temps utilisé (h) [ou Fin opérations], Exceptions (h) (pluie, fériés, pannes…), Taux demurrage (USD/jour), Taux despatch (USD/jour) [, Montant payé/réclamé, Source]
Règles par défaut (modifiables par --config) : laytime autorisé = quantité / cadence × 24 ; temps compté = (fin − début) − exceptions ; despatch = demurrage / 2 si taux despatch absent ; arrondi au centième de jour.
Sortie : classeur « Laytime » : calcul par escale, écart vs montant payé/réclamé, synthèse en formules.
Usage : laytime.py --calls sof.xlsx --out laytime.xlsx [--config alias.json]"""
import argparse, json
import pandas as pd, numpy as np
from openpyxl import Workbook
from xlsx_common import *

ap = argparse.ArgumentParser(); ap.add_argument("--calls", required=True); ap.add_argument("--out", required=True); ap.add_argument("--config"); ap.add_argument("--despatch-ratio", type=float, default=0.5)
a = ap.parse_args()
DEF = {"calls": {"Navire": ["VESSEL", "SHIP"], "Date BL": ["B/L DATE", "BL DATE"], "Port": ["PORT", "PORT OF CALL", "LOAD PORT", "DISCHARGE PORT"], "Opération": ["OPERATION", "OPÉRATION", "LOAD/DISCHARGE", "TYPE"], "Quantité": ["QUANTITY", "QTY", "CARGO QTY", "QUANTITÉ"],
                 "Cadence (t/jour)": ["RATE", "LOADING RATE", "DISCHARGE RATE", "CADENCE", "RATE T/DAY", "T/DAY"], "Laytime autorisé (h)": ["LAYTIME ALLOWED", "LAYTIME ALLOWED (H)", "ALLOWED HOURS", "LAYTIME (H)"], "NOR tendue": ["NOR", "NOR TENDERED", "NOTICE OF READINESS"],
                 "Début comptage": ["LAYTIME COMMENCED", "TIME COUNTING FROM", "DÉBUT COMPTAGE", "START", "COMMENCED"], "Fin opérations": ["COMPLETED", "OPERATIONS COMPLETED", "FIN OPÉRATIONS", "END", "COMPLETION"], "Temps utilisé (h)": ["TIME USED", "TIME USED (H)", "LAYTIME USED", "HOURS USED"],
                 "Exceptions (h)": ["EXCEPTIONS", "DEDUCTIONS", "EXCEPTED TIME", "EXCEPTIONS (H)", "TIME NOT COUNTING"], "Taux demurrage (USD/jour)": ["DEMURRAGE RATE", "DEM RATE", "TAUX DEMURRAGE", "DEMURRAGE USD/DAY"], "Taux despatch (USD/jour)": ["DESPATCH RATE", "DES RATE", "TAUX DESPATCH"],
                 "Montant payé/réclamé": ["AMOUNT PAID", "AMOUNT CLAIMED", "MONTANT PAYÉ", "MONTANT RÉCLAMÉ", "FINAL AMOUNT", "AMOUNT"], "Source": ["SOURCE", "ORIGINE"], "Incoterm": ["INCOTERM", "DELIVERY MODE"]}}
cfg = load_config(a.config, DEF)
S = read_any(a.calls, aliases=cfg["calls"], required=["Navire"])
for c in ["Date BL"]:
    if c in S.columns: S[c] = S[c].map(to_date)
for c in ["NOR tendue", "Début comptage", "Fin opérations"]:
    S[c] = pd.to_datetime(S[c], errors="coerce", dayfirst=True) if c in S.columns else pd.NaT
for c in ["Quantité", "Cadence (t/jour)", "Laytime autorisé (h)", "Temps utilisé (h)", "Exceptions (h)", "Taux demurrage (USD/jour)", "Taux despatch (USD/jour)", "Montant payé/réclamé"]:
    S[c] = S[c].map(to_num) if c in S.columns else np.nan
S["Laytime autorisé (h) calc"] = np.where(S["Laytime autorisé (h)"].notna(), S["Laytime autorisé (h)"], (S["Quantité"] / S["Cadence (t/jour)"] * 24).round(2))
dur = (S["Fin opérations"] - S["Début comptage"]).dt.total_seconds() / 3600
S["Temps brut (h)"] = np.where(S["Temps utilisé (h)"].notna(), S["Temps utilisé (h)"] + S["Exceptions (h)"].fillna(0), dur)
S["Temps compté (h)"] = (S["Temps brut (h)"] - S["Exceptions (h)"].fillna(0)).round(2)
S["Écart (h)"] = (S["Temps compté (h)"] - S["Laytime autorisé (h) calc"]).round(2)
S["Jours de demurrage"] = np.where(S["Écart (h)"] > 0, (S["Écart (h)"] / 24).round(4), 0.0)
S["Jours de despatch"] = np.where(S["Écart (h)"] < 0, (-S["Écart (h)"] / 24).round(4), 0.0)
S["Taux despatch retenu"] = np.where(S["Taux despatch (USD/jour)"].notna(), S["Taux despatch (USD/jour)"], S["Taux demurrage (USD/jour)"] * a.despatch_ratio)
S["Demurrage calculé (USD)"] = (S["Jours de demurrage"] * S["Taux demurrage (USD/jour)"]).round(2)
S["Despatch calculé (USD)"] = (S["Jours de despatch"] * S["Taux despatch retenu"]).round(2)
S["Résultat"] = np.where(S["Écart (h)"] > 0, "Demurrage", np.where(S["Écart (h)"] < 0, "Despatch", "Dans le laytime"))
S["Montant calculé (USD)"] = np.where(S["Résultat"] == "Demurrage", S["Demurrage calculé (USD)"], np.where(S["Résultat"] == "Despatch", S["Despatch calculé (USD)"], 0.0))
S["Écart vs montant payé/réclamé (USD)"] = (S["Montant payé/réclamé"].abs() - S["Montant calculé (USD)"]).round(2)
S["Contrôle"] = np.where(S["Montant payé/réclamé"].isna(), "Pas de montant à comparer", np.where(S["Écart vs montant payé/réclamé (USD)"].abs() <= np.maximum(1, 0.005 * S["Montant calculé (USD)"].abs()), "OK", np.where(S["Écart vs montant payé/réclamé (USD)"] > 0, "Payé/réclamé > calculé", "Payé/réclamé < calculé")))
S["Donnée manquante"] = np.where(S["Laytime autorisé (h) calc"].isna(), "laytime (quantité/cadence)", np.where(S["Temps compté (h)"].isna(), "temps (début/fin ou temps utilisé)", np.where(S["Taux demurrage (USD/jour)"].isna(), "taux demurrage", "")))
cols = [c for c in ["Navire", "Date BL", "Port", "Opération", "Incoterm", "Quantité", "Cadence (t/jour)", "Laytime autorisé (h)", "Laytime autorisé (h) calc", "NOR tendue", "Début comptage", "Fin opérations", "Temps utilisé (h)", "Exceptions (h)", "Temps brut (h)", "Temps compté (h)", "Écart (h)", "Jours de demurrage", "Jours de despatch", "Taux demurrage (USD/jour)", "Taux despatch (USD/jour)", "Taux despatch retenu", "Demurrage calculé (USD)", "Despatch calculé (USD)", "Résultat", "Montant calculé (USD)", "Montant payé/réclamé", "Source", "Écart vs montant payé/réclamé (USD)", "Contrôle", "Donnée manquante"] if c in S.columns]
S = S[cols].reset_index(drop=True)
wb = Workbook(); ws = wb.active; ws.title = "Synthèse"
ws["A1"] = "Recalcul contractuel du laytime : demurrage / despatch par escale"; ws["A1"].font = TITLE
ws["A2"] = f"Laytime autorisé = quantité ÷ cadence × 24 h (ou valeur fournie) ; temps compté = temps brut − exceptions ; despatch = {a.despatch_ratio:.0%} du taux demurrage si non fourni. Vérifier les clauses (SHINC/SHEX, NOR, turn time) avant de conclure."; ws["A2"].font = NOTE
cm = colmap(S); n = len(S); R_ = lambda c: rng("Escales", cm[c], n)
r = kpi_rows(ws, 4, [("Escales", f'=COUNTA({R_("Navire")})', "0"), ("Escales en demurrage", f'=COUNTIF({R_("Résultat")},"Demurrage")', "0"), ("Escales en despatch", f'=COUNTIF({R_("Résultat")},"Despatch")', "0"), ("Demurrage calculé total (USD)", f'=SUM({R_("Demurrage calculé (USD)")})', MONEY0), ("Despatch calculé total (USD)", f'=SUM({R_("Despatch calculé (USD)")})', MONEY0),
                    ("Montant payé/réclamé total (USD, abs)", f'=SUMPRODUCT(ABS({R_("Montant payé/réclamé")}))' if "Montant payé/réclamé" in cm else 0, MONEY0), ("Écart cumulé payé/réclamé − calculé (USD)", f'=SUM({R_("Écart vs montant payé/réclamé (USD)")})', MONEY0),
                    ("Contrôles OK", f'=COUNTIF({R_("Contrôle")},"OK")', "0"), ("Payé/réclamé > calculé", f'=COUNTIF({R_("Contrôle")},"Payé/réclamé > calculé")', "0"), ("Payé/réclamé < calculé", f'=COUNTIF({R_("Contrôle")},"Payé/réclamé < calculé")', "0"), ("Escales avec donnée manquante", f'=COUNTIF({R_("Donnée manquante")},"<>")', "0")], "1. Résultat du recalcul")
ws.column_dimensions["A"].width = 46; ws.column_dimensions["B"].width = 18; ws.sheet_view.showGridLines = False
ws2 = wb.create_sheet("Escales"); write_df(ws2, S, money=[c for c in S.columns if "USD" in c], date=["Date BL"], widths={"Contrôle": 26, "Donnée manquante": 30})
for c in ["NOR tendue", "Début comptage", "Fin opérations"]:
    if c in cm:
        for i in range(2, n + 2): ws2[f"{cm[c]}{i}"].number_format = "DD/MM/YYYY HH:MM"
last = get_column_letter(len(S.columns)); cc = cm["Contrôle"]
ws2.conditional_formatting.add(f"A2:{last}{n + 1}", FormulaRule(formula=[f'LEFT(${cc}2,13)="Payé/réclamé "'], fill=FILL_YEL)); ws2.conditional_formatting.add(f"A2:{last}{n + 1}", FormulaRule(formula=[f'${cc}2="OK"'], fill=FILL_GRN))
summary = {"module": "laytime", "escales": n, "demurrage_calcule": float(S["Demurrage calculé (USD)"].sum()), "despatch_calcule": float(S["Despatch calculé (USD)"].sum()), "ecart_cumule": float(S["Écart vs montant payé/réclamé (USD)"].sum()) if S["Écart vs montant payé/réclamé (USD)"].notna().any() else None,
           "controles": S["Contrôle"].value_counts().to_dict(), "manquants": int((S["Donnée manquante"] != "").sum()), "escales_detail": S[["Navire", "Résultat", "Montant calculé (USD)", "Montant payé/réclamé", "Contrôle"]].head(50).to_dict("records")}
finish(wb, a.out, summary)
print(json.dumps({k: summary[k] for k in ["escales", "demurrage_calcule", "despatch_calcule", "ecart_cumule", "controles"]}, ensure_ascii=False, default=str))
