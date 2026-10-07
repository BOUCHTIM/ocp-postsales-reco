"""Tableau de bord Customer Service (direction) : agrège les résumés JSON produits par les modules (rapprochement / expert, o2c, aging, lc, claims, laytime)
en un classeur « Tableau de bord » (KPI du mois, évolution si plusieurs mois fournis, alertes, plan de relance) + un résumé Markdown.
Usage : dashboard.py --out "Tableau de bord.xlsx" [--md "Tableau de bord.md"] --mois 2026-09 [--summary <fichier_summary.json>]... [--history <dossier des tableaux de bord précédents>]
Les résumés sont reconnus par leur champ "module" (ou "typologie" pour l'analyse experte)."""
import argparse, json, os, glob
import pandas as pd, numpy as np
from openpyxl import Workbook
from xlsx_common import *

ap = argparse.ArgumentParser(); ap.add_argument("--out", required=True); ap.add_argument("--md"); ap.add_argument("--mois", default=pd.Timestamp.today().strftime("%Y-%m")); ap.add_argument("--summary", action="append", default=[]); ap.add_argument("--history")
a = ap.parse_args()
S = {}
for f in a.summary:
    try: j = json.load(open(f))
    except Exception as e: print("ignoré", f, e); continue
    S[j.get("module", "expert" if "typologie" in j else os.path.basename(f))] = j
kpi = []  # (domaine, indicateur, valeur, unité, lecture)
if "expert" in S:
    e = S["expert"]; T = {t["Typologie de fuite"][:3].strip(): t for t in e["typologie"]}; sc = {s["Scénario"]: s["Montant (USD)"] for s in e["scenarios"]}
    tot_recl = sum(t["Réclamé"] for t in e["typologie"]); tot_ref = sum(t["Refacturé"] for t in e["typologie"]); tot_enc = sum(t["Encaissé"] for t in e["typologie"])
    kpi += [("Réclamations (tracker ↔ SAP)", "Dossiers suivis", e["n"], "clés", ""), ("Réclamations (tracker ↔ SAP)", "Montant réclamé", tot_recl, "USD", ""), ("Réclamations (tracker ↔ SAP)", "Taux de refacturation", tot_ref / tot_recl if tot_recl else 0, "%", "refacturé / réclamé"),
            ("Réclamations (tracker ↔ SAP)", "Taux d'encaissement", tot_enc / tot_ref if tot_ref else 0, "%", "encaissé / refacturé"), ("Réclamations (tracker ↔ SAP)", "Récupérable certain (créances)", sc.get("Récupérable certain", 0), "USD", "T4/T5 créances"),
            ("Réclamations (tracker ↔ SAP)", "Risque de prescription (créances non démarrées anciennes)", sc.get("Risque de prescription", 0), "USD", "T1b"), ("Réclamations (tracker ↔ SAP)", "À confirmer (sans document dans l'export)", sc.get("À confirmer (export filtré ?)", 0), "USD", "T2 créances"),
            ("Réclamations (tracker ↔ SAP)", "Dettes à régler / provisionner", sc.get("Dettes à régler ou provisionner", 0), "USD", ""), ("Réclamations (tracker ↔ SAP)", "Délai médian BL → DN/CN", (e["delais"][0] or {}).get("Médiane"), "jours", ""), ("Réclamations (tracker ↔ SAP)", "Délai médian DN/CN → règlement", (e["delais"][1] or {}).get("Médiane"), "jours", ""),
            ("Réclamations (tracker ↔ SAP)", "Contrôles automatiques en échec", sum(1 for c in e.get("controles", []) if c.get("Résultat") == "KO"), "tests", "")]
if "o2c" in S:
    o = S["o2c"]; kpi += [("Order-to-cash", "Cargaisons", o["cargaisons"], "", ""), ("Order-to-cash", "Montant facturé", o["facture"], "USD", ""), ("Order-to-cash", "Solde client ouvert (cargaisons)", o["solde"], "USD", ""), ("Order-to-cash", "Cargaisons livrées non facturées", o["non_factures"], "", "à facturer"),
                          ("Order-to-cash", "Délai médian BL → facture", o.get("delai_bl_facture_median"), "jours", ""), ("Order-to-cash", "Délai médian facture → règlement", o.get("delai_facture_reglement_median"), "jours", "")]
if "aging" in S:
    g = S["aging"]; old = g["tranches"].get("91-180 j", 0) + g["tranches"].get("> 180 j", 0)
    kpi += [("Balance âgée", "Encours débiteur", g["encours_debit"], "USD", ""), ("Balance âgée", "Encours net", g["encours_net"], "USD", "créances − crédits non imputés"), ("Balance âgée", "DSO (count-back)", g.get("dso"), "jours", ""), ("Balance âgée", "Part de l'encours > 90 j", old / g["encours_debit"] if g["encours_debit"] else 0, "%", ""),
            ("Balance âgée", "Documents à relancer", g["relances"], "", ""), ("Balance âgée", "Montant à relancer", g["montant_relance"], "USD", "")]
if "lc" in S:
    l = S["lc"]; al = {k: v for k, v in l["alertes"].items() if k not in ("", "Réglé", "LC ouverte, expédition à venir")}
    kpi += [("Crédits documentaires", "LC suivies", l["lc"], "", ""), ("Crédits documentaires", "Montant non réglé sous LC", l["non_regle"], "USD", ""), ("Crédits documentaires", "LC en alerte", sum(al.values()), "", " ; ".join(f"{k} : {v}" for k, v in al.items())), ("Crédits documentaires", "LC avec réserves", l["reserves"], "", ""),
            ("Crédits documentaires", "Frais bancaires cumulés", l["frais"], "USD", ""), ("Crédits documentaires", "Délai médian BL → présentation", l.get("delai_bl_presentation_median"), "jours", ""), ("Crédits documentaires", "Délai médian présentation → paiement", l.get("delai_presentation_paiement_median"), "jours", "")]
if "claims" in S:
    c = S["claims"]; kpi += [("Réclamations clients (registre)", "Réclamations reçues", c["n"], "", ""), ("Réclamations clients (registre)", "Backlog ouvert", c["ouvertes"], "", ""), ("Réclamations clients (registre)", "Montant réclamé", c["reclame"], "USD", ""), ("Réclamations clients (registre)", "Taux d'acceptation", c.get("taux_acceptation"), "%", "accepté / réclamé"),
                             ("Réclamations clients (registre)", "Délai médian de traitement", c.get("delai_median"), "jours", f"SLA {c['sla']} j"), ("Réclamations clients (registre)", f"Dossiers hors délai", sum(v for k, v in c["hors_delai"].items() if k and k != "Dans le délai"), "", "ouverts ou traités au-delà du SLA")]
if "laytime" in S:
    y = S["laytime"]; kpi += [("Laytime", "Escales recalculées", y["escales"], "", ""), ("Laytime", "Demurrage calculé", y["demurrage_calcule"], "USD", ""), ("Laytime", "Despatch calculé", y["despatch_calcule"], "USD", ""), ("Laytime", "Écart cumulé payé/réclamé − calculé", y.get("ecart_cumule"), "USD", ""), ("Laytime", "Escales avec écart", sum(v for k, v in y["controles"].items() if k.startswith("Payé")), "", "")]
K = pd.DataFrame(kpi, columns=["Domaine", "Indicateur", "Valeur", "Unité", "Lecture"]); K.insert(0, "Mois", a.mois)
# ---- historique (tableaux de bord précédents : onglet KPI de chaque fichier)
hist = [K]
if a.history and os.path.isdir(a.history):
    for f in sorted(glob.glob(os.path.join(a.history, "*.xlsx"))):
        if os.path.abspath(f) == os.path.abspath(a.out): continue
        try:
            h = pd.read_excel(f, sheet_name="KPI"); hist.append(h[["Mois", "Domaine", "Indicateur", "Valeur", "Unité", "Lecture"]])
        except Exception: pass
H = pd.concat(hist, ignore_index=True).drop_duplicates(["Mois", "Domaine", "Indicateur"], keep="first")
evo = H.pivot_table(index=["Domaine", "Indicateur"], columns="Mois", values="Valeur", aggfunc="first").reset_index() if H["Mois"].nunique() > 1 else None
# ---- alertes consolidées
alerts = []
if "expert" in S:
    for c in S["expert"].get("controles", []):
        if c.get("Résultat") == "KO": alerts.append(("Réclamations", c["Contrôle"], c.get("Détail", "")))
if "o2c" in S:
    for k, v in S["o2c"].get("alertes", {}).items():
        if k: alerts.append(("Order-to-cash", k, f"{v} cargaison(s)"))
if "lc" in S:
    for k, v in S["lc"].get("alertes", {}).items():
        if k and k not in ("Réglé", "LC ouverte, expédition à venir"): alerts.append(("LC", k, f"{v} LC"))
if "claims" in S:
    for k, v in S["claims"].get("hors_delai", {}).items():
        if k and k != "Dans le délai": alerts.append(("Réclamations clients", k, f"{v} dossier(s)"))
if "aging" in S and S["aging"].get("relances"): alerts.append(("Balance âgée", "Documents à relancer", f"{S['aging']['relances']} documents, {S['aging']['montant_relance']:,.0f} USD".replace(",", " ")))
A = pd.DataFrame(alerts, columns=["Domaine", "Alerte", "Détail"])
# ---- classeur
wb = Workbook(); ws = wb.active; ws.title = "Tableau de bord"
ws["A1"] = f"Tableau de bord Customer Service — {a.mois}"; ws["A1"].font = TITLE
ws["A2"] = "Une ligne par indicateur ; sources : résumés des modules (rapprochement / expert, order-to-cash, balance âgée, LC, réclamations, laytime). Les valeurs sont celles des classeurs détaillés."; ws["A2"].font = NOTE
r = 4; cur = None
for _, k in K.iterrows():
    if k["Domaine"] != cur:
        cur = k["Domaine"]; r += 1; ws.cell(row=r, column=1, value=cur).font = BOLD; r += 1
    ws.cell(row=r, column=1, value="   " + k["Indicateur"]).font = BASE
    c = ws.cell(row=r, column=2, value=(None if k["Valeur"] is None or (isinstance(k["Valeur"], float) and np.isnan(k["Valeur"])) else k["Valeur"])); c.font = BASE; c.border = BORDER
    c.number_format = PCT if k["Unité"] == "%" else (MONEY0 if k["Unité"] == "USD" else "0")
    ws.cell(row=r, column=3, value=k["Unité"]).font = NOTE; ws.cell(row=r, column=4, value=k["Lecture"]).font = NOTE; r += 1
ws.column_dimensions["A"].width = 58; ws.column_dimensions["B"].width = 16; ws.column_dimensions["C"].width = 8; ws.column_dimensions["D"].width = 60; ws.sheet_view.showGridLines = False
write_df(wb.create_sheet("KPI"), K, widths={"Indicateur": 52, "Lecture": 50})
if evo is not None: write_df(wb.create_sheet("Évolution"), evo, widths={"Indicateur": 52}, autofilter=False)
write_df(wb.create_sheet("Alertes"), A, widths={"Alerte": 60, "Détail": 50})
finish(wb, a.out, {"module": "dashboard", "mois": a.mois, "kpi": K.to_dict("records"), "alertes": A.to_dict("records"), "modules": list(S)})
if a.md:
    fmtv = lambda v, u: ("–" if v is None or (isinstance(v, float) and np.isnan(v)) else (f"{v:.1%}" if u == "%" else f"{v:,.0f} USD".replace(",", " ") if u == "USD" else f"{v:,.0f} {u}".replace(",", " ").strip()))
    lines = [f"# Tableau de bord Customer Service — {a.mois}", ""]
    for dom in K["Domaine"].unique():
        lines += [f"## {dom}", "", "| Indicateur | Valeur | Lecture |", "|---|---|---|"] + [f"| {k['Indicateur']} | {fmtv(k['Valeur'], k['Unité'])} | {k['Lecture']} |" for _, k in K[K["Domaine"] == dom].iterrows()] + [""]
    if len(A): lines += ["## Alertes", ""] + [f"- **{x['Domaine']}** — {x['Alerte']} ({x['Détail']})" for _, x in A.iterrows()] + [""]
    lines += ["Modules présents : " + ", ".join(S) + ". Les modules absents n'apparaissent pas ; fournir leurs fichiers pour compléter le tableau de bord."]
    open(a.md, "w").write("\n".join(lines))
print(json.dumps({"modules": list(S), "kpi": len(K), "alertes": len(A)}, ensure_ascii=False))
