"""Extrait du classeur de rapprochement (produit par reconcile.py) toutes les données dont les livrables ont besoin,
dans un seul JSON : chiffres, listes d'anomalies, qualité des données, fautes de frappe probables, contexte de mission.
Les templates (deck.js, report.js, note.js) ne lisent que ce fichier : aucun chiffre ni nom n'est écrit à la main.

Usage : make_ref.py <rapprochement.xlsx> <ref.json> [--sap <export SAP.xlsx>] [--mission "Nom de la mission"]
        [--trackers "fichier1.xlsx;fichier2.xlsx"] [--fret <factures_fret.xlsx>] [--start 2024-01-01] [--delai-alerte 180]"""
import pandas as pd, json, sys, re, os, argparse
from datetime import date

ap = argparse.ArgumentParser()
ap.add_argument("rapprochement"); ap.add_argument("out")
ap.add_argument("--sap", default=None); ap.add_argument("--mission", default="Rapprochement demurrage / despatch")
ap.add_argument("--trackers", default=""); ap.add_argument("--fret", default=None); ap.add_argument("--famille", default="demdesp"); ap.add_argument("--start", default="2024-01-01"); ap.add_argument("--delai-alerte", type=int, default=180)
a = ap.parse_args()
P = a.rapprochement
r = pd.read_excel(P, sheet_name="Rapprochement"); nt = pd.read_excel(P, sheet_name="SAP non suivi"); det = pd.read_excel(P, sheet_name="Détail tracker"); sap = pd.read_excel(P, sheet_name="Base SAP 2024+")
M = "Montant tracker (FINAL AMOUNT)"; S = "SAP montant refacturé (abs)"; E = "SAP encaissé / imputé (abs)"; SO = "SAP solde client (abs)"
sap["Numéro Facture"] = sap["Numéro Facture"].astype(str)


def J(df): return json.loads(df.to_json(orient="records", force_ascii=False, date_format="iso"))


def blk(col):
    g = r.groupby(col).agg(n=(M, "size"), nonref=("Nb docs SAP", lambda s: (s == 0).sum()), trk=(M, "sum"), sap=(S, "sum"), enc=(E, "sum"), solde=(SO, "sum"))
    g["nonref_amt"] = r[r["Nb docs SAP"] == 0].groupby(col)[M].sum().reindex(g.index).fillna(0)
    return J(g.round(2).reset_index())


nr = r[r["Nb docs SAP"] == 0].copy()
nr["pay"] = nr["Statut paiement (tracker)"].astype(str).map(lambda s: "NOT YET STARTED" if "NOT YET" in s else ("PAID" if "PAID" in s else ("PENDING" if "PENDING" in s else "AUTRE")))
r["Incoterm"] = r["Mode livraison"].astype(str).str.upper().str.extract(r"(FOB|CFR|CIF|DPU|DAP|CPT)")[0].fillna("?") if "Mode livraison" in r.columns else "?"

# --- fautes de frappe probables : n° cité absent de SAP mais n° voisin (même navire, 1 chiffre différent) présent
sap_by_v = sap.groupby(sap["Navire"].astype(str).str.upper().str.strip())["Numéro Facture"].apply(set).to_dict()
typos = []
for _, k in r[r["DN/CN tracker listés"] > r["…dont trouvés dans SAP"]].iterrows():
    v = str(k["Navire"]).upper().strip(); cands = sap_by_v.get(v, set())
    for n in re.findall(r"\b9\d{7}\b", str(k["DN/CN n° (tracker)"])):
        if n in cands: continue
        near = [c for c in cands if len(c) == len(n) and sum(x != y for x, y in zip(c, n)) == 1]
        if near: typos.append({"Navire": k["Navire"], "Date BL": str(k["Date BL (début)"])[:10], "cité": n, "probable": near[0], "montant_doc": float(sap.loc[sap["Numéro Facture"] == near[0], "Montant Facturée"].abs().sum())})

# --- écarts avec lecture automatique
ec = r[r["Contrôle écart"].isin(["SAP < tracker", "SAP > tracker"])].copy()
typo_keys = {(t["Navire"], t["Date BL"]) for t in typos}


def lecture(x):
    if (x["Navire"], str(x["Date BL (début)"])[:10]) in typo_keys: return "N° de DN/CN probablement mal saisi dans le tracker (voir fautes de frappe) : l'écart disparaît avec le bon document."
    if x["Autres docs SAP même voyage (nb)"] > 0: return "Documents SAP supplémentaires non cités par le tracker sur ce voyage (annulation / réémission ?)."
    if x["Contrôle écart"] == "SAP < tracker": return f"SAP facture {x['Écart tracker – SAP']:,.2f} USD de moins que le montant suivi : document manquant dans l'export ou montant à corriger."
    return "SAP facture plus que le montant suivi : vérifier la version du montant final dans le tracker."


ec["Lecture"] = ec.apply(lecture, axis=1) if len(ec) else []
ec_cols = ["Navire", "Date BL (début)", "Type SAP attendu", "Client (tracker)", M, S, "Écart tracker – SAP", "Contrôle écart", "DN/CN n° (tracker)", "Docs SAP (n° facture)", "DN/CN tracker listés", "…dont trouvés dans SAP", "Lecture"]

# --- qualité des données
det_in = det[det["Périmètre"] == "2024+"]
quality = {
    "amounts_text": int(det["FINAL AMOUNT (brut)"].astype(str).str.contains(r"^\s*-?\d{1,3}(?:[  ]\d{3})+(?:[,.]\d+)?\s*$|^\s*\d+,\d{2}\s*$", regex=True).sum()),
    "bl_ranges": int(det["B/L DATE (brut)"].astype(str).str.contains(" - ").sum()),
    "trailing_spaces": int(det["VESSEL"].astype(str).map(lambda v: v != v.strip()).sum()),
    "zero_amounts": sorted(set(r.loc[r[M] == 0, "Navire"].astype(str))),
    "multi_voyage": sorted(set(r.loc[r["Méthode de rapprochement"].astype(str).str.contains("autre BL"), "Navire"].astype(str))),
    "multi_lignes": int((det_in.groupby("Clé").size() > 1).sum()),
    "status_free_text": int((~det["PAYMENT STATUS"].astype(str).str.strip().str.upper().isin(["PAID", "PENDING", "NOT YET STARTED", "CLOSED", "NAN"])).sum()),
    "typos": typos,
}
# --- navires
vk = lambda s: s.astype(str).str.upper().str.replace(r"[^A-Z0-9 ]", " ", regex=True).str.replace(r"\s+", " ", regex=True).str.strip()
tv = set(vk(det_in["VESSEL"])); sv = set(vk(sap["Navire"])); total_export = None
if a.sap and os.path.exists(a.sap):
    try:
        _x = pd.ExcelFile(a.sap); _e = pd.read_excel(a.sap, sheet_name="Feuil1" if "Feuil1" in _x.sheet_names else _x.sheet_names[0])
        _col = next((c for c in _e.columns if str(c).strip().upper() in ("NAVIRE", "VESSEL", "SHIP")), None)
        if _col is not None: sv |= set(vk(_e[_col]))
        total_export = int(len(_e))
    except Exception: pass
vessels = {"tracker": len(tv), "absent_sap": len(tv - sv), "pct_absent": round(100 * len(tv - sv) / max(1, len(tv)), 1)}

# --- listes
paid_pending = nr[nr.pay.isin(["PAID", "PENDING"])]
pp_clients = J(paid_pending.groupby("Client (tracker)").agg(n=(M, "size"), amt=(M, "sum")).round(0).sort_values("amt", ascending=False).head(8).reset_index())
pp_regions = J(paid_pending.groupby("Région (tracker)").agg(n=(M, "size"), amt=(M, "sum")).round(0).sort_values("amt", ascending=False).reset_index())
nonenc = r[r["Statut rapprochement"].astype(str).str.contains("non encaissé|partiellement")].sort_values(SO, ascending=False)
nonenc_top = J(nonenc[["Navire", "Date BL (début)", "Client (tracker)", S, E, SO]].head(5))
nonenc_clients = J(nonenc.groupby("Client (tracker)").agg(n=(SO, "size"), solde=(SO, "sum")).round(0).sort_values("solde", ascending=False).head(5).reset_index())
alertes = J(r[r["Alerte"].notna()][["Navire", "Date BL (début)", "Type SAP attendu", M, S, "Autres docs SAP même voyage (nb)", "Autres docs SAP même voyage (montant)", "Autres docs SAP même voyage (n°)", "Docs SAP – date création"]])
annee = blk("Année BL")
first_year = int(min(x["Année BL"] for x in annee)); last_year = int(max(x["Année BL"] for x in annee))
tr_files = [os.path.basename(f) for f in a.trackers.split(";") if f.strip()]
sources = {str(k): int(v) for k, v in det["Source"].astype(str).value_counts().items()}
apac_subset = (bool((det["Présent fichier APAC"] == "Oui").sum() > 0 and not any(k.startswith("T") for k in sources))) if len(tr_files) > 1 else None

ref = {
    "mission": {"name": a.mission, "famille": a.famille, "famille_label": (r["Famille"].iloc[0] if "Famille" in r.columns and len(r) else "demurrage / despatch"), "date_rapport": date.today().strftime("%d/%m/%Y"), "arrete": pd.to_datetime(sap["Date Création Facture"]).max().strftime("%d/%m/%Y") if "Date Création Facture" in sap.columns else "",
                "start": a.start, "delai_alerte": a.delai_alerte, "periode": f"{first_year}-{last_year}", "fichier_sap": os.path.basename(a.sap) if a.sap else "", "fichiers_tracker": tr_files, "fichier_fret": os.path.basename(a.fret) if a.fret else None,
                "tracker_secondaire_sous_ensemble": apac_subset, "lignes_par_source": sources},
    "total": dict(n=len(r), nonref=int((r["Nb docs SAP"] == 0).sum()), trk=round(r[M].sum(), 2), sap=round(r[S].sum(), 2), enc=round(r[E].sum(), 2), solde=round(r[SO].sum(), 2), nonref_amt=round(nr[M].sum(), 2), matched=int((r["Nb docs SAP"] > 0).sum()),
                  ecarts_ok=int((r["Contrôle écart"] == "OK").sum()), nonenc_n=int(len(nonenc)), nonenc_solde=round(nonenc[SO].sum(), 2)),
    "annee": annee, "type": blk("Type réclamation"), "type_sap": blk("Type SAP attendu"), "statut": blk("Statut rapprochement"), "methode": blk("Méthode de rapprochement"), "region": blk("Région (tracker)"), "incoterm": blk("Incoterm"),
    "nonref_pay": J(nr.groupby("pay").agg(n=(M, "size"), amt=(M, "sum")).round(2).reset_index()),
    "nonref_pay_annee": J(nr.groupby(["Année BL", "pay"]).agg(n=(M, "size"), amt=(M, "sum")).round(2).reset_index()),
    "ecarts": J(ec[ec_cols].astype({"Date BL (début)": str})) if len(ec) else [], "alertes": alertes,
    "nonenc": J(nonenc[["Navire", "Date BL (début)", "Type SAP attendu", "Client (tracker)", S, E, SO, "Statut paiement (tracker)", "Statut rapprochement"]]),
    "nonenc_top": nonenc_top, "nonenc_clients": nonenc_clients,
    "paid_pending_sansdoc": J(paid_pending[["Navire", "Date BL (début)", "Type réclamation", "Client (tracker)", "Région (tracker)", M, "DN/CN n° (tracker)", "pay"]].sort_values("Date BL (début)")),
    "paid_pending_clients": pp_clients, "paid_pending_regions": pp_regions,
    "dncn": dict(listes=int(r["DN/CN tracker listés"].sum()), trouves=int(r["…dont trouvés dans SAP"].sum()), non_trouves=int(r["DN/CN tracker listés"].sum() - r["…dont trouvés dans SAP"].sum()),
                 pct=round(100 * r["…dont trouvés dans SAP"].sum() / max(1, r["DN/CN tracker listés"].sum()), 1),
                 exemples_absents=[{"Navire": x["Navire"], "DN/CN n° (tracker)": (re.findall(r"\b9\d{7}\b", str(x["DN/CN n° (tracker)"])) or ["n° non renseigné"])[0]} for _, x in paid_pending.sort_values(M, ascending=False).head(3).iterrows()]),
    "vessels": vessels, "quality": quality,
    "nonsuivi": dict(n=len(nt), par_type=J(nt.groupby("Type SAP")["Montant SAP"].agg(["size", "sum"]).round(2).reset_index()), par_region=J(nt.groupby("Région SAP")["Montant SAP"].agg(["size", "sum"]).round(2).reset_index().sort_values("sum", ascending=False)), solde=round(nt["Solde client (abs)"].sum(), 2)),
    "detail": dict(total=len(det), perim={str(k): int(v) for k, v in det["Périmètre"].value_counts().items()}, apac={str(k): int(v) for k, v in det["Présent fichier APAC"].value_counts().items()}, amt_perim={str(k): float(v) for k, v in det.groupby("Périmètre")["Montant (num)"].sum().round(2).items()}, plages=quality["bl_ranges"]),
    "sap": dict(n=len(sap), types={str(k): int(v) for k, v in sap["Ord.Type"].value_counts().items()}, amt={str(k): float(v) for k, v in sap.groupby("Ord.Type")["Montant Facturée"].sum().round(2).items()}, rapp={str(k): int(v) for k, v in sap["Rapproché tracker"].value_counts().items()}, annee={int(k): int(v) for k, v in sap["Année BL"].value_counts().sort_index().items()}, total_export=total_export),
}


def clean(o):
    if isinstance(o, dict): return {str(k): clean(v) for k, v in o.items()}
    if isinstance(o, list): return [clean(v) for v in o]
    if isinstance(o, float) and o != o: return None
    if hasattr(o, "item"):
        try: return clean(o.item())
        except Exception: return str(o)
    return o


json.dump(clean(ref), open(a.out, "w"), ensure_ascii=False, indent=1, default=str, allow_nan=False)
print("ref written", a.out, "| clés:", ref["total"]["n"], "| typos:", len(typos), "| écarts:", len(ec), "| alertes:", len(alertes))
