"""Classe les fichiers Excel / CSV d'un dossier, onglet par onglet : export SAP ZDS/ZDP, trackers post-sales (familles), factures de fret,
commandes / livraisons / factures / règlements (order-to-cash), postes ouverts clients (balance âgée), registre LC, registre réclamations, escales (laytime).
Usage : detect_inputs.py <dossier> [--json]"""
import sys, os, json, re
import pandas as pd

KW = {  # (motifs requis ≥ seuil) — en-têtes normalisés en majuscules
    "sap": [r"ORD\.?\s*TYPE", r"ORDER TYPE", r"NUM[ÉE]RO FACTURE", r"BILLING DOCUMENT", r"SOLDE CLIENT", r"MONTANT FACTUR", r"TYPE COMMANDE"],
    "tracker": [r"^VESSEL", r"B/?L DATE", r"CLAIMS? TYPE", r"FINAL AMOUNT", r"DN/?CN", r"PAYMENT STATUS", r"RECEIVER"],
    "fret": [r"DEMURRAGE PAY", r"ARMATEUR", r"LAYTIME \(H\)", r"FACTURE FRET", r"SHIPOWNER", r"FREIGHT INVOICE"],
    "orders": [r"COMMANDE|SALES DOCUMENT|SALES ORDER|^ORDER", r"QUANTIT|QTY", r"PRIX|PRICE|CONTRAT|CONTRACT", r"INCOTERM"],
    "deliveries": [r"LIVRAISON|DELIVERY", r"VESSEL|NAVIRE", r"B/?L DATE|DATE BL|GOODS ISSUE|GI DATE", r"QUANTIT|QTY"],
    "invoices": [r"FACTURE|BILLING|INVOICE", r"CH[ÉE]ANCE|DUE DATE", r"MONTANT|AMOUNT|NET VALUE", r"LIVRAISON|DELIVERY|COMMANDE|ORDER"],
    "payments": [r"R[ÈE]GL|CLEARING|PAID|PAYMENT", r"FACTURE|INVOICE|REFERENCE", r"MONTANT|AMOUNT", r"SOLDE|OPEN"],
    "open_items": [r"OPEN AMOUNT|SOLDE|AMOUNT IN LOCAL", r"CUSTOMER|CLIENT|NAME 1|ACCOUNT", r"DUE|CH[ÉE]ANCE", r"DOCUMENT", r"POSTING|DATE"],
    "lc": [r"\bLC\b|L/C|CREDOC|DOCUMENTARY", r"EXPIR|VALIDITY", r"ISSUING|BANK|BANQUE", r"PRESENTATION|SHIPMENT"],
    "claims": [r"R[ÉE]CLAMATION|CLAIM", r"RECEIVED|R[ÉE]CEPTION|OPENED|OUVERTURE", r"ACCEPT|SETTLED|ACCORD", r"GESTIONNAIRE|OWNER|HANDLER|STATUS|STATUT"],
    "laytime": [r"LAYTIME|NOR|CADENCE|RATE", r"DEMURRAGE RATE|TAUX DEMURRAGE|DEM RATE", r"COMMENC|D[ÉE]BUT|START|TIME USED|TEMPS", r"EXCEPTION|DEDUCT"],
    "sales": [r"^MOIS$|^MONTH$|PERIOD", r"MONTANT|AMOUNT|CA\b|REVENUE"],
}
THRESH = {"sap": 3, "tracker": 3, "fret": 1, "orders": 3, "deliveries": 3, "invoices": 3, "payments": 3, "open_items": 4, "lc": 3, "claims": 3, "laytime": 3, "sales": 2}
PRIORITY = ["sap", "tracker", "fret", "laytime", "lc", "claims", "open_items", "payments", "invoices", "deliveries", "orders", "sales"]


def score(cols):
    cols = [re.sub(r"\s+", " ", str(c)).strip().upper() for c in cols]
    return {k: sum(any(re.search(p, c) for c in cols) for p in pats) for k, pats in KW.items()}


def sheet_role(cols, sheet_name=""):
    sc = score(cols); sn = sheet_name.upper()
    if re.search(r"COMMANDE|VA05|ORDER", sn) and sc["orders"] >= 2: return "orders", sc
    if re.search(r"LIVRAISON|VL06|DELIVER|\bBL\b", sn) and sc["deliveries"] >= 2: return "deliveries", sc
    if re.search(r"FACTURE|VF05|INVOICE", sn) and sc["invoices"] >= 2 and sc["sap"] < 3: return "invoices", sc
    if re.search(r"R[ÈE]GLEMENT|FBL5N|PAYMENT|CLEARING", sn) and sc["payments"] >= 2: return ("open_items" if sc["open_items"] >= 4 and "POSTES" in sn else "payments"), sc
    if re.search(r"POSTES|OPEN ITEMS|BALANCE|AGED", sn) and sc["open_items"] >= 3: return "open_items", sc
    if re.search(r"^CA\b|CHIFFRE|SALES", sn) and sc["sales"] >= 2: return "sales", sc
    for k in PRIORITY:
        if sc[k] >= THRESH[k]: return k, sc
    return "autre", sc


def classify(path):
    try: xl = pd.ExcelFile(path); names = xl.sheet_names
    except Exception as e: return {"file": path, "kind": "illisible", "detail": str(e)}
    sheets = []
    for sh in names:
        try: d = pd.read_excel(path, sheet_name=sh, nrows=3)
        except Exception: continue
        role, sc = sheet_role(d.columns, sh); sheets.append({"sheet": sh, "role": role, "scores": sc})
    roles = [s["role"] for s in sheets]
    kind = next((k for k in PRIORITY if k in roles), "autre")
    if {"orders", "deliveries", "invoices"} <= set(roles): kind = "o2c"
    fams = sorted({("demdesp" if re.search(r"^DEM", s["sheet"].upper()) else "priceadj" if re.search(r"PRICE ADJ|DISCOUNT|REMISE|CLAIM", s["sheet"].upper()) else "") for s in sheets if s["role"] == "tracker"} - {""})
    return {"file": path, "kind": kind, "role": kind, "sheets": [s["sheet"] for s in sheets], "sheet_roles": {s["sheet"]: s["role"] for s in sheets}, "familles": fams}


def detect(folder):
    out = []
    for f in sorted(os.listdir(folder)):
        if f.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")) and not f.startswith("~$"): out.append(classify(os.path.join(folder, f)))
    trackers = [o for o in out if o["kind"] == "tracker"]
    trackers.sort(key=lambda o: os.path.getsize(o["file"]), reverse=True)
    for i, t in enumerate(trackers): t["role"] = "tracker" if i == 0 else "tracker_secondaire"
    return out


def find_sheet(found, role):
    """Premier (fichier, onglet) portant ce rôle, sous la forme 'fichier:onglet'."""
    for o in found:
        for sh, r in o.get("sheet_roles", {}).items():
            if r == role: return f"{o['file']}:{sh}"
    return None


if __name__ == "__main__":
    res = detect(sys.argv[1])
    if "--json" in sys.argv: print(json.dumps(res, ensure_ascii=False, indent=1))
    else:
        for o in res: print(f"{o['kind']:12s} {o.get('role',''):18s} {os.path.basename(o['file'])}  rôles={o.get('sheet_roles')}  familles={o.get('familles')}")
