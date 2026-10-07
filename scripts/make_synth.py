"""Jeu de données synthétique cohérent pour tester tous les modules (o2c, aging, lc, claims, laytime) : 24 cargaisons sur 12 mois,
avec des cas voulus (non facturé, retard de paiement, LC expirée, réclamation hors délai, demurrage payé ≠ calculé). Usage : make_synth.py <dossier>"""
import sys, os, random
import pandas as pd, numpy as np
from datetime import timedelta

random.seed(7); np.random.seed(7)
OUT = sys.argv[1]; os.makedirs(OUT, exist_ok=True)
clients = ["CLIENT A", "CLIENT B", "CLIENT C", "CLIENT D"]; vessels = ["ALPHA STAR", "BETA OCEAN", "GAMMA TRADER", "DELTA WIND", "EPSILON", "ZETA QUEEN", "ETA PRIDE", "THETA SKY", "IOTA BREEZE", "KAPPA SUN", "LAMBDA MOON", "MU HORIZON"]
ARRETE = pd.Timestamp("2026-09-30"); orders = []; dels = []; invs = []; pays = []; opens = []; lcs = []; claims = []; calls = []
for i in range(24):
    bl = pd.Timestamp("2025-10-05") + timedelta(days=15 * i); v = vessels[i % 12]; cli = clients[i % 4]; inc = "CFR" if i % 3 else "FOB"; qty = 25000 + 1000 * (i % 7); price = 500 + 10 * (i % 5); amt = qty * price
    oid = f"600{1000 + i}"; did = f"800{2000 + i}"; iid = f"900{3000 + i}"
    orders.append({"N° commande": oid, "Client": cli, "Contrat": f"CTR-{bl.year}-{i:03d}", "Produit": "DAP STANDARD" if i % 2 else "TSP", "Quantité commandée": qty, "Incoterm": inc, "Prix unitaire": price, "Devise": "USD", "Date commande": bl - timedelta(days=25)})
    dels.append({"N° livraison": did, "N° commande": oid, "Navire": v, "Date BL": bl, "Quantité livrée": qty - (300 if i % 5 == 0 else 0), "Port": "Jorf Lasfar"})
    if i == 22 or i == 23: continue  # livrées non facturées (une récente, une ancienne)
    inv_date = bl + timedelta(days=4 + (20 if i == 10 else 0)); due = inv_date + timedelta(days=90); billed = (qty - (300 if i % 5 == 0 else 0)) * price - (5000 if i == 7 else 0)
    invs.append({"N° facture": iid, "N° livraison": did, "N° commande": oid, "Date facture": inv_date, "Montant": billed, "Échéance": due, "Client": cli, "Devise": "USD"})
    paid = i < 16 and i != 13; part = i == 16
    pays.append({"N° facture": iid, "Montant réglé": billed if paid else (billed * 0.6 if part else 0), "Date règlement": due - timedelta(days=5) + timedelta(days=40 * (i % 2)) if paid else pd.NaT, "Solde": 0 if paid else (billed * 0.4 if part else billed)})
    if not paid:
        opens.append({"Client": cli, "N° document": iid, "Type": "Facture", "Date document": inv_date, "Échéance": due, "Montant": billed * 0.4 if part else billed, "Devise": "USD", "Référence": f"{v} BL {bl:%d/%m/%Y}", "Navire": v, "Texte": ""})
    if i % 4 == 0: lcs.append({"N° LC": f"LC-{bl.year}-{100 + i}", "Client": cli, "Banque émettrice": ["BANK OF X", "BANK Y", "BANK Z"][i % 3], "Banque notificatrice": "ATTIJARI", "Montant": billed, "Devise": "USD", "Date d'émission": bl - timedelta(days=40), "Date limite d'expédition": bl + timedelta(days=(-3 if i == 20 else 10)), "Date d'expiration": bl + timedelta(days=(30 if i != 12 else -2)),
                               "Navire": v, "Date BL": bl, "Incoterm": inc, "N° facture": iid, "Date présentation documents": (bl + timedelta(days=12)) if i not in (12, 20) else pd.NaT, "Date acceptation / paiement": (due - timedelta(days=5)) if paid else pd.NaT, "Tenor (j)": 90, "Réserves": "late shipment" if i == 20 else "", "Frais bancaires": 3500 + 100 * i, "Statut": "PAID" if paid else "OPEN"})
    if i % 3 == 1: claims.append({"N° réclamation": f"CLM-{bl.year}-{i:03d}", "Client": cli, "Navire": v, "Date BL": bl, "Produit": "DAP STANDARD", "Type": ["Qualité", "Quantité", "Documentation"][i % 3], "Cause": "", "Date réception": bl + timedelta(days=20), "Montant réclamé": 60000 + 5000 * i, "Montant accepté": (30000 if i % 2 else 0) if i < 18 else np.nan, "Statut": "CLOSED" if i < 18 else "OPEN", "Date clôture": (bl + timedelta(days=20 + (25 if i < 12 else 50))) if i < 18 else pd.NaT, "Gestionnaire": ["S. ALAMI", "K. BENANI"][i % 2], "Action": "", "Commentaire": ""})
    calls.append({"Navire": v, "Date BL": bl, "Port": "Jorf Lasfar", "Opération": "Chargement", "Incoterm": inc, "Quantité": qty, "Cadence (t/jour)": 10000, "Laytime autorisé (h)": "", "NOR tendue": bl - timedelta(days=4, hours=16), "Début comptage": bl - timedelta(days=4, hours=10), "Fin opérations": bl - timedelta(days=4, hours=10) + timedelta(hours=50 + 12 * (i % 4)), "Temps utilisé (h)": "", "Exceptions (h)": 6 if i % 2 else 0, "Taux demurrage (USD/jour)": 25000, "Taux despatch (USD/jour)": 12500, "Montant payé/réclamé": [48000, 0, 31250, 15000][i % 4] + (9000 if i == 9 else 0), "Source": "tracker"})
with pd.ExcelWriter(f"{OUT}/o2c_sources.xlsx") as w:
    pd.DataFrame(orders).to_excel(w, sheet_name="Commandes (VA05)", index=False); pd.DataFrame(dels).to_excel(w, sheet_name="Livraisons (VL06 - BL)", index=False); pd.DataFrame(invs).to_excel(w, sheet_name="Factures (VF05)", index=False); pd.DataFrame(pays).to_excel(w, sheet_name="Règlements (FBL5N)", index=False)
opens.append({"Client": "CLIENT B", "N° document": "1700001", "Type": "Acompte", "Date document": ARRETE - timedelta(days=10), "Échéance": ARRETE - timedelta(days=10), "Montant": -250000, "Devise": "USD", "Référence": "", "Navire": "", "Texte": "acompte non imputé"})
with pd.ExcelWriter(f"{OUT}/balance_agee.xlsx") as w:
    pd.DataFrame(opens).to_excel(w, sheet_name="Postes ouverts (FBL5N)", index=False); pd.DataFrame([{"Mois": (pd.Timestamp("2025-10-01") + pd.DateOffset(months=k)).strftime("%Y-%m"), "Montant": 2 * 13000000} for k in range(12)]).to_excel(w, sheet_name="CA mensuel (optionnel)", index=False)
pd.DataFrame(lcs).to_excel(f"{OUT}/registre_lc.xlsx", sheet_name="Registre LC", index=False); pd.DataFrame(claims).to_excel(f"{OUT}/registre_reclamations.xlsx", sheet_name="Registre", index=False); pd.DataFrame(calls).to_excel(f"{OUT}/laytime_sof.xlsx", sheet_name="Escales (SOF)", index=False)
print("synth:", {"commandes": len(orders), "livraisons": len(dels), "factures": len(invs), "postes ouverts": len(opens), "lc": len(lcs), "réclamations": len(claims), "escales": len(calls)})
