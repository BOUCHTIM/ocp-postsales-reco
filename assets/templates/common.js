// Helpers partagés par deck.js, report.js, note.js : formats et accès sûrs au ref.json.
const fmt = (n, d = 0) => (n == null || isNaN(Number(n)) ? "–" : Number(n).toLocaleString("fr-FR", { minimumFractionDigits: d, maximumFractionDigits: d }).replace(/ /g, " "));
const fmt2 = (n) => fmt(n, 2);
const musd = (n) => fmt(Number(n || 0) / 1e6, 2) + " MUSD";
const kusd = (n) => fmt(Number(n || 0) / 1e3, 0) + " kUSD";
const pct = (a, b) => (b ? (Math.round((a / b) * 1000) / 10).toLocaleString("fr-FR") + " %" : "–");
const pctn = (a, b) => (b ? Math.round((a / b) * 1000) / 10 : 0);
const dfr = (s) => { const m = String(s ?? "").match(/(\d{4})-(\d{2})-(\d{2})/); return m ? `${m[3]}/${m[2]}/${m[1]}` : String(s ?? ""); };
const by = (arr, key) => Object.fromEntries((arr || []).map((r) => [r[key], r]));
const payOf = (ref) => { const p = by(ref.nonref_pay, "pay"); const z = { n: 0, amt: 0 }; return { NYS: p["NOT YET STARTED"] || z, PAID: p.PAID || z, PENDING: p.PENDING || z, AUTRE: p.AUTRE || z }; };
const lst = (arr, f, max = 3, sep = ", ") => (arr || []).slice(0, max).map(f).join(sep);
const plural = (n, s, p) => (n > 1 ? p : s);
const LABELS = {
  demdesp: { court: "demurrage / despatch", long: "surestaries (demurrage / despatch)", titre: "Surestaries : du paiement à l'armateur à la refacturation client",
    objet: "Lorsqu'un navire immobilise le quai au-delà du temps contractuel (laytime), l'armateur facture des surestaries (demurrage) ; à l'inverse, un chargement ou déchargement plus rapide ouvre droit à une prime de célérité (despatch). Ces montants sont réglés ou perçus dans la facture de fret finale. Contractuellement ils reviennent à l'acheteur : en CFR / CIF le demurrage lui est refacturé par débit note et le despatch lui est restitué par crédit note ; en FOB le sens s'inverse (le vendeur doit le demurrage à l'acheteur, le despatch lui est facturé).",
    maillon1: "Payé à l'armateur", maillon1_src: "Factures de fret finales" },
  priceadj: { court: "remises et ajustements", long: "remises, ajustements de prix et réclamations diverses", titre: "Remises et ajustements de prix : de l'accord commercial à la note émise et imputée",
    objet: "Les remises, ajustements de prix, LC charges, dead freight et réclamations diverses accordés ou réclamés après la vente donnent lieu à une débit note (créance sur l'acheteur) ou une crédit note (dette envers lui) rattachée à la vente. Un accord suivi par le post-sales mais sans document SAP, ou un document émis mais jamais imputé ni réglé, est une fuite ou un risque.",
    maillon1: "Accord commercial", maillon1_src: "Contrats, avenants, courriels d'accord" },
  all: { court: "réclamations post-sales", long: "réclamations post-sales (demurrage, despatch, remises, ajustements de prix, LC charges…)", titre: "Réclamations post-sales : du fait générateur à la note émise et encaissée",
    objet: "Chaque réclamation post-sales (surestaries, remises, ajustements de prix, LC charges, dead freight…) doit donner lieu à une débit note (créance sur l'acheteur) ou une crédit note (dette envers lui) rattachée à la vente, puis être encaissée ou imputée. Tout ce qui est suivi sans document, ou émis sans règlement, est une fuite ou un risque.",
    maillon1: "Fait générateur (fret, accord)", maillon1_src: "Factures de fret, accords commerciaux" } };
const labelsFor = (ref) => LABELS[(ref.mission && ref.mission.famille) || "demdesp"] || LABELS.all;
module.exports = { LABELS, labelsFor, fmt, fmt2, musd, kusd, pct, pctn, dfr, by, payOf, lst, plural };
