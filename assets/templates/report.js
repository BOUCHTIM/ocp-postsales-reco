// Rapport détaillé (docx) — 100 % piloté par ref.json. Usage : node report.js <ref.json> <sortie.docx>
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, HeadingLevel, ShadingType, BorderStyle, LevelFormat, PageBreak, Footer, Header, PageNumber, VerticalAlign } = require("docx");
const { fmt, fmt2, musd, pct, dfr, by, payOf, lst, plural, labelsFor } = require(__dirname + "/common.js");
const ref = JSON.parse(fs.readFileSync(process.argv[2], "utf8")); const OUT = process.argv[3];
const NAVY = "0F4C81", INK = "0B1F3A", GREY = "4A5568", PALE = "EEF3F8", FONT = "Arial";
const mi = ref.mission, tot = ref.total, an = by(ref.annee, "Année BL"), st = by(ref.statut, "Statut rapprochement"), ty = by(ref.type, "Type réclamation"), pay = payOf(ref), q = ref.quality;
const abn_n = pay.PAID.n + pay.PENDING.n, abn_amt = pay.PAID.amt + pay.PENDING.amt;
const years = Object.keys(an).sort();
const TITLE = mi.name; const LB = labelsFor(ref);
const EXCL = Object.keys(ref.detail.perim).find((k) => k.startsWith("Avant")) || "";

const P = (text, o = {}) => new Paragraph({ spacing: { after: o.after ?? 120, before: o.before ?? 0 }, alignment: o.align, children: Array.isArray(text) ? text : [new TextRun({ text, size: o.size ?? 21, bold: o.bold, italics: o.italics, color: o.color ?? INK, font: FONT })] });
const R = (text, o = {}) => new TextRun({ text, size: o.size ?? 21, bold: o.bold, italics: o.italics, color: o.color ?? INK, font: FONT });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 160 }, children: [new TextRun({ text: t, font: FONT, color: NAVY, bold: true, size: 30 })] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 }, children: [new TextRun({ text: t, font: FONT, color: NAVY, bold: true, size: 24 })] });
const B = (text) => new Paragraph({ numbering: { reference: "bul", level: 0 }, spacing: { after: 80 }, children: Array.isArray(text) ? text : [R(text)] });
let numSeq = 0; const numRefs = [];
const NL = () => { numSeq += 1; numRefs.push(`n${numSeq}`); return `n${numSeq}`; };
const N = (text, refName) => new Paragraph({ numbering: { reference: refName, level: 0 }, spacing: { after: 80 }, children: Array.isArray(text) ? text : [R(text)] });
const border = { style: BorderStyle.SINGLE, size: 4, color: "C9D3DE" }; const borders = { top: border, bottom: border, left: border, right: border }; const TW = 9638;
const L = AlignmentType.LEFT, Rt = AlignmentType.RIGHT;
function table(headers, rows, widths, opts = {}) {
  const cw = widths.map((w) => Math.round((w / widths.reduce((a, b) => a + b, 0)) * TW)); const aligns = opts.aligns || headers.map(() => L);
  const mk = (t, i, hdr, fill, bold) => new TableCell({ width: { size: cw[i], type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER, shading: { type: ShadingType.CLEAR, fill, color: "auto" }, margins: { top: 60, bottom: 60, left: 90, right: 90 },
    children: [new Paragraph({ alignment: aligns[i], spacing: { after: 0 }, children: [new TextRun({ text: String(t ?? ""), font: FONT, size: opts.size ?? 18, bold: hdr || bold, color: hdr ? "FFFFFF" : INK })] })] });
  const trs = [new TableRow({ tableHeader: true, children: headers.map((h, i) => mk(h, i, true, NAVY, false)) })];
  rows.forEach((r, ri) => { const last = opts.boldLast && ri === rows.length - 1; trs.push(new TableRow({ children: r.map((v, i) => mk(v, i, false, last ? "DCE6F0" : ri % 2 ? PALE : "FFFFFF", last)) })); });
  return new Table({ width: { size: TW, type: WidthType.DXA }, columnWidths: cw, rows: trs });
}
const kpiHead = ["", "Clés", "Non ref.", "Tracker (USD)", "Refacturé SAP (USD)", "% ref.", "Encaissé (USD)", "% enc.", "Solde client (USD)", "Non refacturé (USD)"];
const kpiW = [1.7, 0.7, 0.8, 1.25, 1.35, 1.0, 1.25, 0.95, 1.25, 1.35]; const kpiA = [L, Rt, Rt, Rt, Rt, Rt, Rt, Rt, Rt, Rt];
const kpiRow = (label, r) => [label, r.n, r.nonref, fmt(r.trk), fmt(r.sap), pct(r.sap, r.trk), fmt(r.enc), pct(r.enc, r.sap), fmt(r.solde), fmt(r.nonref_amt)];
const totRow = ["Total", tot.n, tot.nonref, fmt(tot.trk), fmt(tot.sap), pct(tot.sap, tot.trk), fmt(tot.enc), pct(tot.enc, tot.sap), fmt(tot.solde), fmt(tot.nonref_amt)];
const regionRows = (ref.region || []).filter((r) => r.n >= 3).sort((a, b) => b.trk - a.trk);
const regionOther = (ref.region || []).filter((r) => r.n < 3).reduce((acc, r) => ({ n: acc.n + r.n, nonref: acc.nonref + r.nonref, trk: acc.trk + r.trk, sap: acc.sap + r.sap, enc: acc.enc + r.enc, solde: acc.solde + r.solde, nonref_amt: acc.nonref_amt + r.nonref_amt }), { n: 0, nonref: 0, trk: 0, sap: 0, enc: 0, solde: 0, nonref_amt: 0 });

const c = [];
c.push(new Paragraph({ spacing: { before: 2400, after: 200 }, children: [new TextRun({ text: "Rapport détaillé", font: FONT, size: 28, color: GREY })] }));
c.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: TITLE, font: FONT, size: 44, bold: true, color: NAVY })] }));
c.push(new Paragraph({ spacing: { after: 600 }, children: [new TextRun({ text: `Suivi post-sales (${LB.court}) ↔ débit et crédit notes SAP — connaissements ${mi.periode}`, font: FONT, size: 26, color: INK })] }));
c.push(P("Clé de rapprochement : nom du navire + date BL", { size: 22, color: GREY }));
c.push(P(`Date du rapport : ${mi.date_rapport} · Date d'arrêté de l'export SAP : ${mi.arrete}`, { size: 22, color: GREY }));
c.push(P(`Sources : ${mi.fichier_sap} ; ${mi.fichiers_tracker.join(" ; ")}${mi.fichier_fret ? " ; " + mi.fichier_fret : ""}`, { size: 22, color: GREY }));
c.push(new Paragraph({ children: [new PageBreak()] }));

c.push(H1("Résumé"));
c.push(P(`Sur ${tot.n} clés navire + date BL suivies par le post-sales depuis le ${dfr(mi.start)} (${musd(tot.trk)} de ${LB.court}), ${tot.matched} clés disposent d'une débit ou crédit note dans l'export SAP (${musd(tot.sap)}, soit ${pct(tot.sap, tot.trk)} des montants), et ${musd(tot.enc)} sont effectivement encaissés ou imputés (${pct(tot.enc, tot.sap)} des montants refacturés).`));
c.push(P(`${tot.nonref} clés (${musd(tot.nonref_amt)}) n'ont aucun document ZDS / ZDP dans l'export. ${pay.NYS.n} d'entre elles sont « non démarrées » dans le tracker, ce qui est attendu.${abn_n ? ` En revanche ${abn_n} clés (${musd(abn_amt)}) sont déclarées « payées » ou « en cours » sans document SAP correspondant : ${ref.dncn.trouves} seulement des ${ref.dncn.listes} numéros de DN / CN cités par le tracker existent dans l'export, qui apparaît filtré. Ce point doit être levé avant toute conclusion sur une non-refacturation.` : ""}`));
c.push(P(`Les ${tot.matched} clés rapprochées sont cohérentes : ${ref.ecarts.length} écart${plural(ref.ecarts.length, "", "s")} de montant${ref.ecarts.length ? ` (${fmt2(ref.ecarts.reduce((a, e) => a + Math.abs(Number(e["Écart tracker – SAP"])), 0))} USD en valeur absolue cumulée${q.typos.length ? `, dont ${q.typos.length} probablement dû${plural(q.typos.length, "", "s")} à une faute de frappe dans le tracker` : ""})` : ""}. ${tot.nonenc_n} clés refacturées conservent un solde client ouvert (${musd(tot.nonenc_solde)}).`));
c.push(P(mi.fichier_fret ? "Les factures de fret armateur ont été fournies : le maillon « payé à l'armateur » est rapproché dans le classeur d'analyse experte (colonne ①)." : "Limite majeure : aucun des fichiers ne contient les paiements à l'armateur. Le rapprochement mesure donc « suivi → refacturé → encaissé » ; l'indicateur cible « payé à l'armateur → récupéré auprès du client » nécessite l'extraction des factures de fret finales, à joindre sur la même clé navire + date BL."));

c.push(H1("1. Objet et contexte"));
c.push(P(LB.objet));
c.push(P(`La demande : à partir du ${dfr(mi.start)}, vérifier que les ${LB.court} sont refacturés sous forme de débit note ou de crédit note annexée à la vente, avec pour clé de connexion le nom du navire et la date de connaissement (BL). Le présent rapport décrit les données reçues, la méthode, les résultats, les anomalies, les limites et les actions proposées ; il accompagne le classeur Excel qui contient l'intégralité des rapprochements ligne à ligne.`));

c.push(H1("2. Données reçues"));
const rowsData = [[mi.fichier_sap, "Extraction SAP des commandes ZDS (demurrage → débit note) et ZDP (despatch → crédit note).", `${ref.sap.total_export ? fmt(ref.sap.total_export) + " documents au total ; " : ""}${fmt(ref.sap.n)} avec BL ≥ ${dfr(mi.start)} (${fmt(ref.sap.types.ZDS || 0)} ZDS, ${fmt(ref.sap.types.ZDP || 0)} ZDP)`, "Côté « refacturation » : montant facturé, solde client, dates de création et de règlement."]];
mi.fichiers_tracker.forEach((f, i) => rowsData.push([f, i === 0 ? `Suivi post-sales (onglets ${LB.court}). Fichier maître.` : (mi.tracker_secondaire_sous_ensemble ? "Suivi post-sales, périmètre partiel : sous-ensemble strict du fichier maître, fusionné sans doublon." : "Suivi post-sales, fichier secondaire fusionné sans doublon."), i === 0 ? `${ref.detail.total} lignes au total, dont ${ref.detail.perim["2024+"] || 0} dans le périmètre` : `${ref.detail.apac.Oui || 0} lignes communes avec le maître`, "Côté « suivi » : montant final, n° de DN / CN, statut de paiement, remarques."]));
if (mi.fichier_fret) rowsData.push([mi.fichier_fret, "Factures de fret armateur.", "", "Côté « payé à l'armateur » (maillon ①)."]);
c.push(table(["Fichier", "Nature", "Volume", "Utilisation"], rowsData, [2.6, 3.4, 2.2, 2.6]));
c.push(P(""));
if (mi.famille === "demdesp") c.push(P("Les onglets de remises / ajustements de prix éventuellement présents dans les trackers ne sont pas intégrés dans cette famille : ils se traitent avec la famille « priceadj ».", { italics: true, color: GREY }));

c.push(H1("3. Méthodologie"));
c.push(H2("3.1 Normalisation"));
c.push(B("Nom du navire : majuscules, suppression de la ponctuation, des mentions « M/V » et des espaces multiples ou finaux."));
c.push(B(`Date BL : ${q.bl_ranges} ligne${plural(q.bl_ranges, "", "s")} du tracker porte${plural(q.bl_ranges, "", "nt")} une plage « jj/mm/aaaa - jj/mm/aaaa » ; la plage est conservée comme intervalle, ordre inversé toléré.`));
c.push(B(`Montants : ${q.amounts_text} montant${plural(q.amounts_text, "", "s")} saisi${plural(q.amounts_text, "", "s")} en texte converti${plural(q.amounts_text, "", "s")} en nombre. Statuts de paiement en texte libre normalisés en PAID / PENDING / NOT YET STARTED. Caractères de contrôle retirés.`));
c.push(H2("3.2 Clé de rapprochement"));
c.push(P((mi.famille === "demdesp" ? "Clé = navire normalisé + date BL (ou plage) + type attendu : DEMURRAGES ↔ ZDS, DESPATCH ↔ ZDP, « DESPATCH / DEMURRAGES » ↔ les deux." : "Clé = navire normalisé + date BL (ou plage) + type de réclamation ; tout document SAP du voyage est candidat, les numéros de DN / CN cités tranchent.") + " Le rapprochement est réalisé au niveau de la clé et non de la ligne : plusieurs lignes du tracker partageant la même clé sont regroupées afin qu'un document SAP ne soit jamais compté deux fois. Le détail ligne à ligne reste dans l'onglet « Détail tracker »."));
c.push(H2("3.3 Niveaux de rapprochement"));
{ const n = NL();
  c.push(N("Niveau 1 – navire + date BL : tout document SAP du type attendu dont la date BL est égale à la date du tracker ou comprise dans sa plage.", n));
  c.push(N(`Niveau 2 – numéros de DN / CN : si le tracker cite des numéros de documents, ceux retrouvés dans SAP sur le même navire sont retenus en priorité, y compris sur une autre date BL${q.multi_voyage.length ? ` (cas ${lst(q.multi_voyage, (v) => v)} : une ligne tracker couvre plusieurs voyages)` : ""}.`, n));
  c.push(N("Niveau 3 – tolérance ± 3 jours : utilisé seulement en l'absence des niveaux 1 et 2 ; signalé « probable » pour validation manuelle.", n)); }
c.push(H2("3.4 Montants, statuts et contrôles"));
c.push(B("Montant tracker = FINAL AMOUNT (positif). Montant SAP = Montant facturé (ZDS positif, ZDP négatif) ; « SAP montant refacturé (abs) » = |ZDS| + |ZDP|."));
c.push(B("Encaissé / imputé = |Montant facturé| – |Solde client|, en valeurs absolues afin de ne pas compenser des débits ZDS par des crédits ZDP."));
c.push(B("Statut : « Refacturé – soldé », « Refacturé – partiellement encaissé », « Refacturé – non encaissé », « Non refacturé » (aucun document ZDS / ZDP dans l'export pour la clé)."));
c.push(B("Contrôle d'écart = montant tracker – montant SAP (abs) ; « OK » si |écart| ≤ max(1 USD ; 0,5 % du montant). Alerte lorsque SAP contient, sur le même voyage, des documents non cités par le tracker."));
c.push(H2("3.5 Périmètre"));
c.push(P(`Lignes du tracker avec date BL ≥ ${dfr(mi.start)} : ${ref.detail.perim["2024+"] || 0} lignes, ${tot.n} clés. Les ${ref.detail.perim[EXCL] || 0} lignes antérieures (${musd(ref.detail.amt_perim[EXCL] || 0)}) restent visibles dans « Détail tracker » avec le périmètre « exclu ». Côté SAP, ${fmt(ref.sap.n)} documents avec BL dans le périmètre forment la base de rapprochement.`));
c.push(H2("3.6 Outils et assurance qualité"));
c.push(P("Traitement réalisé par scripts (pandas / openpyxl) à partir des fichiers d'origine, sans modification de ceux-ci. L'onglet « Synthèse » du classeur est entièrement en formules sur l'onglet « Rapprochement ». Les chiffres du présent rapport sont extraits du classeur final, et non ressaisis ; le classeur d'analyse experte exécute en outre une batterie de contrôles automatiques à chaque rejeu."));

c.push(H1("4. Résultats"));
c.push(H2("4.1 Vue d'ensemble"));
c.push(table(kpiHead, [kpiRow(`Périmètre ${mi.periode}`, tot)], kpiW, { aligns: kpiA, size: 16 }));
c.push(P(""));
c.push(P(`Lecture : ${pct(tot.sap, tot.trk)} des montants suivis ont un document SAP ; parmi eux, ${pct(tot.enc, tot.sap)} sont encaissés ou imputés et ${musd(tot.solde)} restent ouverts chez les clients. ${musd(tot.nonref_amt)} n'ont aucun document SAP (section 5).`));
c.push(H2("4.2 Par année de BL"));
c.push(table(kpiHead, [...years.map((y) => kpiRow(y, an[y])), totRow], kpiW, { aligns: kpiA, boldLast: true, size: 16 }));
c.push(P(""));
years.forEach((y) => { const r = an[y]; c.push(B(`${y} : ${pct(r.sap, r.trk)} refacturé, ${pct(r.enc, r.sap)} encaissé ; ${r.nonref} clé${plural(r.nonref, "", "s")} sur ${r.n} sans document SAP ; solde client ${fmt(r.solde)} USD.`)); });
c.push(H2("4.3 Par type de réclamation"));
c.push(table(kpiHead, [...Object.keys(ty).sort((x, y) => ty[y].trk - ty[x].trk).map((k) => kpiRow(k, ty[k])), totRow], kpiW, { aligns: kpiA, boldLast: true, size: 16 }));
c.push(P(""));
{ const dem = Object.keys(ty).find((k) => /DEMURR/i.test(k)), des = Object.keys(ty).find((k) => /DESPATCH|DISPATCH/i.test(k));
  if (dem && des) c.push(P(`Le demurrage (débit notes) pèse ${musd(ty[dem].trk)} et est refacturé à ${pct(ty[dem].sap, ty[dem].trk)} ; le despatch (crédit notes), ${musd(ty[des].trk)}, l'est à ${pct(ty[des].sap, ty[des].trk)}.`));
  else { const top = Object.keys(ty).sort((x, y) => ty[y].trk - ty[x].trk)[0]; if (top) c.push(P(`Le type « ${top} » pèse ${musd(ty[top].trk)} (${pct(ty[top].trk, tot.trk)} du total) et est refacturé à ${pct(ty[top].sap, ty[top].trk)}.`)); } }
c.push(H2("4.4 Par statut de rapprochement"));
c.push(table(["Statut", "Clés", "Tracker (USD)", "Refacturé SAP (USD)", "Encaissé (USD)", "Solde client (USD)"],
  [...Object.keys(st).map((k) => [k, st[k].n, fmt(st[k].trk), fmt(st[k].sap), fmt(st[k].enc), fmt(st[k].solde)]), ["Total", tot.n, fmt(tot.trk), fmt(tot.sap), fmt(tot.enc), fmt(tot.solde)]],
  [3.2, 0.8, 1.4, 1.6, 1.4, 1.5], { aligns: [L, Rt, Rt, Rt, Rt, Rt], boldLast: true }));
c.push(P(""));
if (regionRows.length) {
  c.push(H2("4.5 Par région (tracker)"));
  c.push(table(["Région", "Clés", "Non refact.", "Tracker (USD)", "Refacturé SAP (USD)", "% refact.", "Encaissé (USD)", "Solde client (USD)"],
    [...regionRows.map((r) => [r["Région (tracker)"], r.n, r.nonref, fmt(r.trk), fmt(r.sap), pct(r.sap, r.trk), fmt(r.enc), fmt(r.solde)]),
     ...(regionOther.n ? [["Autres régions (< 3 clés)", regionOther.n, regionOther.nonref, fmt(regionOther.trk), fmt(regionOther.sap), pct(regionOther.sap, regionOther.trk), fmt(regionOther.enc), fmt(regionOther.solde)]] : []),
     ["Total", tot.n, tot.nonref, fmt(tot.trk), fmt(tot.sap), pct(tot.sap, tot.trk), fmt(tot.enc), fmt(tot.solde)]],
    [2.6, 0.75, 0.9, 1.3, 1.5, 0.9, 1.3, 1.4], { aligns: [L, Rt, Rt, Rt, Rt, Rt, Rt, Rt], boldLast: true, size: 17 }));
  c.push(P(""));
  const empty = regionRows.filter((r) => r.sap === 0).map((r) => r["Région (tracker)"]);
  if (empty.length) c.push(P(`Régions sans aucun document SAP rapproché : ${empty.join(", ")}. Ce profil rejoint le constat d'un export filtré (section 5.2).`, { italics: true, color: GREY }));
}

c.push(H1("5. Analyse des clés sans document SAP"));
c.push(H2("5.1 Répartition selon le statut de paiement du tracker"));
const nrRows = ref.nonref_pay_annee.map((r) => [String(r["Année BL"]), r.pay, r.n, fmt(r.amt)]); nrRows.push(["Total", "", tot.nonref, fmt(tot.nonref_amt)]);
c.push(table(["Année BL", "Statut tracker", "Clés", "Montant tracker (USD)"], nrRows, [1.5, 3, 1, 2], { aligns: [L, L, Rt, Rt], boldLast: true }));
c.push(P(""));
c.push(P(`« NOT YET STARTED » (${pay.NYS.n} clés, ${musd(pay.NYS.amt)}) : le post-sales n'a pas encore émis la débit / crédit note. L'absence de document SAP est cohérente ; ces dossiers relèvent d'un suivi de délai de traitement.`));
if (abn_n) c.push(P(`« PAID » et « PENDING » (${abn_n} clés, ${musd(abn_amt)}) : le tracker indique qu'une DN / CN existe et, pour ${pay.PAID.n} d'entre elles, qu'elle est réglée, alors que l'export ne contient aucun document ZDS / ZDP pour la clé. La liste complète figure en annexe A.`));
c.push(H2("5.2 Interprétation : un export SAP incomplet ?"));
c.push(P(`Sur ${ref.dncn.listes} numéros de DN / CN cités par le tracker dans le périmètre, ${ref.dncn.trouves} (${pct(ref.dncn.trouves, ref.dncn.listes)}) existent dans l'export${ref.dncn.exemples_absents.length ? ` ; exemples de numéros absents : ${lst(ref.dncn.exemples_absents, (e) => `${e.Navire} (${String(e["DN/CN n° (tracker)"]).split(/\s*[-;\n]\s*/)[0]})`)}` : ""}. ${ref.vessels.absent_sap} des ${ref.vessels.tracker} navires du tracker n'apparaissent dans aucun document SAP (${pct(ref.vessels.absent_sap, ref.vessels.tracker)}).${ref.paid_pending_clients.length ? ` Les clients les plus concernés par les dossiers réclamés sans document : ${lst(ref.paid_pending_clients, (x) => x["Client (tracker)"], 4)}.` : ""}`));
c.push(P("Un tel profil correspond le plus souvent à une extraction filtrée (entité juridique, organisation commerciale, famille de produit) ou à des flux réglés hors débit note (ordres de paiement en FOB). Tant que l'export complet n'est pas fourni, les clés « PAID / PENDING sans document » doivent être lues comme « hors périmètre de l'extraction » et non comme « non refacturées »."));

c.push(H1("6. Anomalies sur les clés rapprochées"));
c.push(H2(`6.1 Écarts de montant (${ref.ecarts.length} clé${plural(ref.ecarts.length, "", "s")})`));
if (ref.ecarts.length) c.push(table(["Navire", "Date BL", "Type", "Client", "Tracker (USD)", "SAP (USD)", "Écart (USD)", "Lecture"],
  ref.ecarts.map((e) => [e.Navire, dfr(e["Date BL (début)"]), e["Type SAP attendu"], e["Client (tracker)"], fmt2(e["Montant tracker (FINAL AMOUNT)"]), fmt2(e["SAP montant refacturé (abs)"]), fmt2(e["Écart tracker – SAP"]), e.Lecture]),
  [1.5, 1.35, 0.75, 1.6, 1.1, 1.1, 1.1, 2.6], { aligns: [L, L, L, L, Rt, Rt, Rt, L], size: 16 }));
else c.push(P("Aucun écart de montant au-delà de la tolérance."));
c.push(P(""));
c.push(H2(`6.2 Documents SAP en surplus sur un même voyage (${ref.alertes.length} clé${plural(ref.alertes.length, "", "s")})`));
if (ref.alertes.length) {
  c.push(table(["Navire", "Date BL", "Tracker (USD)", "SAP retenu (USD)", "Docs non cités", "Montant non cité (USD)", "N° des documents non cités"],
    ref.alertes.map((a) => [a.Navire, dfr(a["Date BL (début)"]), fmt2(a["Montant tracker (FINAL AMOUNT)"]), fmt2(a["SAP montant refacturé (abs)"]), a["Autres docs SAP même voyage (nb)"], fmt2(a["Autres docs SAP même voyage (montant)"]), a["Autres docs SAP même voyage (n°)"]]),
    [1.6, 1.15, 1.2, 1.2, 0.9, 1.3, 3.0], { aligns: [L, L, Rt, Rt, Rt, Rt, L], size: 16 }));
  c.push(P(""));
  c.push(P("SAP contient sur ces voyages des documents non référencés par le tracker, créés à des dates différentes : vraisemblablement des lots annulés puis réémis. Si le premier lot garde un solde ouvert, l'encours est gonflé. À valider avec la comptabilité clients."));
} else c.push(P("Aucun document SAP en surplus détecté."));
c.push(H2("6.3 Erreurs et fragilités de saisie dans le tracker"));
q.typos.forEach((t) => c.push(B(`${t.Navire} ${dfr(t["Date BL"])} : numéro de DN / CN probablement erroné (${t["cité"]} cité, ${t.probable} présent dans SAP pour ${fmt2(t.montant_doc)} USD).`)));
q.multi_voyage.forEach((v) => c.push(B(`${v} : une seule ligne tracker couvre plusieurs voyages (documents SAP sur des dates BL différentes) ; rattachée via les numéros de DN / CN cités.`)));
c.push(B(`${q.bl_ranges} date${plural(q.bl_ranges, "", "s")} BL en plage ; ${q.amounts_text} montant${plural(q.amounts_text, "", "s")} en texte ; ${q.trailing_spaces} nom${plural(q.trailing_spaces, "", "s")} de navire avec espaces parasites ; ${q.status_free_text} statut${plural(q.status_free_text, "", "s")} de paiement en texte libre${q.zero_amounts.length ? ` ; ${q.zero_amounts.length} ligne${plural(q.zero_amounts.length, "", "s")} à montant nul (${q.zero_amounts.join(", ")})` : ""}.`));

c.push(H1("7. Clés refacturées non ou partiellement encaissées"));
c.push(P(`${tot.nonenc_n} clé${plural(tot.nonenc_n, "", "s")} dispose${plural(tot.nonenc_n, "", "nt")} de documents SAP mais conserve${plural(tot.nonenc_n, "", "nt")} un solde client : ${musd(tot.nonenc_solde)} au total.${ref.nonenc_top.length ? ` Dossiers les plus importants : ${lst(ref.nonenc_top, (x) => `${x.Navire} (${fmt(x["SAP solde client (abs)"])} USD)`, 3)}.` : ""}${ref.nonenc_clients.length ? ` Par client : ${lst(ref.nonenc_clients, (x) => `${x["Client (tracker)"]} ${fmt(x.solde)} USD`, 3, " ; ")}.` : ""}`));
if (ref.nonenc.length) c.push(table(["Navire", "Date BL", "Type", "Client", "Refacturé SAP (USD)", "Encaissé (USD)", "Solde (USD)", "Statut tracker"],
  ref.nonenc.map((r) => [r.Navire, dfr(r["Date BL (début)"]), r["Type SAP attendu"], r["Client (tracker)"], fmt2(r["SAP montant refacturé (abs)"]), fmt2(r["SAP encaissé / imputé (abs)"]), fmt2(r["SAP solde client (abs)"]), r["Statut paiement (tracker)"]]),
  [1.6, 1.3, 0.75, 2.0, 1.3, 1.1, 1.1, 1.0], { aligns: [L, L, L, L, Rt, Rt, Rt, L], size: 16 }));
c.push(P(""));

c.push(H1("8. Contrôle inverse : documents SAP sans ligne de suivi"));
c.push(P(`${ref.nonsuivi.n} clés navire + BL + type de l'export n'ont aucune ligne dans le tracker post-sales. L'onglet « SAP non suivi » du classeur permet de filtrer par région et par client ; une colonne signale les navires qui existent par ailleurs dans le tracker sur un autre voyage.`));
c.push(table(["Type SAP", "Clés", "Montant SAP (USD, signé)"], [...ref.nonsuivi.par_type.map((r) => [r["Type SAP"], r.size, fmt(r.sum)]), ["Total", ref.nonsuivi.n, fmt(ref.nonsuivi.par_type.reduce((a, r) => a + r.sum, 0))]], [2.5, 1, 2.5], { aligns: [L, Rt, Rt], boldLast: true }));
c.push(P(""));
if (ref.nonsuivi.par_region.length) { c.push(table(["Région SAP", "Clés", "Montant SAP (USD, signé)"], ref.nonsuivi.par_region.map((r) => [r["Région SAP"], r.size, fmt(r.sum)]), [3, 1, 2.5], { aligns: [L, Rt, Rt] })); c.push(P("")); }
c.push(P(`Solde client restant sur ces documents non suivis : ${musd(ref.nonsuivi.solde)}.`, { italics: true, color: GREY }));

c.push(H1("9. Limites de l'exercice"));
{ const n = NL();
  if (!mi.fichier_fret) c.push(N("Absence des paiements à l'armateur : l'export SAP contient les documents émis aux clients, pas les factures de fret. Le premier maillon de l'indicateur demandé n'est dans aucun fichier.", n));
  c.push(N("Export SAP possiblement filtré (section 5.2) : les taux de refacturation sont des planchers.", n));
  c.push(N("Sens du flux : en FOB le demurrage est dû par le vendeur à l'acheteur ; les montants « réclamés » FOB demurrage sont des dettes, pas des créances. Le classeur d'analyse experte isole créances et dettes.", n));
  c.push(N("La clé navire + date BL est robuste mais pas infaillible : deux voyages proches ou une plage de BL large peuvent agréger plusieurs opérations ; les cas détectés sont signalés.", n));
  c.push(N("Les onglets de remises / ajustements de prix sont exclus.", n)); }

c.push(H1("10. Recommandations"));
c.push(table(["#", "Action", "Détail", "Effet attendu", "Horizon"], [
  ["1", "Relancer l'export SAP sans filtre", "Toutes entités / familles de produit, types ZDS et ZDP, documents annulés inclus.", abn_n ? `Reclasser les ${abn_n} clés PAID / PENDING (${musd(abn_amt)}) et les ${ref.dncn.non_trouves} n° cités non retrouvés.` : "Confirmer l'exhaustivité de l'export.", "Court terme"],
  ["2", mi.fichier_fret ? "Valider les factures de fret fournies" : "Obtenir l'extraction des factures de fret armateur", "Navire, date BL, montant demurrage / despatch réglé, date de paiement.", "Indicateur « payé → refacturé → encaissé » sur la même clé.", "Court terme"],
  ["3", "Traiter les clés refacturées non encaissées", `${tot.nonenc_n} clés, ${musd(tot.nonenc_solde)} de solde${ref.nonenc_top.length ? ` ; priorité ${lst(ref.nonenc_top, (x) => x.Navire, 2)}` : ""}.`, "Réduction de l'encours client (créances) ; régularisation des crédit notes (dettes).", "Immédiat"],
  ["4", "Corriger les écarts et valider les doublons", `${ref.ecarts.length} écart${plural(ref.ecarts.length, "", "s")}${q.typos.length ? ` (dont ${q.typos.length} n° erroné${plural(q.typos.length, "", "s")})` : ""}, ${ref.alertes.length} voyage${plural(ref.alertes.length, "", "s")} avec documents SAP en surplus${q.multi_voyage.length ? `, scission de ${lst(q.multi_voyage, (v) => v)}` : ""}.`, "Cohérence tracker / SAP ; encours réel.", "Immédiat"],
  ["5", "Fiabiliser la saisie du tracker", "Une date BL par ligne, un n° de DN / CN par cellule, statuts normalisés, navires sans espaces parasites, IMO.", "Rapprochement automatique rejouable à chaque mise à jour.", "Continu"],
  ["6", "Suivre les dossiers « non démarrés »", `${pay.NYS.n} clés (${musd(pay.NYS.amt)}) ; définir un délai cible entre date BL et émission de la DN / CN.`, "Indicateur de délai de traitement.", "Continu"],
], [0.4, 2.4, 3.2, 2.9, 1.1], { size: 17 }));
c.push(P(""));

c.push(H1("11. Guide du classeur Excel livré"));
c.push(table(["Onglet", "Contenu", "Usage"], [
  ["Synthèse", "Indicateurs en formules : par année, type de réclamation, statut, statut de paiement du tracker, méthode, région ; contrôle des montants ; côté SAP non suivi.", "Vue de pilotage ; se recalcule à l'ouverture."],
  ["Rapprochement", `${tot.n} clés : identité, montant tracker, DN / CN citées, statut tracker, méthode, documents SAP, montants, encaissé, solde, écart, alertes, statut de rapprochement.`, "Analyse dossier par dossier. Couleurs : vert soldé, jaune en attente, rouge non refacturé."],
  ["SAP non suivi", `${ref.nonsuivi.n} clés SAP sans ligne tracker.`, "Contrôle inverse."],
  ["Détail tracker", `${ref.detail.total} lignes brutes, périmètre, clé, statut de rapprochement.`, "Traçabilité ligne à ligne."],
  ["Base SAP 2024+", `${fmt(ref.sap.n)} documents ZDS / ZDP avec indicateur « Rapproché tracker ».`, "Source SAP nettoyée."],
  ["Notes & hypothèses", "Sources, règles, hypothèses, points d'attention calculés.", "À lire avant toute réutilisation."],
], [1.8, 5.2, 2.6]));
c.push(P(""));

if (ref.paid_pending_sansdoc.length) {
  c.push(new Paragraph({ children: [new PageBreak()] }));
  c.push(H1(`Annexe A – ${abn_n} clés « PAID » ou « PENDING » sans document SAP`));
  c.push(P("Classées par date BL. Montant = FINAL AMOUNT du tracker. À confronter à un export SAP complet.", { italics: true, color: GREY }));
  c.push(table(["Navire", "Date BL", "Type", "Client", "Région", "Montant (USD)", "DN / CN citées", "Statut"],
    ref.paid_pending_sansdoc.map((r) => [r.Navire, dfr(r["Date BL (début)"]), String(r["Type réclamation"]).replace("DEMURRAGES", "DEM").replace("DESPATCH", "DESP"), r["Client (tracker)"], r["Région (tracker)"], fmt2(r["Montant tracker (FINAL AMOUNT)"]), r["DN/CN n° (tracker)"] == null ? "" : String(r["DN/CN n° (tracker)"]), r.pay]),
    [1.5, 1.1, 0.7, 2.0, 1.3, 1.1, 2.2, 1.05], { aligns: [L, L, L, L, L, Rt, L, L], size: 14 }));
  c.push(P(""));
  c.push(P(`Total : ${abn_n} clés, ${fmt2(abn_amt)} USD.`, { bold: true }));
}
c.push(H1("Annexe B – Hypothèses et conventions"));
["Le premier fichier tracker est le fichier maître ; les suivants sont fusionnés sans doublon.",
 "Les montants tracker sont supposés dans la devise des documents SAP.",
 "Un document SAP ne peut être rattaché qu'à une seule clé ; les documents rattachés par numéro sont exclus des autres clés.",
 "« Encaissé / imputé » suppose que Solde client = 0 signifie règlement (ZDS) ou imputation (ZDP) ; un document annulé non lettré apparaît donc comme non encaissé.",
 "Le seuil d'écart (1 USD ou 0,5 %) absorbe les arrondis ; tout écart au-delà est listé.",
 `Les chiffres de ce rapport sont extraits du classeur final le ${mi.date_rapport} ; toute mise à jour des fichiers sources nécessite de rejouer le traitement.`].forEach((t) => c.push(B(t)));

const doc = new Document({ creator: "Analyse post-sales", title: `Rapport détaillé – ${TITLE}`,
  styles: { default: { document: { run: { font: FONT, size: 21, color: INK } } }, paragraphStyles: [
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 30, bold: true, color: NAVY }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 24, bold: true, color: NAVY }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } }] },
  numbering: { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
    ...numRefs.map((r) => ({ reference: r, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } } } }] }))] },
  sections: [{ properties: { page: { margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: `${TITLE} – Rapport détaillé`, font: FONT, size: 16, color: GREY })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", font: FONT, size: 16, color: GREY }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GREY })] })] }) },
    children: c }] });
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("written", OUT); });
