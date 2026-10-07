// Note d'expertise (docx, ~6 pages) — pilotée par ref.json + <analyse>_summary.json. Usage : node note.js <ref.json> <summary.json> <sortie.docx>
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, HeadingLevel, ShadingType, BorderStyle, LevelFormat, PageBreak, Footer, Header, PageNumber, VerticalAlign } = require("docx");
const { fmt, fmt2, musd, pct, dfr, by, payOf, lst, plural, labelsFor } = require(__dirname + "/common.js");
const ref = JSON.parse(fs.readFileSync(process.argv[2], "utf8")); const sm = JSON.parse(fs.readFileSync(process.argv[3], "utf8")); const OUT = process.argv[4];
const NAVY = "0F4C81", INK = "0B1F3A", GREY = "4A5568", PALE = "EEF3F8", FONT = "Arial";
const mi = ref.mission, tot = ref.total, q = ref.quality, pay = payOf(ref); const LB = labelsFor(ref);
const T = Object.fromEntries(sm.typologie.map((t) => [t["Typologie de fuite"].slice(0, 3).trim(), t])); const g = (k) => T[k] || { Clés: 0, Réclamé: 0, Refacturé: 0, Encaissé: 0, Exposé: 0, Créance: 0, Dette: 0 };
const SC = Object.fromEntries(sm.scenarios.map((s) => [s["Scénario"], s["Montant (USD)"]]));
const expo = sm.typologie.reduce((a, t) => a + t.Exposé, 0), cre = sm.typologie.reduce((a, t) => a + (t.Créance || 0), 0), det = sm.typologie.reduce((a, t) => a + (t.Dette || 0), 0);
const d0 = sm.delais[0] || {}, d1 = sm.delais[1] || {}, d2 = sm.delais[2] || {};
const agingOld = sm.aging.filter((r) => ["91-180 j", "> 180 j"].includes(r["Tranche retard"])).reduce((a, r) => a + r.Solde, 0), agingTot = sm.aging.reduce((a, r) => a + r.Solde, 0);
const nd = sm.nd_age.map((r) => ({ t: r["Tranche âge"], n: Number(r.Clés), amt: Number(r.Montant) }));
const topClient = (sm.aging_cli || [])[0]; const ko = (sm.controles || []).filter((x) => x.Résultat === "KO");
const P = (text, o = {}) => new Paragraph({ spacing: { after: o.after ?? 120, before: o.before ?? 0 }, alignment: o.align, children: Array.isArray(text) ? text : [new TextRun({ text, size: o.size ?? 21, bold: o.bold, italics: o.italics, color: o.color ?? INK, font: FONT })] });
const Rn = (text, o = {}) => new TextRun({ text, size: o.size ?? 21, bold: o.bold, italics: o.italics, color: o.color ?? INK, font: FONT });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 160 }, children: [new TextRun({ text: t, font: FONT, color: NAVY, bold: true, size: 30 })] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 }, children: [new TextRun({ text: t, font: FONT, color: NAVY, bold: true, size: 24 })] });
const B = (text) => new Paragraph({ numbering: { reference: "bul", level: 0 }, spacing: { after: 80 }, children: Array.isArray(text) ? text : [Rn(text)] });
const N = (text) => new Paragraph({ numbering: { reference: "n8", level: 0 }, spacing: { after: 80 }, children: Array.isArray(text) ? text : [Rn(text)] });
const border = { style: BorderStyle.SINGLE, size: 4, color: "C9D3DE" }; const borders = { top: border, bottom: border, left: border, right: border }; const TW = 9638; const L = AlignmentType.LEFT, Rt = AlignmentType.RIGHT;
function table(headers, rows, widths, opts = {}) {
  const cw = widths.map((w) => Math.round((w / widths.reduce((a, b) => a + b, 0)) * TW)); const aligns = opts.aligns || headers.map(() => L);
  const mk = (t, i, hdr, fill, b) => new TableCell({ width: { size: cw[i], type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER, shading: { type: ShadingType.CLEAR, fill, color: "auto" }, margins: { top: 60, bottom: 60, left: 90, right: 90 },
    children: [new Paragraph({ alignment: aligns[i], spacing: { after: 0 }, children: [new TextRun({ text: String(t ?? ""), font: FONT, size: opts.size ?? 18, bold: hdr || b, color: hdr ? "FFFFFF" : INK })] })] });
  const trs = [new TableRow({ tableHeader: true, children: headers.map((h, i) => mk(h, i, true, NAVY, false)) })];
  rows.forEach((r, ri) => { const last = opts.boldLast && ri === rows.length - 1; trs.push(new TableRow({ children: r.map((v, i) => mk(v, i, false, last ? "DCE6F0" : ri % 2 ? PALE : "FFFFFF", last)) })); });
  return new Table({ width: { size: TW, type: WidthType.DXA }, columnWidths: cw, rows: trs });
}
const c = [];
c.push(new Paragraph({ spacing: { before: 2000, after: 200 }, children: [new TextRun({ text: "Note d'expertise", font: FONT, size: 28, color: GREY })] }));
c.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: `${mi.name} : où se perd l'argent, combien est récupérable, comment le reprendre`, font: FONT, size: 40, bold: true, color: NAVY })] }));
c.push(new Paragraph({ spacing: { after: 500 }, children: [new TextRun({ text: `Analyse de la chaîne « ${LB.maillon1.toLowerCase()} → réclamé → refacturé → encaissé » (${LB.court}) sur la base du tracker post-sales et de l'export SAP`, font: FONT, size: 24, color: INK })] }));
c.push(P(`Date d'arrêté des données : ${dfr(sm.arrete)} (extraction SAP) · Note du ${mi.date_rapport}`, { size: 22, color: GREY }));
c.push(P("Complète le rapport détaillé et le classeur « Analyse experte » ; outil rejouable livré dans le dossier 04.", { size: 22, color: GREY }));
c.push(new Paragraph({ children: [new PageBreak()] }));

c.push(H1("1. Synthèse exécutive"));
c.push(P(`Sur ${tot.n} voyages suivis (${musd(tot.trk)} de ${LB.court} réclamés), ${musd(tot.sap)} sont refacturés dans SAP (${pct(tot.sap, tot.trk)}) et ${musd(tot.enc)} encaissés. Le montant exposé — réclamé sans document SAP, ou refacturé sans encaissement — atteint ${musd(expo)}, dont ${musd(cre)} de créances sur les acheteurs et ${musd(det)} de dettes envers eux. Il se décompose en situations très différentes, qui appellent des réponses différentes :`));
c.push(table(["Situation", "Montant (USD)", "Définition", "Réponse"], sm.scenarios.map((s) => [s["Scénario"], fmt(s["Montant (USD)"]), s["Définition"], s.Action]), [2.6, 1.3, 3.6, 2.9], { aligns: [L, Rt, L, L], size: 16 }));
c.push(P(""));
c.push(P([Rn("Sens du flux. ", { bold: true }), Rn("Un demurrage n'est une créance que lorsque le vendeur supporte le fret (CFR, CIF, DPU) : il se refacture par débit note. En FOB, le retard a lieu au port de chargement et c'est le vendeur qui doit le demurrage à l'acheteur, réglé par ordre de paiement ; le despatch FOB est, lui, une débit note. Le despatch CFR est une crédit note, donc une dette. Les scénarios ci-dessus ne retiennent comme récupérables que les créances ; les dettes sont isolées, parce qu'elles ne se récupèrent pas, elles se règlent.")]));
if (sm.sens && sm.sens.length) c.push(table(["Incoterm | sens", "Clés", "Réclamé (USD)", "Refacturé (USD)", "Exposé (USD)"], sm.sens.map((r) => [`${r.Incoterm} | ${r["Sens du flux"]}`, r.Clés, fmt(r.Réclamé), fmt(r.Refacturé), fmt(r.Exposé)]), [3.6, 0.7, 1.5, 1.5, 1.5], { aligns: [L, Rt, Rt, Rt, Rt], size: 16 }));
c.push(P(""));
c.push(P("Constats structurants :", { bold: true }));
const t1b = g("T1b"), t1 = g("T1"), t2 = g("T2"), t45n = g("T4").Clés + g("T5").Clés;
if (t1b.Clés) c.push(B([Rn("Défaut d'émission avant défaut de recouvrement : ", { bold: true }), Rn(`${t1b.Clés} dossiers (${musd(t1b.Exposé)}, dont ${musd(SC["Risque de prescription"] || 0)} de créances) attendent une débit / crédit note depuis plus de ${mi.delai_alerte} jours. Le délai médian entre la date BL et l'émission de la DN / CN, mesuré sur les dossiers déjà traités, est de ${fmt(d0.Médiane)} jours (P90 : ${fmt(d0.P90)} jours).`)]));
else c.push(B([Rn("Délai d'émission : ", { bold: true }), Rn(`délai médian BL → DN / CN de ${fmt(d0.Médiane)} jours (P90 : ${fmt(d0.P90)}) ; aucun dossier non démarré au-delà de ${mi.delai_alerte} jours.`)]));
c.push(B([Rn("Recouvrement : ", { bold: true }), Rn(`${fmt(d1.Médiane)} jours en médiane entre l'émission et le premier règlement ; ${agingTot ? `${pct(agingOld, agingTot)} de l'encours ouvert (${musd(agingTot)}) a dépassé l'échéance de plus de 90 jours` : "aucun encours ouvert"}${topClient ? `, concentré sur ${topClient["Client Name"]} (${fmt(Number(topClient.Total))} USD)` : ""}.`)]));
if (t2.Clés) c.push(B([Rn("Export SAP à compléter : ", { bold: true }), Rn(`${t2.Clés} dossiers (${musd(t2.Exposé)}) sont réclamés sans document dans l'export ; ${ref.vessels.absent_sap} des ${ref.vessels.tracker} navires du tracker n'y apparaissent pas et ${ref.dncn.trouves} n° de DN / CN sur ${ref.dncn.listes} cités sont retrouvés. Extraction filtrée ou flux réglés hors débit note : à confirmer avant toute conclusion.`)]));
c.push(B([Rn(mi.fichier_fret ? "Maillon armateur fourni : " : "Le premier maillon manque : ", { bold: true }), Rn(mi.fichier_fret ? "les factures de fret sont rapprochées voyage par voyage (colonne ① du classeur d'analyse) ; l'écart payé – réclamé est calculable." : "aucun fichier ne contient ce qui a été payé à l'armateur. L'indicateur demandé (« payé → récupéré ») n'est pas calculable ; la chaîne mesurée commence au tracker. L'outil livré intègre ce maillon dès que les factures de fret sont fournies (modèle de fichier joint).")]));

c.push(H1("2. La chaîne de valeur et ses fuites"));
c.push(table(["Maillon", "Source", "Montant (USD)", "Clés", "État"], [
  [`① ${LB.maillon1}`, LB.maillon1_src, mi.fichier_fret ? "fourni" : "non mesuré", "–", mi.fichier_fret ? "Rapproché" : "Fichier absent"],
  ["② Réclamé au client", "Tracker post-sales", fmt(tot.trk), tot.n, "Base de l'analyse"],
  ["③ Refacturé", "SAP ZDS (débit) / ZDP (crédit)", fmt(tot.sap), tot.matched, `${pct(tot.sap, tot.trk)} du réclamé`],
  ["④ Encaissé / imputé", "SAP solde client", fmt(tot.enc), "", `${pct(tot.enc, tot.sap)} du refacturé`],
  ["    Solde client ouvert", "SAP", fmt(tot.solde), tot.nonenc_n, agingTot ? `${pct(agingOld, agingTot)} > 90 j de retard` : ""],
], [2.4, 3.0, 1.5, 0.8, 2.5], { aligns: [L, L, Rt, Rt, L] }));
c.push(P(""));
c.push(H2("2.1 Typologie de fuite par voyage"));
c.push(P("Chaque clé navire + date BL + type reçoit une typologie unique (onglet « Chaîne par voyage » du classeur d'analyse). Le montant exposé est le montant réclamé lorsqu'aucun document SAP n'existe, et le solde client restant sinon ; il est ensuite ventilé en créance ou dette selon l'Incoterm."));
c.push(table(["Typologie", "Clés", "Réclamé (USD)", "Refacturé (USD)", "Exposé (USD)", "dont créances", "dont dettes"],
  [...sm.typologie.map((t) => [t["Typologie de fuite"], t.Clés, fmt(t.Réclamé), fmt(t.Refacturé), fmt(t.Exposé), fmt(t.Créance), fmt(t.Dette)]), ["Total", tot.n, fmt(tot.trk), fmt(tot.sap), fmt(expo), fmt(cre), fmt(det)]],
  [3.3, 0.7, 1.3, 1.3, 1.3, 1.3, 1.3], { aligns: [L, Rt, Rt, Rt, Rt, Rt, Rt], boldLast: true, size: 16 }));
c.push(P(""));
c.push(H2("2.2 Les dossiers non démarrés"));
c.push(P(`${t1.Clés + t1b.Clés} clés (${musd(t1.Exposé + t1b.Exposé)}) sont au statut « NOT YET STARTED » sans aucun document SAP. Leur ancienneté depuis la date BL :`));
c.push(table(["Ancienneté depuis le BL", "Clés", "Montant (USD)"], nd.map((r) => [r.t, r.n, fmt(r.amt)]), [3, 1, 2], { aligns: [L, Rt, Rt] }));
c.push(P(""));
c.push(P(`Plus un dossier vieillit, plus la refacturation devient difficile : contestation du client, perte des pièces (time sheets, statement of facts), clause contractuelle de délai de réclamation, changement d'interlocuteur. Sans connaître la clause applicable à chaque contrat, le seuil de ${mi.delai_alerte} jours est retenu comme alerte ; il est paramétrable dans l'outil.`));
if (t2.Clés) { c.push(H2("2.3 Les dossiers réclamés sans document SAP")); c.push(P(`${t2.Clés} clés (${musd(t2.Exposé)}) sont marquées « payées » ou « en cours » par le post-sales sans aucun document ZDS / ZDP dans l'export. Ce montant est classé « à confirmer » (créances) ou « réglé hors export » (dettes FOB), et non « perdu », tant que l'export complet n'est pas fourni.`)); }

c.push(H1("3. Délais et encours vieilli"));
c.push(table(["Indicateur (jours)", "Médiane", "Moyenne", "P90", "N"], sm.delais.map((d) => [d.Indicateur, fmt(d.Médiane), fmt(d.Moyenne), fmt(d.P90), d.N]), [4, 1, 1, 1, 0.8], { aligns: [L, Rt, Rt, Rt, Rt] }));
c.push(P(""));
if (sm.delai_an && sm.delai_an.length) { c.push(table(["Année BL", "Docs", "Médiane BL→DN (j)", "P90 BL→DN (j)", "Médiane DN→règlement (j)"], sm.delai_an.map((d) => [String(d["Année BL"]).replace(".0", ""), fmt(d.N), fmt(d.Médiane_BL_DN), fmt(d.P90_BL_DN), fmt(d.Médiane_DN_regl)]), [1.2, 0.8, 1.6, 1.6, 2.0], { aligns: [L, Rt, Rt, Rt, Rt] })); c.push(P("")); }
if (d0.Médiane) c.push(P(`Un délai médian de ${fmt(d0.Médiane)} jours entre le BL et l'émission de la DN / CN signifie que la refacturation intervient ${d0.Médiane > 180 ? "plus de six mois" : d0.Médiane > 90 ? "plus de trois mois" : "quelques semaines"} après le voyage. Ramener ce délai sous 60 jours est le levier le plus puissant : il réduit mécaniquement le stock de dossiers non démarrés et améliore le taux d'encaissement (client encore engagé sur le dossier).`));
c.push(H2("Encours vieilli des DN / CN ouvertes (retard par rapport à l'échéance SAP, à la date d'arrêté)"));
c.push(table(["Tranche de retard", "Documents", "Solde (USD)"], sm.aging.map((r) => [r["Tranche retard"], fmt(r.Docs), fmt(r.Solde)]), [2.5, 1.2, 2], { aligns: [L, Rt, Rt] }));
c.push(P(""));
if (sm.aging_cli && sm.aging_cli.length) { c.push(table(["Client", "91-180 j", "> 180 j", "Total (USD)"], sm.aging_cli.map((r) => [r["Client Name"], fmt(Number(r["91-180 j"])), fmt(Number(r["> 180 j"])), fmt(Number(r.Total))]), [3.5, 1.3, 1.3, 1.5], { aligns: [L, Rt, Rt, Rt] })); c.push(P("")); }

c.push(H1("4. Qualité des données et référentiel"));
c.push(B(`Référentiel navires : ${ref.vessels.absent_sap} des ${ref.vessels.tracker} navires du tracker n'existent dans aucun document SAP de l'export ; ${sm.refnav_multi} navire${plural(sm.refnav_multi, "", "s")} saisi${plural(sm.refnav_multi, "", "s")} sous plusieurs graphies ; ${sm.imo_multi} numéro${plural(sm.imo_multi, "", "s")} IMO sous plusieurs noms dans SAP. Le tracker ne porte pas l'IMO : l'ajouter rendrait la jointure fiable sans normalisation de texte.`));
c.push(B(`Réémissions SAP : ${sm.dups} voyage${plural(sm.dups, "", "s")} avec deux lots de documents à des dates différentes et des montants répétés, dont ${sm.dups_nonsolde} avec un premier lot non soldé (annulation non comptabilisée ?). Côté tracker, ${ref.alertes.length} voyage${plural(ref.alertes.length, "", "s")} avec des documents SAP non cités.`));
c.push(B(`Saisie tracker : ${q.typos.length} n° de DN / CN probablement erroné${plural(q.typos.length, "", "s")}${q.typos.length ? ` (${lst(q.typos, (t) => `${t.Navire} : ${t["cité"]} → ${t.probable}`, 2, " ; ")})` : ""} ; ${q.multi_voyage.length} ligne${plural(q.multi_voyage.length, "", "s")} couvrant plusieurs voyages${q.multi_voyage.length ? ` (${q.multi_voyage.join(", ")})` : ""} ; ${q.bl_ranges} date${plural(q.bl_ranges, "", "s")} BL en plage ; ${q.amounts_text} montant${plural(q.amounts_text, "", "s")} en texte ; ${q.zero_amounts.length} ligne${plural(q.zero_amounts.length, "", "s")} à montant nul ; ${q.status_free_text} statut${plural(q.status_free_text, "", "s")} en texte libre.`));
c.push(B(`Contrôles automatiques : ${sm.controles.length} tests exécutés à chaque rejeu (onglet « Contrôles ») ; ${ko.length} en échec${ko.length ? ` : ${ko.map((x) => x.Contrôle).join(" ; ")}` : ""}.`));

c.push(H1("5. Échantillon d'audit sur pièces"));
c.push(P(`${sm.sample} dossiers sont sélectionnés pour un contrôle manuel (onglet « Échantillon audit ») : soldes ouverts les plus élevés, dossiers « payés » sans document dans l'export, écarts de montant, réémissions, dossiers non démarrés anciens à fort enjeu. Pour chacun, les pièces à réunir sont listées ; les colonnes « Conclusion audit » et « Écart constaté » sont à compléter. Ce contrôle calibre le taux d'erreur du rapprochement automatique avant toute communication de montants récupérables.`));

c.push(H1("6. Plan d'action chiffré"));
c.push(table(["#", "Action", "Responsable proposé", "Échéance", "Montant concerné (USD)"], sm.plan.map((p) => [p["#"], p.Action, p["Responsable proposé"], p["Échéance"], p["Montant concerné (USD)"] == null || p["Montant concerné (USD)"] === "nan" ? "–" : fmt(Number(p["Montant concerné (USD)"]))]), [0.4, 4.6, 2.2, 1.2, 1.6], { aligns: [L, L, L, L, Rt], size: 16 }));
c.push(P(""));
c.push(P("Ordre de priorité : les actions qui conditionnent la fiabilité (export complet, factures de fret) d'abord ; puis le gain rapide (relance d'un encours échu) ; puis la protection du montant le plus important (dossiers anciens non émis) ; enfin le pilotage durable."));

c.push(H1("7. Outil rejouable livré"));
c.push(P("Dossier « 04 - Outil rejouable » : scripts et lanceur. Déposer les nouveaux exports dans « 02 - Sources recues », relancer la commande indiquée dans run.txt : rapprochement, analyse experte, livrables et contrôles sont régénérés. Le fichier « modele_factures_fret.xlsx » décrit le format attendu pour le maillon ① ; une fois rempli et placé dans les sources, la colonne « Payé armateur » et l'écart ① – ② se calculent sans modification du code. Paramètres : date de début de périmètre, délai d'alerte, tolérance d'écart, alias de colonnes."));
c.push(H1("8. Ce qu'il manque pour une expertise complète"));
if (!mi.fichier_fret) c.push(N("Factures de fret armateur (maillon ①) : navire, date BL, laytime, temps réel, taux, montant payé, date de paiement."));
c.push(N("Export SAP complet : toutes entités et familles de produit, documents annulés inclus, lettrage des règlements."));
c.push(N("Contrats de vente : Incoterm, clause demurrage / despatch, taux, délai de réclamation — pour juger si la refacturation est due et dans quel délai."));
c.push(N("Time sheets / statements of facts par escale : pour recalculer le demurrage contractuel et le comparer au payé et au refacturé."));
c.push(N("Validation de l'échantillon d'audit par le post-sales et la comptabilité clients."));
c.push(P(""));
c.push(P("Avec ces éléments, l'outil produit l'indicateur cible « payé à l'armateur → récupéré auprès du client » par voyage, par client et par mois, et la liste des fuites par cause racine.", { italics: true, color: GREY }));

const doc = new Document({ creator: "Analyse post-sales", title: `Note d'expertise – ${mi.name}`,
  styles: { default: { document: { run: { font: FONT, size: 21, color: INK } } }, paragraphStyles: [
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 30, bold: true, color: NAVY }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 24, bold: true, color: NAVY }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } }] },
  numbering: { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
    { reference: "n8", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } } } }] }] },
  sections: [{ properties: { page: { margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: `Note d'expertise – ${mi.name}`, font: FONT, size: 16, color: GREY })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", font: FONT, size: 16, color: GREY }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GREY })] })] }) },
    children: c }] });
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("written", OUT); });
