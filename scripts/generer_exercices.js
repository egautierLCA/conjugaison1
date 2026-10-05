// Génère exercices_P3_present_parfait.docx (police sans empattement : Arial)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, HeadingLevel, ShadingType, BorderStyle, ExternalHyperlink,
  PageBreak, LevelFormat,
} = require("docx");

const FONT = "Arial";
const GAFFIOT = "https://www.gaffiot.fr/";

// ---------- Données ----------
// [forme, temps primitifs, sens, temps, traduction]
const ex1 = [
  ["amat", "amo, amas, amare, amavi, amatum", "aimer", "présent", "il (elle) aime"],
  ["monuit", "moneo, mones, monere, monui, monitum", "avertir, conseiller", "parfait", "il (elle) a averti"],
  ["dicit", "dico, dicis, dicere, dixi, dictum", "dire", "présent", "il (elle) dit"],
  ["cepit", "capio, capis, capere, cepi, captum", "prendre", "parfait", "il (elle) a pris"],
  ["audit", "audio, audis, audire, audivi, auditum", "entendre", "présent", "il (elle) entend"],
  ["scripsit", "scribo, scribis, scribere, scripsi, scriptum", "écrire", "parfait", "il (elle) a écrit"],
  ["habet", "habeo, habes, habere, habui, habitum", "avoir", "présent", "il (elle) a"],
  ["fuit", "sum, es, esse, fui, — (part. futur : futurus)", "être", "parfait", "il (elle) a été / fut"],
  ["delevit", "deleo, deles, delere, delevi, deletum", "détruire", "parfait", "il (elle) a détruit"],
  ["vicit", "vinco, vincis, vincere, vici, victum", "vaincre", "parfait", "il (elle) a vaincu"],
  ["legit", "lego, legis, legere, legi, lectum", "lire ; choisir", "présent OU parfait", "il (elle) lit / il (elle) a lu"],
];

const ex2 = [
  ["laudavit", "laudo, laudas, laudare, laudavi, laudatum", "louer, faire l'éloge", "parfait", "il (elle) a loué"],
  ["timet", "timeo, times, timere, timui, — (pas de supin)", "craindre", "présent", "il (elle) craint"],
  ["misit", "mitto, mittis, mittere, misi, missum", "envoyer", "parfait", "il (elle) a envoyé"],
  ["ducit", "duco, ducis, ducere, duxi, ductum", "conduire, mener", "présent", "il (elle) conduit"],
  ["fecit", "facio, facis, facere, feci, factum", "faire", "parfait", "il (elle) a fait"],
  ["pugnat", "pugno, pugnas, pugnare, pugnavi, pugnatum", "combattre", "présent", "il (elle) combat"],
  ["vidit", "video, vides, videre, vidi, visum", "voir", "parfait", "il (elle) a vu"],
  ["tenet", "teneo, tenes, tenere, tenui, tentum", "tenir", "présent", "il (elle) tient"],
  ["rexit", "rego, regis, regere, rexi, rectum", "diriger, gouverner", "parfait", "il (elle) a dirigé"],
  ["dormit", "dormio, dormis, dormire, dormivi, dormitum", "dormir", "présent", "il (elle) dort"],
  ["dedit", "do, das, dare, dedi, datum", "donner", "parfait", "il (elle) a donné"],
  ["posuit", "pono, ponis, ponere, posui, positum", "poser, placer", "parfait", "il (elle) a posé"],
  ["cucurrit", "curro, curris, currere, cucurri, cursum", "courir", "parfait", "il (elle) a couru"],
  ["venit", "venio, venis, venire, veni, ventum", "venir", "présent OU parfait", "il (elle) vient / il (elle) est venu(e)"],
];

// ---------- Helpers ----------
const run = (text, opts = {}) => new TextRun({ text, font: FONT, ...opts });
const p = (children, opts = {}) =>
  new Paragraph({ children: Array.isArray(children) ? children : [run(children)], ...opts });
const border = { style: BorderStyle.SINGLE, size: 4, color: "808080" };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(text, width, { header = false, bold = false, italic = false, center = false } = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders,
    shading: header ? { type: ShadingType.CLEAR, color: "auto", fill: "D9E2F3" } : undefined,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: [p([run(text, { bold: header || bold, italics: italic, size: header ? 20 : 21 })],
      { alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT })],
  });
}

function table(widths, headers, rows) {
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], { header: true, center: true })) }),
      ...rows.map((r) => new TableRow({
        cantSplit: true,
        children: r.map((c, i) => cell(c.text, widths[i], c)),
      })),
    ],
  });
}

const blank = { text: " " };
const spacer = () => p("", { spacing: { after: 120 } });
const bullet = (children) => new Paragraph({ numbering: { reference: "puces", level: 0 }, children, spacing: { after: 60 } });

const W1 = [500, 1200, 3500, 1700, 1300, 1438]; // = 9638
const W2 = [500, 1200, 3300, 1700, 1300, 1638]; // = 9638

// ---------- Contenu ----------
const rappel = [
  p([run("Rappel de méthode", { bold: true, size: 24 })], { spacing: { before: 120, after: 80 } }),
  bullet([run("Présent P3 : ", { bold: true }), run("on remplace le -s du présent P2 par "), run("-t", { bold: true }), run(" (amas → amat, mones → monet, legis → legit, audis → audit).")]),
  bullet([run("Parfait P3 : ", { bold: true }), run("radical du parfait (le parfait P1 sans le -i final : amavi → amav-) + "), run("-it", { bold: true }), run(" (amavit).")]),
  bullet([run("Pour savoir si une forme est au présent ou au parfait, je compare son radical avec celui du présent et celui du parfait.")]),
  bullet([run("Attention : ", { bold: true }), run("pour certains verbes, les deux radicaux sont identiques ; la forme est alors ambiguë et c'est le contexte qui tranche.")]),
];

const enTete = [
  p([run("Latin – Conjugaison", { size: 20, color: "595959" })], { alignment: AlignmentType.RIGHT }),
  p([run("Nom : ................................  Prénom : ................................  Classe : ..........", { size: 20 })], { spacing: { after: 200 } }),
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [run("La 3e personne du singulier (P3) : présent ou parfait de l'indicatif ?", { bold: true, size: 32 })], spacing: { after: 200 } }),
];

const exo1 = [
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [run("Exercice 1 – Avec les temps primitifs", { bold: true, size: 26 })], spacing: { before: 240, after: 120 } }),
  p([run("Pour chaque forme, observe les temps primitifs donnés. Indique si le verbe est conjugué au "), run("présent", { bold: true }), run(" ou au "), run("parfait", { bold: true }), run(" de l'indicatif, puis traduis-le.")], { spacing: { after: 80 } }),
  p([run("Temps primitifs donnés dans l'ordre : présent P1, présent P2, infinitif, parfait P1, supin.", { italics: true, size: 20 })], { spacing: { after: 160 } }),
  table(W1, ["N°", "Forme", "Temps primitifs", "Sens", "Temps", "Traduction"],
    ex1.map((r, i) => [{ text: String(i + 1), center: true }, { text: r[0], bold: true }, { text: r[1] }, { text: r[2], italic: true }, blank, blank])),
  p([run("★ Le n° 11 est un piège : compare bien les deux radicaux…", { italics: true, size: 20 })], { spacing: { before: 80 } }),
];

const exo2 = [
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [run("Exercice 2 – À toi de chercher dans le Gaffiot !", { bold: true, size: 26 })], spacing: { after: 120 } }),
  p([run("Cette fois, les temps primitifs ne sont pas donnés. Cherche chaque verbe dans le dictionnaire Gaffiot en ligne : ")], { spacing: { after: 60 } }),
  p([new ExternalHyperlink({ link: GAFFIOT, children: [run("www.gaffiot.fr", { style: "Hyperlink", bold: true, size: 24 })] })], { alignment: AlignmentType.CENTER, spacing: { after: 120 } }),
  p("Recopie les temps primitifs et le sens, puis indique le temps et traduis.", { spacing: { after: 120 } }),
  p([run("Comment chercher dans le Gaffiot ?", { bold: true })], { spacing: { after: 60 } }),
  bullet([run("Le dictionnaire classe les verbes au "), run("présent P1", { bold: true }), run(" (amo, moneo, lego, audio…). Il faut donc retrouver le présent P1 à partir de la forme conjuguée.")]),
  bullet([run("Essaie d'abord en remplaçant la terminaison -at / -et / -it par -o ou -eo ou -io (pugnat → pugno).")]),
  bullet([run("Si le mot trouvé n'existe pas, la forme est sans doute un parfait au radical différent : utilise le coup de pouce en bas de page.")]),
  spacer(),
  table(W2, ["N°", "Forme", "Temps primitifs (Gaffiot)", "Sens", "Temps", "Traduction"],
    ex2.map((r, i) => [{ text: String(i + 1), center: true }, { text: r[0], bold: true }, blank, blank, blank, blank])),
  p([run("★ Le n° 14 admet deux analyses : lesquelles ?", { italics: true, size: 20 })], { spacing: { before: 80, after: 160 } }),
  p([run("Coup de pouce", { bold: true }), run(" – les présents P1 à chercher pour les formes les plus difficiles (dans le désordre) : "), run("pono – do – curro – facio – rego – mitto – video", { italics: true })], { spacing: { after: 80 } }),
];

const corrige = [
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [run("Corrigé (pour l'enseignant)", { bold: true, size: 26 })], spacing: { after: 160 } }),
  p([run("Exercice 1", { bold: true, size: 24 })], { spacing: { after: 100 } }),
  table(W1, ["N°", "Forme", "Temps primitifs", "Sens", "Temps", "Traduction"],
    ex1.map((r, i) => [{ text: String(i + 1), center: true }, { text: r[0], bold: true }, { text: r[1] }, { text: r[2], italic: true }, { text: r[3], bold: true }, { text: r[4] }])),
  p([run("N° 11 : legit peut être un présent (lĕgit, e bref) ou un parfait (lēgit, e long) ; à l'écrit, sans les quantités, seul le contexte permet de trancher.", { size: 20, italics: true })], { spacing: { before: 80, after: 240 } }),
  p([run("Exercice 2", { bold: true, size: 24 })], { spacing: { after: 100 } }),
  table(W2, ["N°", "Forme", "Temps primitifs (Gaffiot)", "Sens", "Temps", "Traduction"],
    ex2.map((r, i) => [{ text: String(i + 1), center: true }, { text: r[0], bold: true }, { text: r[1] }, { text: r[2], italic: true }, { text: r[3], bold: true }, { text: r[4] }])),
  p([run("N° 14 : venit peut être un présent (vĕnit, e bref) ou un parfait (vēnit, e long).", { size: 20, italics: true })], { spacing: { before: 80 } }),
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: "Title", name: "Title", basedOn: "Normal", run: { font: FONT, size: 32, bold: true, color: "1F3864" } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 26, bold: true, color: "1F3864" }, paragraph: { outlineLevel: 0 } },
    ],
  },
  numbering: { config: [{ reference: "puces", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 260 } } } }] }] },
  sections: [{
    properties: { page: { margin: { top: 1000, bottom: 1000, left: 1134, right: 1134 } } },
    children: [...enTete, ...rappel, ...exo1, ...exo2, ...corrige],
  }],
});

const out = path.join(__dirname, "..", "exercices_P3_present_parfait.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("Écrit :", out); });
