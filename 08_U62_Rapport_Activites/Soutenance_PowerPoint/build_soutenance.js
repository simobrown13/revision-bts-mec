const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "BAHAFID Mohamed";
pres.title = "Rapport d\u2019Activit\u00e9s Professionnelles \u2013 BTS MEC U62 \u2013 Session 2026";

// === PALETTE ===
const TEAL = "007A7F";
const TEAL_LIGHT = "E0F4F5";
const BG_CREAM = "F7F6F2";
const DARK = "1E1E1E";
const BODY = "333333";
const MUTED = "666666";
const WHITE = "FFFFFF";
const ACCENT_ORANGE = "D97706";
const ACCENT_GREEN = "16A34A";
const FONT_TITLE = "Cambria";
const FONT_BODY = "Calibri";

// === HELPERS ===
const mkShadow = () => ({ type: "outer", blur: 4, offset: 1.5, angle: 135, color: "000000", opacity: 0.10 });

function addTopBar(slide) {
  slide.background = { color: BG_CREAM };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: TEAL } });
}

function addBadge(slide, text) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.18, w: 0.85, h: 0.38, fill: { color: TEAL_LIGHT }, line: { color: TEAL, width: 0.8 } });
  slide.addText(text, { x: 0.4, y: 0.18, w: 0.85, h: 0.38, fontSize: 9, bold: true, color: TEAL, fontFace: FONT_BODY, align: "center", valign: "middle", margin: 0 });
}

function addFooter(slide, num) {
  slide.addText(`BAHAFID Mohamed  \u00b7  Rapport U62  \u00b7  BTS MEC 2026  \u00b7  ${num}/15`, { x: 0.4, y: 5.25, w: 9.2, h: 0.3, fontSize: 7, color: MUTED, fontFace: FONT_BODY, align: "right" });
}

function addTitle(slide, title, subtitle) {
  slide.addText(title, { x: 0.4, y: 0.62, w: 9.2, h: 0.45, fontSize: 18, bold: true, color: DARK, fontFace: FONT_TITLE, margin: 0 });
  if (subtitle) {
    slide.addText(subtitle, { x: 0.4, y: 1.02, w: 9.2, h: 0.3, fontSize: 10, color: MUTED, fontFace: FONT_BODY, margin: 0 });
  }
}

function addCard(slide, x, y, w, h, fillColor) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fillColor || WHITE }, shadow: mkShadow() });
}

function addCircle(slide, x, y, text, size) {
  const s = size || 0.38;
  slide.addShape(pres.shapes.OVAL, { x, y, w: s, h: s, fill: { color: TEAL } });
  slide.addText(text, { x, y, w: s, h: s, fontSize: 14, bold: true, color: WHITE, fontFace: FONT_TITLE, align: "center", valign: "middle", margin: 0 });
}

// ============================================================
// SLIDE 1 - TITRE
// ============================================================
let s1 = pres.addSlide();
s1.background = { color: TEAL };
s1.addText("U62", { x: 0.5, y: 0.3, w: 2, h: 0.4, fontSize: 11, charSpacing: 6, color: TEAL_LIGHT, fontFace: FONT_BODY, margin: 0 });
s1.addText([
  { text: "Rapport", options: { breakLine: true, fontSize: 36, bold: true } },
  { text: "d\u2019Activit\u00e9s", options: { breakLine: true, fontSize: 36, bold: true } },
  { text: "Professionnelles", options: { fontSize: 36, bold: true } }
], { x: 0.5, y: 1.0, w: 6, h: 2.5, color: WHITE, fontFace: FONT_TITLE, margin: 0 });

s1.addText("BTS MEC \u00b7 SESSION 2026", { x: 0.5, y: 3.5, w: 5, h: 0.3, fontSize: 10, color: TEAL_LIGHT, fontFace: FONT_BODY, charSpacing: 3, margin: 0 });

// Stats right side
const statsY = 1.2;
s1.addShape(pres.shapes.RECTANGLE, { x: 7, y: statsY, w: 2.6, h: 3.2, fill: { color: WHITE }, shadow: mkShadow() });
s1.addText([
  { text: "8 ans", options: { bold: true, fontSize: 22, breakLine: true } },
  { text: "Exp\u00e9rience BTP", options: { fontSize: 9, color: MUTED, breakLine: true } },
  { text: "", options: { breakLine: true, fontSize: 6 } },
  { text: "2 pays", options: { bold: true, fontSize: 22, breakLine: true } },
  { text: "Maroc + France", options: { fontSize: 9, color: MUTED, breakLine: true } },
  { text: "", options: { breakLine: true, fontSize: 6 } },
  { text: "82,5 M DH", options: { bold: true, fontSize: 22, breakLine: true } },
  { text: "Investissements g\u00e9r\u00e9s", options: { fontSize: 9, color: MUTED, breakLine: true } },
  { text: "", options: { breakLine: true, fontSize: 6 } },
  { text: "7 march\u00e9s", options: { bold: true, fontSize: 22, breakLine: true } },
  { text: "publics", options: { fontSize: 9, color: MUTED } }
], { x: 7.25, y: statsY + 0.15, w: 2.1, h: 2.9, color: DARK, fontFace: FONT_BODY, align: "center", valign: "middle", margin: 0 });

s1.addText([
  { text: "BAHAFID Mohamed", options: { bold: true, breakLine: true } },
  { text: "N\u00b0 02537399911 \u00b7 Acad\u00e9mie de Lyon \u00b7 BIMCO", options: {} }
], { x: 0.5, y: 4.7, w: 6, h: 0.7, fontSize: 10, color: TEAL_LIGHT, fontFace: FONT_BODY, margin: 0 });


s1.addNotes(`SLIDE 1 - TITRE (1 min)
Bonjour, Mohamed BAHAFID, candidat libre BTS MEC session 2026, acad\u00e9mie de Lyon.
En 8 ans dans le BTP, j'ai travaill\u00e9 des deux c\u00f4t\u00e9s de la table : 3 ans c\u00f4t\u00e9 ma\u00eetrise d'ouvrage au Maroc, 5 ans c\u00f4t\u00e9 ex\u00e9cution en France.
C'est cette double lecture des projets qui structure tout ce rapport.
82,5 millions de dirhams d'investissements g\u00e9r\u00e9s, 7 march\u00e9s publics.
Aujourd'hui je dirige BIMCO, micro-entreprise sp\u00e9cialis\u00e9e BIM et \u00e9conomie de la construction.`);

// ============================================================
// SLIDE 2 - FICHE D'IDENTITE + PARCOURS
// ============================================================
let s2 = pres.addSlide();
addTopBar(s2);
addBadge(s2, "Profil");
addTitle(s2, "QUI SUIS-JE ?", "Un profil construit en trois phases : MOA publique, ex\u00e9cution terrain, BIM");
addFooter(s2, 2);

// Left - Fiche
addCard(s2, 0.4, 1.45, 4.3, 3.6, WHITE);
const ficheItems = [
  ["Candidat", "BAHAFID Mohamed"],
  ["N\u00b0 Candidat", "02537399911"],
  ["Acad\u00e9mie", "Lyon"],
  ["Structure", "Conseil R\u00e9gional de B\u00e9ni Mellal-Kh\u00e9nifra"],
  ["Poste", "Technicien \u00e9tudes et suivi des travaux"],
  ["Exp\u00e9rience", "8 ans BTP (3 ans Maroc + 5 ans France)"],
  ["Formation BIM", "Modeleur BIM \u2013 AFPA Colmar (8 mois)"],
  ["Activit\u00e9 actuelle", "BIMCO \u2013 Projeteur BIM / \u00e9conomiste"],
  ["SIREN", "999 580 053 / 7112B"],
];
ficheItems.forEach((item, i) => {
  s2.addText([
    { text: item[0], options: { bold: true, color: TEAL, fontSize: 8 } },
    { text: "  " + item[1], options: { color: BODY, fontSize: 8 } }
  ], { x: 0.6, y: 1.55 + i * 0.38, w: 3.9, h: 0.35, fontFace: FONT_BODY, margin: 0 });
});

// Right - Timeline
addCard(s2, 5.1, 1.45, 4.5, 3.6, WHITE);
const timeline = [
  ["2017\u20132022", "MOA publique \u2013 Maroc (3 ans)", "Conseil R\u00e9gional BMK. 7 march\u00e9s publics, +100 M DH"],
  ["2022\u20132024", "Ex\u00e9cution terrain \u2013 France", "Chef GO Ergalis + Chef chantier Minssieux"],
  ["2024\u20132025", "Formation BIM \u2013 AFPA Colmar", "Titre Modeleur BIM. B\u00e2timent R+2, 78 postes"],
  ["2026", "BIMCO + BTS MEC", "Micro-entreprise + candidat libre"]
];
s2.addText("PARCOURS CHRONOLOGIQUE", { x: 5.3, y: 1.55, w: 4.1, h: 0.3, fontSize: 9, bold: true, charSpacing: 2, color: TEAL, fontFace: FONT_BODY, margin: 0 });
timeline.forEach((t, i) => {
  addCircle(s2, 5.3, 2.0 + i * 0.82, String(i + 1), 0.3);
  s2.addText([
    { text: t[0] + "  ", options: { bold: true, color: TEAL, fontSize: 9 } },
    { text: t[1], options: { bold: true, color: DARK, fontSize: 9, breakLine: true } },
    { text: t[2], options: { color: MUTED, fontSize: 8 } }
  ], { x: 5.75, y: 1.96 + i * 0.82, w: 3.65, h: 0.7, fontFace: FONT_BODY, margin: 0 });
});


s2.addNotes(`SLIDE 2 - QUI SUIS-JE ? (1 min 30)
Mon parcours se d\u00e9compose en 4 phases.
Phase 1 : au Maroc, de 2017 \u00e0 2022, technicien \u00e9tudes et suivi au Conseil R\u00e9gional de B\u00e9ni Mellal. 3 ans d'activit\u00e9 effective sur les march\u00e9s publics \u2013 7 march\u00e9s pour plus de 100 millions de dirhams. M\u00e9tr\u00e9s, estimation confidentielle, commissions d'AO, suivi financier.
Phase 2 : en France depuis 2022, chef d'\u00e9quipe gros \u0153uvre chez Ergalis \u00e0 Feurs puis chef de chantier chez Minssieux \u00e0 Mornant. L\u00e0 j'ai appris les co\u00fbts r\u00e9els de production.
Phase 3 : formation BIM \u00e0 l'AFPA Colmar en 2024-2025. Titre professionnel Modeleur BIM, 8 mois.
Phase 4 : cr\u00e9ation de BIMCO en janvier 2026 et candidature BTS MEC en candidat libre.
Au total : 8 ans dans le BTP, 3 ans Maroc + 5 ans France. Chaque phase a enrichi la suivante.
[Si le jury demande : les 5 ans France incluent l'ex\u00e9cution terrain 2022-2024, la formation BIM 2024-2025 et BIMCO 2025-2026]`);

// ============================================================
// SLIDE 3 - STRUCTURE D'ACCUEIL
// ============================================================
let s3 = pres.addSlide();
addTopBar(s3);
addBadge(s3, "01 \u2013 Cadre");
addTitle(s3, "STRUCTURE D\u2019ACCUEIL", "Conseil R\u00e9gional de B\u00e9ni Mellal-Kh\u00e9nifra \u2013 Agence d\u2019Ex\u00e9cution des Projets");
addFooter(s3, 3);

// Left - Description
addCard(s3, 0.4, 1.45, 5.5, 3.6, WHITE);
s3.addText([
  { text: "Collectivit\u00e9 territoriale couvrant 5 provinces et 2,5 millions d\u2019habitants.", options: { breakLine: true, fontSize: 9 } },
  { text: "Mon poste : Technicien \u00e9tudes et suivi des travaux.", options: { breakLine: true, fontSize: 9, bold: true, color: TEAL } },
  { text: "", options: { breakLine: true, fontSize: 6 } },
  { text: "MISSIONS PRINCIPALES", options: { breakLine: true, fontSize: 9, bold: true, color: TEAL, charSpacing: 2 } },
], { x: 0.6, y: 1.55, w: 5.1, h: 0.9, fontFace: FONT_BODY, color: BODY, margin: 0 });

const missions = [
  "M\u00e9tr\u00e9s avant-projet et estimation confidentielle",
  "R\u00e9daction des DCE : CPS, RC, BPDE",
  "Analyse des offres en Commission AO",
  "Suivi financier mensuel, d\u00e9comptes",
  "Visites terrain, attachements contradictoires"
];
s3.addText(missions.map((m, i) => ({
  text: m,
  options: { bullet: true, breakLine: i < missions.length - 1, fontSize: 9, color: BODY }
})), { x: 0.6, y: 2.5, w: 5.1, h: 2.0, fontFace: FONT_BODY, margin: 0 });

// Right - Cadre reglementaire
addCard(s3, 6.2, 1.45, 3.4, 3.6, TEAL_LIGHT);
s3.addText("CADRE R\u00c9GLEMENTAIRE", { x: 6.4, y: 1.55, w: 3.0, h: 0.3, fontSize: 9, bold: true, charSpacing: 2, color: TEAL, fontFace: FONT_BODY, margin: 0 });
const reglItems = [
  ["Pi\u00e8ces AO", "CPS + RC + BPDE"],
  ["Proc\u00e9dure", "Appel d\u2019offres ouvert"],
  ["Estimation", "Confidentielle obligatoire"],
  ["Normes", "Normes marocaines, RPS 2000"],
  ["Suivi", "Attachements contradictoires"]
];
reglItems.forEach((r, i) => {
  s3.addText([
    { text: r[0], options: { bold: true, color: TEAL, fontSize: 8, breakLine: true } },
    { text: r[1], options: { color: BODY, fontSize: 8 } }
  ], { x: 6.4, y: 2.0 + i * 0.55, w: 3.0, h: 0.5, fontFace: FONT_BODY, margin: 0 });
});


s3.addNotes(`SLIDE 3 - STRUCTURE D'ACCUEIL (1 min 30)
Le Conseil R\u00e9gional de B\u00e9ni Mellal-Kh\u00e9nifra couvre 5 provinces, 2,5 millions d'habitants.
J'\u00e9tais rattach\u00e9 \u00e0 l'Agence d'Ex\u00e9cution des Projets, dirig\u00e9e par M. DOGHMANI.
Mon poste : technicien \u00e9tudes et suivi des travaux. Concr\u00e8tement, quand un march\u00e9 de 53 millions de dirhams devait \u00eatre lanc\u00e9, c'est moi qui \u00e9tablissais l'estimation confidentielle, r\u00e9digeais le DCE \u2013 CPS, r\u00e8glement de consultation, bordereau des prix \u2013 puis suivais les travaux sur le terrain.
Le cadre r\u00e9glementaire marocain impose une estimation confidentielle obligatoire avant tout appel d'offres. C'est cette pi\u00e8ce qui fixe le prix plafond.
5 missions principales que vous voyez sur la slide : m\u00e9tr\u00e9s, DCE, analyse offres en commission, suivi financier mensuel, et visites terrain avec attachements contradictoires.`);

// ============================================================
// SLIDE 4 - BIMCO
// ============================================================
let s4 = pres.addSlide();
addTopBar(s4);
addBadge(s4, "01 \u2013 BIMCO");
addTitle(s4, "BIMCO \u2013 MON ACTIVIT\u00c9 IND\u00c9PENDANTE", "Cr\u00e9\u00e9e en janvier 2026 \u2013 BIM au service de l\u2019\u00e9conomiste de la construction");
addFooter(s4, 4);

// Left - Mission
addCard(s4, 0.4, 1.45, 5.0, 1.8, WHITE);
s4.addText("MISSION", { x: 0.6, y: 1.55, w: 4.6, h: 0.25, fontSize: 9, bold: true, color: TEAL, charSpacing: 2, fontFace: FONT_BODY, margin: 0 });
s4.addText("Appliquer le BIM aux m\u00e9tiers de l\u2019\u00e9conomiste de la construction. M\u00e9tr\u00e9s par extraction de maquette num\u00e9rique, \u00e9tudes de prix ancr\u00e9es dans les co\u00fbts r\u00e9els, plugins Revit/Dynamo pour automatiser la cha\u00eene m\u00e8tre \u2192 DPGF.", {
  x: 0.6, y: 1.85, w: 4.6, h: 1.2, fontSize: 9, color: BODY, fontFace: FONT_BODY, margin: 0
});

// Tech stack pills
const techs = ["Revit API", "C# .NET", "Dynamo", "Python", "IFC / BIM360", "React / Node.js"];
addCard(s4, 0.4, 3.45, 5.0, 1.55, WHITE);
s4.addText("STACK TECHNIQUE", { x: 0.6, y: 3.55, w: 4.6, h: 0.25, fontSize: 9, bold: true, color: TEAL, charSpacing: 2, fontFace: FONT_BODY, margin: 0 });
techs.forEach((t, i) => {
  const col = i % 3;
  const row = Math.floor(i / 3);
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.6 + col * 1.6, y: 3.9 + row * 0.45, w: 1.45, h: 0.35, fill: { color: TEAL_LIGHT }, line: { color: TEAL, width: 0.5 } });
  s4.addText(t, { x: 0.6 + col * 1.6, y: 3.9 + row * 0.45, w: 1.45, h: 0.35, fontSize: 8, bold: true, color: TEAL, fontFace: FONT_BODY, align: "center", valign: "middle", margin: 0 });
});

// Right - Positionnement
addCard(s4, 5.7, 1.45, 3.9, 3.55, TEAL);
s4.addText("POSITIONNEMENT", { x: 5.9, y: 1.6, w: 3.5, h: 0.3, fontSize: 10, bold: true, color: WHITE, charSpacing: 2, fontFace: FONT_BODY, margin: 0 });
s4.addText("Triple comp\u00e9tence rare sur le march\u00e9 :", { x: 5.9, y: 1.95, w: 3.5, h: 0.3, fontSize: 9, color: TEAL_LIGHT, fontFace: FONT_BODY, margin: 0 });
const positions = [
  "MOA publique (Maroc) \u2013 vision globale projet",
  "Ex\u00e9cution terrain (France) \u2013 r\u00e9alit\u00e9 chantier",
  "BIM \u2013 pont entre conception et chiffrage"
];
s4.addText(positions.map((p, i) => ({
  text: p, options: { bullet: true, breakLine: i < 2, fontSize: 9, color: WHITE }
})), { x: 5.9, y: 2.35, w: 3.5, h: 1.5, fontFace: FONT_BODY, margin: 0 });

s4.addText([
  { text: "BIM", options: { bold: true, fontSize: 28, color: WHITE, breakLine: false } },
  { text: " + ", options: { fontSize: 20, color: TEAL_LIGHT, breakLine: false } },
  { text: "CO", options: { bold: true, fontSize: 28, color: WHITE } }
], { x: 5.9, y: 3.9, w: 3.5, h: 0.7, fontFace: FONT_TITLE, align: "center", margin: 0 });
s4.addText("Building Information Modeling + \u00c9conomie de la Construction", { x: 5.9, y: 4.55, w: 3.5, h: 0.3, fontSize: 7, color: TEAL_LIGHT, fontFace: FONT_BODY, align: "center", margin: 0 });


s4.addNotes(`SLIDE 4 - BIMCO (1 min)
BIMCO est n\u00e9 d'un constat simple : les outils BIM sont faits pour les architectes, pas pour celui qui chiffre.
J'ai cr\u00e9\u00e9 BIMCO en janvier 2026 pour corriger ce manque.
La mission : extraire les m\u00e9tr\u00e9s directement de la maquette num\u00e9rique au lieu de les compter sur plan. \u00c9tudes de prix ancr\u00e9es dans les co\u00fbts r\u00e9els du terrain. Et d\u00e9veloppement de plugins Revit/Dynamo pour automatiser la cha\u00eene m\u00e8tre vers DPGF.
Mon positionnement est rare sur le march\u00e9 : triple comp\u00e9tence MOA publique, ex\u00e9cution terrain, et BIM. Peu de professionnels combinent ces trois dimensions.
C\u00f4t\u00e9 technique : Revit API, C#, Dynamo, Python pour les plugins. React et Node.js pour les applications web. IFC et BIM360 pour la collaboration.
BIMCO est le prolongement direct de 8 ans d'exp\u00e9rience.`);

// ============================================================
// SLIDE 5 - PROJET 1 PRESENTATION + BUDGET
// ============================================================
let s5 = pres.addSlide();
addTopBar(s5);
addBadge(s5, "02 \u2013 Projet 1");
addTitle(s5, "MISE \u00c0 NIVEAU DE 4 COMMUNES", "March\u00e9 n\u00b038-RBK-2017 \u2013 53,5 M DH TTC \u2013 Province de Kh\u00e9nifra");
addFooter(s5, 5);

// Big number
s5.addText("53,5 M DH", { x: 0.4, y: 1.5, w: 3.5, h: 0.6, fontSize: 28, bold: true, color: TEAL, fontFace: FONT_TITLE, margin: 0 });
s5.addText("4 communes \u00b7 8 corps d\u2019\u00e9tat \u00b7 18 mois", { x: 0.4, y: 2.05, w: 3.5, h: 0.3, fontSize: 9, color: MUTED, fontFace: FONT_BODY, margin: 0 });

// Corps d'etat table
const corps = [
  ["01", "Assainissement", "22%"], ["02", "Chauss\u00e9e", "19%"],
  ["03", "Trottoirs", "14%"], ["04", "Signalisation", "6%"],
  ["05", "\u00c9clairage public", "16%"], ["06", "Murs & ouvrages", "14%"],
  ["07", "Paysager", "5%"], ["08", "Mobilier urbain", "4%"]
];
const tableRows = [
  [
    { text: "Partie", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 8 } },
    { text: "D\u00e9signation", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 8 } },
    { text: "%", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 8, align: "center" } }
  ]
];
corps.forEach(c => {
  tableRows.push([
    { text: c[0], options: { fontSize: 8, color: TEAL, bold: true } },
    { text: c[1], options: { fontSize: 8, color: BODY } },
    { text: c[2], options: { fontSize: 8, color: BODY, align: "center" } }
  ]);
});
s5.addTable(tableRows, { x: 0.4, y: 2.5, w: 4.5, colW: [0.6, 2.8, 0.6], border: { pt: 0.5, color: "DDDDDD" }, rowH: 0.28 });

// Right - chart
s5.addChart(pres.charts.PIE, [{
  name: "Budget", labels: ["Assainissement", "Chauss\u00e9e", "\u00c9clairage", "Trottoirs", "Murs", "Signalisation", "Paysager", "Mobilier"],
  values: [22, 19, 16, 14, 14, 6, 5, 4]
}], {
  x: 5.3, y: 1.45, w: 4.3, h: 3.8,
  showPercent: true, showTitle: false, showLegend: true, legendPos: "b", legendFontSize: 7,
  chartColors: ["007A7F", "00A3A8", "4DB8BD", "80CED1", "B3E4E6", "E0F4F5", "D97706", "F5DEB3"],
  dataLabelColor: WHITE, dataLabelFontSize: 8
});


s5.addNotes(`SLIDE 5 - PROJET 1 (1 min)
Le Projet 1 : mise \u00e0 niveau de 4 communes de la province de Kh\u00e9nifra.
53,5 millions de dirhams TTC, soit environ 4,8 millions d'euros. March\u00e9 unique couvrant 8 corps d'\u00e9tat, de l'assainissement au mobilier urbain, sur 4 sites distants de 20 \u00e0 80 km.
18 mois de suivi.
Comme vous le voyez sur le graphique, l'assainissement et la chauss\u00e9e concentrent 41% du budget \u00e0 eux seuls.
C'est sur ce projet que se d\u00e9roulent les 4 premi\u00e8res situations de mon rapport.
Le d\u00e9tail des 8 corps d'\u00e9tat est dans le tableau \u00e0 gauche \u2013 je ne vais pas les lire, vous les avez dans le rapport page 10.`);

// ============================================================
// SLIDE 6 - SITUATIONS 1 & 2
// ============================================================
let s6 = pres.addSlide();
addTopBar(s6);
addBadge(s6, "02 \u2013 Sit. 1&2");
addTitle(s6, "ESTIMATION CONFIDENTIELLE & ANALYSE DES OFFRES", "Comp\u00e9tence C18 \u2013 M\u00e9trer, estimer, analyser les offres");
addFooter(s6, 6);

// Situation 1 - Left card
addCard(s6, 0.4, 1.45, 4.5, 3.6, WHITE);
s6.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.45, w: 4.5, h: 0.04, fill: { color: TEAL } });
s6.addText("SITUATION 1 \u2013 ESTIMATION CONFIDENTIELLE", { x: 0.6, y: 1.55, w: 4.1, h: 0.3, fontSize: 8, bold: true, color: TEAL, charSpacing: 1, fontFace: FONT_BODY, margin: 0 });
s6.addText("Ouaoumana \u2013 15,8 M DH HT", { x: 0.6, y: 1.85, w: 4.1, h: 0.25, fontSize: 9, bold: true, color: DARK, fontFace: FONT_BODY, margin: 0 });

s6.addText("3,2%", { x: 3.4, y: 1.55, w: 1.3, h: 0.6, fontSize: 28, bold: true, color: ACCENT_GREEN, fontFace: FONT_TITLE, align: "right", margin: 0 });

const sit1 = [
  ["C", "Prix plafond avant AO. 35% du budget. D\u00e9lai 3 semaines."],
  ["P", "Mercuriale 2014 obsol\u00e8te : prix d\u00e9riv\u00e9s de 15 \u00e0 22%."],
  ["A", "112 lignes AutoCAD + 4 visites terrain + 3 sources crois\u00e9es."],
  ["R", "\u00c9cart 3,2% vs 5-10% norme. M\u00e9thode adopt\u00e9e standard."]
];
sit1.forEach((s, i) => {
  addCircle(s6, 0.6, 2.2 + i * 0.65, s[0], 0.26);
  s6.addText(s[1], { x: 0.95, y: 2.18 + i * 0.65, w: 3.85, h: 0.55, fontSize: 8, color: BODY, fontFace: FONT_BODY, margin: 0 });
});

// Situation 2 - Right card
addCard(s6, 5.1, 1.45, 4.5, 3.6, WHITE);
s6.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.45, w: 4.5, h: 0.04, fill: { color: ACCENT_ORANGE } });
s6.addText("SITUATION 2 \u2013 ANALYSE DES OFFRES", { x: 5.3, y: 1.55, w: 4.1, h: 0.3, fontSize: 8, bold: true, color: ACCENT_ORANGE, charSpacing: 1, fontFace: FONT_BODY, margin: 0 });
s6.addText("Commission CAO \u2013 3 offres", { x: 5.3, y: 1.85, w: 4.1, h: 0.25, fontSize: 9, bold: true, color: DARK, fontFace: FONT_BODY, margin: 0 });

s6.addText("94/100", { x: 8.1, y: 1.55, w: 1.3, h: 0.6, fontSize: 28, bold: true, color: ACCENT_ORANGE, fontFace: FONT_TITLE, align: "right", margin: 0 });

const sit2 = [
  ["C", "Membre technique commission. Analyse comparative 3 dossiers."],
  ["P", "7 erreurs arithm\u00e9tiques + prix bas anormaux sur 42% du montant."],
  ["A", "Grille 100 pts (technique 60 + financier 40). Justification \u00e9crite."],
  ["R", "Attribution 15 jours. Z\u00e9ro recours. Rapport valid\u00e9 sans r\u00e9serve."]
];
sit2.forEach((s, i) => {
  addCircle(s6, 5.3, 2.2 + i * 0.65, s[0], 0.26);
  s6.addText(s[1], { x: 5.65, y: 2.18 + i * 0.65, w: 3.85, h: 0.55, fontSize: 8, color: BODY, fontFace: FONT_BODY, margin: 0 });
});


s6.addNotes(`SLIDE 6 - SITUATIONS 1 & 2 (2 min)
SITUATION 1 \u2013 L'estimation confidentielle d'Ouaoumana.
15,8 millions de dirhams, 35% du budget global. Je devais fixer le prix plafond avant l'appel d'offres, en 3 semaines.
Le probl\u00e8me : la mercuriale de r\u00e9f\u00e9rence datait de 2014. Les prix des enrob\u00e9s avaient d\u00e9riv\u00e9 de 15 \u00e0 22%. Impossible de chiffrer correctement avec des prix obsol\u00e8tes.
J'ai crois\u00e9 3 sources : mercuriale actualis\u00e9e, 5 march\u00e9s similaires r\u00e9cemment adjug\u00e9s, et 8 devis fournisseurs.
R\u00e9sultat : 3,2% d'\u00e9cart avec l'offre retenue. La norme c'est 5 \u00e0 10%. La m\u00e9thode a \u00e9t\u00e9 adopt\u00e9e comme standard par l'Agence.

SITUATION 2 \u2013 L'analyse des offres en commission.
3 entreprises soumissionnaires. Un dossier avec 7 erreurs arithm\u00e9tiques. Un autre avec des prix anormalement bas sur 42% du montant \u2013 des \u00e9carts de 25 \u00e0 33%.
J'ai construit une grille de notation sur 100 points : 60 technique, 40 financier. J'ai demand\u00e9 une justification \u00e9crite pour les prix bas.
R\u00e9sultat : 94/100 pour l'entreprise retenue, attribution en 15 jours, z\u00e9ro recours d\u00e9pos\u00e9.`);

// ============================================================
// SLIDE 7 - SITUATIONS 3 & 4
// ============================================================
let s7 = pres.addSlide();
addTopBar(s7);
addBadge(s7, "02 \u2013 Sit. 3&4");
addTitle(s7, "SUIVI FINANCIER & COORDINATION MULTI-SITES", "Comp\u00e9tences C18 (suivi budget) et C19 (coordonner, communiquer)");
addFooter(s7, 7);

// Situation 3 - Left
addCard(s7, 0.4, 1.45, 4.5, 3.6, WHITE);
s7.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.45, w: 4.5, h: 0.04, fill: { color: TEAL } });
s7.addText("SITUATION 3 \u2013 SUIVI FINANCIER KERROUCHEN", { x: 0.6, y: 1.55, w: 4.1, h: 0.3, fontSize: 8, bold: true, color: TEAL, charSpacing: 1, fontFace: FONT_BODY, margin: 0 });
s7.addText("7,3 M DH \u2013 18 mois", { x: 0.6, y: 1.85, w: 3, h: 0.25, fontSize: 9, bold: true, color: DARK, fontFace: FONT_BODY, margin: 0 });
s7.addText("+0,8%", { x: 3.4, y: 1.55, w: 1.3, h: 0.6, fontSize: 28, bold: true, color: ACCENT_GREEN, fontFace: FONT_TITLE, align: "right", margin: 0 });

const sit3 = [
  ["C", "Suivi mensuel complet : situations, quantit\u00e9s, d\u00e9comptes."],
  ["P", "Chauss\u00e9e +12%, murs +15%. D\u00e9rive +4,8% (seuil avenant = 5%)."],
  ["A", "TdB hebdo 3 indicateurs. Compensation paysager \u221244k + mobilier \u221212k."],
  ["R", "D\u00e9passement +0,8%. Aucun avenant. TdB adopt\u00e9 3 communes."]
];
sit3.forEach((s, i) => {
  addCircle(s7, 0.6, 2.2 + i * 0.65, s[0], 0.26);
  s7.addText(s[1], { x: 0.95, y: 2.18 + i * 0.65, w: 3.85, h: 0.55, fontSize: 8, color: BODY, fontFace: FONT_BODY, margin: 0 });
});

// Situation 4 - Right
addCard(s7, 5.1, 1.45, 4.5, 3.6, WHITE);
s7.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.45, w: 4.5, h: 0.04, fill: { color: ACCENT_ORANGE } });
s7.addText("SITUATION 4 \u2013 COORDINATION 4 CHANTIERS", { x: 5.3, y: 1.55, w: 4.1, h: 0.3, fontSize: 8, bold: true, color: ACCENT_ORANGE, charSpacing: 1, fontFace: FONT_BODY, margin: 0 });
s7.addText("Province de Kh\u00e9nifra \u2013 20 \u00e0 80 km", { x: 5.3, y: 1.85, w: 3, h: 0.25, fontSize: 9, bold: true, color: DARK, fontFace: FONT_BODY, margin: 0 });
s7.addText("48h", { x: 8.1, y: 1.55, w: 1.3, h: 0.6, fontSize: 28, bold: true, color: ACCENT_ORANGE, fontFace: FONT_TITLE, align: "right", margin: 0 });

const sit4 = [
  ["C", "Relais unique Directeur. Interface entreprise/BET/labo/hi\u00e9rarchie."],
  ["P", "S23 : retard enrob\u00e9s + alerte m\u00e9t\u00e9o + litige bordures 15%."],
  ["A", "Note Directeur + OS arr\u00eat photos + re-mesurage contradictoire."],
  ["R", "3 crises r\u00e9solues. Litige 15% \u2192 2,3%. Mod\u00e8le CR adopt\u00e9."]
];
sit4.forEach((s, i) => {
  addCircle(s7, 5.3, 2.2 + i * 0.65, s[0], 0.26);
  s7.addText(s[1], { x: 5.65, y: 2.18 + i * 0.65, w: 3.85, h: 0.55, fontSize: 8, color: BODY, fontFace: FONT_BODY, margin: 0 });
});


s7.addNotes(`SLIDE 7 - SITUATIONS 3 & 4 (2 min 30)
SITUATION 3 \u2013 Le suivi financier de Kerrouchen. C'est le c\u0153ur de mon rapport.
7,3 millions de dirhams sur 18 mois. \u00c0 mi-parcours, la chauss\u00e9e d\u00e9rape de +12% \u00e0 cause du terrain rocheux, les murs de +15%.
En extrapolant : +4,8% de d\u00e9passement. Le seuil d'avenant est \u00e0 5%. Si on le d\u00e9passe, c'est 3 \u00e0 6 mois de blocage administratif.
J'ai mis en place un tableau de bord hebdomadaire \u00e0 3 indicateurs et propos\u00e9 une compensation : r\u00e9duire le paysager de 44 000 DH et le mobilier de 12 000 DH.
R\u00e9sultat : +0,8% final. Aucun avenant. Le tableau de bord a \u00e9t\u00e9 r\u00e9pliqu\u00e9 sur les 3 autres communes.

SITUATION 4 \u2013 La coordination multi-sites.
La semaine 23, tout s'acc\u00e9l\u00e8re. 3 crises en m\u00eame temps : retard enrob\u00e9s \u00e0 Ouaoumana, alerte m\u00e9t\u00e9o \u00e0 Kerrouchen, litige bordures \u00e0 Sebt Ait Rahou avec 15% d'\u00e9cart.
J'ai trait\u00e9 les 3 fronts en 48 heures. Note au Directeur, OS d'arr\u00eat avec photos dat\u00e9es, re-mesurage contradictoire sur place.
Le litige est tomb\u00e9 de 15% \u00e0 2,3%. Le mod\u00e8le de compte-rendu cr\u00e9\u00e9 dans l'urgence est devenu le standard de l'Agence.
C'est cette semaine 23 qui m'a appris que l'\u00e9crit syst\u00e9matique est le seul rempart contre les litiges.`);

// ============================================================
// SLIDE 8 - PROJET 2 + BUDGET
// ============================================================
let s8 = pres.addSlide();
addTopBar(s8);
addBadge(s8, "03 \u2013 Projet 2");
addTitle(s8, "ROUTE LEHRI-KERROUCHEN \u2013 25 KM", "Programme PRR3 \u2013 29 M DH TTC \u2013 Zone montagneuse du Moyen Atlas");
addFooter(s8, 8);

// Stats row
const p2Stats = [
  ["29 M DH", "Budget TTC"], ["25 km", "Zone montagneuse"], ["120 334 m\u00b3", "D\u00e9blais v\u00e9rifi\u00e9s"], ["53 prix", "Au bordereau"]
];
p2Stats.forEach((st, i) => {
  addCard(s8, 0.4 + i * 2.35, 1.45, 2.15, 0.95, WHITE);
  s8.addText(st[0], { x: 0.4 + i * 2.35, y: 1.5, w: 2.15, h: 0.5, fontSize: 18, bold: true, color: TEAL, fontFace: FONT_TITLE, align: "center", valign: "middle", margin: 0 });
  s8.addText(st[1], { x: 0.4 + i * 2.35, y: 2.0, w: 2.15, h: 0.3, fontSize: 8, color: MUTED, fontFace: FONT_BODY, align: "center", margin: 0 });
});

// Budget chart
s8.addChart(pres.charts.BAR, [{
  name: "Budget", labels: ["Corps chauss\u00e9e", "Ouv. hydrauliques", "Terrassement", "Rev\u00eatement", "Bretelles/carrefour", "Sout\u00e8nement"],
  values: [30.1, 28.1, 17.9, 10.7, 11.3, 1.9]
}], {
  x: 0.4, y: 2.6, w: 9.2, h: 2.6, barDir: "col",
  chartColors: ["007A7F"], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: DARK, dataLabelFontSize: 8,
  catAxisLabelColor: MUTED, catAxisLabelFontSize: 7, valAxisHidden: true,
  valGridLine: { style: "none" }, catGridLine: { style: "none" },
  showLegend: false, showTitle: false
});


s8.addNotes(`SLIDE 8 - PROJET 2 (1 min)
Deuxi\u00e8me projet : la route Lehri-Kerrouchen. 25 kilom\u00e8tres en zone montagneuse du Moyen Atlas.
29 millions de dirhams TTC, programme national des routes rurales PRR3.
D\u00e9nivel\u00e9 cumul\u00e9 de 400 m\u00e8tres, pentes \u00e0 12% en lacets. 53 prix au bordereau.
Le poste dominant : corps de chauss\u00e9e \u00e0 30% du budget, suivi des ouvrages hydrauliques \u00e0 28%. Ce ratio drainage/chauss\u00e9e d\u00e9passe de 2,5 fois celui d'une route en plaine \u2013 c'est la signature budg\u00e9taire d'un chantier montagneux.
120 334 m\u00e8tres cubes de d\u00e9blais \u00e0 v\u00e9rifier. C'est l'objet de la situation 5.`);

// ============================================================
// SLIDE 9 - SITUATION 5
// ============================================================
let s9 = pres.addSlide();
addTopBar(s9);
addBadge(s9, "03 \u2013 Sit. 5");
addTitle(s9, "CONTR\u00d4LE DES CUBATURES DE TERRASSEMENT", "Comp\u00e9tence C18 \u2013 V\u00e9rifier les quantit\u00e9s ex\u00e9cut\u00e9es");
addFooter(s9, 9);

// Big stats
s9.addText("+8%", { x: 0.4, y: 1.45, w: 2.5, h: 0.7, fontSize: 36, bold: true, color: ACCENT_ORANGE, fontFace: FONT_TITLE, margin: 0 });
s9.addText("\u00c9cart cumul\u00e9 d\u00e9tect\u00e9 et corrig\u00e9", { x: 0.4, y: 2.1, w: 2.5, h: 0.3, fontSize: 9, color: MUTED, fontFace: FONT_BODY, margin: 0 });

s9.addText("285 000 DH", { x: 3.2, y: 1.45, w: 2.5, h: 0.7, fontSize: 28, bold: true, color: TEAL, fontFace: FONT_TITLE, margin: 0 });
s9.addText("Surco\u00fbt absorb\u00e9 sans avenant", { x: 3.2, y: 2.1, w: 2.5, h: 0.3, fontSize: 9, color: MUTED, fontFace: FONT_BODY, margin: 0 });

// CPAR card
addCard(s9, 0.4, 2.6, 9.2, 2.6, WHITE);
const sit5 = [
  ["C", "V\u00e9rification 120 334 m\u00b3 d\u00e9blais + 76 735 m\u00b3 remblais. D\u00e9nivel\u00e9 400 m, pentes 12%."],
  ["P", "PK 12 : calcaire fractur\u00e9 non d\u00e9tect\u00e9. 5 000 m\u00b3 \u00e0 reclassifier. \u00c9cart cumul\u00e9 +8%."],
  ["A", "Profils en travers /25 m. Relev\u00e9s GPS contradictoires /500 ml. 3 dalots dimensionn\u00e9s avec BET."],
  ["R", "Surco\u00fbt absorb\u00e9 par compensation. R\u00e9ception sans avenant. Tra\u00e7abilit\u00e9 compl\u00e8te."]
];
sit5.forEach((s, i) => {
  addCircle(s9, 0.6, 2.75 + i * 0.58, s[0], 0.26);
  s9.addText(s[1], { x: 0.95, y: 2.73 + i * 0.58, w: 8.45, h: 0.5, fontSize: 9, color: BODY, fontFace: FONT_BODY, margin: 0 });
});


s9.addNotes(`SLIDE 9 - SITUATION 5 (1 min 30)
Le contr\u00f4le des cubatures.
Au kilom\u00e8tre 12, surprise : du calcaire fractur\u00e9 que l'\u00e9tude g\u00e9otechnique n'avait pas d\u00e9tect\u00e9.
5 000 m\u00e8tres cubes \u00e0 reclassifier en d\u00e9blais rocheux. Le co\u00fbt passe de 28 \u00e0 85 dirhams le m\u00e8tre cube. Surco\u00fbt imm\u00e9diat : 285 000 DH. L'\u00e9cart cumul\u00e9 montait \u00e0 +8% du budget.
Ma r\u00e9ponse : profils en travers tous les 25 m\u00e8tres, relev\u00e9s GPS contradictoires tous les 500 ml avec le conducteur de travaux.
En parall\u00e8le, 3 dalots b\u00e9ton arm\u00e9 dimensionn\u00e9s avec le bureau d'\u00e9tudes NOVEC pour les passages en fond de vall\u00e9e.
R\u00e9sultat : surco\u00fbt int\u00e9gralement absorb\u00e9 par compensation. R\u00e9ception sans avenant, sans r\u00e9clamation.
Ce contr\u00f4le m'a appris qu'il faut aller sur le terrain, syst\u00e9matiquement.`);

// ============================================================
// SLIDE 10 - ACTIVITES COMPLEMENTAIRES
// ============================================================
let s10 = pres.addSlide();
addTopBar(s10);
addBadge(s10, "03 \u2013 Compl.");
addTitle(s10, "ACTIVIT\u00c9S COMPL\u00c9MENTAIRES", "5 autres march\u00e9s publics + exp\u00e9rience terrain en France");
addFooter(s10, 10);

// Left - 5 marches table
addCard(s10, 0.4, 1.45, 5.2, 2.8, WHITE);
s10.addText("5 AUTRES MARCH\u00c9S AU CONSEIL R\u00c9GIONAL", { x: 0.6, y: 1.55, w: 4.8, h: 0.25, fontSize: 9, bold: true, color: TEAL, charSpacing: 1, fontFace: FONT_BODY, margin: 0 });
const marches = [
  ["27-RBK", "Route Sidi Bouabbad \u2192 Oued Grou (12 km)", "8,2 M DH"],
  ["28-RBK", "Route Ajdir-Ayoun + Piste Lijon", "6,5 M DH"],
  ["30-RBK", "AEP El Borj \u2013 El Hamam (18 km)", "4,8 M DH"],
  ["39-RBK", "Pistes Hartaf \u2013 Sebt Ait Rahou", "3,1 M DH"],
  ["49-RBK", "Voirie Amghass \u2013 Bouchbel", "5,9 M DH"]
];
const marchRows = [[
  { text: "N\u00b0", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7 } },
  { text: "Objet", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7 } },
  { text: "Montant", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7, align: "right" } }
]];
marches.forEach(m => marchRows.push([
  { text: m[0], options: { fontSize: 7, color: TEAL, bold: true } },
  { text: m[1], options: { fontSize: 7, color: BODY } },
  { text: m[2], options: { fontSize: 7, color: BODY, align: "right" } }
]));
s10.addTable(marchRows, { x: 0.6, y: 1.9, w: 4.8, colW: [0.7, 3.1, 1.0], border: { pt: 0.5, color: "DDDDDD" }, rowH: 0.26 });

// Right - France
addCard(s10, 5.9, 1.45, 3.7, 2.8, WHITE);
s10.addText("EXP\u00c9RIENCE TERRAIN FRANCE", { x: 6.1, y: 1.55, w: 3.3, h: 0.25, fontSize: 9, bold: true, color: ACCENT_ORANGE, charSpacing: 1, fontFace: FONT_BODY, margin: 0 });
s10.addText([
  { text: "Chef d\u2019\u00e9quipe GO \u2013 Ergalis BTP", options: { bold: true, breakLine: true, fontSize: 9, color: DARK } },
  { text: "Feurs (Loire) : banches, armatures HA, coulage", options: { breakLine: true, fontSize: 8, color: BODY } },
  { text: "", options: { breakLine: true, fontSize: 5 } },
  { text: "Chef de chantier \u2013 Minssieux & Fils", options: { bold: true, breakLine: true, fontSize: 9, color: DARK } },
  { text: "Mornant (Rh\u00f4ne) : planning, contr\u00f4le qualit\u00e9", options: { breakLine: true, fontSize: 8, color: BODY } },
  { text: "", options: { breakLine: true, fontSize: 5 } },
  { text: "Apport : ", options: { bold: true, color: TEAL, fontSize: 8 } },
  { text: "co\u00fbts r\u00e9els de production, contraintes d\u2019ex\u00e9cution, rendements terrain", options: { fontSize: 8, color: BODY } }
], { x: 6.1, y: 1.9, w: 3.3, h: 2.2, fontFace: FONT_BODY, margin: 0 });

// Bottom quote
s10.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.45, w: 9.2, h: 0.55, fill: { color: TEAL_LIGHT }, line: { color: TEAL, width: 0.5 } });
s10.addText("4 types d\u2019infrastructures ma\u00eetris\u00e9es : routes, VRD, adduction d\u2019eau potable, voirie urbaine", {
  x: 0.6, y: 4.45, w: 8.8, h: 0.55, fontSize: 9, italic: true, color: TEAL, fontFace: FONT_BODY, align: "center", valign: "middle", margin: 0
});


s10.addNotes(`SLIDE 10 - ACTIVIT\u00c9S COMPL\u00c9MENTAIRES (1 min)
En parall\u00e8le des deux projets principaux, j'ai suivi 5 autres march\u00e9s couvrant des types d'infrastructures diff\u00e9rents : routes rurales, pistes, adduction d'eau potable et voirie urbaine.
Ces missions m'ont permis de construire une vision transversale de la ma\u00eetrise d'ouvrage publique.
C\u00f4t\u00e9 France : chef d'\u00e9quipe gros \u0153uvre chez Ergalis \u00e0 Feurs, puis chef de chantier chez Minssieux \u00e0 Mornant.
L'apport essentiel de cette exp\u00e9rience terrain : comprendre les co\u00fbts r\u00e9els de production. Main-d'\u0153uvre, rendements, consommation mat\u00e9riaux. Conna\u00eetre le chantier de l'int\u00e9rieur change profond\u00e9ment la fa\u00e7on d'estimer.
Au total : 4 types d'infrastructures ma\u00eetris\u00e9es.`);

// ============================================================
// SLIDE 11 - SYNTHESE DES COMPETENCES
// ============================================================
let s11 = pres.addSlide();
addTopBar(s11);
addBadge(s11, "04 \u2013 Bilan");
addTitle(s11, "SYNTH\u00c8SE DES COMP\u00c9TENCES BTS MEC", "Comp\u00e9tences mobilis\u00e9es sur 5 situations professionnelles");
addFooter(s11, 11);

const compRows = [
  [
    { text: "Activit\u00e9", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7 } },
    { text: "Sous-comp\u00e9tence", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7 } },
    { text: "Sit.", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7, align: "center" } },
    { text: "Niveau", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7, align: "center" } }
  ],
  [{ text: "M\u00e9tr\u00e9s 112 prix, 4 communes", options: { fontSize: 7 } }, { text: "R\u00e9aliser des m\u00e9tr\u00e9s TCA", options: { fontSize: 7 } }, { text: "1", options: { fontSize: 7, align: "center" } }, { text: "Ma\u00eetrise", options: { fontSize: 7, align: "center", color: TEAL, bold: true } }],
  [{ text: "Estimation 15,8 M DH, 3 sources", options: { fontSize: 7 } }, { text: "Estimer un ouvrage en AO", options: { fontSize: 7 } }, { text: "1", options: { fontSize: 7, align: "center" } }, { text: "Ma\u00eetrise", options: { fontSize: 7, align: "center", color: TEAL, bold: true } }],
  [{ text: "Analyse 3 offres, grille /100", options: { fontSize: 7 } }, { text: "Analyser les offres en commission", options: { fontSize: 7 } }, { text: "2", options: { fontSize: 7, align: "center" } }, { text: "Ma\u00eetrise", options: { fontSize: 7, align: "center", color: TEAL, bold: true } }],
  [{ text: "TdB hebdo, compensation", options: { fontSize: 7 } }, { text: "Suivre l\u2019ex\u00e9cution financi\u00e8re", options: { fontSize: 7 } }, { text: "3", options: { fontSize: 7, align: "center" } }, { text: "Expert", options: { fontSize: 7, align: "center", color: ACCENT_GREEN, bold: true } }],
  [{ text: "Contr\u00f4le cubatures GPS", options: { fontSize: 7 } }, { text: "V\u00e9rifier les quantit\u00e9s", options: { fontSize: 7 } }, { text: "5", options: { fontSize: 7, align: "center" } }, { text: "Expert", options: { fontSize: 7, align: "center", color: ACCENT_GREEN, bold: true } }],
  [{ text: "CR, OS, notes factuelles", options: { fontSize: 7 } }, { text: "Communiquer par \u00e9crit", options: { fontSize: 7 } }, { text: "4", options: { fontSize: 7, align: "center" } }, { text: "Ma\u00eetrise", options: { fontSize: 7, align: "center", color: TEAL, bold: true } }],
  [{ text: "R\u00e9unions, points quotidiens, CAO", options: { fontSize: 7 } }, { text: "Communiquer oralement", options: { fontSize: 7 } }, { text: "2,4", options: { fontSize: 7, align: "center" } }, { text: "Ma\u00eetrise", options: { fontSize: 7, align: "center", color: TEAL, bold: true } }],
  [{ text: "Convention BIM, export IFC, 78 postes", options: { fontSize: 7 } }, { text: "Collaborer en BIM (Open BIM)", options: { fontSize: 7 } }, { text: "BIM", options: { fontSize: 7, align: "center" } }, { text: "Ma\u00eetrise", options: { fontSize: 7, align: "center", color: TEAL, bold: true } }],
];
s11.addTable(compRows, { x: 0.4, y: 1.4, w: 9.2, colW: [3.0, 2.8, 0.6, 0.9], border: { pt: 0.5, color: "DDDDDD" }, rowH: 0.32, autoPage: false });

// Radar chart
s11.addChart(pres.charts.RADAR, [{
  name: "Niveau", labels: ["M\u00e9tr\u00e9s", "Estimation", "Analyse AO", "Suivi financier", "Coordination", "Cubatures", "BIM"],
  values: [90, 90, 85, 95, 88, 92, 80]
}], { x: 5.5, y: 3.4, w: 4.1, h: 2.0, chartColors: ["007A7F"], showLegend: false, catAxisLabelFontSize: 7, catAxisLabelColor: BODY });


s11.addNotes(`SLIDE 11 - SYNTH\u00c8SE DES COMP\u00c9TENCES (1 min)
Vous avez le d\u00e9tail page 23 du rapport.
Ce que je retiens : 8 sous-comp\u00e9tences du BTS MEC mobilis\u00e9es, toutes pratiqu\u00e9es en situation r\u00e9elle sous contrainte.
Deux comp\u00e9tences au niveau Expert : le suivi financier, gr\u00e2ce au tableau de bord qui est devenu un outil de r\u00e9f\u00e9rence, et le contr\u00f4le des quantit\u00e9s, avec la m\u00e9thode de contr\u00f4le contradictoire GPS que j'ai syst\u00e9matis\u00e9e.
Les autres comp\u00e9tences sont au niveau Ma\u00eetrise \u2013 pratique r\u00e9guli\u00e8re et autonome.
Le radar \u00e0 droite montre que le suivi financier est mon point le plus fort \u00e0 95%, et le BIM \u00e0 80% est en progression, c'est la comp\u00e9tence la plus r\u00e9cente.`);

// ============================================================
// SLIDE 12 - ANALYSE MAROC / FRANCE
// ============================================================
let s12 = pres.addSlide();
addTopBar(s12);
addBadge(s12, "04 \u2013 Analyse");
addTitle(s12, "ANALYSE COMPARATIVE MAROC / FRANCE", "M\u00eames fondamentaux \u2013 deux cadres r\u00e9glementaires \u2013 une seule exigence : la rigueur");
addFooter(s12, 12);

const compTableRows = [
  [
    { text: "Aspect", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 8 } },
    { text: "Maroc", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 8, align: "center" } },
    { text: "France", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 8, align: "center" } }
  ],
  [{ text: "R\u00e9glementation", options: { fontSize: 8, bold: true } }, { text: "D\u00e9cret n\u00b02-12-349", options: { fontSize: 8 } }, { text: "Code commande publique", options: { fontSize: 8 } }],
  [{ text: "Pi\u00e8ces march\u00e9", options: { fontSize: 8, bold: true } }, { text: "CPS + RC + BPDE", options: { fontSize: 8 } }, { text: "CCAP + CCTP + BPU/DQE", options: { fontSize: 8 } }],
  [{ text: "Estimation", options: { fontSize: 8, bold: true } }, { text: "Confidentielle obligatoire", options: { fontSize: 8 } }, { text: "Estimation MOA", options: { fontSize: 8 } }],
  [{ text: "Normes", options: { fontSize: 8, bold: true } }, { text: "Normes marocaines, RPS 2000", options: { fontSize: 8 } }, { text: "DTU, Eurocodes, RE2020", options: { fontSize: 8 } }],
  [{ text: "Suivi financier", options: { fontSize: 8, bold: true } }, { text: "Attachements contradictoires", options: { fontSize: 8 } }, { text: "Situations mensuelles", options: { fontSize: 8 } }],
  [{ text: "Commission", options: { fontSize: 8, bold: true } }, { text: "CAO (Commission AO)", options: { fontSize: 8 } }, { text: "Commission d\u2019Appel d\u2019Offres", options: { fontSize: 8 } }],
];
s12.addTable(compTableRows, { x: 0.4, y: 1.4, w: 9.2, colW: [1.8, 3.7, 3.7], border: { pt: 0.5, color: "DDDDDD" }, rowH: 0.35 });

// Bottom - 2 cards
addCard(s12, 0.4, 3.95, 4.3, 1.2, WHITE);
s12.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.95, w: 0.06, h: 1.2, fill: { color: TEAL } });
s12.addText([
  { text: "C\u00f4t\u00e9 MOA \u2013 Maroc", options: { bold: true, color: TEAL, fontSize: 9, breakLine: true } },
  { text: "Concevoir les march\u00e9s, r\u00e9diger CPS/RC/BPDE, piloter la CAO, analyser les offres. 7 march\u00e9s, 100 M DH.", options: { fontSize: 8, color: BODY } }
], { x: 0.65, y: 4.0, w: 3.85, h: 1.05, fontFace: FONT_BODY, margin: 0 });

addCard(s12, 5.3, 3.95, 4.3, 1.2, WHITE);
s12.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.95, w: 0.06, h: 1.2, fill: { color: ACCENT_ORANGE } });
s12.addText([
  { text: "C\u00f4t\u00e9 ex\u00e9cution \u2013 France", options: { bold: true, color: ACCENT_ORANGE, fontSize: 9, breakLine: true } },
  { text: "Banches, ferraillage, planning, contr\u00f4le qualit\u00e9. Co\u00fbts r\u00e9els, rendements, contraintes terrain.", options: { fontSize: 8, color: BODY } }
], { x: 5.55, y: 4.0, w: 3.85, h: 1.05, fontFace: FONT_BODY, margin: 0 });


s12.addNotes(`SLIDE 12 - ANALYSE MAROC / FRANCE (1 min 30)
Les deux syst\u00e8mes partagent les m\u00eames principes fondamentaux : transparence, \u00e9galit\u00e9 de traitement, mise en concurrence.
CPS au Maroc, CCAP en France \u2013 les noms changent mais la logique est la m\u00eame.
L'estimation confidentielle au Maroc est obligatoire \u2013 c'est elle qui fixe le prix plafond. En France, on retrouve le m\u00eame m\u00e9canisme avec l'estimation du ma\u00eetre d'ouvrage.
Le suivi se fait par attachements contradictoires au Maroc, par situations mensuelles en France. Les deux exigent la m\u00eame rigueur documentaire.
Ce que cette double exp\u00e9rience m'apporte concr\u00e8tement :
C\u00f4t\u00e9 MOA au Maroc : j'ai appris \u00e0 concevoir les march\u00e9s, \u00e0 r\u00e9diger les pi\u00e8ces, \u00e0 analyser les offres.
C\u00f4t\u00e9 ex\u00e9cution en France : j'ai appris les co\u00fbts r\u00e9els, les rendements, les contraintes terrain.
La combinaison des deux rend mes estimations r\u00e9alistes \u2013 ni trop basses, ni trop hautes.`);

// ============================================================
// SLIDE 13 - BILAN REFLEXIF ADM
// ============================================================
let s13 = pres.addSlide();
addTopBar(s13);
addBadge(s13, "04 \u2013 R\u00e9flexif");
addTitle(s13, "BILAN R\u00c9FLEXIF", "Ce que le terrain m\u2019a appris \u2013 8 ans de pratique");
addFooter(s13, 13);

const campItems = [
  ["A", "CE QUE J\u2019AI APPRIS", "L\u2019\u00e9crit syst\u00e9matique \u2013 OS, attachements, CR consolid\u00e9s \u2013 est le seul rempart contre les litiges. Croiser 3 sources de prix donne un \u00e9cart de 3,2% l\u00e0 o\u00f9 les estimations r\u00e9gionales d\u00e9passaient 10%. Voir le projet depuis le terrain rend les estimations plus justes."],
  ["D", "CE QUE JE FERAIS DIFF\u00c9REMMENT", "D\u00e9ployer le tableau de bord d\u00e8s le 1er mois, pas au 9\u00e8me quand le signal \u00e9tait d\u00e9j\u00e0 \u00e0 +4,8%. Insister pour une \u00e9tude g\u00e9otechnique compl\u00e9mentaire avant terrassements. Standardiser les CR d\u00e8s le d\u00e9marrage \u2013 pas dans l\u2019urgence de la semaine 23."],
  ["M", "CE QUE J\u2019APPORTE AU BTS MEC", "L\u2019estimation confidentielle est l\u2019acte fondateur de tout march\u00e9 public. En France, m\u00eames fondamentaux avec le DQE/DPGF. Triple comp\u00e9tence MOA + Ex\u00e9cution + BIM = positionnement rare. BIMCO est n\u00e9 de ce parcours. Le BIM industrialise la m\u00e9thode d\u00e9velopp\u00e9e sur le terrain."],
];

campItems.forEach((item, i) => {
  const col = 0;
  const row = i;
  const x = 0.4;
  const y = 1.4 + row * 1.35;
  addCard(s13, x, y, 9.2, 1.15, WHITE);
  addCircle(s13, x + 0.15, y + 0.12, item[0], 0.35);
  s13.addText(item[1], { x: x + 0.6, y: y + 0.1, w: 8.3, h: 0.3, fontSize: 9, bold: true, color: TEAL, charSpacing: 1, fontFace: FONT_BODY, margin: 0 });
  s13.addText(item[2], { x: x + 0.6, y: y + 0.4, w: 8.3, h: 0.7, fontSize: 8.5, color: BODY, fontFace: FONT_BODY, margin: 0 });
});


s13.addNotes(`SLIDE 13 - BILAN R\u00c9FLEXIF (2 min) \u2013 SLIDE IMPORTANTE POUR LE JURY
Cette slide montre ma capacit\u00e9 de recul. Le jury attend \u00e7a. Structure ADM identique au rapport papier page 26.

A \u2013 CE QUE J'AI APPRIS :
La le\u00e7on la plus forte vient de la semaine 23 : l'\u00e9crit syst\u00e9matique \u2013 ordres de service, attachements contradictoires, CR consolid\u00e9s \u2013 est le seul rempart r\u00e9el contre les litiges. Sans la trace \u00e9crite sign\u00e9e \u00e0 Kerrouchen, le d\u00e9passement de 0,8% aurait pu \u00eatre contest\u00e9.
Croiser 3 sources de prix m'a donn\u00e9 un \u00e9cart de 3,2% l\u00e0 o\u00f9 les estimations r\u00e9gionales d\u00e9passaient 10%.

D \u2013 CE QUE JE FERAIS DIFF\u00c9REMMENT :
Si c'\u00e9tait \u00e0 refaire : tableau de bord d\u00e8s le 1er mois, pas au 9\u00e8me quand le signal \u00e9tait d\u00e9j\u00e0 \u00e0 +4,8%. Le probl\u00e8me d\u00e9tect\u00e9 t\u00f4t co\u00fbte dix fois moins cher \u00e0 corriger.
Et j'aurais insist\u00e9 pour une \u00e9tude g\u00e9otechnique compl\u00e9mentaire : la reclassification de 5 000 m\u00b3 en d\u00e9blais rocheux aurait pu \u00eatre anticip\u00e9e.

M \u2013 CE QUE J'APPORTE AU BTS MEC :
L'estimation confidentielle est l'acte fondateur de tout march\u00e9 public. En France, m\u00eames fondamentaux avec le DQE/DPGF. La triple comp\u00e9tence MOA + Ex\u00e9cution + BIM est rare sur le march\u00e9 \u2013 c'est ce positionnement que traduit BIMCO.`);

// ============================================================
// SLIDE 14 - PROTOCOLE BIM + PERSPECTIVES
// ============================================================
let s14 = pres.addSlide();
addTopBar(s14);
addBadge(s14, "05 \u2013 BIM");
addTitle(s14, "PROTOCOLE BIM & PERSPECTIVES BIMCO", "Formation AFPA Colmar + Projet professionnel 2026\u20132029");
addFooter(s14, 14);

// Left - BIM Protocol
addCard(s14, 0.4, 1.4, 4.5, 3.7, WHITE);
s14.addText("PROTOCOLE BIM \u2013 CAS AFPA", { x: 0.6, y: 1.5, w: 4.1, h: 0.25, fontSize: 9, bold: true, color: TEAL, charSpacing: 1, fontFace: FONT_BODY, margin: 0 });

const bimStats = [["78", "postes extraits"], ["1,8%", "\u00e9cart vs m\u00e9tr\u00e9 manuel"], ["12", "clashs r\u00e9solus"]];
bimStats.forEach((bs, i) => {
  s14.addText(bs[0], { x: 0.6 + i * 1.45, y: 1.85, w: 1.3, h: 0.45, fontSize: 20, bold: true, color: TEAL, fontFace: FONT_TITLE, align: "center", margin: 0 });
  s14.addText(bs[1], { x: 0.6 + i * 1.45, y: 2.3, w: 1.3, h: 0.25, fontSize: 7, color: MUTED, fontFace: FONT_BODY, align: "center", margin: 0 });
});

const bimSteps = [
  "Mod\u00e9lisation Revit Architecture + Structure (LOD 300)",
  "Export IFC 2x3 \u2013 Open BIM \u2013 MVD Coordination View",
  "D\u00e9tection clashs Navisworks (12 conflits r\u00e9solus)",
  "Extraction quantit\u00e9s Revit + Dynamo (2h vs 2 jours)",
  "Chiffrage avec tra\u00e7abilit\u00e9 maquette"
];
bimSteps.forEach((step, i) => {
  addCircle(s14, 0.6, 2.7 + i * 0.43, String(i + 1), 0.22);
  s14.addText(step, { x: 0.9, y: 2.68 + i * 0.43, w: 3.9, h: 0.38, fontSize: 7.5, color: BODY, fontFace: FONT_BODY, margin: 0 });
});

// Right - Perspectives
addCard(s14, 5.1, 1.4, 4.5, 3.7, WHITE);
s14.addText("PROJET PROFESSIONNEL", { x: 5.3, y: 1.5, w: 4.1, h: 0.25, fontSize: 9, bold: true, color: ACCENT_ORANGE, charSpacing: 1, fontFace: FONT_BODY, margin: 0 });

const horizons = [
  ["2026", "Court terme", "BTS MEC + premi\u00e8res prestations BIMCO : m\u00e9tr\u00e9s BIM, \u00e9tudes de prix, DPGF. Premiers plugins Revit/Dynamo."],
  ["2027\u201328", "Moyen terme", "Plugin Revit \u2192 DPGF automatis\u00e9. App web suivi \u00e9conomique. Base prix Batiprix. 3\u20135 clients."],
  ["2029+", "Long terme", "Cabinet ing\u00e9nierie BIM + \u00e9co construction. SaaS MEC. \u00c9quipe 3\u20135 personnes. CA 200\u2013300 k\u20ac/an."]
];
horizons.forEach((h, i) => {
  s14.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.85 + i * 1.1, w: 0.8, h: 0.3, fill: { color: TEAL } });
  s14.addText(h[0], { x: 5.3, y: 1.85 + i * 1.1, w: 0.8, h: 0.3, fontSize: 8, bold: true, color: WHITE, fontFace: FONT_BODY, align: "center", valign: "middle", margin: 0 });
  s14.addText(h[1], { x: 6.2, y: 1.85 + i * 1.1, w: 3.2, h: 0.3, fontSize: 9, bold: true, color: DARK, fontFace: FONT_BODY, margin: 0 });
  s14.addText(h[2], { x: 5.3, y: 2.2 + i * 1.1, w: 4.1, h: 0.8, fontSize: 7.5, color: BODY, fontFace: FONT_BODY, margin: 0 });
});


s14.addNotes(`SLIDE 14 - PROTOCOLE BIM & PERSPECTIVES (1 min 30)
\u00c0 gauche : le protocole BIM que j'ai pratiqu\u00e9 \u00e0 l'AFPA Colmar.
B\u00e2timent R+2, mod\u00e9lis\u00e9 en LOD 300. 78 postes de m\u00e9tr\u00e9s extraits automatiquement en 2 heures \u2013 contre 2 jours en m\u00e9thode traditionnelle.
\u00c9cart avec le m\u00e9tr\u00e9 manuel : seulement 1,8%. Et 12 clashs structure/r\u00e9seaux d\u00e9tect\u00e9s et r\u00e9solus en amont du chantier.
Le workflow en 6 \u00e9tapes : mod\u00e9lisation Revit, export IFC, d\u00e9tection clashs Navisworks, extraction quantit\u00e9s avec Dynamo, puis chiffrage avec tra\u00e7abilit\u00e9 maquette.

\u00c0 droite : mon projet professionnel sur 3 horizons.
Court terme 2026 : obtenir ce BTS MEC, premi\u00e8res prestations BIMCO.
Moyen terme 2027-28 : un plugin Revit qui g\u00e9n\u00e8re automatiquement le DPGF depuis la maquette. App web de suivi \u00e9conomique. 3 \u00e0 5 clients.
Long terme 2029+ : BIMCO devient un cabinet d'ing\u00e9nierie BIM + \u00e9conomie de la construction, avec une \u00e9quipe de 3 \u00e0 5 personnes.`);

// ============================================================
// SLIDE 15 - CONCLUSION / MERCI
// ============================================================
let s15 = pres.addSlide();
s15.background = { color: TEAL };

s15.addText("Merci de votre\nattention", { x: 0.5, y: 0.8, w: 5.5, h: 1.8, fontSize: 32, bold: true, color: WHITE, fontFace: FONT_TITLE, margin: 0 });

s15.addText("Soutenance BTS MEC U62 \u00b7 Session 2026", { x: 0.5, y: 0.4, w: 5, h: 0.3, fontSize: 9, color: TEAL_LIGHT, fontFace: FONT_BODY, charSpacing: 2, margin: 0 });

// Right - 4 takeaways
const takeaways = [
  ["5 comp\u00e9tences terrain", "Estimer, analyser, suivre, coordonner, contr\u00f4ler"],
  ["Double lecture", "MOA Maroc + ex\u00e9cution France"],
  ["Rigueur de l\u2019\u00e9crit", "OS, attachements, CR \u2013 le rempart contre les litiges"],
  ["BIMCO", "La rigueur terrain + les outils num\u00e9riques"]
];
takeaways.forEach((t, i) => {
  s15.addShape(pres.shapes.RECTANGLE, { x: 6, y: 0.8 + i * 1.05, w: 3.6, h: 0.9, fill: { color: WHITE }, shadow: mkShadow() });
  addCircle(s15, 6.15, 0.88 + i * 1.05, String(i + 1), 0.3);
  s15.addText([
    { text: t[0], options: { bold: true, color: TEAL, fontSize: 10, breakLine: true } },
    { text: t[1], options: { color: BODY, fontSize: 8 } }
  ], { x: 6.55, y: 0.88 + i * 1.05, w: 2.9, h: 0.75, fontFace: FONT_BODY, valign: "middle", margin: 0 });
});

s15.addText([
  { text: "BAHAFID Mohamed \u00b7 N\u00b0 02537399911 \u00b7 Acad\u00e9mie de Lyon", options: { breakLine: true } },
  { text: "BIMCO | gestion.bimco-consulting.fr | Bussi\u00e8res, Loire 42510", options: {} }
], { x: 0.5, y: 4.6, w: 6, h: 0.8, fontSize: 9, color: TEAL_LIGHT, fontFace: FONT_BODY, margin: 0 });

s15.addText("8 ans BTP \u00b7 2 pays \u00b7 7 march\u00e9s publics \u00b7 82,5 M DH g\u00e9r\u00e9s \u00b7 Triple comp\u00e9tence MOA + Ex\u00e9cution + BIM", {
  x: 0.5, y: 3.1, w: 5, h: 0.4, fontSize: 8, color: TEAL_LIGHT, fontFace: FONT_BODY, margin: 0
});


s15.addNotes(`SLIDE 15 - CONCLUSION (1 min 30)
Ce rapport d\u00e9montre une chose : je sais r\u00e9soudre un probl\u00e8me sous contrainte \u2013 d\u00e9lai, budget, al\u00e9as g\u00e9ologiques, coordination de crise.

4 points cl\u00e9s :
1. Cinq comp\u00e9tences construites sur le terrain : estimer, analyser, suivre, coordonner, contr\u00f4ler. Acquises par la pratique, pas par la th\u00e9orie.
2. Une double lecture des projets : c\u00f4t\u00e9 MOA au Maroc, c\u00f4t\u00e9 ex\u00e9cution en France. Comprendre les enjeux des deux c\u00f4t\u00e9s de la table change la fa\u00e7on d'estimer.
3. La rigueur de l'\u00e9crit : OS, attachements contradictoires, CR consolid\u00e9s. Le seul rempart r\u00e9el contre les litiges.
4. BIMCO : un bureau o\u00f9 la rigueur terrain rencontre les outils num\u00e9riques.

Le BTS MEC ne valide pas seulement un dipl\u00f4me. Il valide 8 ans de terrain, de chiffres et de chantiers. Et il ouvre la voie \u00e0 ce que je veux construire avec BIMCO.

Merci pour votre attention. Je suis disponible pour r\u00e9pondre \u00e0 vos questions.

[TIMING TOTAL VIS\u00c9 : 18-20 minutes \u2013 garder 2 min de marge pour ne pas d\u00e9passer]`);

// ============================================================
// GENERATE
// ============================================================
pres.writeFile({ fileName: "D:/PREPA BTS MEC/08_U62_Rapport_Activites/Soutenance_PowerPoint/SOUTENANCE_U62_v2.pptx" })
  .then(() => console.log("OK: SOUTENANCE_U62_2026.pptx generated"))
  .catch(err => console.error("ERROR:", err));
