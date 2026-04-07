const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "BAHAFID Mohamed";
pres.title = "Rapport d\u2019Activit\u00e9s Professionnelles \u2013 BTS MEC U62 \u2013 Session 2026";

// === PALETTE — Ocean Gradient + Warm accent ===
const NAVY   = "1A2744";   // dark background
const TEAL   = "007A7F";   // primary accent
const TEAL_D = "005F63";   // dark teal for emphasis
const TEAL_L = "E0F4F5";   // light teal fill
const CREAM  = "FAFAF7";   // slide background — warmer than pure white
const WHITE  = "FFFFFF";
const DARK   = "1E1E1E";   // text headings
const BODY   = "3A3A3A";   // body text
const MUTED  = "777777";   // captions
const ORANGE = "D97706";   // secondary accent
const GREEN  = "16A34A";   // success accent
const CARD   = "FFFFFF";   // card fill

const FT = "Cambria";      // title font
const FB = "Calibri";      // body font

// === HELPERS ===
const shadow = () => ({ type: "outer", blur: 5, offset: 2, angle: 135, color: "000000", opacity: 0.08 });

function topBar(s) {
  s.background = { color: CREAM };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: TEAL } });
}

function badge(s, text) {
  const w = Math.max(1.0, text.length * 0.1 + 0.4);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.15, w, h: 0.36, fill: { color: TEAL_L }, line: { color: TEAL, width: 0.6 }, rectRadius: 0.04 });
  s.addText(text, { x: 0.4, y: 0.15, w, h: 0.36, fontSize: 8.5, bold: true, color: TEAL, fontFace: FB, align: "center", valign: "middle", margin: 0 });
}

function footer(s, n) {
  s.addText(`BAHAFID Mohamed  \u00b7  Rapport U62  \u00b7  BTS MEC 2026  \u00b7  ${n}/15`, {
    x: 0.4, y: 5.28, w: 9.2, h: 0.25, fontSize: 7, color: MUTED, fontFace: FB, align: "right"
  });
}

function title(s, t, sub) {
  s.addText(t, { x: 0.4, y: 0.58, w: 9.2, h: 0.42, fontSize: 18, bold: true, color: DARK, fontFace: FT, margin: 0 });
  if (sub) s.addText(sub, { x: 0.4, y: 0.98, w: 9.2, h: 0.28, fontSize: 9.5, color: MUTED, fontFace: FB, margin: 0 });
}

function card(s, x, y, w, h, fill, accentColor) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill || CARD }, shadow: shadow(), rectRadius: 0.05 });
  if (accentColor) s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.04, fill: { color: accentColor }, rectRadius: 0.02 });
}

function circle(s, x, y, text, sz, bg) {
  const r = sz || 0.32;
  s.addShape(pres.shapes.OVAL, { x, y, w: r, h: r, fill: { color: bg || TEAL } });
  s.addText(text, { x, y, w: r, h: r, fontSize: r > 0.3 ? 13 : 10, bold: true, color: WHITE, fontFace: FT, align: "center", valign: "middle", margin: 0 });
}

function bigStat(s, x, y, value, label, color) {
  s.addText(value, { x, y, w: 2.2, h: 0.55, fontSize: 24, bold: true, color: color || TEAL, fontFace: FT, margin: 0 });
  s.addText(label, { x, y: y + 0.5, w: 2.2, h: 0.22, fontSize: 8, color: MUTED, fontFace: FB, margin: 0 });
}

function cparBlock(s, items, x0, y0, w) {
  items.forEach((item, i) => {
    const colors = { C: TEAL, P: ORANGE, A: TEAL_D, R: GREEN };
    circle(s, x0, y0 + i * 0.6, item[0], 0.24, colors[item[0]] || TEAL);
    s.addText(item[1], { x: x0 + 0.32, y: y0 - 0.02 + i * 0.6, w: w - 0.4, h: 0.52, fontSize: 8, color: BODY, fontFace: FB, margin: 0 });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 1 — TITRE (dark premium)
// ═══════════════════════════════════════════════════════════
let s1 = pres.addSlide();
s1.background = { color: NAVY };

// Left decorative line
s1.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.5, w: 0.04, h: 2.8, fill: { color: TEAL } });

s1.addText("U62", { x: 0.6, y: 0.5, w: 2, h: 0.35, fontSize: 11, charSpacing: 8, color: TEAL, fontFace: FB, margin: 0 });
s1.addText([
  { text: "Rapport", options: { breakLine: true, fontSize: 38, bold: true } },
  { text: "d\u2019Activit\u00e9s", options: { breakLine: true, fontSize: 38, bold: true } },
  { text: "Professionnelles", options: { fontSize: 38, bold: true } }
], { x: 0.6, y: 0.9, w: 6, h: 2.4, color: WHITE, fontFace: FT, margin: 0 });

s1.addText("BTS MEC  \u00b7  SESSION 2026", { x: 0.6, y: 3.4, w: 5, h: 0.3, fontSize: 10, color: TEAL, fontFace: FB, charSpacing: 4, margin: 0 });

// Stats column — right side, stacked vertically with subtle cards
const stats = [["8 ans", "Exp\u00e9rience BTP"], ["2 pays", "Maroc + France"], ["82,5 M DH", "Investissements g\u00e9r\u00e9s"], ["7 march\u00e9s", "publics"]];
stats.forEach((st, i) => {
  const sy = 0.6 + i * 1.1;
  s1.addShape(pres.shapes.RECTANGLE, { x: 7.1, y: sy, w: 2.5, h: 0.9, fill: { color: "2A3A5A" }, line: { color: TEAL, width: 0.5 }, rectRadius: 0.04 });
  s1.addText(st[0], { x: 7.2, y: sy + 0.05, w: 2.3, h: 0.45, fontSize: 22, bold: true, color: WHITE, fontFace: FT, align: "center", margin: 0 });
  s1.addText(st[1], { x: 7.2, y: sy + 0.5, w: 2.3, h: 0.3, fontSize: 8.5, color: TEAL, fontFace: FB, align: "center", margin: 0 });
});

s1.addText([
  { text: "BAHAFID Mohamed", options: { bold: true, breakLine: true, fontSize: 11 } },
  { text: "N\u00b0 02537399911  \u00b7  Acad\u00e9mie de Lyon  \u00b7  BIMCO", options: { fontSize: 9 } }
], { x: 0.6, y: 4.65, w: 6, h: 0.7, color: TEAL, fontFace: FB, margin: 0 });


s1.addNotes(`SLIDE 1 - TITRE (1 min)
Bonjour, Mohamed BAHAFID, candidat libre BTS MEC session 2026, acad\u00e9mie de Lyon.
En 8 ans dans le BTP, j\u2019ai travaill\u00e9 des deux c\u00f4t\u00e9s de la table : 3 ans c\u00f4t\u00e9 ma\u00eetrise d\u2019ouvrage au Maroc, 5 ans c\u00f4t\u00e9 ex\u00e9cution en France.
C\u2019est cette double lecture des projets qui structure tout ce rapport.
82,5 millions de dirhams d\u2019investissements g\u00e9r\u00e9s, 7 march\u00e9s publics.
Aujourd\u2019hui je dirige BIMCO, micro-entreprise sp\u00e9cialis\u00e9e BIM et \u00e9conomie de la construction.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 2 — QUI SUIS-JE
// ═══════════════════════════════════════════════════════════
let s2 = pres.addSlide();
topBar(s2); badge(s2, "Profil"); footer(s2, 2);
title(s2, "QUI SUIS-JE ?", "Un profil construit en trois phases : MOA publique, ex\u00e9cution terrain, BIM");

// Left — Fiche
card(s2, 0.4, 1.4, 4.3, 3.7, CARD, TEAL);
const fiche = [
  ["Candidat", "BAHAFID Mohamed"],
  ["N\u00b0 Candidat", "02537399911"],
  ["Acad\u00e9mie", "Lyon"],
  ["Structure", "Conseil R\u00e9gional de B\u00e9ni Mellal-Kh\u00e9nifra"],
  ["Poste", "Technicien \u00e9tudes et suivi des travaux"],
  ["Exp\u00e9rience", "8 ans BTP (3 ans Maroc + 5 ans France)"],
  ["Formation BIM", "Modeleur BIM \u2013 AFPA Colmar (8 mois)"],
  ["Activit\u00e9 actuelle", "BIMCO \u2013 Projeteur BIM / \u00e9conomiste"],
  ["SIREN", "999 580 053 / 7112B"]
];
fiche.forEach((f, i) => {
  const yy = 1.6 + i * 0.37;
  s2.addText([
    { text: f[0], options: { bold: true, color: TEAL, fontSize: 8 } },
    { text: "  " + f[1], options: { color: BODY, fontSize: 8 } }
  ], { x: 0.6, y: yy, w: 3.9, h: 0.33, fontFace: FB, margin: 0 });
  if (i < fiche.length - 1) s2.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: yy + 0.32, w: 3.7, h: 0.005, fill: { color: "EEEEEE" } });
});

// Right — Timeline
card(s2, 5.1, 1.4, 4.5, 3.7, CARD, ORANGE);
s2.addText("PARCOURS CHRONOLOGIQUE", { x: 5.3, y: 1.55, w: 4.1, h: 0.28, fontSize: 9, bold: true, charSpacing: 2, color: ORANGE, fontFace: FB, margin: 0 });
// Vertical line
s2.addShape(pres.shapes.RECTANGLE, { x: 5.44, y: 2.0, w: 0.02, h: 3.0, fill: { color: "DDDDDD" } });
const tl = [
  ["2017\u20132022", "MOA publique \u2013 Maroc (3 ans)", "Conseil R\u00e9gional BMK. 7 march\u00e9s, +100 M DH"],
  ["2022\u20132024", "Ex\u00e9cution terrain \u2013 France", "Chef GO Ergalis + Chef chantier Minssieux"],
  ["2024\u20132025", "Formation BIM \u2013 AFPA Colmar", "Titre Modeleur BIM. B\u00e2timent R+2, 78 postes"],
  ["2026", "BIMCO + BTS MEC", "Micro-entreprise + candidat libre"]
];
tl.forEach((t, i) => {
  circle(s2, 5.32, 2.05 + i * 0.78, String(i + 1), 0.26);
  s2.addText([
    { text: t[0] + "  ", options: { bold: true, color: TEAL, fontSize: 8.5 } },
    { text: t[1], options: { bold: true, color: DARK, fontSize: 8.5, breakLine: true } },
    { text: t[2], options: { color: MUTED, fontSize: 7.5 } }
  ], { x: 5.7, y: 2.0 + i * 0.78, w: 3.7, h: 0.68, fontFace: FB, margin: 0 });
});

s2.addNotes(`SLIDE 2 - QUI SUIS-JE ? (1 min 30)
Mon parcours se d\u00e9compose en 4 phases.
Phase 1 : au Maroc, de 2017 \u00e0 2022, technicien \u00e9tudes et suivi au Conseil R\u00e9gional de B\u00e9ni Mellal. 3 ans d\u2019activit\u00e9 effective sur les march\u00e9s publics \u2013 7 march\u00e9s pour plus de 100 millions de dirhams.
Phase 2 : en France depuis 2022, chef d\u2019\u00e9quipe GO chez Ergalis \u00e0 Feurs puis chef de chantier chez Minssieux \u00e0 Mornant.
Phase 3 : formation BIM \u00e0 l\u2019AFPA Colmar en 2024-2025. Titre professionnel Modeleur BIM, 8 mois.
Phase 4 : cr\u00e9ation de BIMCO en janvier 2026 et candidature BTS MEC en candidat libre.
Au total : 8 ans dans le BTP, 3 ans Maroc + 5 ans France. Chaque phase a enrichi la suivante.
[Si le jury demande : les 5 ans France incluent l\u2019ex\u00e9cution terrain 2022-2024, la formation BIM 2024-2025 et BIMCO 2025-2026]`);

// ═══════════════════════════════════════════════════════════
// SLIDE 3 — STRUCTURE D'ACCUEIL
// ═══════════════════════════════════════════════════════════
let s3 = pres.addSlide();
topBar(s3); badge(s3, "01 \u2013 Cadre"); footer(s3, 3);
title(s3, "STRUCTURE D\u2019ACCUEIL", "Conseil R\u00e9gional de B\u00e9ni Mellal-Kh\u00e9nifra \u2013 Agence d\u2019Ex\u00e9cution des Projets");

// Left — Description
card(s3, 0.4, 1.4, 5.5, 3.7, CARD, TEAL);
s3.addText([
  { text: "Collectivit\u00e9 territoriale couvrant 5 provinces et 2,5 millions d\u2019habitants.", options: { breakLine: true, fontSize: 9 } },
  { text: "Mon poste : Technicien \u00e9tudes et suivi des travaux.", options: { breakLine: true, fontSize: 9, bold: true, color: TEAL } },
  { text: "", options: { breakLine: true, fontSize: 6 } },
  { text: "MISSIONS PRINCIPALES", options: { breakLine: true, fontSize: 9, bold: true, color: TEAL, charSpacing: 2 } },
], { x: 0.6, y: 1.6, w: 5.1, h: 0.9, fontFace: FB, color: BODY, margin: 0 });

const missions = [
  "M\u00e9tr\u00e9s avant-projet et estimation confidentielle",
  "R\u00e9daction des DCE : CPS, RC, BPDE",
  "Analyse des offres en Commission AO",
  "Suivi financier mensuel, d\u00e9comptes",
  "Visites terrain, attachements contradictoires"
];
s3.addText(missions.map((m, i) => ({
  text: m, options: { bullet: true, breakLine: i < missions.length - 1, fontSize: 9, color: BODY }
})), { x: 0.6, y: 2.5, w: 5.1, h: 2.0, fontFace: FB, margin: 0 });

// Right — Cadre r\u00e9glementaire
card(s3, 6.2, 1.4, 3.4, 3.7, TEAL_L, TEAL_D);
s3.addText("CADRE R\u00c9GLEMENTAIRE", { x: 6.4, y: 1.55, w: 3.0, h: 0.28, fontSize: 9, bold: true, charSpacing: 2, color: TEAL_D, fontFace: FB, margin: 0 });
const regl = [
  ["Pi\u00e8ces AO", "CPS + RC + BPDE"],
  ["Proc\u00e9dure", "Appel d\u2019offres ouvert"],
  ["Estimation", "Confidentielle obligatoire"],
  ["Normes", "Normes marocaines, RPS 2000"],
  ["Suivi", "Attachements contradictoires"]
];
regl.forEach((r, i) => {
  s3.addText([
    { text: r[0], options: { bold: true, color: TEAL_D, fontSize: 8, breakLine: true } },
    { text: r[1], options: { color: BODY, fontSize: 8 } }
  ], { x: 6.4, y: 2.0 + i * 0.55, w: 3.0, h: 0.5, fontFace: FB, margin: 0 });
});

s3.addNotes(`SLIDE 3 - STRUCTURE D'ACCUEIL (1 min 30)
Le Conseil R\u00e9gional de B\u00e9ni Mellal-Kh\u00e9nifra couvre 5 provinces, 2,5 millions d\u2019habitants.
J\u2019\u00e9tais rattach\u00e9 \u00e0 l\u2019Agence d\u2019Ex\u00e9cution des Projets, dirig\u00e9e par M. DOGHMANI.
Mon poste : technicien \u00e9tudes et suivi des travaux. Quand un march\u00e9 de 53 millions de dirhams devait \u00eatre lanc\u00e9, c\u2019est moi qui \u00e9tablissais l\u2019estimation confidentielle, r\u00e9digeais le DCE, puis suivais les travaux sur le terrain.
Le cadre r\u00e9glementaire marocain impose une estimation confidentielle obligatoire avant tout AO. C\u2019est cette pi\u00e8ce qui fixe le prix plafond.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 4 — BIMCO
// ═══════════════════════════════════════════════════════════
let s4 = pres.addSlide();
topBar(s4); badge(s4, "01 \u2013 BIMCO"); footer(s4, 4);
title(s4, "BIMCO \u2013 MON ACTIVIT\u00c9 IND\u00c9PENDANTE", "Cr\u00e9\u00e9e en janvier 2026 \u2013 BIM au service de l\u2019\u00e9conomiste de la construction");

// Left — Mission
card(s4, 0.4, 1.4, 5.0, 1.7, CARD, TEAL);
s4.addText("MISSION", { x: 0.6, y: 1.55, w: 4.6, h: 0.22, fontSize: 9, bold: true, color: TEAL, charSpacing: 2, fontFace: FB, margin: 0 });
s4.addText("Appliquer le BIM aux m\u00e9tiers de l\u2019\u00e9conomiste de la construction. M\u00e9tr\u00e9s par extraction de maquette num\u00e9rique, \u00e9tudes de prix ancr\u00e9es dans les co\u00fbts r\u00e9els, plugins Revit/Dynamo pour automatiser la cha\u00eene m\u00e8tre \u2192 DPGF.", {
  x: 0.6, y: 1.82, w: 4.6, h: 1.1, fontSize: 9, color: BODY, fontFace: FB, margin: 0
});

// Tech stack pills
const techs = ["Revit API", "C# .NET", "Dynamo", "Python", "IFC / BIM360", "React / Node.js"];
card(s4, 0.4, 3.3, 5.0, 1.65, CARD, ORANGE);
s4.addText("STACK TECHNIQUE", { x: 0.6, y: 3.45, w: 4.6, h: 0.22, fontSize: 9, bold: true, color: ORANGE, charSpacing: 2, fontFace: FB, margin: 0 });
techs.forEach((t, i) => {
  const col = i % 3;
  const row = Math.floor(i / 3);
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.6 + col * 1.55, y: 3.8 + row * 0.45, w: 1.4, h: 0.33, fill: { color: TEAL_L }, line: { color: TEAL, width: 0.5 }, rectRadius: 0.03 });
  s4.addText(t, { x: 0.6 + col * 1.55, y: 3.8 + row * 0.45, w: 1.4, h: 0.33, fontSize: 8, bold: true, color: TEAL_D, fontFace: FB, align: "center", valign: "middle", margin: 0 });
});

// Right — Positionnement
card(s4, 5.7, 1.4, 3.9, 3.55, NAVY);
s4.addText("POSITIONNEMENT", { x: 5.9, y: 1.55, w: 3.5, h: 0.28, fontSize: 10, bold: true, color: WHITE, charSpacing: 2, fontFace: FB, margin: 0 });
s4.addText("Triple comp\u00e9tence rare :", { x: 5.9, y: 1.88, w: 3.5, h: 0.25, fontSize: 9, color: TEAL, fontFace: FB, margin: 0 });
const positions = [
  "MOA publique (Maroc) \u2013 vision globale projet",
  "Ex\u00e9cution terrain (France) \u2013 r\u00e9alit\u00e9 chantier",
  "BIM \u2013 pont entre conception et chiffrage"
];
s4.addText(positions.map((p, i) => ({
  text: p, options: { bullet: true, breakLine: i < 2, fontSize: 9, color: WHITE }
})), { x: 5.9, y: 2.25, w: 3.5, h: 1.3, fontFace: FB, margin: 0 });

s4.addText([
  { text: "BIM", options: { bold: true, fontSize: 30, color: WHITE } },
  { text: " + ", options: { fontSize: 22, color: TEAL } },
  { text: "CO", options: { bold: true, fontSize: 30, color: WHITE } }
], { x: 5.9, y: 3.7, w: 3.5, h: 0.65, fontFace: FT, align: "center", margin: 0 });
s4.addText("Building Information Modeling + \u00c9conomie de la Construction", { x: 5.9, y: 4.35, w: 3.5, h: 0.3, fontSize: 7, color: TEAL, fontFace: FB, align: "center", margin: 0 });

s4.addNotes(`SLIDE 4 - BIMCO (1 min)
BIMCO est n\u00e9 d\u2019un constat : les outils BIM sont faits pour les architectes, pas pour celui qui chiffre.
J\u2019ai cr\u00e9\u00e9 BIMCO en janvier 2026 pour corriger ce manque.
La mission : extraire les m\u00e9tr\u00e9s directement de la maquette num\u00e9rique au lieu de les compter sur plan.
Mon positionnement est rare : triple comp\u00e9tence MOA publique, ex\u00e9cution terrain, et BIM.
C\u00f4t\u00e9 technique : Revit API, C#, Dynamo, Python pour les plugins. React et Node.js pour les applications web.
BIMCO est le prolongement direct de 8 ans d\u2019exp\u00e9rience.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 5 — PROJET 1 PRESENTATION + BUDGET
// ═══════════════════════════════════════════════════════════
let s5 = pres.addSlide();
topBar(s5); badge(s5, "02 \u2013 Projet 1"); footer(s5, 5);
title(s5, "MISE \u00c0 NIVEAU DE 4 COMMUNES", "March\u00e9 n\u00b038-RBK-2017 \u2013 53,5 M DH TTC \u2013 Province de Kh\u00e9nifra");

// Big stat left
s5.addText("53,5", { x: 0.4, y: 1.45, w: 2.5, h: 0.65, fontSize: 36, bold: true, color: TEAL, fontFace: FT, margin: 0 });
s5.addText("M DH TTC", { x: 2.8, y: 1.55, w: 1.5, h: 0.45, fontSize: 14, color: TEAL, fontFace: FB, margin: 0 });
s5.addText("4 communes  \u00b7  8 corps d\u2019\u00e9tat  \u00b7  18 mois", { x: 0.4, y: 2.1, w: 4.5, h: 0.25, fontSize: 9, color: MUTED, fontFace: FB, margin: 0 });

// Corps d'etat table
const corps = [
  ["01", "Assainissement", "22%"], ["02", "Chauss\u00e9e", "19%"],
  ["03", "Trottoirs", "14%"], ["04", "Signalisation", "6%"],
  ["05", "\u00c9clairage public", "16%"], ["06", "Murs & ouvrages", "14%"],
  ["07", "Paysager", "5%"], ["08", "Mobilier urbain", "4%"]
];
const tRows = [[
  { text: "Lot", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7.5 } },
  { text: "D\u00e9signation", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7.5 } },
  { text: "%", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7.5, align: "center" } }
]];
corps.forEach(c => tRows.push([
  { text: c[0], options: { fontSize: 7.5, color: TEAL, bold: true } },
  { text: c[1], options: { fontSize: 7.5, color: BODY } },
  { text: c[2], options: { fontSize: 7.5, color: BODY, align: "center" } }
]));
s5.addTable(tRows, { x: 0.4, y: 2.5, w: 4.5, colW: [0.5, 2.9, 0.6], border: { pt: 0.5, color: "E0E0E0" }, rowH: 0.27, autoPage: false });

// Right — Pie chart
s5.addChart(pres.charts.PIE, [{
  name: "Budget", labels: ["Assainissement", "Chauss\u00e9e", "\u00c9clairage", "Trottoirs", "Murs", "Signalisation", "Paysager", "Mobilier"],
  values: [22, 19, 16, 14, 14, 6, 5, 4]
}], {
  x: 5.2, y: 1.4, w: 4.4, h: 3.8,
  showPercent: true, showTitle: false, showLegend: true, legendPos: "b", legendFontSize: 7,
  chartColors: [TEAL, "00A3A8", "4DB8BD", "80CED1", "B3E4E6", TEAL_L, ORANGE, "F5DEB3"],
  dataLabelColor: WHITE, dataLabelFontSize: 8
});

s5.addNotes(`SLIDE 5 - PROJET 1 (1 min)
Le Projet 1 : mise \u00e0 niveau de 4 communes de la province de Kh\u00e9nifra.
53,5 millions de dirhams TTC, soit environ 4,8 millions d\u2019euros. March\u00e9 unique couvrant 8 corps d\u2019\u00e9tat, de l\u2019assainissement au mobilier urbain, sur 4 sites distants de 20 \u00e0 80 km.
18 mois de suivi.
L\u2019assainissement et la chauss\u00e9e concentrent 41% du budget \u00e0 eux seuls.
C\u2019est sur ce projet que se d\u00e9roulent les 4 premi\u00e8res situations de mon rapport.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 6 — SITUATIONS 1 & 2
// ═══════════════════════════════════════════════════════════
let s6 = pres.addSlide();
topBar(s6); badge(s6, "02 \u2013 Sit. 1 & 2"); footer(s6, 6);
title(s6, "ESTIMATION CONFIDENTIELLE & ANALYSE DES OFFRES", "Comp\u00e9tence C18 \u2013 M\u00e9trer, estimer, analyser les offres");

// Sit 1 — Left
card(s6, 0.4, 1.4, 4.5, 3.7, CARD, TEAL);
s6.addText("SITUATION 1 \u2013 ESTIMATION CONFIDENTIELLE", { x: 0.6, y: 1.55, w: 3.3, h: 0.28, fontSize: 8, bold: true, color: TEAL, charSpacing: 1, fontFace: FB, margin: 0 });
s6.addText("Ouaoumana \u2013 15,8 M DH HT", { x: 0.6, y: 1.85, w: 3.0, h: 0.22, fontSize: 9, bold: true, color: DARK, fontFace: FB, margin: 0 });
s6.addText("3,2%", { x: 3.5, y: 1.5, w: 1.2, h: 0.55, fontSize: 28, bold: true, color: GREEN, fontFace: FT, align: "right", margin: 0 });

cparBlock(s6, [
  ["C", "Prix plafond avant AO. 35% du budget. D\u00e9lai 3 semaines."],
  ["P", "Mercuriale 2014 obsol\u00e8te : prix d\u00e9riv\u00e9s de 15 \u00e0 22%."],
  ["A", "112 lignes AutoCAD + 4 visites terrain + 3 sources crois\u00e9es."],
  ["R", "\u00c9cart 3,2% vs 5-10% norme. M\u00e9thode adopt\u00e9e standard."]
], 0.6, 2.2, 4.1);

// Sit 2 — Right
card(s6, 5.1, 1.4, 4.5, 3.7, CARD, ORANGE);
s6.addText("SITUATION 2 \u2013 ANALYSE DES OFFRES", { x: 5.3, y: 1.55, w: 3.3, h: 0.28, fontSize: 8, bold: true, color: ORANGE, charSpacing: 1, fontFace: FB, margin: 0 });
s6.addText("Commission CAO \u2013 3 offres", { x: 5.3, y: 1.85, w: 3.0, h: 0.22, fontSize: 9, bold: true, color: DARK, fontFace: FB, margin: 0 });
s6.addText("94/100", { x: 8.1, y: 1.5, w: 1.3, h: 0.55, fontSize: 28, bold: true, color: ORANGE, fontFace: FT, align: "right", margin: 0 });

cparBlock(s6, [
  ["C", "Membre technique commission. Analyse comparative 3 dossiers."],
  ["P", "7 erreurs arithm\u00e9tiques + prix bas anormaux sur 42% du montant."],
  ["A", "Grille 100 pts (technique 60 + financier 40). Justification \u00e9crite."],
  ["R", "Attribution 15 jours. Z\u00e9ro recours. Rapport valid\u00e9 sans r\u00e9serve."]
], 5.3, 2.2, 4.1);

s6.addNotes(`SLIDE 6 - SITUATIONS 1 & 2 (2 min)
SITUATION 1 \u2013 L\u2019estimation confidentielle d\u2019Ouaoumana.
15,8 millions de dirhams, 35% du budget global. Le probl\u00e8me : la mercuriale de r\u00e9f\u00e9rence datait de 2014. Les prix avaient d\u00e9riv\u00e9 de 15 \u00e0 22%.
J\u2019ai crois\u00e9 3 sources : mercuriale actualis\u00e9e, 5 march\u00e9s similaires, 8 devis fournisseurs.
R\u00e9sultat : 3,2% d\u2019\u00e9cart. La norme c\u2019est 5 \u00e0 10%.

SITUATION 2 \u2013 L\u2019analyse des offres en commission.
3 soumissionnaires. Un dossier avec 7 erreurs arithm\u00e9tiques. Un autre avec des prix anormalement bas sur 42% du montant.
Grille de notation sur 100 : 60 technique, 40 financier.
R\u00e9sultat : 94/100, attribution en 15 jours, z\u00e9ro recours.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 7 — SITUATIONS 3 & 4
// ═══════════════════════════════════════════════════════════
let s7 = pres.addSlide();
topBar(s7); badge(s7, "02 \u2013 Sit. 3 & 4"); footer(s7, 7);
title(s7, "SUIVI FINANCIER & COORDINATION MULTI-SITES", "Comp\u00e9tences C18 (suivi budget) et C19 (coordonner, communiquer)");

// Sit 3 — Left
card(s7, 0.4, 1.4, 4.5, 3.7, CARD, TEAL);
s7.addText("SITUATION 3 \u2013 SUIVI FINANCIER KERROUCHEN", { x: 0.6, y: 1.55, w: 3.3, h: 0.28, fontSize: 8, bold: true, color: TEAL, charSpacing: 1, fontFace: FB, margin: 0 });
s7.addText("7,3 M DH \u2013 18 mois", { x: 0.6, y: 1.85, w: 2.5, h: 0.22, fontSize: 9, bold: true, color: DARK, fontFace: FB, margin: 0 });
s7.addText("+0,8%", { x: 3.5, y: 1.5, w: 1.2, h: 0.55, fontSize: 28, bold: true, color: GREEN, fontFace: FT, align: "right", margin: 0 });

cparBlock(s7, [
  ["C", "Suivi mensuel complet : situations, quantit\u00e9s, d\u00e9comptes."],
  ["P", "Chauss\u00e9e +12%, murs +15%. D\u00e9rive +4,8% (seuil avenant = 5%)."],
  ["A", "TdB hebdo 3 indicateurs. Compensation paysager \u221244k + mobilier \u221212k."],
  ["R", "D\u00e9passement +0,8%. Aucun avenant. TdB adopt\u00e9 3 communes."]
], 0.6, 2.2, 4.1);

// Sit 4 — Right
card(s7, 5.1, 1.4, 4.5, 3.7, CARD, ORANGE);
s7.addText("SITUATION 4 \u2013 COORDINATION 4 CHANTIERS", { x: 5.3, y: 1.55, w: 3.3, h: 0.28, fontSize: 8, bold: true, color: ORANGE, charSpacing: 1, fontFace: FB, margin: 0 });
s7.addText("Province de Kh\u00e9nifra \u2013 20 \u00e0 80 km", { x: 5.3, y: 1.85, w: 3.0, h: 0.22, fontSize: 9, bold: true, color: DARK, fontFace: FB, margin: 0 });
s7.addText("48h", { x: 8.1, y: 1.5, w: 1.3, h: 0.55, fontSize: 28, bold: true, color: ORANGE, fontFace: FT, align: "right", margin: 0 });

cparBlock(s7, [
  ["C", "Relais unique Directeur. Interface entreprise/BET/labo/hi\u00e9rarchie."],
  ["P", "S23 : retard enrob\u00e9s + alerte m\u00e9t\u00e9o + litige bordures 15%."],
  ["A", "Note Directeur + OS arr\u00eat photos + re-mesurage contradictoire."],
  ["R", "3 crises r\u00e9solues. Litige 15% \u2192 2,3%. Mod\u00e8le CR adopt\u00e9."]
], 5.3, 2.2, 4.1);

s7.addNotes(`SLIDE 7 - SITUATIONS 3 & 4 (2 min 30)
SITUATION 3 \u2013 Le suivi financier de Kerrouchen. C\u2019est le c\u0153ur de mon rapport.
7,3 M DH sur 18 mois. \u00c0 mi-parcours, chauss\u00e9e +12%, murs +15%. Extrapolation : +4,8%. Seuil avenant = 5%.
TdB hebdomadaire \u00e0 3 indicateurs. Compensation paysager \u221244k, mobilier \u221212k.
R\u00e9sultat : +0,8% final. Aucun avenant. TdB r\u00e9pliqu\u00e9 sur 3 communes.

SITUATION 4 \u2013 La semaine 23. 3 crises simultan\u00e9es. Retard enrob\u00e9s, alerte m\u00e9t\u00e9o, litige bordures 15%.
Trait\u00e9 en 48h. Litige tomb\u00e9 de 15% \u00e0 2,3%. CR devenu standard de l\u2019Agence.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 8 — PROJET 2 + BUDGET
// ═══════════════════════════════════════════════════════════
let s8 = pres.addSlide();
topBar(s8); badge(s8, "03 \u2013 Projet 2"); footer(s8, 8);
title(s8, "ROUTE LEHRI-KERROUCHEN \u2013 25 KM", "Programme PRR3 \u2013 29 M DH TTC \u2013 Zone montagneuse du Moyen Atlas");

// Stats row
const p2Stats = [
  ["29 M DH", "Budget TTC", TEAL], ["25 km", "Zone montagneuse", TEAL_D], ["120 334 m\u00b3", "D\u00e9blais v\u00e9rifi\u00e9s", ORANGE], ["53 prix", "Au bordereau", NAVY]
];
p2Stats.forEach((st, i) => {
  card(s8, 0.4 + i * 2.35, 1.4, 2.15, 0.95, CARD, st[2]);
  s8.addText(st[0], { x: 0.4 + i * 2.35, y: 1.5, w: 2.15, h: 0.45, fontSize: 18, bold: true, color: st[2], fontFace: FT, align: "center", valign: "middle", margin: 0 });
  s8.addText(st[1], { x: 0.4 + i * 2.35, y: 1.95, w: 2.15, h: 0.28, fontSize: 8, color: MUTED, fontFace: FB, align: "center", margin: 0 });
});

// Bar chart
s8.addChart(pres.charts.BAR, [{
  name: "Budget", labels: ["Corps chauss\u00e9e", "Ouv. hydrauliques", "Terrassement", "Rev\u00eatement", "Bretelles", "Sout\u00e8nement"],
  values: [30.1, 28.1, 17.9, 10.7, 11.3, 1.9]
}], {
  x: 0.4, y: 2.55, w: 9.2, h: 2.6, barDir: "col",
  chartColors: [TEAL], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: DARK, dataLabelFontSize: 8,
  catAxisLabelColor: MUTED, catAxisLabelFontSize: 7, valAxisHidden: true,
  valGridLine: { style: "none" }, catGridLine: { style: "none" },
  showLegend: false, showTitle: false
});

s8.addNotes(`SLIDE 8 - PROJET 2 (1 min)
Route Lehri-Kerrouchen. 25 km en zone montagneuse du Moyen Atlas.
29 M DH TTC, programme PRR3. D\u00e9nivel\u00e9 400 m, pentes \u00e0 12% en lacets. 53 prix au bordereau.
Le poste dominant : corps de chauss\u00e9e \u00e0 30%, suivi des ouvrages hydrauliques \u00e0 28%. Ce ratio drainage/chauss\u00e9e d\u00e9passe de 2,5x celui d\u2019une route en plaine.
120 334 m\u00b3 de d\u00e9blais \u00e0 v\u00e9rifier. C\u2019est l\u2019objet de la situation 5.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 9 — SITUATION 5
// ═══════════════════════════════════════════════════════════
let s9 = pres.addSlide();
topBar(s9); badge(s9, "03 \u2013 Sit. 5"); footer(s9, 9);
title(s9, "CONTR\u00d4LE DES CUBATURES DE TERRASSEMENT", "Comp\u00e9tence C18 \u2013 V\u00e9rifier les quantit\u00e9s ex\u00e9cut\u00e9es");

// Two big stats
bigStat(s9, 0.4, 1.45, "+8%", "\u00c9cart cumul\u00e9 d\u00e9tect\u00e9 et corrig\u00e9", ORANGE);
bigStat(s9, 3.2, 1.45, "285 000 DH", "Surco\u00fbt absorb\u00e9 sans avenant", TEAL);

// CPAR card
card(s9, 0.4, 2.5, 9.2, 2.7, CARD, TEAL_D);
cparBlock(s9, [
  ["C", "V\u00e9rification 120 334 m\u00b3 d\u00e9blais + 76 735 m\u00b3 remblais. D\u00e9nivel\u00e9 400 m, pentes 12%."],
  ["P", "PK 12 : calcaire fractur\u00e9 non d\u00e9tect\u00e9. 5 000 m\u00b3 \u00e0 reclassifier. \u00c9cart cumul\u00e9 +8%."],
  ["A", "Profils en travers /25 m. Relev\u00e9s GPS contradictoires /500 ml. 3 dalots dimensionn\u00e9s avec BET."],
  ["R", "Surco\u00fbt absorb\u00e9 par compensation. R\u00e9ception sans avenant. Tra\u00e7abilit\u00e9 compl\u00e8te."]
], 0.6, 2.7, 8.8);

s9.addNotes(`SLIDE 9 - SITUATION 5 (1 min 30)
Au PK 12, surprise : calcaire fractur\u00e9 que l\u2019\u00e9tude g\u00e9otechnique n\u2019avait pas d\u00e9tect\u00e9.
5 000 m\u00b3 \u00e0 reclassifier en d\u00e9blais rocheux. Co\u00fbt 28 \u2192 85 DH/m\u00b3. Surco\u00fbt 285 000 DH. \u00c9cart cumul\u00e9 +8%.
R\u00e9ponse : profils en travers /25m, relev\u00e9s GPS contradictoires /500ml. 3 dalots BA avec NOVEC.
R\u00e9sultat : surco\u00fbt absorb\u00e9 par compensation. R\u00e9ception sans avenant.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 10 — ACTIVITES COMPLEMENTAIRES
// ═══════════════════════════════════════════════════════════
let s10 = pres.addSlide();
topBar(s10); badge(s10, "03 \u2013 Compl."); footer(s10, 10);
title(s10, "ACTIVIT\u00c9S COMPL\u00c9MENTAIRES", "5 autres march\u00e9s publics + exp\u00e9rience terrain en France");

// Left — 5 march\u00e9s
card(s10, 0.4, 1.4, 5.2, 2.8, CARD, TEAL);
s10.addText("5 AUTRES MARCH\u00c9S", { x: 0.6, y: 1.55, w: 4.8, h: 0.22, fontSize: 9, bold: true, color: TEAL, charSpacing: 1, fontFace: FB, margin: 0 });
const marches = [
  ["27-RBK", "Route Sidi Bouabbad \u2192 Oued Grou (12 km)", "8,2 M DH"],
  ["28-RBK", "Route Ajdir-Ayoun + Piste Lijon", "6,5 M DH"],
  ["30-RBK", "AEP El Borj \u2013 El Hamam (18 km)", "4,8 M DH"],
  ["39-RBK", "Pistes Hartaf \u2013 Sebt Ait Rahou", "3,1 M DH"],
  ["49-RBK", "Voirie Amghass \u2013 Bouchbel", "5,9 M DH"]
];
const mRows = [[
  { text: "N\u00b0", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7 } },
  { text: "Objet", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7 } },
  { text: "Montant", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7, align: "right" } }
]];
marches.forEach(m => mRows.push([
  { text: m[0], options: { fontSize: 7, color: TEAL_D, bold: true } },
  { text: m[1], options: { fontSize: 7, color: BODY } },
  { text: m[2], options: { fontSize: 7, color: BODY, align: "right" } }
]));
s10.addTable(mRows, { x: 0.6, y: 1.9, w: 4.8, colW: [0.7, 3.1, 1.0], border: { pt: 0.5, color: "E0E0E0" }, rowH: 0.26 });

// Right — France
card(s10, 5.9, 1.4, 3.7, 2.8, CARD, ORANGE);
s10.addText("EXP\u00c9RIENCE TERRAIN FRANCE", { x: 6.1, y: 1.55, w: 3.3, h: 0.22, fontSize: 9, bold: true, color: ORANGE, charSpacing: 1, fontFace: FB, margin: 0 });
s10.addText([
  { text: "Chef d\u2019\u00e9quipe GO \u2013 Ergalis BTP", options: { bold: true, breakLine: true, fontSize: 9, color: DARK } },
  { text: "Feurs (Loire) : banches, armatures HA, coulage", options: { breakLine: true, fontSize: 8, color: BODY } },
  { text: "", options: { breakLine: true, fontSize: 5 } },
  { text: "Chef de chantier \u2013 Minssieux & Fils", options: { bold: true, breakLine: true, fontSize: 9, color: DARK } },
  { text: "Mornant (Rh\u00f4ne) : planning, contr\u00f4le qualit\u00e9", options: { breakLine: true, fontSize: 8, color: BODY } },
  { text: "", options: { breakLine: true, fontSize: 5 } },
  { text: "Apport : ", options: { bold: true, color: TEAL, fontSize: 8 } },
  { text: "co\u00fbts r\u00e9els, rendements, contraintes terrain", options: { fontSize: 8, color: BODY } }
], { x: 6.1, y: 1.9, w: 3.3, h: 2.1, fontFace: FB, margin: 0 });

// Bottom banner
s10.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.4, w: 9.2, h: 0.55, fill: { color: TEAL_L }, rectRadius: 0.04 });
s10.addText("4 types d\u2019infrastructures : routes, VRD, adduction d\u2019eau potable, voirie urbaine", {
  x: 0.6, y: 4.4, w: 8.8, h: 0.55, fontSize: 9, italic: true, color: TEAL_D, fontFace: FB, align: "center", valign: "middle", margin: 0
});

s10.addNotes(`SLIDE 10 - ACTIVIT\u00c9S COMPL\u00c9MENTAIRES (1 min)
5 autres march\u00e9s couvrant routes rurales, pistes, AEP et voirie urbaine.
C\u00f4t\u00e9 France : chef d\u2019\u00e9quipe GO chez Ergalis, puis chef de chantier chez Minssieux.
L\u2019apport : comprendre les co\u00fbts r\u00e9els de production. Conna\u00eetre le chantier de l\u2019int\u00e9rieur change la fa\u00e7on d\u2019estimer.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 11 — SYNTHESE DES COMPETENCES
// ═══════════════════════════════════════════════════════════
let s11 = pres.addSlide();
topBar(s11); badge(s11, "04 \u2013 Bilan"); footer(s11, 11);
title(s11, "SYNTH\u00c8SE DES COMP\u00c9TENCES BTS MEC", "Comp\u00e9tences mobilis\u00e9es sur 5 situations professionnelles");

const cRows = [
  [
    { text: "Activit\u00e9", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7.5 } },
    { text: "Sous-comp\u00e9tence", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7.5 } },
    { text: "Sit.", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7.5, align: "center" } },
    { text: "Niveau", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7.5, align: "center" } }
  ],
  [{ text: "M\u00e9tr\u00e9s 112 prix, 4 communes", options: { fontSize: 7.5 } }, { text: "R\u00e9aliser des m\u00e9tr\u00e9s TCA", options: { fontSize: 7.5 } }, { text: "1", options: { fontSize: 7.5, align: "center" } }, { text: "\u25cf Ma\u00eetrise", options: { fontSize: 7.5, align: "center", color: TEAL, bold: true } }],
  [{ text: "Estimation 15,8 M DH, 3 sources", options: { fontSize: 7.5 } }, { text: "Estimer un ouvrage en AO", options: { fontSize: 7.5 } }, { text: "1", options: { fontSize: 7.5, align: "center" } }, { text: "\u25cf Ma\u00eetrise", options: { fontSize: 7.5, align: "center", color: TEAL, bold: true } }],
  [{ text: "Analyse 3 offres, grille /100", options: { fontSize: 7.5 } }, { text: "Analyser les offres en commission", options: { fontSize: 7.5 } }, { text: "2", options: { fontSize: 7.5, align: "center" } }, { text: "\u25cf Ma\u00eetrise", options: { fontSize: 7.5, align: "center", color: TEAL, bold: true } }],
  [{ text: "TdB hebdo, compensation", options: { fontSize: 7.5 } }, { text: "Suivre l\u2019ex\u00e9cution financi\u00e8re", options: { fontSize: 7.5 } }, { text: "3", options: { fontSize: 7.5, align: "center" } }, { text: "\u2605 Expert", options: { fontSize: 7.5, align: "center", color: GREEN, bold: true } }],
  [{ text: "Contr\u00f4le cubatures GPS", options: { fontSize: 7.5 } }, { text: "V\u00e9rifier les quantit\u00e9s", options: { fontSize: 7.5 } }, { text: "5", options: { fontSize: 7.5, align: "center" } }, { text: "\u2605 Expert", options: { fontSize: 7.5, align: "center", color: GREEN, bold: true } }],
  [{ text: "CR, OS, notes factuelles", options: { fontSize: 7.5 } }, { text: "Communiquer par \u00e9crit", options: { fontSize: 7.5 } }, { text: "4", options: { fontSize: 7.5, align: "center" } }, { text: "\u25cf Ma\u00eetrise", options: { fontSize: 7.5, align: "center", color: TEAL, bold: true } }],
  [{ text: "R\u00e9unions, points quotidiens, CAO", options: { fontSize: 7.5 } }, { text: "Communiquer oralement", options: { fontSize: 7.5 } }, { text: "2,4", options: { fontSize: 7.5, align: "center" } }, { text: "\u25cf Ma\u00eetrise", options: { fontSize: 7.5, align: "center", color: TEAL, bold: true } }],
  [{ text: "Convention BIM, export IFC, 78 postes", options: { fontSize: 7.5 } }, { text: "Collaborer en BIM (Open BIM)", options: { fontSize: 7.5 } }, { text: "BIM", options: { fontSize: 7.5, align: "center" } }, { text: "\u25cf Ma\u00eetrise", options: { fontSize: 7.5, align: "center", color: TEAL, bold: true } }],
];
s11.addTable(cRows, { x: 0.4, y: 1.35, w: 9.2, colW: [2.9, 2.8, 0.6, 1.0], border: { pt: 0.5, color: "E0E0E0" }, rowH: 0.33, autoPage: false });

// Bottom summary boxes instead of radar
const summBoxes = [
  ["8", "sous-comp\u00e9tences", TEAL], ["2", "niveau Expert", GREEN], ["5", "situations r\u00e9elles", ORANGE], ["100%", "sous contrainte", NAVY]
];
summBoxes.forEach((b, i) => {
  const bx = 0.4 + i * 2.35;
  card(s11, bx, 4.5, 2.15, 0.65, CARD, b[2]);
  s11.addText(b[0], { x: bx, y: 4.52, w: 0.7, h: 0.55, fontSize: 22, bold: true, color: b[2], fontFace: FT, align: "center", valign: "middle", margin: 0 });
  s11.addText(b[1], { x: bx + 0.7, y: 4.52, w: 1.35, h: 0.55, fontSize: 9, color: BODY, fontFace: FB, valign: "middle", margin: 0 });
});

s11.addNotes(`SLIDE 11 - SYNTH\u00c8SE DES COMP\u00c9TENCES (1 min)
8 sous-comp\u00e9tences du BTS MEC mobilis\u00e9es, toutes en situation r\u00e9elle sous contrainte.
2 au niveau Expert : suivi financier (TdB devenu r\u00e9f\u00e9rence) et contr\u00f4le quantit\u00e9s (m\u00e9thode GPS contradictoire syst\u00e9matis\u00e9e).
Le reste au niveau Ma\u00eetrise \u2013 pratique r\u00e9guli\u00e8re et autonome.
D\u00e9tail page 23 du rapport.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 12 — ANALYSE MAROC / FRANCE
// ═══════════════════════════════════════════════════════════
let s12 = pres.addSlide();
topBar(s12); badge(s12, "04 \u2013 Analyse"); footer(s12, 12);
title(s12, "ANALYSE COMPARATIVE MAROC / FRANCE", "M\u00eames fondamentaux \u2013 deux cadres r\u00e9glementaires \u2013 une seule exigence : la rigueur");

const cmpRows = [
  [
    { text: "Aspect", options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 8 } },
    { text: "\ud83c\uddf2\ud83c\udde6  Maroc", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 8, align: "center" } },
    { text: "\ud83c\uddeb\ud83c\uddf7  France", options: { bold: true, color: WHITE, fill: { color: ORANGE }, fontSize: 8, align: "center" } }
  ],
  [{ text: "R\u00e9glementation", options: { fontSize: 8, bold: true } }, { text: "D\u00e9cret n\u00b02-12-349", options: { fontSize: 8 } }, { text: "Code commande publique", options: { fontSize: 8 } }],
  [{ text: "Pi\u00e8ces march\u00e9", options: { fontSize: 8, bold: true } }, { text: "CPS + RC + BPDE", options: { fontSize: 8 } }, { text: "CCAP + CCTP + BPU/DQE", options: { fontSize: 8 } }],
  [{ text: "Estimation", options: { fontSize: 8, bold: true } }, { text: "Confidentielle obligatoire", options: { fontSize: 8 } }, { text: "Estimation MOA", options: { fontSize: 8 } }],
  [{ text: "Normes", options: { fontSize: 8, bold: true } }, { text: "Normes marocaines, RPS 2000", options: { fontSize: 8 } }, { text: "DTU, Eurocodes, RE2020", options: { fontSize: 8 } }],
  [{ text: "Suivi financier", options: { fontSize: 8, bold: true } }, { text: "Attachements contradictoires", options: { fontSize: 8 } }, { text: "Situations mensuelles", options: { fontSize: 8 } }],
  [{ text: "Commission", options: { fontSize: 8, bold: true } }, { text: "CAO", options: { fontSize: 8 } }, { text: "Commission d\u2019AO", options: { fontSize: 8 } }],
];
s12.addTable(cmpRows, { x: 0.4, y: 1.35, w: 9.2, colW: [1.8, 3.7, 3.7], border: { pt: 0.5, color: "E0E0E0" }, rowH: 0.36 });

// Bottom cards
card(s12, 0.4, 3.95, 4.3, 1.15, CARD, TEAL);
s12.addText([
  { text: "C\u00f4t\u00e9 MOA \u2013 Maroc", options: { bold: true, color: TEAL, fontSize: 9, breakLine: true } },
  { text: "Concevoir les march\u00e9s, r\u00e9diger CPS/RC/BPDE, piloter la CAO. 7 march\u00e9s, 100 M DH.", options: { fontSize: 8, color: BODY } }
], { x: 0.6, y: 4.05, w: 3.9, h: 0.95, fontFace: FB, margin: 0 });

card(s12, 5.3, 3.95, 4.3, 1.15, CARD, ORANGE);
s12.addText([
  { text: "C\u00f4t\u00e9 ex\u00e9cution \u2013 France", options: { bold: true, color: ORANGE, fontSize: 9, breakLine: true } },
  { text: "Banches, ferraillage, planning, contr\u00f4le qualit\u00e9. Co\u00fbts r\u00e9els et rendements.", options: { fontSize: 8, color: BODY } }
], { x: 5.5, y: 4.05, w: 3.9, h: 0.95, fontFace: FB, margin: 0 });

s12.addNotes(`SLIDE 12 - ANALYSE MAROC / FRANCE (1 min 30)
Les deux syst\u00e8mes partagent les m\u00eames principes : transparence, \u00e9galit\u00e9, mise en concurrence.
CPS au Maroc, CCAP en France \u2013 les noms changent, la logique est la m\u00eame.
L\u2019estimation confidentielle au Maroc est obligatoire. En France, m\u00eame m\u00e9canisme avec l\u2019estimation MOA.
La combinaison des deux rend mes estimations r\u00e9alistes \u2013 ni trop basses, ni trop hautes.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 13 — BILAN REFLEXIF ADM
// ═══════════════════════════════════════════════════════════
let s13 = pres.addSlide();
topBar(s13); badge(s13, "04 \u2013 R\u00e9flexif"); footer(s13, 13);
title(s13, "BILAN R\u00c9FLEXIF", "Ce que le terrain m\u2019a appris \u2013 8 ans de pratique");

const adm = [
  ["A", "CE QUE J\u2019AI APPRIS", "L\u2019\u00e9crit syst\u00e9matique \u2013 OS, attachements, CR consolid\u00e9s \u2013 est le seul rempart contre les litiges. Croiser 3 sources de prix donne un \u00e9cart de 3,2% l\u00e0 o\u00f9 les estimations r\u00e9gionales d\u00e9passaient 10%. Voir le projet depuis le terrain rend les estimations plus justes.", TEAL],
  ["D", "CE QUE JE FERAIS DIFF\u00c9REMMENT", "D\u00e9ployer le TdB d\u00e8s le 1er mois, pas au 9\u00e8me quand le signal \u00e9tait d\u00e9j\u00e0 \u00e0 +4,8%. Insister pour une \u00e9tude g\u00e9otechnique compl\u00e9mentaire avant terrassements. Standardiser les CR d\u00e8s le d\u00e9marrage \u2013 pas dans l\u2019urgence de la semaine 23.", ORANGE],
  ["M", "CE QUE J\u2019APPORTE AU BTS MEC", "L\u2019estimation confidentielle est l\u2019acte fondateur de tout march\u00e9 public. En France, m\u00eames fondamentaux avec le DQE/DPGF. Triple comp\u00e9tence MOA + Ex\u00e9cution + BIM = positionnement rare. BIMCO est n\u00e9 de ce parcours.", NAVY]
];
adm.forEach((item, i) => {
  const y = 1.35 + i * 1.3;
  card(s13, 0.4, y, 9.2, 1.12, CARD, item[3]);
  circle(s13, 0.55, y + 0.12, item[0], 0.34, item[3]);
  s13.addText(item[1], { x: 1.0, y: y + 0.1, w: 8.3, h: 0.28, fontSize: 9, bold: true, color: item[3], charSpacing: 1, fontFace: FB, margin: 0 });
  s13.addText(item[2], { x: 1.0, y: y + 0.4, w: 8.3, h: 0.65, fontSize: 8.5, color: BODY, fontFace: FB, margin: 0 });
});

s13.addNotes(`SLIDE 13 - BILAN R\u00c9FLEXIF (2 min) \u2013 SLIDE IMPORTANTE POUR LE JURY
Structure ADM identique au rapport papier page 26.

A \u2013 CE QUE J\u2019AI APPRIS :
La semaine 23 : l\u2019\u00e9crit syst\u00e9matique est le seul rempart contre les litiges.
Croiser 3 sources de prix = \u00e9cart 3,2% vs 10% de norme r\u00e9gionale.

D \u2013 CE QUE JE FERAIS DIFF\u00c9REMMENT :
TdB d\u00e8s le 1er mois. \u00c9tude g\u00e9otech compl\u00e9mentaire. CR standardis\u00e9s d\u00e8s le d\u00e9marrage.

M \u2013 CE QUE J\u2019APPORTE AU BTS MEC :
Triple comp\u00e9tence MOA + Ex\u00e9cution + BIM. BIMCO est n\u00e9 de ce parcours.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 14 — PROTOCOLE BIM + PERSPECTIVES
// ═══════════════════════════════════════════════════════════
let s14 = pres.addSlide();
topBar(s14); badge(s14, "05 \u2013 BIM"); footer(s14, 14);
title(s14, "PROTOCOLE BIM & PERSPECTIVES BIMCO", "Formation AFPA Colmar + Projet professionnel 2026\u20132029");

// Left — BIM Protocol
card(s14, 0.4, 1.35, 4.5, 3.75, CARD, TEAL);
s14.addText("PROTOCOLE BIM \u2013 CAS AFPA", { x: 0.6, y: 1.5, w: 4.1, h: 0.22, fontSize: 9, bold: true, color: TEAL, charSpacing: 1, fontFace: FB, margin: 0 });

const bimStats = [["78", "postes"], ["1,8%", "\u00e9cart"], ["12", "clashs"]];
bimStats.forEach((bs, i) => {
  s14.addText(bs[0], { x: 0.6 + i * 1.4, y: 1.82, w: 1.25, h: 0.4, fontSize: 20, bold: true, color: TEAL, fontFace: FT, align: "center", margin: 0 });
  s14.addText(bs[1], { x: 0.6 + i * 1.4, y: 2.22, w: 1.25, h: 0.2, fontSize: 7, color: MUTED, fontFace: FB, align: "center", margin: 0 });
});

const bimSteps = [
  "Mod\u00e9lisation Revit Architecture + Structure (LOD 300)",
  "Export IFC 2x3 \u2013 Open BIM \u2013 MVD Coordination View",
  "D\u00e9tection clashs Navisworks (12 conflits r\u00e9solus)",
  "Extraction quantit\u00e9s Revit + Dynamo (2h vs 2 jours)",
  "Chiffrage avec tra\u00e7abilit\u00e9 maquette"
];
bimSteps.forEach((step, i) => {
  circle(s14, 0.6, 2.6 + i * 0.42, String(i + 1), 0.2);
  s14.addText(step, { x: 0.88, y: 2.58 + i * 0.42, w: 3.85, h: 0.36, fontSize: 7.5, color: BODY, fontFace: FB, margin: 0 });
});

// Right — Perspectives
card(s14, 5.1, 1.35, 4.5, 3.75, CARD, ORANGE);
s14.addText("PROJET PROFESSIONNEL", { x: 5.3, y: 1.5, w: 4.1, h: 0.22, fontSize: 9, bold: true, color: ORANGE, charSpacing: 1, fontFace: FB, margin: 0 });

const horizons = [
  ["2026", "Court terme", "BTS MEC + premi\u00e8res prestations BIMCO : m\u00e9tr\u00e9s BIM, \u00e9tudes de prix, DPGF. Premiers plugins Revit/Dynamo."],
  ["2027\u201328", "Moyen terme", "Plugin Revit \u2192 DPGF automatis\u00e9. App web suivi \u00e9conomique. Base prix Batiprix. 3\u20135 clients."],
  ["2029+", "Long terme", "Cabinet ing\u00e9nierie BIM + \u00e9co construction. SaaS MEC. \u00c9quipe 3\u20135 pers. CA 200\u2013300 k\u20ac/an."]
];
horizons.forEach((h, i) => {
  s14.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.85 + i * 1.08, w: 0.8, h: 0.28, fill: { color: i === 0 ? TEAL : i === 1 ? ORANGE : NAVY }, rectRadius: 0.03 });
  s14.addText(h[0], { x: 5.3, y: 1.85 + i * 1.08, w: 0.8, h: 0.28, fontSize: 8, bold: true, color: WHITE, fontFace: FB, align: "center", valign: "middle", margin: 0 });
  s14.addText(h[1], { x: 6.2, y: 1.85 + i * 1.08, w: 3.2, h: 0.28, fontSize: 9, bold: true, color: DARK, fontFace: FB, margin: 0 });
  s14.addText(h[2], { x: 5.3, y: 2.18 + i * 1.08, w: 4.1, h: 0.75, fontSize: 7.5, color: BODY, fontFace: FB, margin: 0 });
});

s14.addNotes(`SLIDE 14 - PROTOCOLE BIM & PERSPECTIVES (1 min 30)
\u00c0 gauche : protocole BIM pratiqu\u00e9 \u00e0 l\u2019AFPA Colmar.
B\u00e2timent R+2, LOD 300. 78 postes extraits en 2h vs 2 jours. \u00c9cart 1,8%. 12 clashs r\u00e9solus.
\u00c0 droite : projet professionnel sur 3 horizons.
2026 : BTS MEC + d\u00e9but BIMCO.
2027-28 : plugin Revit \u2192 DPGF, app web, 3-5 clients.
2029+ : cabinet ing\u00e9nierie BIM, SaaS, \u00e9quipe 3-5 pers.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 15 — CONCLUSION (dark premium)
// ═══════════════════════════════════════════════════════════
let s15 = pres.addSlide();
s15.background = { color: NAVY };

s15.addText("Soutenance BTS MEC U62  \u00b7  Session 2026", { x: 0.5, y: 0.4, w: 5, h: 0.3, fontSize: 9, color: TEAL, fontFace: FB, charSpacing: 2, margin: 0 });

// Left decorative line
s15.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.85, w: 0.04, h: 1.5, fill: { color: TEAL } });

s15.addText("Merci de votre\nattention", { x: 0.6, y: 0.85, w: 5.5, h: 1.5, fontSize: 34, bold: true, color: WHITE, fontFace: FT, margin: 0 });

// Right — takeaways
const takes = [
  ["5 comp\u00e9tences terrain", "Estimer, analyser, suivre, coordonner, contr\u00f4ler", TEAL],
  ["Double lecture", "MOA Maroc + ex\u00e9cution France", ORANGE],
  ["Rigueur de l\u2019\u00e9crit", "OS, attachements, CR \u2013 le rempart contre les litiges", GREEN],
  ["BIMCO", "La rigueur terrain + les outils num\u00e9riques", NAVY]
];
takes.forEach((t, i) => {
  s15.addShape(pres.shapes.RECTANGLE, { x: 6, y: 0.7 + i * 1.05, w: 3.6, h: 0.88, fill: { color: WHITE }, shadow: shadow(), rectRadius: 0.05 });
  s15.addShape(pres.shapes.RECTANGLE, { x: 6, y: 0.7 + i * 1.05, w: 0.06, h: 0.88, fill: { color: t[2] }, rectRadius: 0.02 });
  circle(s15, 6.15, 0.82 + i * 1.05, String(i + 1), 0.28, t[2]);
  s15.addText([
    { text: t[0], options: { bold: true, color: DARK, fontSize: 10, breakLine: true } },
    { text: t[1], options: { color: MUTED, fontSize: 8 } }
  ], { x: 6.55, y: 0.82 + i * 1.05, w: 2.9, h: 0.68, fontFace: FB, valign: "middle", margin: 0 });
});

s15.addText([
  { text: "BAHAFID Mohamed  \u00b7  N\u00b0 02537399911  \u00b7  Acad\u00e9mie de Lyon", options: { breakLine: true } },
  { text: "BIMCO  |  gestion.bimco-consulting.fr  |  Bussi\u00e8res, Loire 42510", options: {} }
], { x: 0.6, y: 4.55, w: 6, h: 0.8, fontSize: 9, color: TEAL, fontFace: FB, margin: 0 });

s15.addNotes(`SLIDE 15 - CONCLUSION (1 min)
En synth\u00e8se, 4 messages :
1. 5 comp\u00e9tences du BTS MEC pratiqu\u00e9es en situation r\u00e9elle.
2. Double lecture : MOA publique au Maroc, ex\u00e9cution terrain en France.
3. La rigueur de l\u2019\u00e9crit \u2013 le seul rempart contre les litiges.
4. BIMCO \u2013 la rigueur terrain + les outils num\u00e9riques.
Je suis \u00e0 votre disposition pour vos questions.`);

// ═══════════════════════════════════════════════════════════
// GENERATE
// ═══════════════════════════════════════════════════════════
pres.writeFile({ fileName: "D:/PREPA BTS MEC/08_U62_Rapport_Activites/Soutenance_PowerPoint/SOUTENANCE_U62_v3.pptx" })
  .then(() => console.log("OK: SOUTENANCE_U62_v3.pptx generated"))
  .catch(e => console.error("ERROR:", e));
