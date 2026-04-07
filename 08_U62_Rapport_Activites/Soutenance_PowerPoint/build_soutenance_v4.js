const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "BAHAFID Mohamed";
pres.title = "Rapport d\u2019Activit\u00e9s Professionnelles \u2013 BTS MEC U62 \u2013 Session 2026";

// ── PALETTE ──
const NAVY    = "1A2744";
const NAVY_L  = "2A3D5E";
const TEAL    = "007A7F";
const TEAL_D  = "005F63";
const TEAL_L  = "E0F4F5";
const CREAM   = "FAFAF7";
const WHITE   = "FFFFFF";
const DARK    = "1C1C1C";
const BODY    = "3A3A3A";
const MUTED   = "888888";
const ORANGE  = "D97706";
const ORANGE_L= "FEF3C7";
const GREEN   = "059669";
const GREEN_L = "D1FAE5";
const RED_L   = "FEE2E2";
const CARD    = WHITE;
const FT = "Cambria", FB = "Calibri";

// ── HELPERS ──
const shd = () => ({ type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.07 });

function topBar(s) { s.background = { color: CREAM }; s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: TEAL } }); }
function footer(s, n) { s.addText(`BAHAFID Mohamed  \u00b7  Rapport U62  \u00b7  BTS MEC 2026  \u00b7  ${n}/15`, { x: 0.4, y: 5.3, w: 9.2, h: 0.22, fontSize: 6.5, color: MUTED, fontFace: FB, align: "right" }); }

function badge(s, text) {
  const w = Math.max(1.05, text.length * 0.09 + 0.45);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.14, w, h: 0.33, fill: { color: TEAL_L }, rectRadius: 0.04 });
  s.addText(text, { x: 0.4, y: 0.14, w, h: 0.33, fontSize: 8, bold: true, color: TEAL_D, fontFace: FB, align: "center", valign: "middle", margin: 0 });
}

function heading(s, t, sub) {
  s.addText(t, { x: 0.4, y: 0.55, w: 9.2, h: 0.4, fontSize: 17, bold: true, color: DARK, fontFace: FT, margin: 0 });
  if (sub) s.addText(sub, { x: 0.4, y: 0.93, w: 9.2, h: 0.26, fontSize: 9, color: MUTED, fontFace: FB, margin: 0 });
}

function card(s, x, y, w, h, fill, accent) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill || CARD }, shadow: shd(), rectRadius: 0.06 });
  if (accent) s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.045, fill: { color: accent }, rectRadius: 0.02 });
}

function iconCircle(s, x, y, icon, sz, bg) {
  const r = sz || 0.38;
  s.addShape(pres.shapes.OVAL, { x, y, w: r, h: r, fill: { color: bg || TEAL } });
  s.addText(icon, { x, y, w: r, h: r, fontSize: r > 0.35 ? 15 : 11, color: WHITE, fontFace: FB, align: "center", valign: "middle", margin: 0 });
}

function progressBar(s, x, y, w, pct, color, label, valueText) {
  // background bar
  s.addShape(pres.shapes.RECTANGLE, { x, y: y + 0.01, w, h: 0.14, fill: { color: "EEEEEE" }, rectRadius: 0.07 });
  // filled bar
  s.addShape(pres.shapes.RECTANGLE, { x, y: y + 0.01, w: w * pct / 100, h: 0.14, fill: { color }, rectRadius: 0.07 });
  // label left
  s.addText(label, { x: x - 2.6, y: y - 0.02, w: 2.5, h: 0.2, fontSize: 7.5, color: BODY, fontFace: FB, align: "right", margin: 0 });
  // value right
  s.addText(valueText, { x: x + w + 0.08, y: y - 0.02, w: 0.8, h: 0.2, fontSize: 7.5, bold: true, color, fontFace: FB, margin: 0 });
}

function arrow(s, x1, y1, x2, y2, color) {
  s.addShape(pres.shapes.LINE, { x: x1, y: y1, w: x2 - x1, h: y2 - y1, line: { color: color || TEAL, width: 1.5, endArrowType: "triangle" } });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 1 — TITRE (dark premium + geometric motif)
// ═══════════════════════════════════════════════════════════
let s1 = pres.addSlide();
s1.background = { color: NAVY };

// Geometric accent circles — decorative background elements
s1.addShape(pres.shapes.OVAL, { x: -0.5, y: -0.5, w: 2.5, h: 2.5, fill: { color: NAVY_L } });
s1.addShape(pres.shapes.OVAL, { x: 8.0, y: 3.5, w: 3.0, h: 3.0, fill: { color: NAVY_L } });

// Vertical accent line
s1.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: 0.6, w: 0.05, h: 2.6, fill: { color: TEAL } });

s1.addText("U62", { x: 0.65, y: 0.6, w: 2, h: 0.3, fontSize: 11, charSpacing: 8, bold: true, color: TEAL, fontFace: FB, margin: 0 });
s1.addText([
  { text: "Rapport", options: { breakLine: true, fontSize: 40, bold: true } },
  { text: "d\u2019Activit\u00e9s", options: { breakLine: true, fontSize: 40, bold: true } },
  { text: "Professionnelles", options: { fontSize: 40, bold: true } }
], { x: 0.65, y: 0.95, w: 6, h: 2.3, color: WHITE, fontFace: FT, margin: 0 });

s1.addText("BTS MEC  \u00b7  SESSION 2026", { x: 0.65, y: 3.35, w: 5, h: 0.3, fontSize: 10, color: TEAL, fontFace: FB, charSpacing: 4, margin: 0 });

// Stat boxes with icons — right side
const s1Stats = [
  ["\u23F3", "8 ans", "Exp\u00e9rience BTP"],
  ["\uD83C\uDF0D", "2 pays", "Maroc + France"],
  ["\uD83D\uDCB0", "82,5 M DH", "Investissements"],
  ["\uD83D\uDCC4", "7 march\u00e9s", "publics"]
];
s1Stats.forEach((st, i) => {
  const sy = 0.55 + i * 1.12;
  s1.addShape(pres.shapes.RECTANGLE, { x: 7.0, y: sy, w: 2.6, h: 0.95, fill: { color: NAVY_L }, line: { color: TEAL, width: 0.8 }, rectRadius: 0.06 });
  s1.addText(st[0], { x: 7.08, y: sy + 0.1, w: 0.5, h: 0.5, fontSize: 18, fontFace: FB, align: "center", valign: "middle", margin: 0 });
  s1.addText(st[1], { x: 7.55, y: sy + 0.08, w: 1.9, h: 0.42, fontSize: 20, bold: true, color: WHITE, fontFace: FT, margin: 0 });
  s1.addText(st[2], { x: 7.55, y: sy + 0.52, w: 1.9, h: 0.3, fontSize: 8.5, color: TEAL, fontFace: FB, margin: 0 });
});

s1.addText([
  { text: "BAHAFID Mohamed", options: { bold: true, breakLine: true, fontSize: 11 } },
  { text: "N\u00b0 02537399911  \u00b7  Acad\u00e9mie de Lyon  \u00b7  BIMCO", options: { fontSize: 9 } }
], { x: 0.65, y: 4.6, w: 6, h: 0.7, color: TEAL, fontFace: FB, margin: 0 });

s1.addNotes(`SLIDE 1 - TITRE (1 min)
Bonjour, Mohamed BAHAFID, candidat libre BTS MEC session 2026, acad\u00e9mie de Lyon.
En 8 ans dans le BTP, j\u2019ai travaill\u00e9 des deux c\u00f4t\u00e9s de la table : 3 ans c\u00f4t\u00e9 ma\u00eetrise d\u2019ouvrage au Maroc, 5 ans c\u00f4t\u00e9 ex\u00e9cution en France.
C\u2019est cette double lecture des projets qui structure tout ce rapport.
82,5 millions de dirhams d\u2019investissements g\u00e9r\u00e9s, 7 march\u00e9s publics.
Aujourd\u2019hui je dirige BIMCO, micro-entreprise sp\u00e9cialis\u00e9e BIM et \u00e9conomie de la construction.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 2 — QUI SUIS-JE (icon fiche + visual timeline)
// ═══════════════════════════════════════════════════════════
let s2 = pres.addSlide();
topBar(s2); badge(s2, "Profil"); footer(s2, 2);
heading(s2, "QUI SUIS-JE ?", "Un profil construit en trois phases : MOA publique, ex\u00e9cution terrain, BIM");

// Left — Fiche with icons
card(s2, 0.4, 1.32, 4.35, 3.8, CARD, TEAL);
const ficheI = [
  ["\uD83D\uDC64", "Candidat", "BAHAFID Mohamed"],
  ["\uD83C\uDFAB", "N\u00b0 Candidat", "02537399911"],
  ["\uD83C\uDFEB", "Acad\u00e9mie", "Lyon"],
  ["\uD83C\uDFE2", "Structure", "Conseil R\u00e9gional BMK"],
  ["\uD83D\uDCBC", "Poste", "Technicien \u00e9tudes et suivi"],
  ["\u23F1\uFE0F", "Exp\u00e9rience", "8 ans (3 Maroc + 5 France)"],
  ["\uD83D\uDDA5\uFE0F", "Formation BIM", "Modeleur BIM \u2013 AFPA (8 mois)"],
  ["\uD83D\uDE80", "Activit\u00e9", "BIMCO \u2013 Projeteur BIM / \u00e9co."],
  ["\uD83D\uDD11", "SIREN", "999 580 053 / 7112B"]
];
ficheI.forEach((f, i) => {
  const yy = 1.5 + i * 0.385;
  s2.addText(f[0], { x: 0.52, y: yy, w: 0.35, h: 0.32, fontSize: 11, fontFace: FB, align: "center", valign: "middle", margin: 0 });
  s2.addText([
    { text: f[1], options: { bold: true, color: TEAL_D, fontSize: 7.5 } },
    { text: "  " + f[2], options: { color: BODY, fontSize: 7.5 } }
  ], { x: 0.9, y: yy, w: 3.65, h: 0.32, fontFace: FB, margin: 0 });
  if (i < ficheI.length - 1) s2.addShape(pres.shapes.RECTANGLE, { x: 0.9, y: yy + 0.34, w: 3.5, h: 0.004, fill: { color: "EEEEEE" } });
});

// Right — Visual timeline with connecting line
card(s2, 5.05, 1.32, 4.55, 3.8, CARD, ORANGE);
s2.addText("PARCOURS", { x: 5.25, y: 1.45, w: 2.5, h: 0.28, fontSize: 9, bold: true, charSpacing: 3, color: ORANGE, fontFace: FB, margin: 0 });

// Vertical line
s2.addShape(pres.shapes.RECTANGLE, { x: 5.43, y: 1.92, w: 0.025, h: 2.95, fill: { color: "DDDDDD" } });

const tlIcons = ["\uD83C\uDDF2\uD83C\uDDE6", "\uD83C\uDDEB\uD83C\uDDF7", "\uD83C\uDFAF", "\uD83D\uDE80"];
const tlData = [
  ["2017\u20132022", "MOA publique \u2013 Maroc (3 ans)", "7 march\u00e9s, +100 M DH"],
  ["2022\u20132024", "Ex\u00e9cution terrain \u2013 France", "Chef GO Ergalis + Minssieux"],
  ["2024\u20132025", "Formation BIM \u2013 AFPA Colmar", "Titre Modeleur BIM, 78 postes"],
  ["2026", "BIMCO + BTS MEC", "Micro-entreprise + candidat libre"]
];
tlData.forEach((t, i) => {
  const ty = 1.92 + i * 0.74;
  iconCircle(s2, 5.3, ty, tlIcons[i], 0.3, i < 2 ? TEAL : (i === 2 ? ORANGE : NAVY));
  s2.addText([
    { text: t[0] + "  ", options: { bold: true, color: TEAL_D, fontSize: 8.5 } },
    { text: t[1], options: { bold: true, color: DARK, fontSize: 8.5, breakLine: true } },
    { text: t[2], options: { color: MUTED, fontSize: 7.5 } }
  ], { x: 5.72, y: ty - 0.04, w: 3.7, h: 0.65, fontFace: FB, margin: 0 });
});

s2.addNotes(`SLIDE 2 - QUI SUIS-JE ? (1 min 30)
Mon parcours en 4 phases.
Phase 1 : Maroc, 2017-2022. Technicien \u00e9tudes au Conseil R\u00e9gional BMK. 3 ans sur 7 march\u00e9s publics, +100 M DH.
Phase 2 : France, 2022-2024. Chef GO Ergalis \u00e0 Feurs, puis chef chantier Minssieux \u00e0 Mornant.
Phase 3 : BIM \u00e0 l\u2019AFPA Colmar, 2024-2025. Titre Modeleur BIM, 8 mois.
Phase 4 : BIMCO janvier 2026, candidat libre BTS MEC.
Au total : 8 ans BTP, 3 Maroc + 5 France.
[Si le jury demande : les 5 ans France incluent ex\u00e9cution 2022-2024, BIM 2024-2025, BIMCO 2025-2026]`);

// ═══════════════════════════════════════════════════════════
// SLIDE 3 — STRUCTURE D'ACCUEIL (icons pour missions)
// ═══════════════════════════════════════════════════════════
let s3 = pres.addSlide();
topBar(s3); badge(s3, "01 \u2013 Cadre"); footer(s3, 3);
heading(s3, "STRUCTURE D\u2019ACCUEIL", "Conseil R\u00e9gional de B\u00e9ni Mellal-Kh\u00e9nifra \u2013 Agence d\u2019Ex\u00e9cution des Projets");

// Left — Description + icon missions
card(s3, 0.4, 1.32, 5.6, 3.8, CARD, TEAL);
s3.addText([
  { text: "Collectivit\u00e9 territoriale  \u00b7  5 provinces  \u00b7  2,5 millions d\u2019habitants", options: { fontSize: 9, color: BODY, breakLine: true } },
  { text: "Poste : Technicien \u00e9tudes et suivi des travaux", options: { fontSize: 9, bold: true, color: TEAL } }
], { x: 0.6, y: 1.48, w: 5.2, h: 0.55, fontFace: FB, margin: 0 });

s3.addText("MISSIONS PRINCIPALES", { x: 0.6, y: 2.15, w: 3, h: 0.25, fontSize: 8, bold: true, color: TEAL, charSpacing: 2, fontFace: FB, margin: 0 });

const missionIcons = [
  ["\uD83D\uDCCF", "M\u00e9tr\u00e9s avant-projet", "Estimation confidentielle", TEAL],
  ["\uD83D\uDCDD", "R\u00e9daction DCE", "CPS, RC, BPDE", TEAL_D],
  ["\uD83D\uDD0D", "Analyse des offres", "Commission AO", ORANGE],
  ["\uD83D\uDCCA", "Suivi financier", "Mensuel, d\u00e9comptes", GREEN],
  ["\uD83D\uDEE0\uFE0F", "Visites terrain", "Attachements contradictoires", NAVY]
];
missionIcons.forEach((m, i) => {
  const my = 2.5 + i * 0.52;
  iconCircle(s3, 0.6, my, m[0], 0.34, m[3]);
  s3.addText([
    { text: m[1], options: { bold: true, color: DARK, fontSize: 8.5, breakLine: true } },
    { text: m[2], options: { color: MUTED, fontSize: 7.5 } }
  ], { x: 1.04, y: my - 0.02, w: 4.7, h: 0.46, fontFace: FB, margin: 0 });
});

// Right — Cadre r\u00e9glementaire (dark card)
card(s3, 6.25, 1.32, 3.35, 3.8, NAVY);
s3.addText("\uD83D\uDCDC  CADRE R\u00c9GLEMENTAIRE", { x: 6.45, y: 1.5, w: 2.95, h: 0.28, fontSize: 8.5, bold: true, charSpacing: 1, color: TEAL, fontFace: FB, margin: 0 });
const reglI = [
  ["\u2022", "Pi\u00e8ces AO", "CPS + RC + BPDE"],
  ["\u2022", "Proc\u00e9dure", "Appel d\u2019offres ouvert"],
  ["\u2022", "Estimation", "Confidentielle obligatoire"],
  ["\u2022", "Normes", "Normes marocaines, RPS 2000"],
  ["\u2022", "Suivi", "Attachements contradictoires"]
];
reglI.forEach((r, i) => {
  s3.addText([
    { text: r[1], options: { bold: true, color: WHITE, fontSize: 8.5, breakLine: true } },
    { text: r[2], options: { color: TEAL, fontSize: 8 } }
  ], { x: 6.65, y: 1.95 + i * 0.52, w: 2.75, h: 0.46, fontFace: FB, margin: 0 });
});

s3.addNotes(`SLIDE 3 - STRUCTURE D'ACCUEIL (1 min 30)
Conseil R\u00e9gional BMK. 5 provinces, 2,5 M habitants.
Rattach\u00e9 \u00e0 l\u2019Agence d\u2019Ex\u00e9cution des Projets, dirig\u00e9e par M. DOGHMANI.
Poste : technicien \u00e9tudes et suivi. Concr\u00e8tement : estimation confidentielle, DCE, suivi terrain.
5 missions principales visible sur la slide.
Cadre r\u00e9glementaire : estimation confidentielle obligatoire avant tout AO \u2013 fixe le prix plafond.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 4 — BIMCO (workflow diagram)
// ═══════════════════════════════════════════════════════════
let s4 = pres.addSlide();
topBar(s4); badge(s4, "01 \u2013 BIMCO"); footer(s4, 4);
heading(s4, "BIMCO \u2013 MON ACTIVIT\u00c9 IND\u00c9PENDANTE", "Cr\u00e9\u00e9e en janvier 2026 \u2013 BIM au service de l\u2019\u00e9conomiste de la construction");

// Mission box
card(s4, 0.4, 1.32, 5.0, 1.3, CARD, TEAL);
s4.addText("\uD83C\uDFAF  MISSION", { x: 0.6, y: 1.45, w: 4.6, h: 0.22, fontSize: 9, bold: true, color: TEAL, fontFace: FB, margin: 0 });
s4.addText("Appliquer le BIM aux m\u00e9tiers de l\u2019\u00e9conomiste. M\u00e9tr\u00e9s par extraction de maquette, \u00e9tudes de prix ancr\u00e9es co\u00fbts r\u00e9els, plugins Revit/Dynamo pour automatiser m\u00e8tre \u2192 DPGF.", {
  x: 0.6, y: 1.72, w: 4.6, h: 0.75, fontSize: 8.5, color: BODY, fontFace: FB, margin: 0
});

// Workflow diagram — horizontal process flow
card(s4, 0.4, 2.82, 9.2, 1.2, CARD, ORANGE);
s4.addText("\u26A1  WORKFLOW BIMCO", { x: 0.6, y: 2.92, w: 4, h: 0.22, fontSize: 9, bold: true, color: ORANGE, fontFace: FB, margin: 0 });
const wfSteps = [
  ["\uD83C\uDFD7\uFE0F", "Maquette\nRevit"],
  ["\uD83D\uDD04", "Export\nIFC"],
  ["\uD83D\uDCCF", "Extraction\nQuantit\u00e9s"],
  ["\uD83D\uDCB0", "\u00c9tude\nde Prix"],
  ["\uD83D\uDCC4", "DPGF\nAutomatis\u00e9"]
];
wfSteps.forEach((ws, i) => {
  const wx = 0.65 + i * 1.8;
  iconCircle(s4, wx, 3.25, ws[0], 0.36, i === 4 ? GREEN : TEAL);
  s4.addText(ws[1], { x: wx - 0.15, y: 3.65, w: 0.68, h: 0.32, fontSize: 6.5, color: BODY, fontFace: FB, align: "center", margin: 0 });
  if (i < 4) {
    s4.addShape(pres.shapes.LINE, { x: wx + 0.42, y: 3.43, w: 1.32, h: 0, line: { color: TEAL, width: 1.2, endArrowType: "triangle" } });
  }
});

// Right — Positionnement (dark vertical card)
card(s4, 5.7, 1.32, 3.9, 1.3, NAVY);
s4.addText("\uD83D\uDCA1  POSITIONNEMENT", { x: 5.9, y: 1.42, w: 3.5, h: 0.22, fontSize: 9, bold: true, color: WHITE, fontFace: FB, margin: 0 });
const posI = [
  ["\uD83C\uDDF2\uD83C\uDDE6", "MOA publique \u2013 vision globale"],
  ["\uD83C\uDDEB\uD83C\uDDF7", "Ex\u00e9cution \u2013 r\u00e9alit\u00e9 chantier"],
  ["\uD83D\uDDA5\uFE0F", "BIM \u2013 conception \u2194 chiffrage"]
];
posI.forEach((p, i) => {
  s4.addText(p[0] + "  " + p[1], { x: 5.9, y: 1.72 + i * 0.27, w: 3.5, h: 0.24, fontSize: 8.5, color: WHITE, fontFace: FB, margin: 0 });
});

// Tech pills
card(s4, 0.4, 4.2, 9.2, 0.85, CARD, TEAL_D);
s4.addText("\uD83D\uDD27  STACK", { x: 0.6, y: 4.3, w: 1, h: 0.22, fontSize: 8, bold: true, color: TEAL_D, fontFace: FB, margin: 0 });
const techs = ["Revit API", "C# .NET", "Dynamo", "Python", "IFC / BIM360", "React / Node.js"];
techs.forEach((t, i) => {
  s4.addShape(pres.shapes.RECTANGLE, { x: 1.7 + i * 1.3, y: 4.55, w: 1.18, h: 0.3, fill: { color: TEAL_L }, rectRadius: 0.15 });
  s4.addText(t, { x: 1.7 + i * 1.3, y: 4.55, w: 1.18, h: 0.3, fontSize: 7, bold: true, color: TEAL_D, fontFace: FB, align: "center", valign: "middle", margin: 0 });
});

s4.addNotes(`SLIDE 4 - BIMCO (1 min)
BIMCO est n\u00e9 d\u2019un constat : les outils BIM sont faits pour les architectes, pas pour celui qui chiffre.
Cr\u00e9\u00e9 en janvier 2026. Mission : extraction m\u00e9tr\u00e9s depuis la maquette, \u00e9tudes de prix, plugins Revit/Dynamo pour automatiser la cha\u00eene m\u00e8tre \u2192 DPGF.
Le workflow en 5 \u00e9tapes : Maquette Revit \u2192 Export IFC \u2192 Extraction quantit\u00e9s \u2192 \u00c9tude de prix \u2192 DPGF automatis\u00e9.
Triple comp\u00e9tence rare : MOA + Ex\u00e9cution + BIM.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 5 — PROJET 1 + BUDGET
// ═══════════════════════════════════════════════════════════
let s5 = pres.addSlide();
topBar(s5); badge(s5, "02 \u2013 Projet 1"); footer(s5, 5);
heading(s5, "MISE \u00c0 NIVEAU DE 4 COMMUNES", "March\u00e9 n\u00b038-RBK-2017 \u2013 Province de Kh\u00e9nifra");

// Stat row with icons
const p1Stats = [
  ["\uD83D\uDCB0", "53,5 M DH", "Budget TTC", TEAL],
  ["\uD83C\uDFD8\uFE0F", "4 communes", "Province Kh\u00e9nifra", ORANGE],
  ["\uD83D\uDD27", "8 corps", "d\u2019\u00e9tat", TEAL_D],
  ["\u23F3", "18 mois", "Dur\u00e9e suivi", NAVY]
];
p1Stats.forEach((st, i) => {
  card(s5, 0.4 + i * 2.35, 1.32, 2.15, 0.9, CARD, st[3]);
  s5.addText(st[0], { x: 0.45 + i * 2.35, y: 1.4, w: 0.4, h: 0.35, fontSize: 16, fontFace: FB, align: "center", margin: 0 });
  s5.addText(st[1], { x: 0.85 + i * 2.35, y: 1.4, w: 1.6, h: 0.38, fontSize: 15, bold: true, color: st[3], fontFace: FT, margin: 0 });
  s5.addText(st[2], { x: 0.85 + i * 2.35, y: 1.78, w: 1.6, h: 0.25, fontSize: 7.5, color: MUTED, fontFace: FB, margin: 0 });
});

// Table
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
  { text: c[0], options: { fontSize: 7.5, color: TEAL_D, bold: true } },
  { text: c[1], options: { fontSize: 7.5, color: BODY } },
  { text: c[2], options: { fontSize: 7.5, color: BODY, align: "center" } }
]));
s5.addTable(tRows, { x: 0.4, y: 2.4, w: 4.3, colW: [0.5, 2.8, 0.55], border: { pt: 0.5, color: "E8E8E8" }, rowH: 0.27, autoPage: false });

// Pie chart
s5.addChart(pres.charts.DOUGHNUT, [{
  name: "Budget", labels: ["Assainissement", "Chauss\u00e9e", "\u00c9clairage", "Trottoirs", "Murs", "Signal.", "Paysager", "Mobilier"],
  values: [22, 19, 16, 14, 14, 6, 5, 4]
}], {
  x: 5.1, y: 2.3, w: 4.5, h: 2.8,
  showPercent: true, showTitle: false, showLegend: true, legendPos: "b", legendFontSize: 7,
  chartColors: [TEAL, "00A3A8", ORANGE, "80CED1", "B3E4E6", TEAL_L, NAVY_L, "F5DEB3"],
  dataLabelColor: DARK, dataLabelFontSize: 8, holeSize: 55
});

s5.addNotes(`SLIDE 5 - PROJET 1 (1 min)
Mise \u00e0 niveau de 4 communes, province de Kh\u00e9nifra.
53,5 M DH TTC \u2248 4,8 M\u20ac. 8 corps d\u2019\u00e9tat, 18 mois, 4 sites distants de 20 \u00e0 80 km.
Assainissement + chauss\u00e9e = 41% du budget.
C\u2019est sur ce projet que portent les 4 premi\u00e8res situations.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 6 — SITUATIONS 1 & 2 (enhanced CPAR)
// ═══════════════════════════════════════════════════════════
let s6 = pres.addSlide();
topBar(s6); badge(s6, "02 \u2013 Sit. 1 & 2"); footer(s6, 6);
heading(s6, "ESTIMATION CONFIDENTIELLE & ANALYSE DES OFFRES", "Comp\u00e9tence C18 \u2013 M\u00e9trer, estimer, analyser les offres");

function sitCard(s, x, sit, subtitle, bigVal, bigColor, accent, items) {
  card(s, x, 1.32, 4.5, 3.8, CARD, accent);
  s.addText(sit, { x: x + 0.2, y: 1.45, w: 3.3, h: 0.26, fontSize: 8, bold: true, color: accent, charSpacing: 1, fontFace: FB, margin: 0 });
  s.addText(subtitle, { x: x + 0.2, y: 1.73, w: 3.0, h: 0.22, fontSize: 9, bold: true, color: DARK, fontFace: FB, margin: 0 });
  s.addText(bigVal, { x: x + 3.0, y: 1.4, w: 1.3, h: 0.55, fontSize: 26, bold: true, color: bigColor, fontFace: FT, align: "right", margin: 0 });
  const labels = ["CONTEXTE", "PROBL\u00c8ME", "ACTION", "R\u00c9SULTAT"];
  const colors = [TEAL, ORANGE, TEAL_D, GREEN];
  const icons  = ["\uD83D\uDCCB", "\u26A0\uFE0F", "\u2699\uFE0F", "\u2705"];
  items.forEach((item, i) => {
    const iy = 2.08 + i * 0.58;
    s.addShape(pres.shapes.RECTANGLE, { x: x + 0.2, y: iy, w: 4.1, h: 0.5, fill: { color: i % 2 === 0 ? WHITE : CREAM }, rectRadius: 0.04 });
    s.addText(icons[i], { x: x + 0.22, y: iy + 0.04, w: 0.3, h: 0.3, fontSize: 12, fontFace: FB, align: "center", margin: 0 });
    s.addText([
      { text: labels[i] + "  ", options: { bold: true, color: colors[i], fontSize: 6.5 } },
      { text: item, options: { color: BODY, fontSize: 7.5 } }
    ], { x: x + 0.55, y: iy + 0.02, w: 3.65, h: 0.44, fontFace: FB, margin: 0 });
  });
}

sitCard(s6, 0.4, "SITUATION 1 \u2013 ESTIMATION CONFIDENTIELLE", "Ouaoumana \u2013 15,8 M DH HT", "3,2%", GREEN, TEAL, [
  "Prix plafond avant AO. 35% du budget. D\u00e9lai 3 semaines.",
  "Mercuriale 2014 obsol\u00e8te : prix d\u00e9riv\u00e9s de 15 \u00e0 22%.",
  "112 lignes AutoCAD + 4 visites terrain + 3 sources crois\u00e9es.",
  "\u00c9cart 3,2% vs 5-10% norme. M\u00e9thode adopt\u00e9e standard."
]);

sitCard(s6, 5.1, "SITUATION 2 \u2013 ANALYSE DES OFFRES", "Commission CAO \u2013 3 offres", "94/100", ORANGE, ORANGE, [
  "Membre technique commission. Analyse comparative 3 dossiers.",
  "7 erreurs arithm\u00e9tiques + prix bas anormaux sur 42% du montant.",
  "Grille 100 pts (tech 60 + fin 40). Justification \u00e9crite.",
  "Attribution 15 jours. Z\u00e9ro recours. Rapport valid\u00e9 sans r\u00e9serve."
]);

s6.addNotes(`SLIDE 6 - SITUATIONS 1 & 2 (2 min)
SIT 1 : Estimation confidentielle Ouaoumana. 15,8 M DH, 35% budget, 3 semaines.
Probl\u00e8me : mercuriale 2014 obsol\u00e8te, \u00e9carts 15-22%.
Action : 3 sources crois\u00e9es. R\u00e9sultat : 3,2% d\u2019\u00e9cart vs 5-10% norme.

SIT 2 : Analyse des offres en commission. 3 soumissionnaires.
Probl\u00e8me : 7 erreurs + prix bas anormaux 42%. Grille 100 pts.
R\u00e9sultat : 94/100, 15 jours, z\u00e9ro recours.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 7 — SITUATIONS 3 & 4
// ═══════════════════════════════════════════════════════════
let s7 = pres.addSlide();
topBar(s7); badge(s7, "02 \u2013 Sit. 3 & 4"); footer(s7, 7);
heading(s7, "SUIVI FINANCIER & COORDINATION MULTI-SITES", "Comp\u00e9tences C18 (suivi budget) et C19 (coordonner, communiquer)");

sitCard(s7, 0.4, "SITUATION 3 \u2013 SUIVI FINANCIER", "Kerrouchen \u2013 7,3 M DH \u2013 18 mois", "+0,8%", GREEN, TEAL, [
  "Suivi mensuel complet : situations, quantit\u00e9s, d\u00e9comptes.",
  "Chauss\u00e9e +12%, murs +15%. D\u00e9rive +4,8% (seuil avenant = 5%).",
  "TdB hebdo 3 indicateurs. Compensation paysager \u221244k + mobilier \u221212k.",
  "D\u00e9passement final +0,8%. Aucun avenant. TdB adopt\u00e9 3 communes."
]);

sitCard(s7, 5.1, "SITUATION 4 \u2013 COORDINATION 4 CHANTIERS", "Province de Kh\u00e9nifra \u2013 20 \u00e0 80 km", "48h", ORANGE, ORANGE, [
  "Relais unique Directeur. Interface entreprise/BET/labo/hi\u00e9rarchie.",
  "S23 : retard enrob\u00e9s + alerte m\u00e9t\u00e9o + litige bordures 15%.",
  "Note Directeur + OS arr\u00eat photos + re-mesurage contradictoire.",
  "3 crises r\u00e9solues. Litige 15% \u2192 2,3%. Mod\u00e8le CR adopt\u00e9."
]);

s7.addNotes(`SLIDE 7 - SITUATIONS 3 & 4 (2 min 30)
SIT 3 : Suivi financier Kerrouchen. 7,3 M DH, 18 mois.
Chauss\u00e9e +12%, murs +15%. Extrapolation +4,8% (seuil 5%).
TdB hebdo + compensation \u221256k. R\u00e9sultat : +0,8%, aucun avenant.

SIT 4 : Semaine 23 \u2013 3 crises en 48h.
Retard enrob\u00e9s + m\u00e9t\u00e9o + litige bordures 15%.
R\u00e9sultat : litige 15% \u2192 2,3%. CR devenu standard.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 8 — PROJET 2 + BUDGET
// ═══════════════════════════════════════════════════════════
let s8 = pres.addSlide();
topBar(s8); badge(s8, "03 \u2013 Projet 2"); footer(s8, 8);
heading(s8, "ROUTE LEHRI-KERROUCHEN \u2013 25 KM", "Programme PRR3 \u2013 29 M DH TTC \u2013 Zone montagneuse du Moyen Atlas");

// Stats row
const p2S = [
  ["\uD83D\uDCB0", "29 M DH", "Budget TTC", TEAL],
  ["\u26F0\uFE0F", "25 km", "Zone montagneuse", ORANGE],
  ["\uD83D\uDE9C", "120 334 m\u00b3", "D\u00e9blais v\u00e9rifi\u00e9s", TEAL_D],
  ["\uD83D\uDCC4", "53 prix", "Au bordereau", NAVY]
];
p2S.forEach((st, i) => {
  card(s8, 0.4 + i * 2.35, 1.32, 2.15, 0.9, CARD, st[3]);
  s8.addText(st[0], { x: 0.45 + i * 2.35, y: 1.4, w: 0.4, h: 0.35, fontSize: 16, fontFace: FB, align: "center", margin: 0 });
  s8.addText(st[1], { x: 0.85 + i * 2.35, y: 1.4, w: 1.6, h: 0.38, fontSize: 15, bold: true, color: st[3], fontFace: FT, margin: 0 });
  s8.addText(st[2], { x: 0.85 + i * 2.35, y: 1.78, w: 1.6, h: 0.25, fontSize: 7.5, color: MUTED, fontFace: FB, margin: 0 });
});

// Bar chart
s8.addChart(pres.charts.BAR, [{
  name: "Budget %", labels: ["Corps chauss\u00e9e", "Ouv. hydrauliques", "Terrassement", "Bretelles", "Rev\u00eatement", "Sout\u00e8nement"],
  values: [30.1, 28.1, 17.9, 11.3, 10.7, 1.9]
}], {
  x: 0.4, y: 2.45, w: 9.2, h: 2.7, barDir: "col",
  chartColors: [TEAL], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: DARK, dataLabelFontSize: 8,
  catAxisLabelColor: BODY, catAxisLabelFontSize: 7, valAxisHidden: true,
  valGridLine: { style: "none" }, catGridLine: { style: "none" },
  showLegend: false, showTitle: false
});

s8.addNotes(`SLIDE 8 - PROJET 2 (1 min)
Route Lehri-Kerrouchen. 25 km, Moyen Atlas. 29 M DH, PRR3.
D\u00e9nivel\u00e9 400 m, pentes 12%. 53 prix, 120 334 m\u00b3 de d\u00e9blais.
Ratio drainage/chauss\u00e9e 2,5x sup\u00e9rieur \u00e0 une route en plaine.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 9 — SITUATION 5 (visual problem-action-result flow)
// ═══════════════════════════════════════════════════════════
let s9 = pres.addSlide();
topBar(s9); badge(s9, "03 \u2013 Sit. 5"); footer(s9, 9);
heading(s9, "CONTR\u00d4LE DES CUBATURES DE TERRASSEMENT", "Comp\u00e9tence C18 \u2013 V\u00e9rifier les quantit\u00e9s ex\u00e9cut\u00e9es");

// Big stat banner
card(s9, 0.4, 1.32, 4.4, 0.85, ORANGE_L, ORANGE);
s9.addText("\u26A0\uFE0F", { x: 0.5, y: 1.4, w: 0.5, h: 0.5, fontSize: 22, fontFace: FB, align: "center", margin: 0 });
s9.addText("+8%", { x: 1.0, y: 1.38, w: 1.5, h: 0.5, fontSize: 30, bold: true, color: ORANGE, fontFace: FT, margin: 0 });
s9.addText("\u00c9cart cumul\u00e9\nd\u00e9tect\u00e9 et corrig\u00e9", { x: 2.5, y: 1.38, w: 2.2, h: 0.6, fontSize: 8.5, color: BODY, fontFace: FB, margin: 0 });

card(s9, 5.2, 1.32, 4.4, 0.85, GREEN_L, GREEN);
s9.addText("\u2705", { x: 5.3, y: 1.4, w: 0.5, h: 0.5, fontSize: 22, fontFace: FB, align: "center", margin: 0 });
s9.addText("285 000 DH", { x: 5.8, y: 1.38, w: 2.2, h: 0.5, fontSize: 22, bold: true, color: GREEN, fontFace: FT, margin: 0 });
s9.addText("Surco\u00fbt absorb\u00e9\nsans avenant", { x: 8.0, y: 1.38, w: 1.5, h: 0.6, fontSize: 8.5, color: BODY, fontFace: FB, margin: 0 });

// Visual flow : Problem → Action → Result (3 columns)
const flowCols = [
  { x: 0.4, w: 3.0, title: "\uD83D\uDD34  PROBL\u00c8ME", bg: RED_L, accent: "DC2626",
    items: ["Calcaire fractur\u00e9 au PK 12", "5 000 m\u00b3 \u00e0 reclassifier", "Co\u00fbt 28 \u2192 85 DH/m\u00b3", "\u00c9cart cumul\u00e9 +8%"] },
  { x: 3.6, w: 3.0, title: "\uD83D\uDD35  ACTION", bg: TEAL_L, accent: TEAL,
    items: ["Profils en travers /25 m", "Relev\u00e9s GPS /500 ml", "3 dalots BA avec BET NOVEC", "Compensation inter-postes"] },
  { x: 6.8, w: 2.8, title: "\uD83D\uDFE2  R\u00c9SULTAT", bg: GREEN_L, accent: GREEN,
    items: ["Surco\u00fbt 100% absorb\u00e9", "R\u00e9ception sans avenant", "Tra\u00e7abilit\u00e9 compl\u00e8te", "Z\u00e9ro r\u00e9clamation"] }
];
flowCols.forEach(col => {
  card(s9, col.x, 2.4, col.w, 2.7, col.bg, col.accent);
  s9.addText(col.title, { x: col.x + 0.15, y: 2.52, w: col.w - 0.3, h: 0.25, fontSize: 8.5, bold: true, color: col.accent, fontFace: FB, margin: 0 });
  col.items.forEach((item, i) => {
    s9.addText("\u2022  " + item, { x: col.x + 0.15, y: 2.85 + i * 0.42, w: col.w - 0.3, h: 0.35, fontSize: 8, color: BODY, fontFace: FB, margin: 0 });
  });
});
// Arrows between columns
s9.addShape(pres.shapes.LINE, { x: 3.45, y: 3.6, w: 0.12, h: 0, line: { color: TEAL, width: 2, endArrowType: "triangle" } });
s9.addShape(pres.shapes.LINE, { x: 6.65, y: 3.6, w: 0.12, h: 0, line: { color: GREEN, width: 2, endArrowType: "triangle" } });

s9.addNotes(`SLIDE 9 - SITUATION 5 (1 min 30)
PK 12 : calcaire fractur\u00e9 non d\u00e9tect\u00e9 par la g\u00e9otechnique.
5 000 m\u00b3 \u00e0 reclassifier. 28 \u2192 85 DH/m\u00b3. Surco\u00fbt 285 000 DH. \u00c9cart +8%.
Action : profils /25m, GPS contradictoire /500ml, 3 dalots BA avec NOVEC.
R\u00e9sultat : 100% absorb\u00e9 par compensation. Z\u00e9ro avenant.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 10 — ACTIVITES COMPLEMENTAIRES
// ═══════════════════════════════════════════════════════════
let s10 = pres.addSlide();
topBar(s10); badge(s10, "03 \u2013 Compl."); footer(s10, 10);
heading(s10, "ACTIVIT\u00c9S COMPL\u00c9MENTAIRES", "5 autres march\u00e9s publics + exp\u00e9rience terrain en France");

// Left — Table
card(s10, 0.4, 1.32, 5.2, 2.8, CARD, TEAL);
s10.addText("\uD83C\uDDF2\uD83C\uDDE6  5 AUTRES MARCH\u00c9S", { x: 0.6, y: 1.45, w: 4.8, h: 0.22, fontSize: 9, bold: true, color: TEAL, fontFace: FB, margin: 0 });
const marches = [
  ["27-RBK", "Route Sidi Bouabbad \u2192 Oued Grou (12 km)", "8,2 M DH"],
  ["28-RBK", "Route Ajdir-Ayoun + Piste Lijon", "6,5 M DH"],
  ["30-RBK", "AEP El Borj \u2013 El Hamam (18 km)", "4,8 M DH"],
  ["39-RBK", "Pistes Hartaf \u2013 Sebt Ait Rahou", "3,1 M DH"],
  ["49-RBK", "Voirie Amghass \u2013 Bouchbel", "5,9 M DH"]
];
const mRows = [[
  { text: "N\u00b0", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7.5 } },
  { text: "Objet", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7.5 } },
  { text: "Montant", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 7.5, align: "right" } }
]];
marches.forEach(m => mRows.push([
  { text: m[0], options: { fontSize: 7.5, color: TEAL_D, bold: true } },
  { text: m[1], options: { fontSize: 7.5, color: BODY } },
  { text: m[2], options: { fontSize: 7.5, color: BODY, align: "right" } }
]));
s10.addTable(mRows, { x: 0.6, y: 1.8, w: 4.8, colW: [0.7, 3.1, 1.0], border: { pt: 0.5, color: "E8E8E8" }, rowH: 0.27 });

// Right — France
card(s10, 5.9, 1.32, 3.7, 2.8, CARD, ORANGE);
s10.addText("\uD83C\uDDEB\uD83C\uDDF7  EXP\u00c9RIENCE FRANCE", { x: 6.1, y: 1.45, w: 3.3, h: 0.22, fontSize: 9, bold: true, color: ORANGE, fontFace: FB, margin: 0 });
const frJobs = [
  ["\uD83D\uDEE0\uFE0F", "Chef d\u2019\u00e9quipe GO", "Ergalis BTP \u2013 Feurs (Loire)", "Banches, armatures, coulage"],
  ["\uD83D\uDC77", "Chef de chantier", "Minssieux & Fils \u2013 Mornant", "Planning, contr\u00f4le qualit\u00e9"]
];
frJobs.forEach((j, i) => {
  const jy = 1.82 + i * 0.95;
  s10.addText(j[0], { x: 6.1, y: jy, w: 0.4, h: 0.35, fontSize: 16, fontFace: FB, align: "center", margin: 0 });
  s10.addText([
    { text: j[1], options: { bold: true, color: DARK, fontSize: 8.5, breakLine: true } },
    { text: j[2], options: { color: ORANGE, fontSize: 7.5, breakLine: true } },
    { text: j[3], options: { color: MUTED, fontSize: 7.5 } }
  ], { x: 6.5, y: jy - 0.02, w: 2.9, h: 0.7, fontFace: FB, margin: 0 });
});

// Apport box
s10.addShape(pres.shapes.RECTANGLE, { x: 6.1, y: 3.55, w: 3.3, h: 0.4, fill: { color: ORANGE_L }, rectRadius: 0.04 });
s10.addText("\uD83D\uDCA1 Apport : co\u00fbts r\u00e9els, rendements, contraintes", { x: 6.1, y: 3.55, w: 3.3, h: 0.4, fontSize: 7.5, bold: true, color: ORANGE, fontFace: FB, align: "center", valign: "middle", margin: 0 });

// Bottom banner
s10.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.35, w: 9.2, h: 0.55, fill: { color: TEAL_L }, rectRadius: 0.06 });
s10.addText("\uD83C\uDFD7\uFE0F  4 types d\u2019infrastructures : routes  \u00b7  VRD  \u00b7  AEP  \u00b7  voirie urbaine", {
  x: 0.6, y: 4.35, w: 8.8, h: 0.55, fontSize: 9.5, bold: true, color: TEAL_D, fontFace: FB, align: "center", valign: "middle", margin: 0
});

s10.addNotes(`SLIDE 10 - ACTIVIT\u00c9S COMPL\u00c9MENTAIRES (1 min)
5 march\u00e9s suppl\u00e9mentaires : routes, pistes, AEP, voirie.
France : chef GO Ergalis + chef chantier Minssieux.
Apport cl\u00e9 : co\u00fbts r\u00e9els de production. 4 types d\u2019infra ma\u00eetris\u00e9es.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 11 — SYNTHESE (progress bars instead of table)
// ═══════════════════════════════════════════════════════════
let s11 = pres.addSlide();
topBar(s11); badge(s11, "04 \u2013 Bilan"); footer(s11, 11);
heading(s11, "SYNTH\u00c8SE DES COMP\u00c9TENCES BTS MEC", "Comp\u00e9tences mobilis\u00e9es sur 5 situations professionnelles");

// Left — Competences with progress bars
card(s11, 0.4, 1.32, 5.8, 3.8, CARD, TEAL);
s11.addText("\uD83D\uDCCA  NIVEAUX DE COMP\u00c9TENCE", { x: 0.6, y: 1.45, w: 5.4, h: 0.25, fontSize: 9, bold: true, color: TEAL, fontFace: FB, margin: 0 });

const compBars = [
  ["M\u00e9tr\u00e9s TCA (112 prix, 4 communes)", 90, TEAL, "Sit.1", "Ma\u00eetrise"],
  ["Estimation en AO (15,8 M DH, 3 sources)", 90, TEAL, "Sit.1", "Ma\u00eetrise"],
  ["Analyse des offres (grille /100)", 85, TEAL, "Sit.2", "Ma\u00eetrise"],
  ["Suivi financier (TdB hebdo)", 95, GREEN, "Sit.3", "Expert \u2605"],
  ["V\u00e9rification quantit\u00e9s (GPS)", 92, GREEN, "Sit.5", "Expert \u2605"],
  ["Communication \u00e9crite (OS, CR)", 88, TEAL, "Sit.4", "Ma\u00eetrise"],
  ["Communication orale (CAO, r\u00e9unions)", 85, TEAL, "Sit.2,4", "Ma\u00eetrise"],
  ["Collaboration BIM (IFC, 78 postes)", 80, ORANGE, "BIM", "Ma\u00eetrise"]
];
compBars.forEach((cb, i) => {
  const by = 1.88 + i * 0.39;
  // Label
  s11.addText(cb[0], { x: 0.6, y: by - 0.02, w: 3.3, h: 0.18, fontSize: 7, color: BODY, fontFace: FB, margin: 0 });
  // Situation badge
  s11.addShape(pres.shapes.RECTANGLE, { x: 3.95, y: by - 0.02, w: 0.45, h: 0.17, fill: { color: TEAL_L }, rectRadius: 0.08 });
  s11.addText(cb[3], { x: 3.95, y: by - 0.02, w: 0.45, h: 0.17, fontSize: 5.5, bold: true, color: TEAL_D, fontFace: FB, align: "center", valign: "middle", margin: 0 });
  // Bar bg
  s11.addShape(pres.shapes.RECTANGLE, { x: 4.5, y: by, w: 1.4, h: 0.12, fill: { color: "EEEEEE" }, rectRadius: 0.06 });
  // Bar fill
  s11.addShape(pres.shapes.RECTANGLE, { x: 4.5, y: by, w: 1.4 * cb[1] / 100, h: 0.12, fill: { color: cb[2] }, rectRadius: 0.06 });
  // Level text
  s11.addText(cb[4], { x: 5.95, y: by - 0.04, w: 0.7, h: 0.2, fontSize: 6.5, bold: true, color: cb[2], fontFace: FB, margin: 0 });
});

// Right — Summary stats
card(s11, 6.5, 1.32, 3.1, 3.8, NAVY);
s11.addText("EN CHIFFRES", { x: 6.7, y: 1.48, w: 2.7, h: 0.25, fontSize: 9, bold: true, color: TEAL, charSpacing: 2, fontFace: FB, margin: 0 });

const summStats = [
  ["8", "sous-comp\u00e9tences\nmobilis\u00e9es", TEAL],
  ["2", "niveau Expert\n(suivi + cubatures)", GREEN],
  ["5", "situations\nprofessionnelles", ORANGE],
  ["100%", "pratiqu\u00e9es sous\ncontrainte r\u00e9elle", WHITE]
];
summStats.forEach((ss, i) => {
  const sy = 1.9 + i * 0.72;
  s11.addText(ss[0], { x: 6.7, y: sy, w: 0.7, h: 0.55, fontSize: 26, bold: true, color: ss[2], fontFace: FT, align: "center", valign: "middle", margin: 0 });
  s11.addText(ss[1], { x: 7.4, y: sy + 0.05, w: 2.0, h: 0.5, fontSize: 8, color: TEAL, fontFace: FB, margin: 0 });
});

s11.addNotes(`SLIDE 11 - SYNTH\u00c8SE (1 min)
8 sous-comp\u00e9tences BTS MEC, toutes en situation r\u00e9elle sous contrainte.
2 au niveau Expert : suivi financier (95%) et contr\u00f4le quantit\u00e9s (92%).
BIM \u00e0 80% = comp\u00e9tence la plus r\u00e9cente, en progression.
D\u00e9tail page 23 du rapport.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 12 — ANALYSE MAROC / FRANCE
// ═══════════════════════════════════════════════════════════
let s12 = pres.addSlide();
topBar(s12); badge(s12, "04 \u2013 Analyse"); footer(s12, 12);
heading(s12, "ANALYSE COMPARATIVE MAROC / FRANCE", "M\u00eames fondamentaux \u2013 deux cadres r\u00e9glementaires \u2013 une seule exigence : la rigueur");

// Table with flag headers
const cmpRows = [
  [
    { text: "", options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 8 } },
    { text: "\uD83C\uDDF2\uD83C\uDDE6  Maroc", options: { bold: true, color: WHITE, fill: { color: TEAL }, fontSize: 8.5, align: "center" } },
    { text: "\uD83C\uDDEB\uD83C\uDDF7  France", options: { bold: true, color: WHITE, fill: { color: ORANGE }, fontSize: 8.5, align: "center" } }
  ],
  [{ text: "\uD83D\uDCDC R\u00e9glementation", options: { fontSize: 8, bold: true } }, { text: "D\u00e9cret n\u00b02-12-349", options: { fontSize: 8 } }, { text: "Code commande publique", options: { fontSize: 8 } }],
  [{ text: "\uD83D\uDCC4 Pi\u00e8ces march\u00e9", options: { fontSize: 8, bold: true } }, { text: "CPS + RC + BPDE", options: { fontSize: 8 } }, { text: "CCAP + CCTP + BPU/DQE", options: { fontSize: 8 } }],
  [{ text: "\uD83D\uDCB0 Estimation", options: { fontSize: 8, bold: true } }, { text: "Confidentielle obligatoire", options: { fontSize: 8 } }, { text: "Estimation MOA", options: { fontSize: 8 } }],
  [{ text: "\uD83D\uDCCF Normes", options: { fontSize: 8, bold: true } }, { text: "Normes marocaines, RPS 2000", options: { fontSize: 8 } }, { text: "DTU, Eurocodes, RE2020", options: { fontSize: 8 } }],
  [{ text: "\uD83D\uDCCA Suivi", options: { fontSize: 8, bold: true } }, { text: "Attachements contradictoires", options: { fontSize: 8 } }, { text: "Situations mensuelles", options: { fontSize: 8 } }],
  [{ text: "\uD83D\uDC65 Commission", options: { fontSize: 8, bold: true } }, { text: "CAO", options: { fontSize: 8 } }, { text: "Commission d\u2019AO", options: { fontSize: 8 } }],
];
s12.addTable(cmpRows, { x: 0.4, y: 1.3, w: 9.2, colW: [2.0, 3.6, 3.6], border: { pt: 0.5, color: "E8E8E8" }, rowH: 0.37, autoPage: false });

// Bottom summary
card(s12, 0.4, 3.95, 4.3, 1.1, TEAL_L, TEAL);
s12.addText([
  { text: "\uD83C\uDDF2\uD83C\uDDE6  C\u00f4t\u00e9 MOA \u2013 Maroc", options: { bold: true, color: TEAL_D, fontSize: 9, breakLine: true } },
  { text: "Concevoir les march\u00e9s, r\u00e9diger les pi\u00e8ces, piloter la CAO. 7 march\u00e9s, 100 M DH.", options: { fontSize: 8, color: BODY } }
], { x: 0.6, y: 4.05, w: 3.9, h: 0.9, fontFace: FB, margin: 0 });

card(s12, 5.3, 3.95, 4.3, 1.1, ORANGE_L, ORANGE);
s12.addText([
  { text: "\uD83C\uDDEB\uD83C\uDDF7  C\u00f4t\u00e9 ex\u00e9cution \u2013 France", options: { bold: true, color: ORANGE, fontSize: 9, breakLine: true } },
  { text: "Banches, ferraillage, planning, contr\u00f4le qualit\u00e9. Co\u00fbts r\u00e9els et rendements.", options: { fontSize: 8, color: BODY } }
], { x: 5.5, y: 4.05, w: 3.9, h: 0.9, fontFace: FB, margin: 0 });

s12.addNotes(`SLIDE 12 - ANALYSE MAROC / FRANCE (1 min 30)
M\u00eames principes : transparence, \u00e9galit\u00e9, mise en concurrence.
CPS \u2194 CCAP, estimation confidentielle \u2194 estimation MOA, attachements \u2194 situations.
MOA Maroc : concevoir, r\u00e9diger, analyser. Ex\u00e9cution France : co\u00fbts r\u00e9els.
La combinaison rend les estimations r\u00e9alistes.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 13 — BILAN REFLEXIF ADM (colored cards with icons)
// ═══════════════════════════════════════════════════════════
let s13 = pres.addSlide();
topBar(s13); badge(s13, "04 \u2013 R\u00e9flexif"); footer(s13, 13);
heading(s13, "BILAN R\u00c9FLEXIF", "Ce que le terrain m\u2019a appris \u2013 8 ans de pratique");

const adm = [
  { letter: "A", icon: "\uD83C\uDF93", title: "CE QUE J\u2019AI APPRIS", color: TEAL, bg: TEAL_L,
    text: "L\u2019\u00e9crit syst\u00e9matique \u2013 OS, attachements, CR \u2013 est le seul rempart contre les litiges. Croiser 3 sources de prix = \u00e9cart 3,2% vs 10% norme r\u00e9gionale. Voir le projet depuis le terrain rend les estimations justes." },
  { letter: "D", icon: "\uD83D\uDD04", title: "CE QUE JE FERAIS DIFF\u00c9REMMENT", color: ORANGE, bg: ORANGE_L,
    text: "TdB d\u00e8s le 1er mois, pas au 9\u00e8me (+4,8%). \u00c9tude g\u00e9otechnique compl\u00e9mentaire avant terrassements. Standardiser les CR d\u00e8s le d\u00e9marrage \u2013 pas dans l\u2019urgence de la semaine 23." },
  { letter: "M", icon: "\uD83C\uDFAF", title: "CE QUE J\u2019APPORTE AU BTS MEC", color: NAVY, bg: CREAM,
    text: "Estimation confidentielle = acte fondateur du march\u00e9 public. En France, m\u00eames fondamentaux (DQE/DPGF). Triple comp\u00e9tence MOA + Ex\u00e9cution + BIM = positionnement rare. BIMCO est n\u00e9 de ce parcours." }
];
adm.forEach((a, i) => {
  const ay = 1.3 + i * 1.28;
  card(s13, 0.4, ay, 9.2, 1.1, a.bg, a.color);
  iconCircle(s13, 0.55, ay + 0.12, a.icon, 0.38, a.color);
  s13.addText(a.title, { x: 1.05, y: ay + 0.08, w: 8.3, h: 0.26, fontSize: 9, bold: true, color: a.color, charSpacing: 1, fontFace: FB, margin: 0 });
  s13.addText(a.text, { x: 1.05, y: ay + 0.38, w: 8.3, h: 0.62, fontSize: 8.5, color: BODY, fontFace: FB, margin: 0 });
});

s13.addNotes(`SLIDE 13 - BILAN R\u00c9FLEXIF (2 min) \u2013 SLIDE IMPORTANTE POUR LE JURY
Structure ADM identique au rapport page 26.

A : La semaine 23 \u2013 l\u2019\u00e9crit syst\u00e9matique est le seul rempart. Croiser 3 sources = 3,2%.
D : TdB d\u00e8s le 1er mois. G\u00e9otech compl\u00e9mentaire. CR standardis\u00e9s d\u00e8s le d\u00e9part.
M : Triple comp\u00e9tence = BIMCO. Le BIM industrialise la m\u00e9thode du terrain.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 14 — PROTOCOLE BIM + PERSPECTIVES (visual pipeline)
// ═══════════════════════════════════════════════════════════
let s14 = pres.addSlide();
topBar(s14); badge(s14, "05 \u2013 BIM"); footer(s14, 14);
heading(s14, "PROTOCOLE BIM & PERSPECTIVES BIMCO", "Formation AFPA Colmar + Projet professionnel 2026\u20132029");

// Left — BIM
card(s14, 0.4, 1.3, 4.5, 3.8, CARD, TEAL);
s14.addText("\uD83D\uDDA5\uFE0F  PROTOCOLE BIM \u2013 CAS AFPA", { x: 0.6, y: 1.42, w: 4.1, h: 0.22, fontSize: 9, bold: true, color: TEAL, fontFace: FB, margin: 0 });

// Mini stats
const bs = [["78", "postes"], ["1,8%", "\u00e9cart"], ["12", "clashs"]];
bs.forEach((b, i) => {
  s14.addShape(pres.shapes.RECTANGLE, { x: 0.6 + i * 1.4, y: 1.75, w: 1.25, h: 0.55, fill: { color: TEAL_L }, rectRadius: 0.04 });
  s14.addText(b[0], { x: 0.6 + i * 1.4, y: 1.76, w: 1.25, h: 0.3, fontSize: 18, bold: true, color: TEAL, fontFace: FT, align: "center", margin: 0 });
  s14.addText(b[1], { x: 0.6 + i * 1.4, y: 2.08, w: 1.25, h: 0.18, fontSize: 7, color: MUTED, fontFace: FB, align: "center", margin: 0 });
});

// Steps with connected line
s14.addShape(pres.shapes.RECTANGLE, { x: 0.72, y: 2.55, w: 0.02, h: 2.3, fill: { color: TEAL_L } });
const bimSteps = [
  ["\u2776", "Mod\u00e9lisation Revit (LOD 300)"],
  ["\u2777", "Export IFC 2x3 \u2013 MVD Coord. View"],
  ["\u2778", "D\u00e9tection clashs Navisworks"],
  ["\u2779", "Extraction Revit + Dynamo (2h vs 2j)"],
  ["\u277A", "Chiffrage avec tra\u00e7abilit\u00e9 maquette"]
];
bimSteps.forEach((step, i) => {
  const sy = 2.5 + i * 0.46;
  iconCircle(s14, 0.58, sy + 0.04, step[0], 0.28, TEAL);
  s14.addText(step[1], { x: 0.95, y: sy + 0.02, w: 3.8, h: 0.34, fontSize: 8, color: BODY, fontFace: FB, margin: 0 });
});

// Right — Perspectives
card(s14, 5.1, 1.3, 4.5, 3.8, CARD, ORANGE);
s14.addText("\uD83D\uDE80  PROJET PROFESSIONNEL", { x: 5.3, y: 1.42, w: 4.1, h: 0.22, fontSize: 9, bold: true, color: ORANGE, fontFace: FB, margin: 0 });

const horizons = [
  { year: "2026", label: "Court terme", color: TEAL, text: "BTS MEC + premi\u00e8res prestations BIMCO : m\u00e9tr\u00e9s BIM, \u00e9tudes de prix, DPGF. Premiers plugins Revit/Dynamo." },
  { year: "2027\u201328", label: "Moyen terme", color: ORANGE, text: "Plugin Revit \u2192 DPGF automatis\u00e9. App web suivi \u00e9conomique. Base prix Batiprix. 3\u20135 clients." },
  { year: "2029+", label: "Long terme", color: NAVY, text: "Cabinet ing\u00e9nierie BIM + \u00e9co construction. SaaS MEC. \u00c9quipe 3\u20135 pers. CA 200\u2013300 k\u20ac/an." }
];
// Connecting line
s14.addShape(pres.shapes.RECTANGLE, { x: 5.62, y: 1.85, w: 0.02, h: 2.95, fill: { color: "DDDDDD" } });
horizons.forEach((h, i) => {
  const hy = 1.78 + i * 1.05;
  s14.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: hy, w: 0.65, h: 0.28, fill: { color: h.color }, rectRadius: 0.14 });
  s14.addText(h.year, { x: 5.3, y: hy, w: 0.65, h: 0.28, fontSize: 7.5, bold: true, color: WHITE, fontFace: FB, align: "center", valign: "middle", margin: 0 });
  s14.addText(h.label, { x: 6.05, y: hy, w: 3.3, h: 0.28, fontSize: 9, bold: true, color: DARK, fontFace: FB, margin: 0 });
  s14.addText(h.text, { x: 5.3, y: hy + 0.32, w: 4.1, h: 0.65, fontSize: 7.5, color: BODY, fontFace: FB, margin: 0 });
});

s14.addNotes(`SLIDE 14 - BIM & PERSPECTIVES (1 min 30)
Protocole BIM AFPA : LOD 300, 78 postes, 2h vs 2j, \u00e9cart 1,8%, 12 clashs.
Workflow 5 \u00e9tapes : mod\u00e9lisation \u2192 IFC \u2192 clashs \u2192 extraction \u2192 chiffrage.
Perspectives : 2026 BTS+BIMCO, 2027-28 plugin+app+3-5 clients, 2029+ cabinet SaaS.`);

// ═══════════════════════════════════════════════════════════
// SLIDE 15 — CONCLUSION (dark premium)
// ═══════════════════════════════════════════════════════════
let s15 = pres.addSlide();
s15.background = { color: NAVY };

// Decorative
s15.addShape(pres.shapes.OVAL, { x: -1, y: -1, w: 3, h: 3, fill: { color: NAVY_L } });
s15.addShape(pres.shapes.OVAL, { x: 7.5, y: 3, w: 4, h: 4, fill: { color: NAVY_L } });

s15.addText("Soutenance BTS MEC U62  \u00b7  Session 2026", { x: 0.5, y: 0.35, w: 5, h: 0.3, fontSize: 9, color: TEAL, fontFace: FB, charSpacing: 2, margin: 0 });
s15.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.8, w: 0.05, h: 1.5, fill: { color: TEAL } });
s15.addText("Merci de votre\nattention", { x: 0.7, y: 0.8, w: 5, h: 1.5, fontSize: 36, bold: true, color: WHITE, fontFace: FT, margin: 0 });

// Takeaways with icons
const takes = [
  ["\uD83C\uDFAF", "5 comp\u00e9tences terrain", "Estimer, analyser, suivre, coordonner, contr\u00f4ler", TEAL],
  ["\uD83C\uDF0D", "Double lecture", "MOA Maroc + ex\u00e9cution France", ORANGE],
  ["\u270D\uFE0F", "Rigueur de l\u2019\u00e9crit", "OS, attachements, CR \u2013 le rempart", GREEN],
  ["\uD83D\uDE80", "BIMCO", "Rigueur terrain + outils num\u00e9riques", TEAL]
];
takes.forEach((t, i) => {
  const ty = 0.65 + i * 1.05;
  s15.addShape(pres.shapes.RECTANGLE, { x: 6, y: ty, w: 3.6, h: 0.88, fill: { color: WHITE }, shadow: shd(), rectRadius: 0.06 });
  s15.addShape(pres.shapes.RECTANGLE, { x: 6, y: ty, w: 0.06, h: 0.88, fill: { color: t[3] } });
  iconCircle(s15, 6.15, ty + 0.12, t[0], 0.32, t[3]);
  s15.addText([
    { text: t[1], options: { bold: true, color: DARK, fontSize: 10, breakLine: true } },
    { text: t[2], options: { color: MUTED, fontSize: 8 } }
  ], { x: 6.55, y: ty + 0.1, w: 2.9, h: 0.68, fontFace: FB, valign: "middle", margin: 0 });
});

s15.addText([
  { text: "BAHAFID Mohamed  \u00b7  N\u00b0 02537399911  \u00b7  Acad\u00e9mie de Lyon", options: { breakLine: true } },
  { text: "BIMCO  |  gestion.bimco-consulting.fr  |  Bussi\u00e8res, Loire 42510", options: {} }
], { x: 0.7, y: 4.55, w: 5, h: 0.8, fontSize: 9, color: TEAL, fontFace: FB, margin: 0 });

s15.addNotes(`SLIDE 15 - CONCLUSION (1 min)
4 messages :
1. 5 comp\u00e9tences du BTS MEC pratiqu\u00e9es en situation r\u00e9elle.
2. Double lecture : MOA Maroc + ex\u00e9cution France.
3. La rigueur de l\u2019\u00e9crit \u2013 le seul rempart contre les litiges.
4. BIMCO \u2013 la rigueur terrain + les outils num\u00e9riques.
Je suis \u00e0 votre disposition pour vos questions.`);

// ═══════════════════════════════════════════════════════════
pres.writeFile({ fileName: "D:/PREPA BTS MEC/08_U62_Rapport_Activites/Soutenance_PowerPoint/SOUTENANCE_U62_v4.pptx" })
  .then(() => console.log("OK: SOUTENANCE_U62_v4.pptx generated"))
  .catch(e => console.error("ERROR:", e));
