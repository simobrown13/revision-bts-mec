const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "BAHAFID Mohamed";
pres.title = "Soutenance U62 - BTS MEC 2026";

const C = {
  navy: "1E3A5F", darkNavy: "0F2B50", deepNavy: "002F6C",
  orange: "F39200", orangeDark: "E98800",
  teal: "5CC8C0", tealDark: "48B4AC", tealMid: "52BEB6",
  white: "FFFFFF", offWhite: "F5F5F5", lightGray: "EEEEEE",
  textDark: "2D2D2D", textMid: "555555", textLight: "888888",
};
const mkShadow = () => ({ type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.12 });
const TOTAL = 10;

function footer(s, n) {
  s.addText(`BAHAFID Mohamed  |  BTS MEC 2026  |  ${n}/${TOTAL}`, { x: 0.5, y: 5.28, w: 9, h: 0.25, fontSize: 7, color: C.textLight, fontFace: "Inter" });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.5, w: 10, h: 0.125, fill: { color: C.orange } });
}
function heading(s, title) {
  s.background = { color: C.offWhite };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.navy } });
  s.addText(title, { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 22, fontFace: "Montserrat Bold", color: C.navy, bold: true, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.72, w: 1.5, h: 0.05, fill: { color: C.orange } });
}

// ============================================================
// SLIDE 1 - COUVERTURE
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.orange } });
  s.addText("SESSION 2026  |  Académie de Lyon", { x: 0.8, y: 0.5, w: 5, h: 0.35, fontSize: 11, fontFace: "Inter", color: C.teal, charSpacing: 2 });
  s.addText("RAPPORT\nD'ACTIVITÉS\nPROFESSIONNELLES", { x: 0.8, y: 1.3, w: 6, h: 2.5, fontSize: 42, fontFace: "Montserrat Bold", color: C.white, bold: true, lineSpacingMultiple: 1.05 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 3.95, w: 2.5, h: 0.06, fill: { color: C.orange } });
  s.addText("BAHAFID Mohamed", { x: 0.8, y: 4.2, w: 5, h: 0.45, fontSize: 20, fontFace: "Montserrat Bold", color: C.orange, bold: true });
  s.addText("BTS Management Économique de la Construction\nN° 02537399911", { x: 0.8, y: 4.65, w: 5, h: 0.5, fontSize: 11, fontFace: "Inter", color: C.tealMid });
}

// ============================================================
// SLIDE 2 - MON PARCOURS
// ============================================================
{
  const s = pres.addSlide();
  heading(s, "MON PARCOURS");

  // Timeline horizontal
  const phases = [
    { period: "2017 - 2022", role: "Maîtrise d'ouvrage", place: "Conseil Régional BMK · Maroc", color: C.tealDark },
    { period: "2022 - 2024", role: "Exécution terrain", place: "Chef de chantier GO · France", color: C.orange },
    { period: "2024", role: "Formation BIM", place: "AFPA Colmar · 8 mois", color: C.navy },
    { period: "2026", role: "BIMCO", place: "BIM + Économie construction", color: C.deepNavy },
  ];
  phases.forEach((p, i) => {
    const xx = 0.5 + i * 2.35;
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: 0.95, w: 2.15, h: 1.35, fill: { color: C.white }, shadow: mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: 0.95, w: 2.15, h: 0.06, fill: { color: p.color } });
    s.addText(p.period, { x: xx + 0.12, y: 1.1, w: 1.9, h: 0.3, fontSize: 12, fontFace: "Montserrat Bold", color: p.color, bold: true, margin: 0 });
    s.addText(p.role, { x: xx + 0.12, y: 1.4, w: 1.9, h: 0.3, fontSize: 11, fontFace: "Montserrat Bold", color: C.textDark, bold: true, margin: 0 });
    s.addText(p.place, { x: xx + 0.12, y: 1.7, w: 1.9, h: 0.4, fontSize: 9, fontFace: "Inter", color: C.textMid, margin: 0 });
  });

  // Les 2 projets
  s.addText("DEUX PROJETS MAJEURS", { x: 0.5, y: 2.55, w: 9, h: 0.35, fontSize: 14, fontFace: "Montserrat Bold", color: C.navy, bold: true, margin: 0 });

  // Projet 1
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.0, w: 4.3, h: 1.2, fill: { color: C.white }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.0, w: 0.08, h: 1.2, fill: { color: C.orange } });
  s.addText("Projet 1 · Aménagement urbain", { x: 0.75, y: 3.05, w: 3.9, h: 0.35, fontSize: 13, fontFace: "Montserrat Bold", color: C.navy, bold: true, margin: 0 });
  s.addText("53,5 M DH TTC (~4,8 M€)\n4 communes · 8 corps d'état · Province de Khénifra\nSituations 1 à 4", { x: 0.75, y: 3.4, w: 3.9, h: 0.7, fontSize: 10, fontFace: "Inter", color: C.textMid, margin: 0 });

  // Projet 2
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.0, w: 4.3, h: 1.2, fill: { color: C.white }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.0, w: 0.08, h: 1.2, fill: { color: C.teal } });
  s.addText("Projet 2 · Route Lehri-Kerrouchen", { x: 5.45, y: 3.05, w: 3.9, h: 0.35, fontSize: 13, fontFace: "Montserrat Bold", color: C.navy, bold: true, margin: 0 });
  s.addText("29 M DH TTC (~2,6 M€)\n25 km en zone montagneuse · Moyen Atlas\nSituation 5", { x: 5.45, y: 3.4, w: 3.9, h: 0.7, fontSize: 10, fontFace: "Inter", color: C.textMid, margin: 0 });

  // Méthode
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.45, w: 9, h: 0.5, fill: { color: C.navy } });
  s.addText("5 situations professionnelles · Méthode CPAR : Contexte · Problème · Action · Résultat", {
    x: 0.6, y: 4.47, w: 8.8, h: 0.46, fontSize: 12, fontFace: "Montserrat Bold", color: C.white, bold: true, valign: "middle", margin: 0,
  });

  footer(s, 2);
}

// ============================================================
// SLIDE 3 - CADRE PROFESSIONNEL
// ============================================================
{
  const s = pres.addSlide();
  heading(s, "CADRE PROFESSIONNEL");

  // LEFT - Structure
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.95, w: 4.3, h: 3.7, fill: { color: C.white }, shadow: mkShadow() });
  s.addText("CONSEIL RÉGIONAL BMK", { x: 0.7, y: 1.05, w: 3.9, h: 0.35, fontSize: 14, fontFace: "Montserrat Bold", color: C.navy, bold: true, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 1.4, w: 1, h: 0.04, fill: { color: C.orange } });
  s.addText("Agence d'Exécution des Projets\nMaîtrise d'ouvrage des infrastructures régionales", {
    x: 0.7, y: 1.55, w: 3.9, h: 0.5, fontSize: 10, fontFace: "Inter", color: C.textMid, margin: 0,
  });
  const missions = ["Métrés et estimation confidentielle", "Rédaction des DCE (CPS, RC, BPDE)", "Analyse des offres en commission AO", "Suivi financier mensuel", "Visites terrain et attachements contradictoires"];
  s.addText("Mes missions :", { x: 0.7, y: 2.15, w: 3.9, h: 0.25, fontSize: 10, fontFace: "Montserrat Bold", color: C.navy, bold: true, margin: 0 });
  s.addText(missions.map((m, i) => ({ text: m, options: { bullet: true, breakLine: i < missions.length - 1, fontSize: 10, color: C.textDark } })),
    { x: 0.7, y: 2.45, w: 3.9, h: 2.0, fontFace: "Inter", paraSpaceAfter: 6 });

  // RIGHT - BIMCO
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 0.95, w: 4.3, h: 3.7, fill: { color: C.white }, shadow: mkShadow() });
  s.addText("BIMCO", { x: 5.4, y: 1.05, w: 2, h: 0.35, fontSize: 14, fontFace: "Montserrat Bold", color: C.orange, bold: true, margin: 0 });
  s.addText("Mon activité indépendante", { x: 7.2, y: 1.07, w: 2.2, h: 0.3, fontSize: 10, fontFace: "Inter", color: C.textMid, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 1.4, w: 1, h: 0.04, fill: { color: C.teal } });
  s.addText("Le BIM au service de l'économiste de la construction", {
    x: 5.4, y: 1.55, w: 3.9, h: 0.35, fontSize: 10, fontFace: "Inter", color: C.textMid, margin: 0,
  });
  const bimItems = ["Métrés par extraction de maquette numérique", "Études de prix ancrées dans les coûts réels", "Développement plugins Revit / Dynamo", "Automatisation de la chaîne métré → DPGF"];
  s.addText(bimItems.map((m, i) => ({ text: m, options: { bullet: true, breakLine: i < bimItems.length - 1, fontSize: 10, color: C.textDark } })),
    { x: 5.4, y: 2.0, w: 3.9, h: 1.4, fontFace: "Inter", paraSpaceAfter: 6 });

  // Tools row
  const tools = [
    { t: "BIM & CAO", v: "Revit · Navisworks · Dynamo · AutoCAD" },
    { t: "Développement", v: "Python · C# · React · Node.js" },
    { t: "Gestion", v: "MS Project · Excel · BIM360" },
  ];
  tools.forEach((t, i) => {
    const yy = 3.5 + i * 0.36;
    s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: yy, w: 3.9, h: 0.32, fill: { color: i === 0 ? C.navy : i === 1 ? C.tealDark : C.darkNavy } });
    s.addText(t.t, { x: 5.5, y: yy, w: 1.3, h: 0.32, fontSize: 8, fontFace: "Montserrat Bold", color: C.orange, bold: true, valign: "middle", margin: 0 });
    s.addText(t.v, { x: 6.8, y: yy, w: 2.4, h: 0.32, fontSize: 8, fontFace: "Inter", color: C.white, valign: "middle", margin: 0 });
  });
  footer(s, 3);
}

// ============================================================
// SLIDE 4 - PROJET 1 + BUDGET (graphique)
// ============================================================
{
  const s = pres.addSlide();
  heading(s, "PROJET 1 : MISE À NIVEAU DE 4 COMMUNES");

  s.addText("Aménagement urbain et VRD · 4 communes · 8 corps d'état · 44,6 M DH HT (4,86 M€)\nDe l'assainissement à l'éclairage public, sur 4 sites distants de 20 à 80 km", {
    x: 0.5, y: 0.9, w: 9, h: 0.5, fontSize: 10, fontFace: "Inter", color: C.textMid, margin: 0,
  });

  // Bar chart
  s.addChart(pres.charts.BAR, [{
    name: "Budget par commune (M DH HT)",
    labels: ["Ouaoumana", "Sebt Ait Rahou", "Kerrouchen", "El Hammam"],
    values: [15.8, 14.8, 7.3, 6.7],
  }], {
    x: 0.3, y: 1.55, w: 5, h: 2.8, barDir: "col",
    chartColors: [C.orange],
    chartArea: { fill: { color: C.white }, roundedCorners: true },
    catAxisLabelColor: C.textMid, catAxisLabelFontSize: 9, catAxisLabelFontFace: "Inter",
    valAxisLabelColor: C.textLight, valAxisLabelFontSize: 8,
    valGridLine: { color: C.lightGray, size: 0.5 }, catGridLine: { style: "none" },
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.navy,
    dataLabelFontSize: 11, dataLabelFontFace: "Montserrat Bold", showLegend: false,
  });

  // Right - 8 corps d'état (simple list)
  s.addText("8 corps d'état", { x: 5.6, y: 1.55, w: 4, h: 0.35, fontSize: 14, fontFace: "Montserrat Bold", color: C.navy, bold: true, margin: 0 });
  const corps = ["Assainissement & réseaux", "Travaux de chaussée", "Aménagement trottoirs", "Signalisation", "Éclairage public", "Murs & ouvrages", "Aménagement paysager", "Mobilier urbain"];
  corps.forEach((c, i) => {
    const yy = 2.0 + i * 0.32;
    s.addShape(pres.shapes.RECTANGLE, { x: 5.6, y: yy, w: 0.28, h: 0.25, fill: { color: C.navy } });
    s.addText(String(i + 1).padStart(2, "0"), { x: 5.6, y: yy, w: 0.28, h: 0.25, fontSize: 8, fontFace: "Montserrat Bold", color: C.orange, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(c, { x: 6.0, y: yy, w: 3.5, h: 0.25, fontSize: 10, fontFace: "Inter", color: C.textDark, valign: "middle", margin: 0 });
  });

  // Bottom callout
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.6, w: 9, h: 0.4, fill: { color: C.navy } });
  s.addText("Mes missions : avant-métrés · estimation confidentielle · commission AO · suivi financier 18 mois", {
    x: 0.6, y: 4.62, w: 8.8, h: 0.36, fontSize: 10, fontFace: "Inter", color: C.white, valign: "middle", margin: 0,
  });
  footer(s, 4);
}

// ============================================================
// SLIDE 5 - SITUATIONS 1 & 2
// ============================================================
{
  const s = pres.addSlide();
  heading(s, "SITUATIONS 1 & 2");

  // S1 LEFT
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.95, w: 4.3, h: 4.0, fill: { color: C.white }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.95, w: 4.3, h: 0.5, fill: { color: C.navy } });
  s.addText("S1 · Estimation confidentielle", { x: 0.65, y: 0.95, w: 4, h: 0.5, fontSize: 13, fontFace: "Montserrat Bold", color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText("Réaliser des métrés TCE et estimer un ouvrage", { x: 0.65, y: 1.55, w: 4, h: 0.25, fontSize: 9, fontFace: "Inter", color: C.tealDark, italic: true, margin: 0 });
  // Simple CPAR
  const s1 = [
    { l: "C", t: "Commune d'Ouaoumana · 15,8 M DH HT\nEstimation à livrer en 3 semaines" },
    { l: "P", t: "Plans incohérents du BET\nMercuriale de référence obsolète" },
    { l: "A", t: "Métrés AutoCAD + visites terrain\nCroisement de 3 sources de prix" },
    { l: "R", t: "Écart de seulement 3,2% avec l'offre retenue\nMéthode adoptée pour les 3 autres communes" },
  ];
  s1.forEach((c, i) => {
    const yy = 1.95 + i * 0.7;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.65, y: yy, w: 0.35, h: 0.35, fill: { color: c.l === "R" ? C.tealDark : C.navy } });
    s.addText(c.l, { x: 0.65, y: yy, w: 0.35, h: 0.35, fontSize: 13, fontFace: "Montserrat Bold", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(c.t, { x: 1.1, y: yy, w: 3.55, h: 0.6, fontSize: 10, fontFace: "Inter", color: C.textDark, valign: "top", margin: 0 });
  });

  // S2 RIGHT
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 0.95, w: 4.3, h: 4.0, fill: { color: C.white }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 0.95, w: 4.3, h: 0.5, fill: { color: C.navy } });
  s.addText("S2 · Analyse des offres", { x: 5.35, y: 0.95, w: 4, h: 0.5, fontSize: 13, fontFace: "Montserrat Bold", color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText("Analyser les offres et préparer la décision en commission", { x: 5.35, y: 1.55, w: 4, h: 0.25, fontSize: 9, fontFace: "Inter", color: C.tealDark, italic: true, margin: 0 });
  const s2 = [
    { l: "C", t: "3 entreprises candidates\nAnalyse comparative pour la commission" },
    { l: "P", t: "Erreurs arithmétiques dans un dossier\nPrix anormalement bas sur des postes majeurs" },
    { l: "A", t: "Vérification de conformité\nGrille comparative technique et financière" },
    { l: "R", t: "Entreprise retenue notée 94/100\nAttribution sans aucun recours" },
  ];
  s2.forEach((c, i) => {
    const yy = 1.95 + i * 0.7;
    s.addShape(pres.shapes.RECTANGLE, { x: 5.35, y: yy, w: 0.35, h: 0.35, fill: { color: c.l === "R" ? C.tealDark : C.navy } });
    s.addText(c.l, { x: 5.35, y: yy, w: 0.35, h: 0.35, fontSize: 13, fontFace: "Montserrat Bold", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(c.t, { x: 5.8, y: yy, w: 3.55, h: 0.6, fontSize: 10, fontFace: "Inter", color: C.textDark, valign: "top", margin: 0 });
  });
  footer(s, 5);
}

// ============================================================
// SLIDE 6 - SITUATIONS 3 & 4
// ============================================================
{
  const s = pres.addSlide();
  heading(s, "SITUATIONS 3 & 4");

  // S3 LEFT
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.95, w: 4.3, h: 4.0, fill: { color: C.white }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.95, w: 4.3, h: 0.5, fill: { color: C.navy } });
  s.addText("S3 · Suivi financier", { x: 0.65, y: 0.95, w: 4, h: 0.5, fontSize: 13, fontFace: "Montserrat Bold", color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText("Suivre l'exécution financière et anticiper les dérives", { x: 0.65, y: 1.55, w: 4, h: 0.25, fontSize: 9, fontFace: "Inter", color: C.tealDark, italic: true, margin: 0 });

  // Line chart inside S3
  s.addChart(pres.charts.LINE, [
    { name: "Consommé", labels: ["M1", "M6", "M9", "M12", "M18"], values: [5, 35, 52, 70, 100.8] },
    { name: "Prévu", labels: ["M1", "M6", "M9", "M12", "M18"], values: [5.5, 33, 50, 66, 100] },
  ], {
    x: 0.6, y: 1.85, w: 4, h: 1.5,
    chartColors: [C.orange, C.teal], lineSize: 2, lineSmooth: true,
    chartArea: { fill: { color: C.white }, roundedCorners: true },
    catAxisLabelColor: C.textMid, catAxisLabelFontSize: 7,
    valAxisLabelColor: C.textLight, valAxisLabelFontSize: 7,
    valGridLine: { color: C.lightGray, size: 0.5 }, catGridLine: { style: "none" },
    showLegend: true, legendPos: "b", legendFontSize: 7, legendColor: C.textMid,
  });

  const s3 = [
    { l: "C", t: "Chantier de Kerrouchen · 7,3 M DH sur 18 mois\nTerrain rocheux imprévu" },
    { l: "P", t: "Dépassement extrapolé +4,8% · 349 000 DH\nSeuil d'avenant à 5%" },
    { l: "A", t: "Tableau de bord hebdomadaire\nCompensation inter-postes · 56 000 DH dégagés" },
    { l: "R", t: "Dépassement maîtrisé, avenant évité\nOutil adopté par l'Agence" },
  ];
  s3.forEach((c, i) => {
    const yy = 3.45 + i * 0.37;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.65, y: yy, w: 0.28, h: 0.28, fill: { color: c.l === "R" || c.l === "A" ? C.tealDark : C.navy } });
    s.addText(c.l, { x: 0.65, y: yy, w: 0.28, h: 0.28, fontSize: 10, fontFace: "Montserrat Bold", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(c.t, { x: 1.02, y: yy - 0.02, w: 3.65, h: 0.35, fontSize: 9, fontFace: "Inter", color: C.textDark, valign: "top", margin: 0 });
  });

  // S4 RIGHT
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 0.95, w: 4.3, h: 4.0, fill: { color: C.white }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 0.95, w: 4.3, h: 0.5, fill: { color: C.navy } });
  s.addText("S4 · Coordination multi-sites", { x: 5.35, y: 0.95, w: 4, h: 0.5, fontSize: 13, fontFace: "Montserrat Bold", color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText("Communiquer et coordonner en contexte de crise", { x: 5.35, y: 1.55, w: 4, h: 0.25, fontSize: 9, fontFace: "Inter", color: C.tealDark, italic: true, margin: 0 });

  const s4 = [
    { l: "C", t: "4 chantiers simultanés\nRelais unique du Directeur" },
    { l: "P", t: "3 crises simultanées en semaine 23\nRetard · Météo · Litige quantités" },
    { l: "A", t: "Traitement des 3 fronts en 48h\nNote factuelle · OS d'arrêt · Re-mesurage" },
    { l: "R", t: "Retard rattrapé · Litige résolu\nModèle de CR adopté par l'Agence" },
  ];
  s4.forEach((c, i) => {
    const yy = 1.95 + i * 0.7;
    s.addShape(pres.shapes.RECTANGLE, { x: 5.35, y: yy, w: 0.35, h: 0.35, fill: { color: c.l === "R" || c.l === "A" ? C.tealDark : C.navy } });
    s.addText(c.l, { x: 5.35, y: yy, w: 0.35, h: 0.35, fontSize: 13, fontFace: "Montserrat Bold", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(c.t, { x: 5.8, y: yy, w: 3.55, h: 0.6, fontSize: 10, fontFace: "Inter", color: C.textDark, valign: "top", margin: 0 });
  });
  footer(s, 6);
}

// ============================================================
// SLIDE 7 - PROJET 2 + SITUATION 5
// ============================================================
{
  const s = pres.addSlide();
  heading(s, "PROJET 2 : ROUTE LEHRI-KERROUCHEN");

  s.addText("Route rurale de 25 km en zone montagneuse · Moyen Atlas · 24,2 M DH HT (2,6 M€)\nDénivelé important · Pentes raides · 3 sections · Programme National Routes Rurales", {
    x: 0.5, y: 0.9, w: 5.5, h: 0.5, fontSize: 10, fontFace: "Inter", color: C.textMid, margin: 0,
  });

  // Pie chart
  s.addChart(pres.charts.PIE, [{
    name: "Budget", labels: ["Terrassement", "Corps chaussée", "Revêtement", "Ouvrages hydr.", "Soutènement", "Bretelles"],
    values: [17.9, 30.1, 10.7, 28.1, 1.9, 11.3],
  }], {
    x: 5.5, y: 0.85, w: 4.2, h: 2.8,
    chartColors: [C.navy, C.orange, C.teal, C.tealDark, C.textMid, C.darkNavy],
    showPercent: true, showLegend: true, legendPos: "b", legendFontSize: 7, legendColor: C.textMid,
    dataLabelColor: C.white, dataLabelFontSize: 8,
  });

  // Situation 5
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.55, w: 4.7, h: 0.45, fill: { color: C.navy } });
  s.addText("S5 · Contrôle des cubatures de terrassement", { x: 0.65, y: 1.57, w: 4.4, h: 0.4, fontSize: 12, fontFace: "Montserrat Bold", color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText("Vérifier les quantités exécutées et contrôler les écarts", { x: 0.65, y: 2.1, w: 4.5, h: 0.25, fontSize: 9, fontFace: "Inter", color: C.tealDark, italic: true, margin: 0 });

  const s5 = [
    { l: "C", t: "Vérifier les cubatures du BET\nsur 25 km de route en montagne" },
    { l: "P", t: "Calcaire fracturé au PK 12 · surcoût 285 000 DH\n3 ouvrages hydrauliques non prévus" },
    { l: "A", t: "Contrôle contradictoire GPS tous les 500 m\nDimensionnement de 3 dalots avec le BET" },
    { l: "R", t: "Surcoût absorbé par compensation\nRéception dans les délais, sans avenant" },
  ];
  s5.forEach((c, i) => {
    const yy = 2.5 + i * 0.6;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.65, y: yy, w: 0.3, h: 0.3, fill: { color: c.l === "R" || c.l === "A" ? C.tealDark : C.navy } });
    s.addText(c.l, { x: 0.65, y: yy, w: 0.3, h: 0.3, fontSize: 11, fontFace: "Montserrat Bold", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(c.t, { x: 1.05, y: yy - 0.02, w: 4.1, h: 0.55, fontSize: 10, fontFace: "Inter", color: C.textDark, valign: "top", margin: 0 });
  });

  footer(s, 7);
}

// ============================================================
// SLIDE 8 - SYNTHÈSE + COMPARAISON MAROC/FRANCE
// ============================================================
{
  const s = pres.addSlide();
  heading(s, "SYNTHÈSE ET DOUBLE EXPÉRIENCE");

  // Compétences acquises - simple visual
  s.addText("COMPÉTENCES ACQUISES", { x: 0.5, y: 0.9, w: 9, h: 0.3, fontSize: 12, fontFace: "Montserrat Bold", color: C.navy, bold: true, margin: 0 });
  const comps = [
    { title: "Métrés TCE", sub: "Maîtrise", sit: "S1" },
    { title: "Estimation", sub: "Maîtrise", sit: "S1" },
    { title: "Analyse offres", sub: "Maîtrise", sit: "S2" },
    { title: "Suivi financier", sub: "Expert", sit: "S3" },
    { title: "Contrôle quantités", sub: "Expert", sit: "S5" },
    { title: "Communication", sub: "Maîtrise", sit: "S4" },
    { title: "Marchés publics", sub: "Maîtrise", sit: "S1-4" },
    { title: "BIM (Open BIM)", sub: "Maîtrise", sit: "BIM" },
  ];
  comps.forEach((c, i) => {
    const col = i % 4, row = Math.floor(i / 4);
    const xx = 0.5 + col * 2.35, yy = 1.3 + row * 0.75;
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: yy, w: 2.15, h: 0.65, fill: { color: C.white }, shadow: mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: yy, w: 2.15, h: 0.05, fill: { color: c.sub === "Expert" ? C.orange : C.tealDark } });
    s.addText(c.title, { x: xx + 0.1, y: yy + 0.1, w: 1.95, h: 0.25, fontSize: 10, fontFace: "Montserrat Bold", color: C.textDark, bold: true, margin: 0 });
    s.addText(`${c.sub} · ${c.sit}`, { x: xx + 0.1, y: yy + 0.35, w: 1.95, h: 0.2, fontSize: 8, fontFace: "Inter", color: c.sub === "Expert" ? C.orange : C.tealDark, margin: 0 });
  });

  // Maroc vs France - simple 3 cards
  s.addText("DOUBLE EXPÉRIENCE MAROC / FRANCE", { x: 0.5, y: 3.0, w: 9, h: 0.3, fontSize: 12, fontFace: "Montserrat Bold", color: C.navy, bold: true, margin: 0 });

  const dbl = [
    { title: "Côté MOA · Maroc", desc: "Concevoir les marchés, rédiger les pièces,\npiloter la commission, émettre les OS", color: C.tealDark },
    { title: "Côté Exécution · France", desc: "Comprendre les coûts réels, les rendements,\nles contraintes terrain", color: C.orange },
    { title: "Triple compétence", desc: "MOA + Exécution + BIM\n→ Estimations réalistes, analyse pertinente", color: C.navy },
  ];
  dbl.forEach((d, i) => {
    const xx = 0.5 + i * 3.1;
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: 3.4, w: 2.9, h: 1.2, fill: { color: C.white }, shadow: mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: 3.4, w: 2.9, h: 0.05, fill: { color: d.color } });
    s.addText(d.title, { x: xx + 0.12, y: 3.5, w: 2.65, h: 0.3, fontSize: 11, fontFace: "Montserrat Bold", color: d.color, bold: true, margin: 0 });
    s.addText(d.desc, { x: xx + 0.12, y: 3.85, w: 2.65, h: 0.65, fontSize: 9, fontFace: "Inter", color: C.textDark, margin: 0 });
  });
  footer(s, 8);
}

// ============================================================
// SLIDE 9 - BIM + PROJET PROFESSIONNEL
// ============================================================
{
  const s = pres.addSlide();
  heading(s, "BIM ET PROJET PROFESSIONNEL");

  // BIM section
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.95, w: 4.3, h: 2.8, fill: { color: C.white }, shadow: mkShadow() });
  s.addText("PROTOCOLE BIM", { x: 0.7, y: 1.0, w: 3.9, h: 0.3, fontSize: 13, fontFace: "Montserrat Bold", color: C.navy, bold: true, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 1.3, w: 1, h: 0.04, fill: { color: C.orange } });
  s.addText("Cas appliqué : Bâtiment R+2 (Formation AFPA Colmar)\nMaquette Revit · Logements collectifs", {
    x: 0.7, y: 1.45, w: 3.9, h: 0.45, fontSize: 9, fontFace: "Inter", color: C.textMid, margin: 0,
  });
  const bimRes = [
    "78 postes de métrés extraits automatiquement",
    "Écart de seulement 1,8% vs métré manuel",
    "12 conflits structure/réseaux détectés en amont",
    "Extraction en 2h vs 2 jours en traditionnel",
  ];
  s.addText(bimRes.map((m, i) => ({ text: m, options: { bullet: true, breakLine: i < bimRes.length - 1, fontSize: 10, color: C.textDark } })),
    { x: 0.7, y: 2.0, w: 3.9, h: 1.5, fontFace: "Inter", paraSpaceAfter: 6 });

  // Projet professionnel
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 0.95, w: 4.3, h: 2.8, fill: { color: C.white }, shadow: mkShadow() });
  s.addText("PROJET BIMCO", { x: 5.4, y: 1.0, w: 3.9, h: 0.3, fontSize: 13, fontFace: "Montserrat Bold", color: C.orange, bold: true, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 1.3, w: 1, h: 0.04, fill: { color: C.teal } });

  const timeline = [
    { year: "2026", items: "Obtenir le BTS MEC\nPremières prestations BIMCO" },
    { year: "2027-28", items: "Métrés BIM · Études de prix\nPremiers plugins Revit/Dynamo" },
    { year: "2029+", items: "Cabinet BIM + Économie\nSaaS pour économistes MEC" },
  ];
  timeline.forEach((t, i) => {
    const yy = 1.45 + i * 0.72;
    s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: yy, w: 3.9, h: 0.65, fill: { color: C.offWhite } });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: yy, w: 0.06, h: 0.65, fill: { color: i === 0 ? C.tealDark : i === 1 ? C.orange : C.navy } });
    s.addText(t.year, { x: 5.55, y: yy + 0.03, w: 0.8, h: 0.25, fontSize: 11, fontFace: "Montserrat Bold", color: C.navy, bold: true, margin: 0 });
    s.addText(t.items, { x: 6.4, y: yy + 0.03, w: 2.8, h: 0.58, fontSize: 9, fontFace: "Inter", color: C.textDark, margin: 0 });
  });

  // Bottom callout
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.0, w: 9, h: 0.7, fill: { color: C.navy } });
  s.addText("Les outils numériques doivent être au service de l'économiste de la construction, et non l'inverse.", {
    x: 0.7, y: 4.02, w: 8.6, h: 0.3, fontSize: 12, fontFace: "Montserrat Bold", color: C.white, bold: true, margin: 0, italic: true,
  });
  s.addText("Ma double compétence  :  Économiste terrain  +  Développeur BIM  ·  Un avantage rare sur le marché", {
    x: 0.7, y: 4.35, w: 8.6, h: 0.3, fontSize: 10, fontFace: "Inter", color: C.teal, margin: 0,
  });

  footer(s, 9);
}

// ============================================================
// SLIDE 10 - CONCLUSION
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.orange } });
  s.addText("CONCLUSION", { x: 0.5, y: 0.4, w: 9, h: 0.5, fontSize: 30, fontFace: "Montserrat Bold", color: C.white, bold: true, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.95, w: 1.5, h: 0.05, fill: { color: C.orange } });

  const points = [
    { num: "1", title: "Cinq compétences construites sur le terrain", desc: "Estimer, analyser, suivre, coordonner, contrôler — des réflexes acquis par la pratique" },
    { num: "2", title: "Une double lecture des projets", desc: "Côté MOA au Maroc, côté exécution en France — deux visions qui enrichissent l'estimation" },
    { num: "3", title: "La rigueur de l'écrit comme protection", desc: "OS, attachements, comptes-rendus — la trace écrite est le rempart contre les litiges" },
    { num: "4", title: "BIMCO : la traduction de ce parcours", desc: "La rigueur du terrain rencontre les outils numériques pour un chiffrage plus précis" },
  ];
  points.forEach((p, i) => {
    const yy = 1.25 + i * 0.95;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yy, w: 9, h: 0.8, fill: { color: C.darkNavy } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yy, w: 0.55, h: 0.8, fill: { color: C.orange } });
    s.addText(p.num, { x: 0.5, y: yy, w: 0.55, h: 0.8, fontSize: 24, fontFace: "Montserrat Bold", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.title, { x: 1.2, y: yy + 0.08, w: 8.2, h: 0.35, fontSize: 13, fontFace: "Montserrat Bold", color: C.orange, bold: true, margin: 0 });
    s.addText(p.desc, { x: 1.2, y: yy + 0.43, w: 8.2, h: 0.3, fontSize: 10, fontFace: "Inter", color: C.white, margin: 0 });
  });

  // Bottom
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.35, w: 10, h: 0.275, fill: { color: C.orange } });
  s.addText("BAHAFID Mohamed  |  BIMCO  |  BTS MEC Session 2026  |  Académie de Lyon", {
    x: 0.5, y: 5.1, w: 9, h: 0.2, fontSize: 9, fontFace: "Inter", color: C.teal, align: "center", margin: 0,
  });
}

// ============================================================
const out = "D:/PREPA BTS MEC/08_U62_Rapport_Activites/Soutenance_PowerPoint/v22/Soutenance_U62_16x9_v6.pptx";
pres.writeFile({ fileName: out }).then(() => console.log("OK: " + out)).catch(e => console.error(e));
