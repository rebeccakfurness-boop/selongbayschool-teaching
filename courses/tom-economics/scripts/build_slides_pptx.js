// Builds slides.pptx for a lesson, reproducing the content of that lesson's slides.html
// as a real presentable PowerPoint deck. Palette: dark navy anchor + coral accent + gold
// "reward" tag color — kept consistent across lessons as the course's visual identity.
const pptxgen = require("pptxgenjs");
const path = require("path");

const NAVY = "1E2761";
const CORAL = "F96167";
const GOLD = "D98C00";
const WHITE = "FFFFFF";
const CARD = "F4F5F7";
const INK = "1A1A1A";
const MUTED = "5B5B5B";

function titleSlide(pres, eyebrow, title, subtitle) {
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addText(eyebrow.toUpperCase(), { x: 0.7, y: 2.3, w: 11.9, h: 0.5, fontFace: "Calibri", fontSize: 14, color: "CADCFC", charSpacing: 2, bold: true });
  s.addText(title, { x: 0.7, y: 2.8, w: 11.9, h: 1.6, fontFace: "Cambria", fontSize: 44, color: WHITE, bold: true });
  s.addText(subtitle, { x: 0.7, y: 4.3, w: 10.5, h: 0.8, fontFace: "Calibri", fontSize: 18, color: "CADCFC" });
  return s;
}

function closingSlide(pres, eyebrow, title, subtitle) {
  return titleSlide(pres, eyebrow, title, subtitle);
}

function iconCircle(s, x, y, d, emoji, fill) {
  s.addShape("ellipse", { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" } });
  s.addText(emoji, { x, y, w: d, h: d, align: "center", valign: "middle", fontSize: Math.round(d * 28) });
}

function factorSlide(pres, { eyebrow, emoji, title, body, reward }) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  iconCircle(s, 0.7, 0.6, 1.1, emoji, CORAL);
  s.addText(eyebrow, { x: 2.0, y: 0.65, w: 6, h: 0.4, fontFace: "Calibri", fontSize: 13, color: MUTED, bold: true, charSpacing: 1 });
  s.addText(title, { x: 2.0, y: 1.0, w: 8, h: 0.8, fontFace: "Cambria", fontSize: 34, color: NAVY, bold: true });
  s.addShape("roundRect", { x: 0.7, y: 2.15, w: 8.6, h: 2.9, rectRadius: 0.12, fill: { color: CARD }, line: { type: "none" } });
  s.addText(body, { x: 1.05, y: 2.45, w: 7.9, h: 2.4, fontFace: "Calibri", fontSize: 16, color: INK, valign: "top", lineSpacingMultiple: 1.3 });
  s.addShape("roundRect", { x: 9.6, y: 2.15, w: 3.0, h: 2.9, rectRadius: 0.12, fill: { color: NAVY }, line: { type: "none" } });
  s.addText("REWARD", { x: 9.6, y: 2.5, w: 3.0, h: 0.4, align: "center", fontFace: "Calibri", fontSize: 11, color: "CADCFC", bold: true, charSpacing: 2 });
  s.addText(reward, { x: 9.6, y: 2.9, w: 3.0, h: 1.0, align: "center", fontFace: "Cambria", fontSize: 26, color: GOLD, bold: true });
  return s;
}

function twoColSlide(pres, title, leftHeader, leftBody, rightHeader, rightBody) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText(title, { x: 0.7, y: 0.55, w: 11.9, h: 0.8, fontFace: "Cambria", fontSize: 30, color: NAVY, bold: true });
  s.addShape("roundRect", { x: 0.7, y: 1.7, w: 5.7, h: 4.3, rectRadius: 0.12, fill: { color: CARD }, line: { type: "none" } });
  s.addText(leftHeader, { x: 1.0, y: 1.95, w: 5.1, h: 0.5, fontFace: "Calibri", fontSize: 18, bold: true, color: CORAL });
  s.addText(leftBody, { x: 1.0, y: 2.5, w: 5.1, h: 3.3, fontFace: "Calibri", fontSize: 15, color: INK, valign: "top", lineSpacingMultiple: 1.3 });
  s.addShape("roundRect", { x: 6.7, y: 1.7, w: 5.7, h: 4.3, rectRadius: 0.12, fill: { color: CARD }, line: { type: "none" } });
  s.addText(rightHeader, { x: 7.0, y: 1.95, w: 5.1, h: 0.5, fontFace: "Calibri", fontSize: 18, bold: true, color: GOLD });
  s.addText(rightBody, { x: 7.0, y: 2.5, w: 5.1, h: 3.3, fontFace: "Calibri", fontSize: 15, color: INK, valign: "top", lineSpacingMultiple: 1.3 });
  return s;
}

function bodySlide(pres, title, bodyLines, opts = {}) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText(title, { x: 0.7, y: 0.6, w: 11.9, h: 0.9, fontFace: "Cambria", fontSize: 32, color: NAVY, bold: true });
  const paras = bodyLines.map((t, i) => ({ text: t, options: { bullet: opts.bullets ? { code: "25CF" } : false, breakLine: true, paraSpaceAfter: 14, color: INK, fontSize: 18, fontFace: "Calibri" } }));
  s.addText(paras, { x: 0.9, y: 2.0, w: 10.8, h: 4.5, valign: "top", lineSpacingMultiple: 1.3 });
  return s;
}

// ---- a simple demand/supply diagram built from shapes (for Lesson 2) ----
function dsDiagram(s, x, y, w, h, mode) {
  const axisColor = INK;
  // axes
  s.addShape("line", { x, y, w: 0, h, line: { color: axisColor, width: 1.5 } });
  s.addShape("line", { x, y: y + h, w, h: 0, line: { color: axisColor, width: 1.5 } });
  // D line (downward, blue-ish navy) and S line (upward, coral) drawn corner-to-corner of their bounding box
  s.addShape("line", { x: x + 0.3, y: y + 0.1, w: w - 0.6, h: h - 0.3, line: { color: NAVY, width: 3 }, flipV: true });
  s.addShape("line", { x: x + 0.3, y: y + 0.1, w: w - 0.6, h: h - 0.3, line: { color: CORAL, width: 3 } });
  s.addText("D", { x: x + w - 0.55, y: y + h - 0.55, w: 0.5, h: 0.35, color: NAVY, bold: true, fontSize: 14 });
  s.addText("S", { x: x + w - 0.55, y: y + 0.05, w: 0.5, h: 0.35, color: CORAL, bold: true, fontSize: 14 });

  const midY = mode === "shortage" ? y + h * 0.62 : y + h * 0.3;
  s.addShape("line", { x, y: midY, w, h: 0, line: { color: MUTED, width: 1.25, dashType: "dash" } });
  s.addText("P1", { x: x - 0.55, y: midY - 0.18, w: 0.5, h: 0.35, fontSize: 12, color: INK, align: "right" });

  // intersection points: compute x along D and S lines at midY (approx, for visual only)
  const frac = mode === "shortage" ? [0.22, 0.72] : [0.72, 0.22]; // [Sfrac, Dfrac] positions along width
  const leftX = x + 0.3 + (w - 0.6) * frac[0];
  const rightX = x + 0.3 + (w - 0.6) * frac[1];
  s.addShape("ellipse", { x: leftX - 0.06, y: midY - 0.06, w: 0.12, h: 0.12, fill: { color: mode === "shortage" ? CORAL : NAVY }, line: { type: "none" } });
  s.addShape("ellipse", { x: rightX - 0.06, y: midY - 0.06, w: 0.12, h: 0.12, fill: { color: mode === "shortage" ? NAVY : CORAL }, line: { type: "none" } });
  s.addShape("line", { x: Math.min(leftX, rightX), y: y + h + 0.15, w: Math.abs(rightX - leftX), h: 0, line: { color: INK, width: 1.5 } });
  s.addText(mode === "shortage" ? "shortage" : "surplus", { x: Math.min(leftX, rightX), y: y + h + 0.22, w: Math.abs(rightX - leftX), h: 0.3, align: "center", fontSize: 12, color: INK });
  const leftLabel = mode === "shortage" ? "Qs" : "Qd";
  const rightLabel = mode === "shortage" ? "Qd" : "Qs";
  s.addText(leftLabel, { x: leftX - 0.3, y: y + h + 0.15, w: 0.5, h: 0.3, fontSize: 11, color: MUTED, align: mode === "shortage" ? "right" : "left" });
  s.addText(rightLabel, { x: rightX - 0.2, y: y + h + 0.15, w: 0.5, h: 0.3, fontSize: 11, color: MUTED });
}

function diagramSlide(pres, title, mode, caption) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText(title, { x: 0.7, y: 0.55, w: 11.9, h: 0.8, fontFace: "Cambria", fontSize: 28, color: NAVY, bold: true });
  dsDiagram(s, 1.2, 1.7, 4.6, 3.6, mode);
  s.addShape("roundRect", { x: 6.6, y: 2.1, w: 5.9, h: 2.9, rectRadius: 0.12, fill: { color: CARD }, line: { type: "none" } });
  s.addText(caption, { x: 6.95, y: 2.4, w: 5.2, h: 2.4, fontFace: "Calibri", fontSize: 16, color: INK, valign: "top", lineSpacingMultiple: 1.3 });
  return s;
}

function buildLesson1() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  titleSlide(pres, "Lesson 1 · Syllabus 1.2", "Factors of Production", "How does Tom's Print Co. actually make stuff?");

  bodySlide(pres, "Every business needs 4 ingredients", [
    "🖨️  🧩  🎾",
    "Tom's Print Co. 3D-prints custom padel grips and Lego-compatible parts.",
    "Before we name the ingredients — what do you think you'd actually need to start it this weekend?",
  ]);

  factorSlide(pres, {
    eyebrow: "FACTOR 1",
    emoji: "🌍",
    title: "Land",
    body: "Natural resources used in production — not just soil.\n\nThe filament spool is made from petroleum. That's land.",
    reward: "Rent",
  });
  factorSlide(pres, {
    eyebrow: "FACTOR 2",
    emoji: "🧑‍🔧",
    title: "Labour",
    body: "Human physical or mental effort.\n\nTom designing CAD files, running the printer, packaging orders.",
    reward: "Wages",
  });
  factorSlide(pres, {
    eyebrow: "FACTOR 3",
    emoji: "🖨️",
    title: "Capital",
    body: "Manufactured goods used to produce other goods. The printer, the laptop, the tools.\n\nNot money! Cash isn't capital in economics.",
    reward: "Interest",
  });
  factorSlide(pres, {
    eyebrow: "FACTOR 4",
    emoji: "🎲",
    title: "Enterprise",
    body: "Risk-taking + organising the other three factors.\n\nTom's decision to start the business at all. Unlike the other three rewards, profit isn't guaranteed.",
    reward: "Profit",
  });

  twoColSlide(pres, "Quantity vs Quality",
    "📈  Quantity", "Tom buys a second 3D printer.\n\nMore units of the same capital.",
    "⭐  Quality", "Tom upgrades to a faster, more precise printer.\n\nSame number, better ones.");

  bodySlide(pres, "Quick check", [
    "🎾  The padel club buys 3 more courts. Quantity or quality of capital?",
    "🎾  The head coach completes an advanced coaching qualification. Quantity or quality of labour?",
  ]);

  bodySlide(pres, "Recap", [
    "🌍  Land → Rent",
    "🧑‍🔧  Labour → Wages",
    "🖨️  Capital → Interest (not money!)",
    "🎲  Enterprise → Profit (not guaranteed)",
  ], { bullets: false });

  closingSlide(pres, "Next lesson", "Market disequilibrium & price changes",
    "What happens when everyone wants a padel court at 6pm and there aren't enough? Shortages, surpluses, and how prices react.");

  return pres;
}

function buildLesson2() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  titleSlide(pres, "Lesson 2 · Syllabus 2.4.3 / 2.5", "Disequilibrium & Price Changes", "Why can't you ever get a padel court at 6pm?");

  bodySlide(pres, "Quick recap", [
    "At equilibrium, quantity demanded = quantity supplied.",
    "You already know this cold — today's about what happens when it isn't equal.",
  ]);

  diagramSlide(pres, "🎾 6–8pm slot: SHORTAGE", "shortage",
    "Club's normal price (P1) is below the true equilibrium for this slot.\n\nQuantity demanded (Qd) > quantity supplied (Qs, fixed courts).\n\nThat gap is the shortage.");

  diagramSlide(pres, "🎾 9am slot: SURPLUS", "surplus",
    "Same price (P1) is now above the true equilibrium for this quiet slot.\n\nQuantity supplied (Qs) > quantity demanded (Qd).\n\nEmpty courts = surplus.");

  bodySlide(pres, "Clearing it: pricing", [
    "Raise the evening price → quantity demanded falls back (movement along D) until it matches supply.",
    "Cut the morning price → quantity demanded rises until it matches supply.",
    "Both cases: price moves toward that slot's own equilibrium.",
  ]);

  bodySlide(pres, "Same idea: retired Lego sets", [
    "🧩 📈",
    "Set is discontinued → supply fixed/shrinking, demand persists → shortage at old retail price.",
    "Resale price rises until quantity demanded falls to match the fixed supply.",
  ]);

  twoColSlide(pres, "Causes vs consequences",
    "Cause", "A price change is caused by a shift in demand and/or supply.",
    "Consequence", "The consequence is that quantity bought & sold changes.\n\nWatch out: a price change alone only moves you along a curve — it doesn't shift it.");

  closingSlide(pres, "Next lesson", "Price elasticity of demand (PED)",
    "Why does a price rise crush sales of some things but barely dent others?");

  return pres;
}

async function run() {
  const BASE = path.join(__dirname, "..", "lessons");
  await buildLesson1().writeFile({ fileName: path.join(BASE, "lesson-01", "slides.pptx") });
  console.log("wrote lesson-01/slides.pptx");
  await buildLesson2().writeFile({ fileName: path.join(BASE, "lesson-02", "slides.pptx") });
  console.log("wrote lesson-02/slides.pptx");
}
run();
