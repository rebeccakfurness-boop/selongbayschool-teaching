// Builds slides.pptx for a lesson, reproducing the content of that lesson's slides.html
// as a real presentable PowerPoint deck. Palette: dark navy anchor + coral accent + gold
// "reward" tag color — kept consistent across lessons as the course's visual identity.
const pptxgen = require("pptxgenjs");
const path = require("path");

const NAVY = "1E2761";
const CORAL = "F96167";
const GOLD = "D98C00";
const AQUA = "1BAF7A";
const WHITE = "FFFFFF";
const CARD = "F4F5F7";
const INK = "1A1A1A";
const MUTED = "5B5B5B";

// NB: pptxgenjs fill colors must be plain 6-digit hex, never 8-digit/alpha-baked (corrupts the
// file) — use `transparency` for a light tint instead of concatenating an alpha suffix.
function tagPill(s, label, color) {
  const w = 0.35 + label.length * 0.11;
  s.addShape("roundRect", { x: 0.7, y: 0.55, w, h: 0.4, rectRadius: 0.2, fill: { color, transparency: 82 }, line: { type: "none" } });
  s.addText(label, { x: 0.7, y: 0.55, w, h: 0.4, align: "center", valign: "middle", fontFace: "Calibri", fontSize: 12, bold: true, color });
}

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

function calcSlide(pres, eyebrow, title, lines, result, opts = {}) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  if (opts.tag) tagPill(s, opts.tag, opts.tagColor || CORAL);
  s.addText(eyebrow, { x: 0.7, y: opts.tag ? 1.15 : 0.7, w: 11.9, h: 0.6, fontFace: "Calibri", fontSize: 22, color: NAVY, bold: true });
  s.addShape("roundRect", { x: 3.9, y: 2.1, w: 5.5, h: 3.4, rectRadius: 0.14, fill: { color: CARD }, line: { type: "none" } });
  s.addText(title, { x: 4.2, y: 2.4, w: 4.9, h: 0.4, fontFace: "Calibri", fontSize: 14, color: MUTED, bold: true });
  const paras = lines.map(t => ({ text: t, options: { breakLine: true, paraSpaceAfter: 6, color: INK, fontSize: 15, fontFace: "Calibri" } }));
  s.addText(paras, { x: 4.2, y: 2.9, w: 4.9, h: 1.6, valign: "top" });
  s.addText(result, { x: 4.2, y: 4.6, w: 4.9, h: 0.8, fontFace: "Cambria", fontSize: 18, bold: true, color: opts.resultColor || NAVY });
  return s;
}

function ruleTableSlide(pres, title, rows) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText(title, { x: 0.7, y: 0.6, w: 11.9, h: 0.8, fontFace: "Cambria", fontSize: 32, color: NAVY, bold: true });
  const tableRows = rows.map((r, i) => r.map((cell, ci) => {
    const isHeader = i === 0;
    let color = INK;
    if (!isHeader && ci > 0) color = cell.toLowerCase().includes("rise") ? AQUA : cell.toLowerCase().includes("fall") ? CORAL : INK;
    return {
      text: cell,
      options: {
        bold: isHeader || ci === 0,
        color: isHeader ? MUTED : color,
        fill: { color: isHeader ? CARD : WHITE },
        fontFace: "Calibri",
        fontSize: 15,
        align: ci === 0 ? "left" : "center",
      },
    };
  }));
  s.addTable(tableRows, { x: 2.9, y: 2.0, w: 7.5, colW: [2.5, 2.5, 2.5], border: { type: "solid", color: "D8D8D8", pt: 1 }, autoPage: false });
  return s;
}

function blankCalcSlide(pres, eyebrow, title, lines, opts = {}) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  if (opts.tag) tagPill(s, opts.tag, opts.tagColor || CORAL);
  s.addText(eyebrow, { x: 0.7, y: opts.tag ? 1.15 : 0.7, w: 11.9, h: 0.6, fontFace: "Calibri", fontSize: 22, color: NAVY, bold: true });
  s.addShape("roundRect", { x: 3.9, y: 2.1, w: 5.5, h: 3.4, rectRadius: 0.14, fill: { color: CARD }, line: { type: "none" } });
  s.addText(title, { x: 4.2, y: 2.4, w: 4.9, h: 0.4, fontFace: "Calibri", fontSize: 14, color: MUTED, bold: true });
  const paras = lines.map(t => ({ text: t, options: { breakLine: true, paraSpaceAfter: 8, color: INK, fontSize: 15, fontFace: "Calibri" } }));
  s.addText(paras, { x: 4.2, y: 2.9, w: 4.9, h: 1.4, valign: "top" });
  s.addShape("line", { x: 4.2, y: 4.75, w: 2.0, h: 0, line: { color: MUTED, width: 1.5 } });
  s.addText("TR before = $____", { x: 4.2, y: 4.35, w: 4.6, h: 0.35, fontFace: "Calibri", fontSize: 14, color: INK });
  s.addShape("line", { x: 4.2, y: 5.15, w: 2.0, h: 0, line: { color: MUTED, width: 1.5 } });
  s.addText("TR after = $____", { x: 4.2, y: 4.75, w: 4.6, h: 0.35, fontFace: "Calibri", fontSize: 14, color: INK });
  return s;
}

// ---- PED spectrum: 5 small axis+line diagrams in a row ----
function pedMiniDiagram(s, x, y, size, slopeType, color) {
  s.addShape("line", { x, y, w: 0, h: size, line: { color: INK, width: 1.25 } });
  s.addShape("line", { x, y: y + size, w: size, h: 0, line: { color: INK, width: 1.25 } });
  const pad = size * 0.15;
  if (slopeType === "vertical") {
    s.addShape("line", { x: x + size * 0.5, y: y + pad, w: 0, h: size - pad * 2, line: { color, width: 2.5 } });
  } else if (slopeType === "horizontal") {
    s.addShape("line", { x: x + pad, y: y + size * 0.5, w: size - pad * 2, h: 0, line: { color, width: 2.5 } });
  } else {
    // diagonal, steepness varies: "steep" | "mid" | "shallow"
    const insetMap = { steep: 0.32, mid: 0.15, shallow: 0.32 };
    const inset = insetMap[slopeType] || 0.15;
    if (slopeType === "shallow") {
      s.addShape("line", { x: x + pad, y: y + size * 0.35, w: size - pad * 2, h: size * 0.3, line: { color, width: 2.5 } });
    } else {
      s.addShape("line", { x: x + size * inset, y: y + pad, w: size - size * inset * 2, h: size - pad * 2, line: { color, width: 2.5 }, flipV: true });
    }
  }
}

function pedSpectrumSlide(pres) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText("The PED spectrum", { x: 0.7, y: 0.6, w: 11.9, h: 0.8, fontFace: "Cambria", fontSize: 32, color: NAVY, bold: true });
  const items = [
    { slope: "vertical", label: "Perfectly inelastic", sub: "PED = 0", color: NAVY },
    { slope: "steep", label: "Inelastic", sub: "0 < PED < 1", color: NAVY },
    { slope: "mid", label: "Unitary", sub: "PED = 1", color: CORAL },
    { slope: "shallow", label: "Elastic", sub: "PED > 1", color: "1BAF7A" },
    { slope: "horizontal", label: "Perfectly elastic", sub: "PED = ∞", color: "1BAF7A" },
  ];
  const dSize = 1.7, gap = 0.5;
  const totalW = items.length * dSize + (items.length - 1) * gap;
  let x = (13.33 - totalW) / 2;
  const y = 2.2;
  for (const it of items) {
    pedMiniDiagram(s, x, y, dSize, it.slope, it.color);
    s.addText(it.label, { x: x - 0.3, y: y + dSize + 0.15, w: dSize + 0.6, h: 0.35, align: "center", fontFace: "Calibri", fontSize: 12, bold: true, color: NAVY });
    s.addText(it.sub, { x: x - 0.3, y: y + dSize + 0.5, w: dSize + 0.6, h: 0.3, align: "center", fontFace: "Calibri", fontSize: 11, color: MUTED });
    x += dSize + gap;
  }
  return s;
}

function cardRowSlide(pres, title, cards) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText(title, { x: 0.7, y: 0.6, w: 11.9, h: 0.8, fontFace: "Cambria", fontSize: 30, color: NAVY, bold: true });
  const n = cards.length;
  const cardW = 3.6, gap = 0.4;
  const totalW = n * cardW + (n - 1) * gap;
  let x = (13.33 - totalW) / 2;
  const y = 2.0;
  cards.forEach(c => {
    s.addShape("roundRect", { x, y, w: cardW, h: 3.4, rectRadius: 0.12, fill: { color: CARD }, line: { type: "none" } });
    s.addText(c.name, { x: x + 0.3, y: y + 0.3, w: cardW - 0.6, h: 0.6, fontFace: "Calibri", fontSize: 16, bold: true, color: NAVY });
    s.addText(c.desc, { x: x + 0.3, y: y + 0.95, w: cardW - 0.6, h: 2.2, fontFace: "Calibri", fontSize: 12.5, color: INK, valign: "top", lineSpacingMultiple: 1.25 });
    x += cardW + gap;
  });
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

function buildLesson3() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  titleSlide(pres, "Lesson 3 · Syllabus 2.6", "Price Elasticity of Demand", "If Tom raises his prices, does he lose a few customers — or a lot?");

  bodySlide(pres, "Same price rise, different reactions", [
    "Tom puts up the price of custom padel grips by 10%.",
    "Would he lose a few customers, or a lot? What would that actually depend on?",
  ]);

  const s3 = pres.addSlide();
  s3.background = { color: WHITE };
  s3.addText("Definition & formula", { x: 0.7, y: 0.7, w: 11.9, h: 0.8, fontFace: "Cambria", fontSize: 30, color: NAVY, bold: true });
  s3.addText("PED measures how responsive quantity demanded is to a change in price.", { x: 0.9, y: 1.8, w: 10.5, h: 0.7, fontFace: "Calibri", fontSize: 17, color: INK });
  s3.addShape("roundRect", { x: 2.9, y: 2.7, w: 7.5, h: 1.3, rectRadius: 0.14, fill: { color: NAVY }, line: { type: "none" } });
  s3.addText("PED = %ΔQd ÷ %ΔP", { x: 2.9, y: 2.7, w: 7.5, h: 1.3, align: "center", valign: "middle", fontFace: "Cambria", fontSize: 30, bold: true, color: "CADCFC" });
  s3.addText("PED is technically always negative (law of demand) — economists usually just discuss its size, ignoring the minus sign.", { x: 0.9, y: 4.3, w: 10.5, h: 0.9, fontFace: "Calibri", fontSize: 15, color: MUTED, italic: true });

  calcSlide(pres, "🎾 Worked example: elastic", "Padel club casual court booking",
    ["Price: $20 → $22  (+10%)", "Quantity: 100 → 80 bookings/week  (−20%)"],
    "PED = −20 ÷ 10 = −2 → Elastic");

  calcSlide(pres, "🧩 Worked example: inelastic", "Retired Lego set, resale market",
    ["Price: $50 → $55  (+10%)", "Quantity: 40 → 38 sets/month  (−5%)"],
    "PED = −5 ÷ 10 = −0.5 → Inelastic");

  pedSpectrumSlide(pres);

  cardRowSlide(pres, "Determinants of PED (1/2)", [
    { name: "Substitutes", desc: "More/closer substitutes = more elastic. A generic Lego set (many alternatives) vs. a specific retired set (no substitute)." },
    { name: "Necessity vs luxury", desc: "Necessities are more inelastic. Court resurfacing (has to happen) vs. novelty keychains (impulse buy)." },
    { name: "Share of income", desc: "Bigger share of income = more elastic. A cheap grip (barely noticed) vs. an annual club membership (very noticeable)." },
  ]);

  cardRowSlide(pres, "Determinants of PED (2/2)", [
    { name: "Habit-forming", desc: "Addictive goods are more inelastic (classic example: cigarettes)." },
    { name: "Time period", desc: "More elastic in the long run — customers keep buying out of habit short-term, then switch supplier over time." },
    { name: "Width of definition", desc: "\"3D-printed products\" broadly (hard to avoid) vs. \"Tom's specific grip design\" narrowly (easy to switch away from)." },
  ]);

  bodySlide(pres, "Quick check", [
    "A 10% price rise causes quantity demanded to fall by 25%. Elastic or inelastic?",
    "Which makes demand more elastic: many substitutes, or being a necessity?",
  ]);

  bodySlide(pres, "Recap", [
    "PED = %ΔQd ÷ %ΔP",
    "0 → perfectly inelastic  ·  0–1 → inelastic  ·  1 → unitary  ·  >1 → elastic  ·  ∞ → perfectly elastic",
    "Determinants: substitutes, necessity, income share, habit, time, width of definition.",
  ]);

  closingSlide(pres, "Next lesson", "PED: revenue & significance",
    "What PED means for how much money Tom actually makes when he changes his prices.");

  return pres;
}

function buildLesson4() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  titleSlide(pres, "Lesson 4 · Syllabus 2.6", "PED: Revenue & Significance", "If Tom raises his prices, does he end up richer — or poorer?");

  const fix1 = pres.addSlide();
  fix1.background = { color: WHITE };
  tagPill(fix1, "🔧 Quick fix", NAVY);
  fix1.addText("Land vs. free goods", { x: 0.7, y: 1.4, w: 11.9, h: 0.9, fontFace: "Cambria", fontSize: 30, color: NAVY, bold: true });
  fix1.addText("Your notes had land down as a free good. What's the actual difference between land as a factor of production, and a free good?", { x: 0.9, y: 2.6, w: 10.8, h: 1.5, fontFace: "Calibri", fontSize: 18, color: INK });

  const fix2 = pres.addSlide();
  fix2.background = { color: WHITE };
  tagPill(fix2, "🔧 Quick fix", NAVY);
  fix2.addText("Causes vs. consequences", { x: 0.7, y: 1.4, w: 11.9, h: 0.9, fontFace: "Cambria", fontSize: 30, color: NAVY, bold: true });
  fix2.addText("Last lesson, a shortage/surplus got named as the cause of a price change. What actually causes a price to change in a market?", { x: 0.9, y: 2.6, w: 10.8, h: 1.5, fontFace: "Calibri", fontSize: 18, color: INK });

  const ref = pres.addSlide();
  ref.background = { color: WHITE };
  ref.addText("Quick PED refresher", { x: 0.7, y: 0.7, w: 11.9, h: 0.8, fontFace: "Cambria", fontSize: 30, color: NAVY, bold: true });
  ref.addShape("roundRect", { x: 2.9, y: 1.8, w: 7.5, h: 1.3, rectRadius: 0.14, fill: { color: NAVY }, line: { type: "none" } });
  ref.addText("PED = %ΔQd ÷ %ΔP", { x: 2.9, y: 1.8, w: 7.5, h: 1.3, align: "center", valign: "middle", fontFace: "Cambria", fontSize: 30, bold: true, color: "CADCFC" });
  ref.addText("Size < 1 = inelastic. Size > 1 = elastic. Size = 1 = unitary. Let's get a few reps in before we look at revenue.", { x: 0.9, y: 3.4, w: 10.5, h: 1.0, fontFace: "Calibri", fontSize: 16, color: INK });

  calcSlide(pres, "🎾 Racket restring service", "Practice: calculate the PED",
    ["Price: $15 → $18  (+20%)", "Quantity: 50 → 45 per week  (−10%)"],
    "PED = −10 ÷ 20 = −0.5 → Inelastic", { tag: "✍️ Practice", tagColor: CORAL });

  calcSlide(pres, "🏓 Limited-edition padel paddle", "Practice: calculate the PED",
    ["Price: $150 → $165  (+10%)", "Quantity: 20 → 18 sold  (−10%)"],
    "PED = −10 ÷ 10 = −1 → Unitary", { tag: "✍️ Practice", tagColor: CORAL });

  calcSlide(pres, "🧱 Generic Lego minifigure", "Practice: calculate the PED",
    ["Price: $5 → $6  (+20%)", "Quantity: 100 → 70 sold  (−30%)"],
    "PED = −30 ÷ 20 = −1.5 → Elastic", { tag: "✍️ Practice", tagColor: CORAL });

  bodySlide(pres, "Today's real question", [
    "Tom puts up the price of his padel grips.",
    "More money per grip sold... but does he end up with more total money?",
    "Guess — and tell me why.",
  ]);

  const f = pres.addSlide();
  f.background = { color: WHITE };
  f.addText("Total revenue", { x: 0.7, y: 0.7, w: 11.9, h: 0.8, fontFace: "Cambria", fontSize: 30, color: NAVY, bold: true });
  f.addShape("roundRect", { x: 2.9, y: 1.8, w: 7.5, h: 1.3, rectRadius: 0.14, fill: { color: NAVY }, line: { type: "none" } });
  f.addText("TR = P × Q", { x: 2.9, y: 1.8, w: 7.5, h: 1.3, align: "center", valign: "middle", fontFace: "Cambria", fontSize: 30, bold: true, color: "CADCFC" });
  f.addText("A price change moves P. But it also moves Q, because of PED. Those two moves can fight each other — that's the whole lesson.", { x: 0.9, y: 3.4, w: 10.5, h: 1.0, fontFace: "Calibri", fontSize: 16, color: INK });

  calcSlide(pres, "🎾 Padel bookings (elastic)", "Guess before you check",
    ["Price: $20 → $22  (+10%)", "Quantity: 100 → 80 bookings/week  (−20%)", "PED = −2 (elastic)"],
    "Revenue: up, down, or the same?", { tag: "🤔 Guess first", tagColor: CORAL });

  calcSlide(pres, "🎾 Padel bookings — answer", "The reveal",
    ["TR before = $20 × 100 = $2000", "TR after = $22 × 80 = $1760"],
    "Revenue FELL, despite the price rise", { tag: "✅ Revealed", tagColor: AQUA, resultColor: AQUA });

  calcSlide(pres, "🧩 Retired Lego set (inelastic)", "Guess before you check",
    ["Price: $50 → $55  (+10%)", "Quantity: 40 → 38 sets/month  (−5%)", "PED = −0.5 (inelastic)"],
    "Revenue: up, down, or the same?", { tag: "🤔 Guess first", tagColor: CORAL });

  calcSlide(pres, "🧩 Retired Lego set — answer", "The reveal",
    ["TR before = $50 × 40 = $2000", "TR after = $55 × 38 = $2090"],
    "Revenue ROSE this time", { tag: "✅ Revealed", tagColor: AQUA, resultColor: AQUA });

  blankCalcSlide(pres, "🖨️ Tom's padel grips", "Your turn — work it out",
    ["Price: $8 → $10", "Quantity: 60 → 54", "(PED = −0.4, from last lesson)"],
    { tag: "✍️ Your turn", tagColor: CORAL });

  calcSlide(pres, "🖨️ Tom's grips — answer", "Check your answer",
    ["TR before = $8 × 60 = $480", "TR after = $10 × 54 = $540"],
    "Revenue rose — inelastic, again", { tag: "✅ Revealed", tagColor: AQUA, resultColor: AQUA });

  ruleTableSlide(pres, "The rule", [
    ["PED", "Price rises", "Price falls"],
    ["Inelastic (<1)", "Revenue rises", "Revenue falls"],
    ["Elastic (>1)", "Revenue falls", "Revenue rises"],
    ["Unitary (=1)", "No change", "No change"],
  ]);

  cardRowSlide(pres, "Why it matters", [
    { name: "Firms", desc: "Use PED to set prices for maximum revenue — raise prices on inelastic products, avoid raising (or cut) prices on elastic ones." },
    { name: "Consumers", desc: "Inelastic necessities hurt household budgets most when prices rise — you can't easily cut back." },
    { name: "Government", desc: "Taxes work best on inelastic goods (fuel, cigarettes) — quantity barely falls, so tax revenue stays high and reliable." },
  ]);

  bodySlide(pres, "Quick check", [
    "A firm's product has PED = −3. It's considering a price rise. Good idea or bad idea?",
    "A firm's product has PED = −0.2. Same question.",
  ]);

  bodySlide(pres, "Recap", [
    "TR = P × Q",
    "Inelastic demand: price & revenue move together.",
    "Elastic demand: price & revenue move opposite.",
    "That's why PED matters to firms' pricing, consumers' budgets, and government tax choices.",
  ]);

  closingSlide(pres, "Next lesson", "Price elasticity of supply (PES)",
    "Same structure, new letter — how responsive is quantity supplied to a price change?");

  return pres;
}

async function run() {
  const BASE = path.join(__dirname, "..", "lessons");
  await buildLesson1().writeFile({ fileName: path.join(BASE, "lesson-01", "slides.pptx") });
  console.log("wrote lesson-01/slides.pptx");
  await buildLesson2().writeFile({ fileName: path.join(BASE, "lesson-02", "slides.pptx") });
  console.log("wrote lesson-02/slides.pptx");
  await buildLesson3().writeFile({ fileName: path.join(BASE, "lesson-03", "slides.pptx") });
  console.log("wrote lesson-03/slides.pptx");
  await buildLesson4().writeFile({ fileName: path.join(BASE, "lesson-04", "slides.pptx") });
  console.log("wrote lesson-04/slides.pptx");
}
run();
