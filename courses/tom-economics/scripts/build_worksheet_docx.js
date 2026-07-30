// Builds worksheet.docx for a lesson from a structured content object (not from HTML/markdown -
// Word documents don't map cleanly from arbitrary markdown, so worksheets are authored directly
// in this structured shape). Add a new lesson by adding an entry to LESSONS below.
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType,
} = require("docx");
const fs = require("fs");
const path = require("path");

const BLUE = "2A78D6";
const ORANGE = "EB6834";
const AMBER_LIGHT = "FDF1DE";
const GRID = "E1E0D9";

// ---- tiny inline **bold** parser ----
function runs(text, opts = {}) {
  const parts = text.split(/(\*\*[^*]+\*\*)/g).filter(Boolean);
  return parts.map(p => {
    if (p.startsWith("**") && p.endsWith("**")) {
      return new TextRun({ text: p.slice(2, -2), bold: true, ...opts });
    }
    return new TextRun({ text: p, ...opts });
  });
}

function para(text, opts = {}) {
  return new Paragraph({ children: runs(text, opts.runOpts || {}), spacing: { after: 160 }, shading: opts.shading });
}

function heading(text, level = HeadingLevel.HEADING_2, color = BLUE) {
  return new Paragraph({
    heading: level,
    spacing: { before: 280, after: 140 },
    children: [new TextRun({ text, bold: true, color })],
  });
}

function makeTable(headerRow, rows) {
  const colCount = headerRow.length;
  const colWidth = Math.floor(9000 / colCount);
  const widths = new Array(colCount).fill(colWidth);
  const headerCells = headerRow.map(h => new TableCell({
    width: { size: colWidth, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: GRID },
    children: [new Paragraph({ children: [new TextRun({ text: h, bold: true })] })],
  }));
  const bodyRows = rows.map(r => new TableRow({
    children: r.map(c => new TableCell({
      width: { size: colWidth, type: WidthType.DXA },
      children: [new Paragraph({ children: runs(c || "") })],
    })),
  }));
  return new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: widths,
    rows: [new TableRow({ children: headerCells }), ...bodyRows],
  });
}

function buildDoc(data) {
  const children = [];
  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text: data.title, bold: true })],
    spacing: { after: 80 },
  }));
  children.push(para(`**Syllabus ref: ${data.syllabusRef}** · Complete during your 30 min self-directed time.`));

  for (const section of data.sections) {
    children.push(heading(section.heading));
    if (section.intro) children.push(para(section.intro));
    for (const item of section.items) {
      if (item.type === "table") {
        children.push(makeTable(item.headers, item.rows));
        children.push(new Paragraph({ text: "", spacing: { after: 120 } }));
      } else if (item.type === "qa") {
        children.push(para(item.q));
        children.push(para("Answer: " + (item.blank || "")));
      } else if (item.type === "mcq") {
        children.push(para(item.q));
        item.options.forEach(o => children.push(para(o)));
        children.push(para("Answer: ____"));
      } else if (item.type === "text") {
        children.push(para(item.text));
      }
    }
  }

  // Answer key — visually boxed with shading + orange heading
  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 360, after: 140 },
    children: [new TextRun({ text: "Answer key (tutor copy — not for Tom)", bold: true, color: ORANGE })],
    shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT },
  }));
  for (const item of data.answerKey) {
    if (item.type === "table") {
      children.push(makeTable(item.headers, item.rows));
      children.push(new Paragraph({ text: "", spacing: { after: 120 }, shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }));
    } else {
      children.push(para(item.text, { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }));
    }
  }

  return new Document({
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 } } },
      children,
    }],
  });
}

const LESSON1 = {
  title: "Lesson 1 Worksheet — Factors of Production",
  syllabusRef: "1.2",
  sections: [
    {
      heading: "Section A — Definitions",
      items: [
        { type: "text", text: "**1a.** Fill in the table: definition of each factor of production, and the reward it earns." },
        { type: "table", headers: ["Factor", "Definition", "Reward"], rows: [["Land", "", ""], ["Labour", "", ""], ["Capital", "", ""], ["Enterprise", "", ""]] },
      ],
    },
    {
      heading: "Section B — Apply it: Tom's Print Co.",
      intro: "Tom runs a small 3D-printing business making custom padel paddle overgrips and personalised Lego-compatible minifigure accessories. For each item, state **which factor of production it is** and **explain why**.",
      items: [
        { type: "qa", q: "**2a.** The 3D printer itself." },
        { type: "qa", q: "**2b.** The spool of PLA plastic filament (made from petroleum, a natural resource)." },
        { type: "qa", q: "**2c.** Tom's own time spent designing the CAD files and running the printer." },
        { type: "qa", q: "**2d.** Tom's decision to start the business and risk his savings on the first printer." },
      ],
    },
    {
      heading: "Section C — Quantity vs quality",
      items: [
        { type: "qa", q: "**3a.** Tom buys a second 3D printer to keep up with demand. Is this a change in the **quantity** or **quality** of capital? Explain." },
        { type: "qa", q: "**3b.** Tom completes an online CAD design course and becomes much faster at designing new products. Is this a change in the **quantity** or **quality** of labour? Explain." },
        { type: "qa", q: "**3c.** Give one example (any business — a padel club, a Lego resale shop, anything) of a change in the **quantity** of a factor of production, and one example of a change in **quality**.\nQuantity example:\nQuality example:" },
      ],
    },
    {
      heading: "Section D — Exam-style practice",
      items: [
        { type: "mcq", q: "**4.** *(Paper 1 style MCQ)* Which of the following is an example of capital in economics?", options: ["A. The money Tom uses to buy filament", "B. Tom's 3D printer", "C. Tom's decision to start the business", "D. The rent Tom pays for his workshop"] },
        { type: "mcq", q: "**5.** *(Paper 1 style MCQ)* A worker's reward for supplying labour is called:", options: ["A. Rent", "B. Profit", "C. Interest", "D. Wages"] },
        { type: "qa", q: "**6.** *(Paper 2 style, 4 marks)* Explain **two** reasons why the quality of capital available to a firm like Tom's Print Co. might improve over time." },
      ],
    },
  ],
  answerKey: [
    { type: "table", headers: ["Factor", "Definition", "Reward"], rows: [
      ["Land", "Natural resources used in production", "Rent"],
      ["Labour", "Human physical or mental effort used in production", "Wages"],
      ["Capital", "Manufactured goods used to produce other goods/services", "Interest"],
      ["Enterprise", "Willingness/ability to take risks and organise the other factors, in pursuit of profit", "Profit"],
    ]},
    { type: "text", text: "**2a.** Capital — a manufactured good used to produce other goods (padel grips, Lego accessories)." },
    { type: "text", text: "**2b.** Land — a raw natural resource (petroleum-derived), not manufactured." },
    { type: "text", text: "**2c.** Labour — human physical/mental effort." },
    { type: "text", text: "**2d.** Enterprise — risk-taking and organising the other three factors; rewarded by (uncertain) profit, not a guaranteed wage." },
    { type: "text", text: "**3a.** Quantity — it's more units of the same capital good, not a better one." },
    { type: "text", text: "**3b.** Quality — the same amount of labour is now more productive/skilled, not more workers." },
    { type: "text", text: "**3c.** Open-ended; accept any valid business example, e.g. padel club buying 3 more courts (quantity) vs resurfacing existing courts with better material (quality)." },
    { type: "text", text: "**4.** B — the printer is a manufactured good used to produce other goods; A is finance, not capital; C is enterprise; D is a cost linked to land." },
    { type: "text", text: "**5.** D — wages are the reward to labour." },
    { type: "text", text: "**6.** Look for **two developed points** (AO2 chains), e.g.: (i) technological improvement — a newer printer model with finer print resolution raises the quality/precision of output; (ii) investment in R&D or firmware upgrades improves reliability and reduces waste; (iii) training in maintenance improves how effectively existing capital is used. Full marks need identification + explanation of the effect on output/quality, not just a list." },
  ],
};

const LESSON2 = {
  title: "Lesson 2 Worksheet — Market Disequilibrium & Price Changes",
  syllabusRef: "2.4.3 / 2.5",
  sections: [
    {
      heading: "Section A — Definitions",
      items: [
        { type: "qa", q: "**1a.** Define market disequilibrium." },
        { type: "qa", q: "**1b.** Define a shortage and a surplus, in terms of quantity demanded and quantity supplied.\nShortage:\nSurplus:" },
      ],
    },
    {
      heading: "Section B — Apply it: the padel club",
      intro: "The padel club has a fixed number of courts. At its normal flat price: the 6–8pm weekday slot is fully booked out weeks in advance, and people are still asking for more slots; the 9am weekday slot regularly has empty courts.",
      items: [
        { type: "qa", q: "**2a.** Is the 6–8pm slot a shortage or a surplus at the current price? Explain using quantity demanded and quantity supplied." },
        { type: "qa", q: "**2b.** Is the 9am slot a shortage or a surplus at the current price? Explain." },
        { type: "qa", q: "**2c.** The club decides to charge more for the 6–8pm slot and less for the 9am slot. Explain how each price change helps move that time slot's market toward its own equilibrium." },
      ],
    },
    {
      heading: "Section C — A second example: Lego resale",
      intro: "A Lego set is discontinued (retired) by the manufacturer. New supply stops, but plenty of collectors still want it.",
      items: [
        { type: "qa", q: "**3a.** What happens to the balance between quantity demanded and quantity supplied at the old retail price once the set is retired?" },
        { type: "qa", q: "**3b.** What do you predict happens to the resale price of the set over time, and why?" },
      ],
    },
    {
      heading: "Section D — Causes vs consequences",
      items: [
        { type: "qa", q: "**4a.** State what *causes* a price change in a market (in general terms, not just from the examples above)." },
        { type: "qa", q: "**4b.** State the *consequence* of a price change for the quantity bought and sold." },
        { type: "qa", q: "**4c.** True or false, with a reason: \"A price change on its own causes the whole demand curve to shift.\"" },
      ],
    },
    {
      heading: "Section E — Exam-style practice",
      items: [
        { type: "mcq", q: "**5.** *(Paper 1 style MCQ)* At the padel club's 9am weekday price, quantity supplied exceeds quantity demanded. This is an example of:", options: ["A. Market equilibrium", "B. A shortage", "C. A surplus", "D. Price elasticity"] },
        { type: "mcq", q: "**6.** *(Paper 1 style MCQ)* Which of these would most directly explain why the price of a retired Lego set rises on the resale market?", options: ["A. Quantity supplied is fixed while demand persists, causing a shortage at the old price", "B. The government has introduced a new tax on Lego", "C. The manufacturer has increased production", "D. Quantity demanded has fallen sharply"] },
        { type: "qa", q: "**7.** *(Paper 2 style, 6 marks)* Using a demand and supply diagram, explain how a padel club could use pricing to deal with a shortage at its 6–8pm slot. Your answer should refer to quantity demanded, quantity supplied, and the new equilibrium." },
      ],
    },
  ],
  answerKey: [
    { type: "text", text: "**1a.** A situation where quantity demanded and quantity supplied are not equal at the current price (the market is not at equilibrium)." },
    { type: "text", text: "**1b.** Shortage: quantity demanded exceeds quantity supplied at the current price. Surplus: quantity supplied exceeds quantity demanded at the current price." },
    { type: "text", text: "**2a.** Shortage — quantity demanded (people wanting the slot) exceeds quantity supplied (fixed number of courts) at the current price." },
    { type: "text", text: "**2b.** Surplus — quantity supplied (courts sitting empty) exceeds quantity demanded at the current price." },
    { type: "text", text: "**2c.** Raising the evening price reduces quantity demanded (movement along the demand curve) until it matches the fixed quantity supplied — clearing the shortage at a new, higher equilibrium for that slot. Lowering the morning price increases quantity demanded until it matches quantity supplied — clearing the surplus at a new, lower equilibrium for that slot." },
    { type: "text", text: "**3a.** A shortage emerges — quantity demanded now exceeds the fixed/shrinking quantity supplied at the old retail price." },
    { type: "text", text: "**3b.** The resale price rises, because as price rises, quantity demanded falls back until it matches the fixed quantity available — the market moves to a new, higher equilibrium." },
    { type: "text", text: "**4a.** A change in demand and/or a change in supply (a shift of one or both curves), driven by changes in their non-price determinants." },
    { type: "text", text: "**4b.** The quantity bought and sold changes (the exact size of the change depends on elasticity — covered next lesson)." },
    { type: "text", text: "**4c.** False — a price change alone causes a *movement along* an existing curve, not a shift of the curve itself. A shift is caused by a change in a non-price determinant (income, tastes, cost of production, etc.)." },
    { type: "text", text: "**5.** C" },
    { type: "text", text: "**6.** A" },
    { type: "text", text: "**7.** Look for: diagram showing the original price below equilibrium with quantity demanded > quantity supplied (shortage) at that price; explanation that raising price causes quantity demanded to contract (movement along the demand curve) while quantity supplied stays fixed (courts can't easily be added); new equilibrium reached where quantity demanded = quantity supplied at the higher price. Full marks need the diagram AND the explanation linked together, not just one or the other." },
  ],
};

async function run() {
  const BASE = path.join(__dirname, "..", "lessons");
  const jobs = [
    { lesson: "lesson-01", data: LESSON1 },
    { lesson: "lesson-02", data: LESSON2 },
  ];
  for (const job of jobs) {
    const doc = buildDoc(job.data);
    const buf = await Packer.toBuffer(doc);
    const out = path.join(BASE, job.lesson, "worksheet.docx");
    fs.writeFileSync(out, buf);
    console.log("wrote", out);
  }
}
run();
