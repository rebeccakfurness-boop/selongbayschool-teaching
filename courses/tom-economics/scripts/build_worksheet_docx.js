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

const LESSON3 = {
  title: "Lesson 3 Worksheet — Price Elasticity of Demand (PED)",
  syllabusRef: "2.6",
  sections: [
    {
      heading: "Section A — Definitions",
      items: [
        { type: "qa", q: "**1a.** Define price elasticity of demand (PED)." },
        { type: "qa", q: "**1b.** Write out the formula for PED." },
        { type: "text", text: "**1c.** Fill in the table: match each PED value range to its category." },
        { type: "table", headers: ["PED value", "Category"], rows: [["0", ""], ["Between 0 and 1", ""], ["Exactly 1", ""], ["Greater than 1", ""], ["Infinite", ""]] },
      ],
    },
    {
      heading: "Section B — Worked calculations",
      items: [
        { type: "qa", q: "**2a.** The padel club raises the price of a casual court booking from $20 to $22. Weekly bookings fall from 100 to 80. Calculate PED and state whether demand is elastic or inelastic." },
        { type: "qa", q: "**2b.** A retired Lego set's resale price rises from $50 to $55. Monthly sales fall from 40 to 38. Calculate PED and state whether demand is elastic or inelastic." },
        { type: "qa", q: "**2c.** Tom raises the price of his custom padel grips from $8 to $10. Weekly sales fall from 60 to 54. Calculate PED and state whether demand is elastic or inelastic." },
      ],
    },
    {
      heading: "Section C — Determinants",
      items: [
        { type: "qa", q: "**3a.** Explain why a specific, retired Lego set tends to have inelastic demand, using the idea of substitutes." },
        { type: "qa", q: "**3b.** Explain why a padel club's annual membership fee is likely to have more elastic demand than a single cheap accessory, using the idea of proportion of income." },
        { type: "qa", q: "**3c.** Give one reason why demand for \"3D-printed products\" in general is more inelastic than demand for \"Tom's specific grip design\" alone." },
      ],
    },
    {
      heading: "Section D — Exam-style practice",
      items: [
        { type: "mcq", q: "**4.** *(Paper 1 style MCQ)* A 10% rise in price causes quantity demanded to fall by 25%. This good's demand is:", options: ["A. Perfectly inelastic", "B. Inelastic", "C. Unitary", "D. Elastic"] },
        { type: "mcq", q: "**5.** *(Paper 1 style MCQ)* Which of the following would most likely make demand for a good MORE elastic?", options: ["A. The good is addictive", "B. The good has many close substitutes", "C. The good takes up a tiny share of a buyer's income", "D. The good is a necessity"] },
        { type: "qa", q: "**6.** *(Paper 2 style, 6 marks)* Explain **three** factors that determine whether demand for a good is elastic or inelastic. Use examples from Tom's Print Co., the padel club, or Lego in your answer." },
      ],
    },
  ],
  answerKey: [
    { type: "text", text: "**1a.** A measure of how responsive quantity demanded is to a change in price." },
    { type: "text", text: "**1b.** PED = % change in quantity demanded ÷ % change in price." },
    { type: "text", text: "**1c.**" },
    { type: "table", headers: ["PED value", "Category"], rows: [["0", "Perfectly inelastic"], ["Between 0 and 1", "Inelastic"], ["Exactly 1", "Unitary"], ["Greater than 1", "Elastic"], ["Infinite", "Perfectly elastic"]] },
    { type: "text", text: "**2a.** % change in price = +10%. % change in quantity = −20%. PED = −20/10 = −2. Elastic (ignoring the sign, size is 2, greater than 1)." },
    { type: "text", text: "**2b.** % change in price = +10%. % change in quantity = −5%. PED = −5/10 = −0.5. Inelastic (size is 0.5, between 0 and 1)." },
    { type: "text", text: "**2c.** % change in price = +25%. % change in quantity = −10%. PED = −10/25 = −0.4. Inelastic (size is 0.4, between 0 and 1)." },
    { type: "text", text: "**3a.** Collectors have no substitute for that exact retired set — nothing else will do — so they keep buying even at a higher price, making demand inelastic." },
    { type: "text", text: "**3b.** The membership fee is a much larger share of a family's leisure budget than a cheap accessory, so a price rise is far more noticeable and more likely to change behaviour — making demand more elastic. A cheap accessory's price change barely registers, so demand stays inelastic." },
    { type: "text", text: "**3c.** A broadly defined good (all 3D-printed products) has few or no substitutes — you'd have to go without entirely — whereas a narrowly defined good (one specific design) has an easy substitute in a competitor's version, so demand for the broad category is more inelastic." },
    { type: "text", text: "**4.** D — the percentage change in quantity demanded (25%) is larger than the percentage change in price (10%), so PED's size is greater than 1." },
    { type: "text", text: "**5.** B — more substitutes make it easier for buyers to switch away when price rises, making demand more responsive (elastic). A, C and D all point toward inelastic demand." },
    { type: "text", text: "**6.** Look for **three** developed points (AO2 chains), each identifying a determinant and explaining why it affects responsiveness, e.g.: (i) availability of substitutes — Tom's specific grip design has rival products buyers can switch to, making its demand elastic; (ii) necessity vs luxury — the padel club's court resurfacing is close to essential for the club to keep operating, making it inelastic; (iii) proportion of income — a Lego minifigure accessory is a tiny share of pocket money, so a price rise barely changes buying behaviour, making it inelastic. Full marks need identification + explanation of the effect on responsiveness, not just a list of factors." },
  ],
};

const LESSON4 = {
  title: "Lesson 4 Worksheet — PED: Revenue & Significance",
  syllabusRef: "2.6 (2.6.4 / 2.6.5)",
  sections: [
    {
      heading: "Section A — Definitions",
      items: [
        { type: "qa", q: "**1a.** Write out the formula for total revenue (TR)." },
        { type: "text", text: "**1b.** Fill in the table: for each PED category, state what happens to revenue when price rises." },
        { type: "table", headers: ["PED", "Effect on revenue when price RISES"], rows: [["Inelastic (PED < 1)", ""], ["Elastic (PED > 1)", ""], ["Unitary (PED = 1)", ""]] },
      ],
    },
    {
      heading: "Section B — New calculations",
      items: [
        { type: "qa", q: "**2a.** Tom's Print Co. raises the price of its Lego-compatible minifigure accessories from $5 to $6. Weekly sales fall from 200 to 140. Calculate PED, state whether demand is elastic or inelastic, then calculate TR before and after. Did revenue rise or fall?" },
        { type: "qa", q: "**2b.** The padel club raises its annual membership price from $300 to $330. Membership numbers fall from 150 to 145. Calculate PED, state whether demand is elastic or inelastic, then calculate TR before and after. Did revenue rise or fall?" },
        { type: "qa", q: "**2c.** Looking at 2a and 2b together: why did the same direction of price change (a rise) lead to opposite effects on revenue?" },
      ],
    },
    {
      heading: "Section C — Significance",
      items: [
        { type: "qa", q: "**3a.** Explain why a firm would want to know whether its product has elastic or inelastic demand before deciding to raise prices." },
        { type: "qa", q: "**3b.** Explain why a rise in the price of an inelastic necessity (like electricity) affects a household's budget more than a rise in the price of an elastic luxury (like a padel club's casual walk-in rate)." },
        { type: "qa", q: "**3c.** Explain why governments often place taxes on goods with inelastic demand, such as fuel or cigarettes." },
      ],
    },
    {
      heading: "Section D — Exam-style practice",
      items: [
        { type: "mcq", q: "**4.** *(Paper 1 style MCQ)* A firm's product has PED = −0.3. If the firm raises its price, total revenue will:", options: ["A. Rise", "B. Fall", "C. Stay exactly the same", "D. Become impossible to calculate"] },
        { type: "mcq", q: "**5.** *(Paper 1 style MCQ)* A firm wants to increase its total revenue and knows its product has highly elastic demand. It should:", options: ["A. Raise the price", "B. Lower the price", "C. Leave the price unchanged, since PED doesn't affect revenue", "D. Double the price"] },
        { type: "qa", q: "**6.** *(Paper 2 style, 6 marks)* Explain how a government's choice of which goods to tax might be influenced by price elasticity of demand. Use an example in your answer." },
      ],
    },
  ],
  answerKey: [
    { type: "text", text: "**1a.** TR = Price × Quantity." },
    { type: "text", text: "**1b.**" },
    { type: "table", headers: ["PED", "Effect on revenue when price RISES"], rows: [["Inelastic (PED < 1)", "Revenue rises"], ["Elastic (PED > 1)", "Revenue falls"], ["Unitary (PED = 1)", "No change"]] },
    { type: "text", text: "**2a.** % change in price = +20%. % change in quantity = −30%. PED = −30/20 = −1.5, elastic. TR before = $5 × 200 = $1000. TR after = $6 × 140 = $840. Revenue fell — consistent with elastic demand." },
    { type: "text", text: "**2b.** % change in price = +10%. % change in quantity ≈ −3.3% (145 is about a 3.3% fall from 150). PED ≈ −0.33, inelastic. TR before = $300 × 150 = $45,000. TR after = $330 × 145 = $47,850. Revenue rose — consistent with inelastic demand." },
    { type: "text", text: "**2c.** Because the two goods have different PED: the accessories have elastic demand (many substitute toys/accessories), so quantity fell proportionally more than price rose, cutting revenue. The membership has inelastic demand (members are committed, few close substitutes for their specific club), so quantity barely fell, and revenue rose. Same direction of price change, opposite effect on revenue, because of the underlying elasticity." },
    { type: "text", text: "**3a.** Knowing PED lets a firm predict what will actually happen to revenue, not just to price — a firm with inelastic demand can safely raise prices to raise revenue, but a firm with elastic demand would see revenue fall if it raised prices, so it should look at other strategies (e.g. improving quality, cutting costs) instead." },
    { type: "text", text: "**3b.** A household can't easily cut back on an inelastic necessity like electricity, so a price rise there directly increases their spending with no easy substitute — it eats into the budget. An elastic luxury like a casual padel booking is easy to cut back on or substitute away from, so a price rise there has much less impact on the household's overall spending, since they simply buy less of it." },
    { type: "text", text: "**3c.** Taxing an inelastic good means quantity demanded barely falls when the tax pushes the price up, so the government collects a large, reliable amount of tax revenue. Taxing an elastic good would cause a big fall in quantity demanded, undermining the amount of revenue the tax actually raises. Example: governments commonly tax fuel or cigarettes (both fairly inelastic, due to few substitutes/habit) rather than, say, luxury holidays (elastic, easily cut back on)." },
    { type: "text", text: "**4.** A — PED size is 0.3, inelastic, so a price rise increases revenue." },
    { type: "text", text: "**5.** B — with elastic demand, lowering price brings in proportionally more extra sales than the discount costs, raising total revenue." },
    { type: "text", text: "**6.** Look for: identification that governments prefer taxing goods with inelastic demand; explanation that this is because quantity demanded falls only a little when price rises due to the tax, so the tax raises substantial and predictable revenue; a named example (fuel, cigarettes, alcohol) with a brief reason why its demand is inelastic (necessity, habit-forming, few substitutes); ideally a contrast with what would happen taxing an elastic good instead (revenue undermined by a bigger fall in quantity). Full marks need the mechanism explained, not just \"governments tax inelastic goods.\"" },
  ],
};

async function run() {
  const BASE = path.join(__dirname, "..", "lessons");
  const jobs = [
    { lesson: "lesson-01", data: LESSON1 },
    { lesson: "lesson-02", data: LESSON2 },
    { lesson: "lesson-03", data: LESSON3 },
    { lesson: "lesson-04", data: LESSON4 },
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
