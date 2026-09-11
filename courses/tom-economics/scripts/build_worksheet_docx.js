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
  children.push(para(`**Syllabus ref: ${data.syllabusRef}** · ${data.subtitle || "Complete during your 30 min self-directed time."}`));

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

const LESSON5 = {
  title: "Lesson 5 Worksheet — Price Elasticity of Supply (PES)",
  syllabusRef: "2.7",
  sections: [
    {
      heading: "Section A — Definitions",
      items: [
        { type: "qa", q: "**1a.** Define price elasticity of supply (PES)." },
        { type: "qa", q: "**1b.** Write out the formula for PES." },
        { type: "qa", q: "**1c.** Is PES usually positive or negative? Explain why, using the law of supply." },
        { type: "text", text: "**1d.** Fill in the table: match each PES value range to its category." },
        { type: "table", headers: ["PES value", "Category"], rows: [["0", ""], ["Between 0 and 1", ""], ["Exactly 1", ""], ["Greater than 1", ""], ["Infinite", ""]] },
      ],
    },
    {
      heading: "Section B — Worked calculations",
      items: [
        { type: "qa", q: "**2a.** An order surge pushes the price of Tom's padel grips from $8 to $10. He ramps up weekly production from 60 to 90. Calculate PES and state whether supply is elastic or inelastic." },
        { type: "qa", q: "**2b.** Court hire prices at the padel club rise 15% for the peak evening slot, but the club cannot add courts before next year — quantity supplied stays unchanged. Calculate PES and state whether supply is elastic or inelastic." },
        { type: "qa", q: "**2c.** The price of a limited-edition padel paddle rises from $150 to $180 in the resale market. Because no more will ever be made, quantity supplied stays fixed. Calculate PES and state whether supply is elastic or inelastic." },
      ],
    },
    {
      heading: "Section C — Determinants",
      items: [
        { type: "qa", q: "**3a.** Explain why Tom's grip supply is more elastic than the padel club's court supply, using the idea of spare capacity." },
        { type: "qa", q: "**3b.** Explain why a padel coaching session (a service) is harder to supply elastically than a padel grip (a physical good), using the idea of ability to store stock." },
        { type: "qa", q: "**3c.** Give one reason why supply tends to become more elastic over a longer time period." },
      ],
    },
    {
      heading: "Section D — Exam-style practice",
      items: [
        { type: "mcq", q: "**4.** *(Paper 1 style MCQ)* A 5% rise in price causes quantity supplied to rise by 15%. This good's supply is:", options: ["A. Perfectly inelastic", "B. Inelastic", "C. Unitary", "D. Elastic"] },
        { type: "mcq", q: "**5.** *(Paper 1 style MCQ)* Which of the following would most likely make a firm's supply MORE elastic?", options: ["A. The firm is already running at full capacity", "B. The firm has spare machinery capacity it isn't using", "C. The good takes many months to produce", "D. The firm has no way to store finished stock"] },
        { type: "qa", q: "**6.** *(Paper 2 style, 6 marks)* Explain **three** factors that determine whether supply of a good is elastic or inelastic. Use examples from Tom's Print Co., the padel club, or Lego in your answer." },
      ],
    },
  ],
  answerKey: [
    { type: "text", text: "**1a.** A measure of how responsive quantity supplied is to a change in price." },
    { type: "text", text: "**1b.** PES = % change in quantity supplied ÷ % change in price." },
    { type: "text", text: "**1c.** Almost always positive. The law of supply says a higher price gives producers a stronger incentive to supply more, so price and quantity supplied move in the same direction — unlike demand, there's no sign to strip out." },
    { type: "text", text: "**1d.**" },
    { type: "table", headers: ["PES value", "Category"], rows: [["0", "Perfectly inelastic"], ["Between 0 and 1", "Inelastic"], ["Exactly 1", "Unitary"], ["Greater than 1", "Elastic"], ["Infinite", "Perfectly elastic"]] },
    { type: "text", text: "**2a.** % change in price = +25%. % change in quantity supplied = +50%. PES = 50/25 = 2. Elastic (greater than 1)." },
    { type: "text", text: "**2b.** % change in price = +15%. % change in quantity supplied = 0%. PES = 0/15 = 0. Perfectly inelastic." },
    { type: "text", text: "**2c.** % change in price = +20%. % change in quantity supplied = 0%. PES = 0/20 = 0. Perfectly inelastic." },
    { type: "text", text: "**3a.** Tom has an idle second printer he can switch on immediately, so a price rise lets him increase output fast with no new investment needed. The padel club has no spare courts sitting unused — every court is already in use — so it has no equivalent way to respond quickly." },
    { type: "text", text: "**3b.** A grip can be 3D-printed in advance and held as stock, ready to sell the moment price rises. A coaching session can't be produced ahead of time and stored — it only exists when a coach is actually delivering it — so supply is limited to however many coaches are available right now, whatever the price." },
    { type: "text", text: "**3c.** Over a longer time period, producers have more opportunity to build new capacity, retrain or hire staff, and new firms have time to enter the market — all of which let quantity supplied respond more fully to a price change." },
    { type: "text", text: "**4.** D — the percentage change in quantity supplied (15%) is larger than the percentage change in price (5%), so PES's size is greater than 1." },
    { type: "text", text: "**5.** B — spare machinery capacity lets a firm increase output quickly when price rises, without needing new investment, making supply more responsive (elastic). A, C and D all point toward inelastic supply." },
    { type: "text", text: "**6.** Look for **three** developed points (AO2 chains), each identifying a determinant and explaining why it affects responsiveness, e.g.: (i) spare capacity — Tom's idle second printer lets him ramp up output fast, making his supply elastic; (ii) length of production/gestation period — building a new padel court takes months, making court supply inelastic in the short run; (iii) ability to store stock — a Lego minifigure accessory can be produced in advance and stockpiled, making its supply more elastic than a service like coaching that can't be stored. Full marks need identification + explanation of the effect on responsiveness, not just a list of factors." },
  ],
};

const LESSON6 = {
  title: "Lesson 6 Worksheet — The Market Economic System",
  syllabusRef: "2.8",
  sections: [
    {
      heading: "Section A — Definitions",
      items: [
        { type: "qa", q: "**1a.** Define the price mechanism." },
        { type: "qa", q: "**1b.** Define a market economic system." },
        { type: "qa", q: "**1c.** Define effective demand." },
        { type: "qa", q: "**1d.** List the five characteristics of a market economic system." },
      ],
    },
    {
      heading: "Section B — The three basic questions",
      items: [
        { type: "qa", q: "**2a.** In a market economy, what decides *what* gets produced?" },
        { type: "qa", q: "**2b.** In a market economy, what decides *how* it gets produced?" },
        { type: "qa", q: "**2c.** In a market economy, what decides *for whom* it gets produced?" },
      ],
    },
    {
      heading: "Section C — Apply it: Tom's Print Co scenarios",
      items: [
        { type: "qa", q: "**3a.** Three rival 3D-print shops open near the padel club and compete hard for Tom's customers. Which characteristic of a market economy does this show, and what effect would you expect it to have on price and quality?" },
        { type: "qa", q: "**3b.** Tom's 3D printer produces plastic waste that ends up in a local stream, but this cost isn't reflected in the $8 price of a grip. Which disadvantage of the market system does this show? Name the concept." },
        { type: "qa", q: "**3c.** A family that wants to play padel cannot afford the court fees. Explain, using the idea of effective demand, why the market does not respond to their want to play." },
      ],
    },
    {
      heading: "Section D — Exam-style practice",
      items: [
        { type: "mcq", q: "**4.** *(Paper 1 style MCQ)* In a market economic system, resources are allocated mainly through:", options: ["A. Government planning committees", "B. The price mechanism", "C. A public vote", "D. Random allocation"] },
        { type: "mcq", q: "**5.** *(Paper 1 style MCQ)* Which of the following is a genuine disadvantage of a market economic system?", options: ["A. It always produces perfectly equal outcomes", "B. It automatically accounts for the cost of pollution", "C. It can leave genuine needs unmet if buyers cannot afford to pay", "D. It requires a large government planning department"] },
        { type: "qa", q: "**6.** *(Paper 2 style, 6 marks)* Explain **three** advantages and/or disadvantages of relying on the price mechanism to allocate resources. Use examples from Tom's Print Co. or the padel club in your answer." },
      ],
    },
  ],
  answerKey: [
    { type: "text", text: "**1a.** The system where prices, driven by supply and demand, coordinate what gets produced, how, and for whom — without central planning." },
    { type: "text", text: "**1b.** An economy where resources are allocated mainly through the price mechanism, with private ownership of resources and minimal government intervention." },
    { type: "text", text: "**1c.** Wanting a good or service *and* having the money to actually buy it — only this combination counts in the market; want alone is not enough." },
    { type: "text", text: "**1d.** Private ownership, profit motive, freedom of choice, competition, minimal government role." },
    { type: "text", text: "**2a.** Whatever earns a profit — signalled by what consumers are willing to pay for." },
    { type: "text", text: "**2b.** However keeps costs lowest — competition punishes producers who waste resources with lower profit." },
    { type: "text", text: "**2c.** Whoever is willing and able to pay the market price — not decided by need, by lottery, or by the state." },
    { type: "text", text: "**3a.** Competition. Rival producers competing for the same customers tends to push prices down and push quality/service up, since each shop has to work to win business rather than being guaranteed sales." },
    { type: "text", text: "**3b.** A negative externality. The cost of the pollution falls on people who live near the stream, who were not part of the transaction between Tom and his customer, so the market price does not reflect it." },
    { type: "text", text: "**3c.** The market only responds to effective demand — wanting something *and* being able to pay for it. The family wants to play but doesn't have the money to back that want up, so as far as the price mechanism is concerned, that demand doesn't exist — genuine need with no money behind it goes unmet." },
    { type: "text", text: "**4.** B — the price mechanism, driven by supply and demand, is what coordinates resource allocation in a market system." },
    { type: "text", text: "**5.** C — effective demand requires ability to pay, so real need with no money behind it goes unmet; A, B and D all describe planned-economy features or false claims about markets." },
    { type: "text", text: "**6.** Look for **three** developed points (AO2 chains), each identifying an advantage or disadvantage and explaining why it follows from relying on the price mechanism, e.g.: (i) efficient allocation — resources flow toward whatever padel players and Lego collectors actually want, without a planner needing to guess demand in advance; (ii) inequality — a family that cannot afford court fees is priced out even if their need to be active is genuine, since the market responds only to effective demand; (iii) ignored externalities — the plastic waste from Tom's printer isn't priced into a grip's cost, so the market allocates resources without accounting for that pollution. Full marks need identification + explanation of the effect, not just a list of terms." },
  ],
};

const LESSON7 = {
  title: "Lesson 7 Worksheet — Market Failure: Definitions, Causes & Consequences",
  syllabusRef: "2.9",
  subtitle: "This worksheet covers the full topic in one go — allow around 40 minutes of self-directed time.",
  sections: [
    {
      heading: "Section A — Definitions",
      items: [
        { type: "qa", q: "**1a.** Define market failure." },
        { type: "qa", q: "**1b.** Define social cost, using the ideas of private cost and external cost." },
        { type: "qa", q: "**1c.** Define externality." },
        { type: "qa", q: "**1d.** Define a public good, using excludability and rivalry." },
        { type: "qa", q: "**1e.** Define a merit good and a demerit good." },
        { type: "qa", q: "**1f.** Define socially efficient output, and explain how it differs from market output." },
      ],
    },
    {
      heading: "Section B — Which externality is this?",
      items: [
        { type: "qa", q: "**2a.** A local factory burns waste materials while manufacturing, and the smoke drifts over the padel courts during a match. Which type of externality is this?" },
        { type: "qa", q: "**2b.** A new bakery's fresh-bread smell each morning, a side effect of baking, makes the whole street more pleasant for everyone who walks past. Which type of externality is this?" },
        { type: "qa", q: "**2c.** Someone playing loud music from their car stereo outside the padel club disturbs residents trying to relax at home. Which type of externality is this?" },
      ],
    },
    {
      heading: "Section C — Which consequence is this?",
      items: [
        { type: "qa", q: "**3a.** A dry-cleaning shop's chemical fumes irritate nearby office workers, a cost never reflected in its prices, so the shop keeps operating at full capacity. Over-provided, under-provided, or not provided at all?" },
        { type: "qa", q: "**3b.** A private tutoring company could offer subsidised classes for struggling students (a clear benefit to the wider community), but doesn't, since it can't charge for that wider benefit. Over-provided, under-provided, or not provided at all?" },
        { type: "qa", q: "**3c.** No private company installs flood barriers along a public riverbank, even though every nearby resident would benefit, because no single resident will pay for something everyone gets free. Over-provided, under-provided, or not provided at all?" },
      ],
    },
    {
      heading: "Section D — Apply it: Tom's Print Co & the padel club",
      items: [
        { type: "qa", q: "**4a.** A padel club installs floodlights so members can play after dark, and charges a small extra fee for evening slots. Is floodlit court time a private good or a public good? Explain using excludability and rivalry." },
        { type: "qa", q: "**4b.** Explain, using the ideas of private cost, external cost, and social cost, why the market price of a padel grip does not reflect the true cost to society of producing it — and why this leads to overproduction." },
        { type: "qa", q: "**4c.** Tom is thinking about running a free basic coding workshop for local kids using his 3D printer, but is unsure it's worth his time since he can't charge much for it. Using the idea of a merit good, explain why the market alone might under-provide something like this." },
        { type: "qa", q: "**4d.** Explain why a producer like Tom has no built-in incentive to fix a negative externality on his own, without government involvement." },
      ],
    },
    {
      heading: "Section E — Exam-style practice",
      items: [
        { type: "mcq", q: "**5.** *(Paper 1 style MCQ)* Which of the following best describes a positive externality?", options: ["A. A cost imposed on a third party", "B. A benefit received only by the producer", "C. A benefit that spills over onto a third party not involved in the transaction", "D. A cost that is always reflected in the market price"] },
        { type: "mcq", q: "**6.** *(Paper 1 style MCQ)* A good is non-excludable and non-rivalrous. This is best described as:", options: ["A. A private good", "B. A demerit good", "C. A public good", "D. A merit good"] },
        { type: "mcq", q: "**7.** *(Paper 1 style MCQ)* When a negative externality exists, the market tends to produce:", options: ["A. Exactly the socially efficient amount", "B. More than the socially efficient amount", "C. Less than the socially efficient amount", "D. Nothing at all"] },
        { type: "mcq", q: "**8.** *(Paper 1 style MCQ)* The free-rider problem is a direct cause of which consequence?", options: ["A. Overproduction of demerit goods", "B. Zero private provision of public goods", "C. Perfectly efficient markets", "D. Falling prices for merit goods"] },
        { type: "qa", q: "**9.** *(Paper 2 style, 6 marks)* Explain, using examples, the causes of market failure and their consequences for **two** different groups (choose from consumers, producers, and government). Use examples from Tom's Print Co or the padel club in your answer." },
      ],
    },
  ],
  answerKey: [
    { type: "text", text: "**1a.** Market failure is when the price mechanism fails to allocate resources efficiently, leading to over- or under-production of a good or service compared to what's best for society." },
    { type: "text", text: "**1b.** Social cost is the full cost to society of an economic activity — private cost (what the producer/consumer pays) plus external cost (the cost dumped on third parties not involved in the transaction)." },
    { type: "text", text: "**1c.** An externality is a cost or benefit of a transaction that spills over onto a third party not involved in it." },
    { type: "text", text: "**1d.** A public good is non-excludable (nobody can be stopped from benefiting even without paying) and non-rivalrous (one person's use doesn't reduce how much is available to others)." },
    { type: "text", text: "**1e.** A merit good is under-provided by the market because people undervalue its benefit to themselves, or ignore the positive externality it creates for others (e.g. education, healthcare). A demerit good is over-provided because people undervalue its cost to themselves, or ignore the negative externality it creates for others (e.g. cigarettes)." },
    { type: "text", text: "**1f.** Socially efficient output is the amount of a good or service that's actually best for society once all costs and benefits are counted. Market output is how much actually gets produced based only on private cost and private benefit — the two differ whenever an externality exists." },
    { type: "text", text: "**2a.** Negative production externality — the pollution cost is created during the factory's manufacturing process and imposed on people uninvolved in the transaction." },
    { type: "text", text: "**2b.** Positive production externality — the pleasant smell is a side effect of the bakery's production process, benefiting passersby who paid nothing for it." },
    { type: "text", text: "**2c.** Negative consumption externality — the disturbance is created while the good (the music) is being consumed, imposed on residents nearby." },
    { type: "text", text: "**3a.** Over-provided — a negative externality (the fumes) isn't reflected in the shop's costs, so it keeps producing past the socially efficient point." },
    { type: "text", text: "**3b.** Under-provided — a positive externality (community benefit) isn't captured as revenue, so less gets provided than would be socially efficient." },
    { type: "text", text: "**3c.** Not provided at all — a classic public good; the free-rider problem means no private firm can profit from providing it." },
    { type: "text", text: "**4a.** A private good. Excludable — the club can refuse entry to evening slots to anyone who won't pay the extra fee. Rivalrous — if one group is booked onto the court for that slot, no other group can use the same court at the same time." },
    { type: "text", text: "**4b.** The $8 price only reflects Tom's private cost (materials and electricity). It leaves out the external cost of the plastic pollution affecting people near the stream. Social cost = private cost + external cost, so the true cost to society of producing that grip is higher than the market price shows. Since the market only \"sees\" private cost, production keeps going past the point where it's actually efficient for society — leading to overproduction." },
    { type: "text", text: "**4c.** It's a merit good. Kids and parents may undervalue the long-term benefit of coding skills, and there's a positive externality to the wider community (more skilled future workers) that Tom isn't paid for. Because Tom can't capture the full value of what he creates, the free market alone tends to under-provide it." },
    { type: "text", text: "**4d.** There's no market signal telling Tom to change — nobody is actually charging him for the pollution he creates, so from his own private cost-and-revenue perspective, nothing looks wrong. Only an outside intervention (like government regulation or a tax) would change that signal." },
    { type: "text", text: "**5.** C — a positive externality is a benefit that spills over onto a third party not involved in the transaction; A and D describe costs, not benefits, and B describes a purely private benefit with no spillover." },
    { type: "text", text: "**6.** C — non-excludable and non-rivalrous is the definition of a public good." },
    { type: "text", text: "**7.** B — a negative externality means social cost exceeds private cost, so the market overproduces relative to the socially efficient amount." },
    { type: "text", text: "**8.** B — since non-payers can't be excluded from a public good, no private firm has a profit incentive to supply it, so the market provides none at all." },
    { type: "text", text: "**9.** Look for **two** developed points (AO2 chains), each naming a cause of market failure, identifying the consequence it produces, and explaining the effect on a specific group, e.g.: (i) consumers — a family near Tom's stream suffers the effects of pollution (a negative externality causing overproduction) they never agreed to and aren't compensated for; (ii) producers — Tom has no financial incentive to change his production process, since the externality cost never appears in his own accounts; (iii) government — must judge whether the gap between market output and socially efficient output (e.g. for the padel club's community taster sessions, a case of underproduction) is large enough to justify a policy response. Full marks need identification of the cause + the consequence + explanation of the effect on that specific group, not just a list of terms." },
  ],
};

const REVIEW1 = {
  title: "Review Session 1 Worksheet — PED, Total Revenue & PES",
  syllabusRef: "2.6 / 2.7",
  sections: [
    {
      heading: "Section A — Definitions",
      items: [
        { type: "qa", q: "**1a.** Define PED, in your own words (not just the formula)." },
        { type: "qa", q: "**1b.** Define total revenue and give its formula." },
        { type: "qa", q: "**1c.** Define PES, in your own words." },
        { type: "qa", q: "**1d.** PED and PES have the same shape of formula, but one key difference. What is it, and why?" },
      ],
    },
    {
      heading: "Section B — PED practice",
      items: [
        { type: "qa", q: "**2a.** Padel club sun hats: price $10 → $8, weekly sales 40 → 60. Calculate the PED and classify it." },
        { type: "qa", q: "**2b.** Tom's Print Co keychain multipack: price $6 → $9, weekly sales 90 → 30. Calculate the PED and classify it." },
        { type: "qa", q: "**2c.** Padel club bottled water: price $2 → $2.20, weekly sales 300 → 285. Calculate the PED and classify it." },
      ],
    },
    {
      heading: "Section C — PED + total revenue",
      items: [
        { type: "qa", q: "**3a.** Padel club adult casual entry: price $15 → $18, weekly visits 200 → 150. Calculate the PED, then TR before and after. Did revenue rise or fall?" },
        { type: "qa", q: "**3b.** Tom's Print Co desk lamps: price $30 → $24, weekly sales 20 → 30. Calculate the PED, then TR before and after. Did revenue rise or fall?" },
      ],
    },
    {
      heading: "Section D — PES practice",
      items: [
        { type: "qa", q: "**4a.** Padel club ball-machine rental: price $8 → $10, quantity supplied 50 → 60 per week. Calculate the PES and classify it." },
        { type: "qa", q: "**4b.** Tom's Print Co custom keyring output: price $5 → $6, quantity supplied 100 → 140 per week. Calculate the PES and classify it." },
        { type: "qa", q: "**4c.** A vintage Lego minifigure (no more can ever be made): price $40 → $60, quantity supplied stays at 15 either way. Calculate the PES and classify it." },
      ],
    },
    {
      heading: "Section E — PED or PES? Decide first",
      items: [
        { type: "qa", q: "**5a.** Padel club coaching price rises $25 → $30, and weekly attendance falls 80 → 64. Is this a PED or PES scenario? Calculate it." },
        { type: "qa", q: "**5b.** Tom's grip price rises $8 → $9, and his own weekly output rises 50 → 65. Is this a PED or PES scenario? Calculate it." },
      ],
    },
    {
      heading: "Section F — Exam-style practice",
      items: [
        { type: "mcq", q: "**6.** *(Paper 1 style MCQ)* A firm with elastic demand wants to increase its total revenue. It should:", options: ["A. Raise its price", "B. Cut its price", "C. Leave its price unchanged", "D. It is impossible to increase revenue"] },
        { type: "mcq", q: "**7.** *(Paper 1 style MCQ)* Which of the following would make a firm's PES more elastic?", options: ["A. No spare production capacity", "B. A long, slow production process", "C. Spare machinery capacity ready to switch on", "D. Being the only producer in the market"] },
        { type: "qa", q: "**8.** *(Paper 2 style, 6 marks)* Explain, using PED, total revenue, and PES, why a business owner and a government might each care about elasticity when making a decision. Use examples from Tom's Print Co or the padel club." },
      ],
    },
  ],
  answerKey: [
    { type: "text", text: "**1a.** A measure of how much quantity demanded changes when price changes, how strongly buyers react." },
    { type: "text", text: "**1b.** The total money a firm receives from sales, before costs. TR = Price × Quantity." },
    { type: "text", text: "**1c.** A measure of how much quantity supplied changes when price changes, how strongly sellers can react." },
    { type: "text", text: "**1d.** PED is (almost) always negative, since price and quantity demanded move in opposite directions. PES is (almost) always positive, since price and quantity supplied move in the same direction (the law of supply)." },
    { type: "text", text: "**2a.** % price = −20%, % quantity = +50%. PED = 50 ÷ −20 = −2.5 → Elastic." },
    { type: "text", text: "**2b.** % price = +50%, % quantity = −66.7%. PED = −66.7 ÷ 50 = −1.33 → Elastic." },
    { type: "text", text: "**2c.** % price = +10%, % quantity = −5%. PED = −5 ÷ 10 = −0.5 → Inelastic." },
    { type: "text", text: "**3a.** % price = +20%, % quantity = −25%. PED = −25 ÷ 20 = −1.25 → Elastic. TR before = $15 × 200 = $3,000. TR after = $18 × 150 = $2,700. Revenue FELL, elastic demand, price rise, revenue falls." },
    { type: "text", text: "**3b.** % price = −20%, % quantity = +50%. PED = 50 ÷ −20 = −2.5 → Elastic. TR before = $30 × 20 = $600. TR after = $24 × 30 = $720. Revenue ROSE, elastic demand, price cut, revenue rises." },
    { type: "text", text: "**4a.** % price = +25%, % quantity supplied = +20%. PES = 20 ÷ 25 = 0.8 → Inelastic." },
    { type: "text", text: "**4b.** % price = +20%, % quantity supplied = +40%. PES = 40 ÷ 20 = 2.0 → Elastic." },
    { type: "text", text: "**4c.** % price = +50%, % quantity supplied = 0%. PES = 0 ÷ 50 = 0 → Perfectly inelastic. Once production stops, no price can increase how many exist." },
    { type: "text", text: "**5a.** PED, this describes buyers' reaction (attendance) to a price change. % price = +20%, % quantity = −20%. PED = −20 ÷ 20 = −1.0 → Unitary." },
    { type: "text", text: "**5b.** PES, this describes the seller's (Tom's) own output response to a price change. % price = +12.5%, % quantity supplied = +30%. PES = 30 ÷ 12.5 = 2.4 → Elastic." },
    { type: "text", text: "**6.** B, with elastic demand, a price cut brings in proportionally more extra sales than the discount costs, growing revenue." },
    { type: "text", text: "**7.** C, spare capacity lets a producer ramp up output quickly when price rises, making supply more elastic." },
    { type: "text", text: "**8.** Look for **two** developed points (AO2 chains), each naming a decision-maker and explaining how elasticity informs it, e.g.: (i) a business owner like Tom uses PED and TR together to decide whether raising or cutting price grows revenue; (ii) a government taxing an inelastic good like cigarettes knows demand won't fall much, so the tax mainly raises revenue rather than changing behaviour; (iii) a producer uses PES to judge how fast they can respond to a price change, e.g. spare printer capacity versus a padel court that takes months to build. Full marks need identification + explanation of the effect on the decision, not just definitions." },
  ],
};

async function run() {
  const BASE = path.join(__dirname, "..", "lessons");
  const jobs = [
    { lesson: "lesson-01", data: LESSON1 },
    { lesson: "lesson-02", data: LESSON2 },
    { lesson: "lesson-03", data: LESSON3 },
    { lesson: "lesson-04", data: LESSON4 },
    { lesson: "lesson-05", data: LESSON5 },
    { lesson: "lesson-06", data: LESSON6 },
    { lesson: "lesson-07", data: LESSON7 },
    { lesson: "review-01-ped-tr-pes", data: REVIEW1 },
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
