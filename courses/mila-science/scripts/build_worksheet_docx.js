// Builds worksheet.docx for Lesson 2 (1.2 The respiratory system) from a structured
// content object, matching the same shape/approach as tom-economics/scripts/build_worksheet_docx.js.
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType,
} = require("docx");
const fs = require("fs");
const path = require("path");

const BLUE = "1C93A6";
const ORANGE = "F0663B";
const AMBER_LIGHT = "FDEED8";
const GRID = "DCEAEC";

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

function wordBox(words) {
  return new Paragraph({
    spacing: { after: 200 },
    shading: { type: ShadingType.CLEAR, fill: GRID },
    children: [new TextRun({ text: words.join("    "), bold: true })],
  });
}

function makeTable(headerRow, rows) {
  const colCount = headerRow.length;
  const colWidth = Math.floor(9000 / colCount);
  const widths = new Array(colCount).fill(colWidth);
  const headerCells = headerRow.map(h => new TableCell({
    width: { size: colWidth, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: GRID },
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, bold: true })] })],
  }));
  const bodyRows = rows.map(r => new TableRow({
    children: r.map(c => new TableCell({
      width: { size: colWidth, type: WidthType.DXA },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: runs(c || "") })],
    })),
  }));
  return new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: widths,
    rows: [new TableRow({ children: headerCells }), ...bodyRows],
  });
}

function drawBox(label) {
  return new Table({
    width: { size: 9000, type: WidthType.DXA },
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: 9000, type: WidthType.DXA },
        borders: {
          top: { style: BorderStyle.DASHED, size: 6, color: "7B939D" },
          bottom: { style: BorderStyle.DASHED, size: 6, color: "7B939D" },
          left: { style: BorderStyle.DASHED, size: 6, color: "7B939D" },
          right: { style: BorderStyle.DASHED, size: 6, color: "7B939D" },
        },
        children: [
          new Paragraph({ text: "" }), new Paragraph({ text: "" }), new Paragraph({ text: "" }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: label, italics: true, color: "7B939D" })] }),
          new Paragraph({ text: "" }), new Paragraph({ text: "" }), new Paragraph({ text: "" }),
        ],
      })],
    })],
  });
}

function flowRow(words) {
  const cells = [];
  words.forEach((w, i) => {
    cells.push(new TableCell({
      width: { size: 2000, type: WidthType.DXA },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: " " })] })],
    }));
    if (i < words.length - 1) {
      cells.push(new TableCell({
        width: { size: 500, type: WidthType.DXA },
        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "→" })] })],
      }));
    }
  });
  return new Table({ width: { size: 9000, type: WidthType.DXA }, rows: [new TableRow({ children: cells })] });
}

const doc1 = new Document({
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 } } },
    children: [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "Lesson 1 Worksheet — The Circulatory System", bold: true })],
        spacing: { after: 80 },
      }),
      para("**Cambridge Primary Science, Stage 6 · Unit 1.1**"),

      heading("Focus"),
      para("**1.** Name the three parts of the circulatory system."),
      para("a _______________     b _______________     c _______________"),
      para("**2.** Use the words in the box to complete the sentences. You will use some words more than once."),
      wordBox(["blood vessels", "oxygen", "blood", "waste products", "food", "lungs"]),
      para("The heart pumps _______________ through the body."),
      para("The left side of the heart pumps _______________ that contains _______________."),
      para("The right side of the heart pumps _______________ without _______________ to the _______________."),
      para("Blood is carried in the _______________."),
      para("Blood carries _______________ and _______________ to all parts of the body and takes away _______________."),
      para("**3.** In your own words:"),
      para("**a** What two things does blood carry *to* your cells, and why do your cells need them?\nAnswer:"),
      para("**b** What does blood carry *away* from your cells, and why does it need to get rid of it?\nAnswer:"),

      heading("Practice"),
      para("**4.** Match each word in List 1 with its clue in List 2. Draw a line to link each one (the order is scrambled on purpose!)."),
      makeTable(["List 1", "List 2"], [
        ["Artery", "Tiny tube with super-thin walls, connects the other two"],
        ["Vein", "Carries blood away from the heart"],
        ["Capillary", "Carries blood back to the heart"],
      ]),
      new Paragraph({ text: "", spacing: { after: 160 } }),
      para("**5.** Circle the letter of the correct answer."),
      para("**a** Your heart pumping blood through the body is called...\nA. heartbeat    B. circulation    C. pulsing"),
      para("**b** The circulatory system is made up of...\nA. the heart only    B. the heart and blood vessels    C. the heart, blood vessels, and blood"),
      para("**c** Which blood vessels bring oxygen to body cells and carry away waste?\nA. arteries    B. veins    C. capillaries"),
      para("**d** The organ that helps you think and controls your whole body is the...\nA. heart    B. brain    C. stomach"),

      heading("Challenge"),
      para("**6.** Priya measured her pulse rate sitting still, then after five different activities. Here are her results."),
      makeTable(
        ["Activity", "Pulse rate (bpm)"],
        [["Sitting still", "78"], ["Walking", "90"], ["Climbing stairs", "105"], ["Jumping jacks", "118"], ["Dancing", "130"]]
      ),
      new Paragraph({ text: "", spacing: { after: 160 } }),
      para("**a** Draw a bar chart of Priya's results. Use a different colour for each activity."),
      drawBox("Bar chart"),
      new Paragraph({ text: "", spacing: { after: 200 } }),
      para("**b** When was Priya's pulse rate lowest? Why do you think that is?\nAnswer:"),
      para("**c** Which activity caused the highest pulse rate?\nAnswer:"),
      para("**d** Predict what would happen to Priya's pulse rate if she danced for even longer. Explain your answer.\nAnswer:"),
      para("**e** Write one sentence to conclude what Priya's results show.\nAnswer:"),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 360, after: 140 },
        children: [new TextRun({ text: "Answer key (tutor copy — not for Mila)", bold: true, color: ORANGE })],
        shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT },
      }),
      para("**1.** a. the heart   b. blood vessels   c. blood", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**2.** The heart pumps **blood** through the body. The left side of the heart pumps **blood** that contains **oxygen**. The right side of the heart pumps **blood** without **oxygen** to the **lungs**. Blood is carried in the **blood vessels**. Blood carries **food** and **oxygen** to all parts of the body and takes away **waste products**.", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**3a.** Blood carries food and oxygen to your cells. Cells need food for fuel/energy, and oxygen to release the energy from that food.", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**3b.** Blood carries away waste, such as carbon dioxide. It needs to get rid of it because it's a waste gas the body doesn't need — the lungs and kidneys are the organs that help get rid of it.", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**4.** Artery → carries blood away from the heart. Vein → carries blood back to the heart. Capillary → tiny tube with super-thin walls, connects the other two.", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**5a.** B — circulation. **5b.** C — the heart, blood vessels, and blood. **5c.** C — capillaries (their thin walls let oxygen and food pass through to cells, and waste pass back). **5d.** B — the brain.", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**6a.** A bar chart with five bars rising roughly in order: sitting (lowest) → walking → climbing stairs → jumping jacks → dancing (highest).", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**6b.** Sitting still, at 78 bpm — her body wasn't working hard, so her muscles didn't need much extra oxygen delivered.", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**6c.** Dancing, at 130 bpm — likely her most physically demanding activity in this set.", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**6d.** Accept any reasonable answer with a reason, e.g. it would likely stay high or rise a little further while dancing continues, since her muscles keep needing extra oxygen delivered quickly, then fall back down once she stops and rests.", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**6e.** The harder Priya's body works, the faster her pulse rate gets, because her heart needs to deliver oxygen and food to her muscles more quickly.", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
    ],
  }],
});

const doc2 = new Document({
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 } } },
    children: [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "Lesson 2 Worksheet — The Respiratory System", bold: true })],
        spacing: { after: 80 },
      }),
      para("**Cambridge Primary Science, Stage 6 · Unit 1.2**"),

      heading("Focus"),
      para("**1.** Use the words in the box to complete the sentences. You will use some words more than once."),
      wordBox(["blood", "ribs", "lungs", "windpipe", "nose", "carbon dioxide", "oxygen"]),
      para("We breathe in air through our _______________. The air we breathe in contains _______________ gas. The air moves down the _______________ and into our _______________."),
      para("The _______________ in the air then moves from the _______________ into the _______________. We breathe out air that contains _______________ gas. The _______________ protect our respiratory system."),

      heading("Practice"),
      para("**2.** The box on the left shows the lungs when breathing **OUT**. In the empty box, draw and label what the lungs look like when breathing **IN**. Think about: does the chest get bigger or smaller? Do the ribs move in or out?"),
      drawBox("Breathing OUT: ribs in, diaphragm domed up, lungs smaller"),
      new Paragraph({ text: "", spacing: { after: 120 } }),
      drawBox("Draw & label: Breathing IN"),
      new Paragraph({ text: "", spacing: { after: 200 } }),
      para("**3.** Put these words in order to show the path of oxygen when we breathe in: **lungs, nose, blood, windpipe**."),
      flowRow(["", "", "", ""]),
      new Paragraph({ text: "", spacing: { after: 200 } }),

      heading("Challenge"),
      para("**4.** A class measured the breathing rate and pulse rate of 8 people straight after light exercise. Here are their results."),
      makeTable(
        ["Person", "Breathing rate (breaths/min)", "Pulse rate (bpm)"],
        [["1", "20", "80"], ["2", "24", "92"], ["3", "28", "100"], ["4", "22", "88"],
         ["5", "32", "112"], ["6", "26", "96"], ["7", "30", "145"], ["8", "34", "118"]]
      ),
      new Paragraph({ text: "", spacing: { after: 160 } }),
      para("**a** Draw a scatter graph of the results (breathing rate along the bottom, pulse rate up the side)."),
      drawBox("Scatter graph"),
      new Paragraph({ text: "", spacing: { after: 200 } }),
      para("**b** Describe the pattern you observe in the results.\nAnswer:"),
      para("**c i** Identify any result that does not fit the pattern.\nAnswer:"),
      para("**ii** Suggest a reason for this.\nAnswer:"),
      para("**d** Use the graph to predict the pulse rate of a person whose breathing rate is 36 breaths per minute.\nAnswer:"),
      para("**e** Suggest a conclusion the class can make from these results.\nAnswer:"),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 360, after: 140 },
        children: [new TextRun({ text: "Answer key (tutor copy — not for Mila)", bold: true, color: ORANGE })],
        shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT },
      }),
      para("**1.** We breathe in air through our **nose**. The air we breathe in contains **oxygen** gas. The air moves down the **windpipe** and into our **lungs**. The **oxygen** in the air then moves from the **lungs** into the **blood**. We breathe out air that contains **carbon dioxide** gas. The **ribs** protect our respiratory system.", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**2.** Breathing IN: the ribs move outward and upward, the diaphragm flattens and pulls down, and the lungs get bigger, taking up more space in the chest than in the \"breathing out\" picture.", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**3.** nose → windpipe → lungs → blood.", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**4a.** A rising trend: as breathing rate increases, pulse rate generally increases too, except for person 7.", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**4b.** As breathing rate goes up, pulse rate also tends to go up — the two are linked, because harder-working muscles need more oxygen delivered faster (higher pulse) and taken in faster (higher breathing rate).", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**4c i.** Person 7 (30 breaths/min, 145 bpm) does not fit the pattern — their pulse rate is much higher than other people with a similar breathing rate.", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**4c ii.** Accept any sensible suggestion, e.g. this person may have been more nervous or anxious, may have measured their pulse incorrectly, may have a different fitness level, or found the exercise personally harder.", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**4d.** Accept any value that reasonably continues the trend, e.g. around 125–130 bpm (following the pattern set by persons 3, 5, 6 and 8, ignoring the anomaly at person 7).", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
      para("**4e.** Breathing rate and pulse rate rise together during exercise, because the body needs to take in more oxygen and deliver it faster to working muscles. One result (person 7) did not fit this pattern and may need to be checked or re-measured.", { shading: { type: ShadingType.CLEAR, fill: AMBER_LIGHT } }),
    ],
  }],
});

const jobs = [
  { doc: doc1, out: path.join(__dirname, "..", "lessons", "lesson-01-circulatory-system", "worksheet.docx") },
  { doc: doc2, out: path.join(__dirname, "..", "lessons", "lesson-02-respiratory-system", "worksheet.docx") },
];
(async () => {
  for (const job of jobs) {
    const buf = await Packer.toBuffer(job.doc);
    fs.writeFileSync(job.out, buf);
    console.log("wrote", job.out);
  }
})();
