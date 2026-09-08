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

const doc = new Document({
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

const OUT = path.join(__dirname, "..", "lessons", "lesson-02-respiratory-system", "worksheet.docx");
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("wrote", OUT);
});
