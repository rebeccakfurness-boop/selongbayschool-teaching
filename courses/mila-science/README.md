# Mila — Cambridge Primary Science, Stage 6

No dashboard yet (only started with Lesson 2) — open a lesson's `lesson.html` directly in
a browser, or use the published interactive artifact.

## Layout

```
mila-science/
  assets/
    lesson-kit.css       <- shared styling (same structure as tom-economics' kit, re-themed:
                             sky blue/leaf green/coral for primary science instead of teal/coral)
    lesson-kit.js        <- shared interactive widgets — identical to tom-economics' kit,
                             copied as-is since it's subject-agnostic
  lessons/
    lesson-02-respiratory-system/
      guide.html          <- teaching script for the tutor
      lesson.html         <- the interactive single-page tool, worked through live with Mila
      worksheet.html      <- practice questions + answer key (boxed off), quick-view
      worksheet.docx       <- same worksheet as an editable/printable Word doc
  scripts/
    build_worksheet_docx.js  <- regenerates worksheet.docx (node build_worksheet_docx.js;
                                 npm install once first — docx package, not committed)
```

Lesson numbering starts at 2 because Lesson 1 (1.1 The circulatory system) was an existing
worksheet made outside this system — it isn't reproduced here, only referenced as prior
knowledge in Lesson 2's recap.

## Content alignment

Lessons are aligned to **Cambridge Primary Science Workbook 6** (Fiona Baxter & Liz Dilley) —
matching its unit numbering (1.2, 1.3, ...) and the shape of its Focus/Practice/Challenge
questions — but every diagram, sentence, and data set is written fresh, not reproduced from
the workbook. The workbook itself has no teaching content (it's practice-only, meant to
accompany a separate Coursebook), so each lesson's explanations, diagrams, and interactive
activities are original.

## Diagrams

The respiratory-system illustrations in Lesson 2 (`lesson.html`) are original inline SVG —
simple labelled shapes (ellipses, rounded-cap stroke paths) rather than photorealistic
anatomy, matching the plain, friendly illustration style appropriate for this age group.

## Adding the next lesson

Follow 1.3 The reproductive system next, continuing the same numbering and folder pattern as
`lesson-02-respiratory-system`. No `dashboard.html`/`data/*.json` pipeline exists yet for this
course (see `tom-economics/` for that pattern, if it's ever worth building here too).
