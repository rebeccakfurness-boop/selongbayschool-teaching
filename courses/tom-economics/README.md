# Tom — Cambridge IGCSE Economics 0455

**Open `dashboard.html` in any browser — no install, no server.** Your progress (done/not-done, re-teach flags) is saved in the browser's local storage tied to that file, so it keeps working offline and survives closing the tab.

## Layout

```
tom-economics/
  dashboard.html          <- open this
  data/
    calendar.json         <- school holidays & public holidays (26/27)
    syllabus.json          <- full 0455 topic tree, with known/unknown flags
    lessons.json           <- the 74-lesson dated sequence
  lessons/
    lesson-01/
      guide.html           <- teaching script for the tutor
      worksheet.html       <- practice questions + answer key (boxed off at the bottom), quick-view
      worksheet.docx       <- same worksheet as an editable/printable Word doc
      slides.html           <- presentable slide deck (arrow keys / on-screen buttons to navigate), quick-view
      slides.pptx            <- same deck as a real PowerPoint file you can present from
      flashcards.txt        <- paste into NotebookLM > Studio > Flashcards
    lesson-02/ ...
  scripts/                 <- regeneration tooling (see below)
```

The dashboard always links to the fast `.html` versions for between-lesson viewing, plus small `(.docx)` / `(.pptx)` links next to Worksheet/Slides whenever those exist, for editing, printing, or presenting from PowerPoint/Word directly.

## Asking Claude Code to plan a lesson

Say **"plan lesson 7"** (or whichever number) and Claude Code will generate that lesson's `guide.md`/`worksheet.md` (converted to `.html` + `.docx`), `slides.html` + `slides.pptx`, and `flashcards.txt` in `lessons/lesson-07/`, all tied to that lesson's syllabus reference from `data/lessons.json`, then re-run the scripts below so the dashboard picks it up automatically.

## Regenerating things

- **`scripts/build_data.py`** — rebuilds `data/*.json`. Run after changing the holiday dates, the syllabus known/unknown flags, or the lesson sequence in the script itself. Each lesson's `files` flags (which formats exist) are auto-detected by checking what's actually on disk in that lesson's folder, so you never need to hand-edit them.
- **`scripts/inject_data.py`** — rebuilds `dashboard.html` from `scripts/dashboard_template.html` + the current `data/*.json`. Run this after `build_data.py`, or after editing the template itself.
- **`scripts/convert_docs.py <lesson-folder>`** — converts a lesson's `guide.md`/`worksheet.md` into styled, printable `guide.html`/`worksheet.html` (the answer key is auto-boxed off). Has a few markdown formatting gotchas noted at the top of the file — read them before writing new lesson markdown.
- **`scripts/build_worksheet_docx.js`** (`node build_worksheet_docx.js`) — generates `worksheet.docx` for a lesson from a structured content object defined at the top of the file (Word docs don't map cleanly from arbitrary markdown/HTML, so worksheets are authored directly in this shape — add a new lesson by adding an entry).
- **`scripts/build_slides_pptx.js`** (`node build_slides_pptx.js`) — generates `slides.pptx` for a lesson using reusable slide-layout helpers (title slide, factor/icon slide, two-column, demand-supply diagram, closing slide) — add a new lesson by calling these helpers with that lesson's content.

Both `.js` scripts need `npm install` once inside `scripts/` (docx, pptxgenjs — not committed, see `.gitignore`).

Typical flow for a new lesson: write `guide.md` + `worksheet.md` + `slides.html` + `flashcards.txt` into a new `lessons/lesson-NN/` folder → run `convert_docs.py` on that folder → delete the now-redundant `.md` files → add that lesson's content to `build_worksheet_docx.js` and `build_slides_pptx.js` and run both → run `build_data.py` then `inject_data.py`.

## Adding a second student/subject later

Duplicate this whole `tom-economics/` folder as a sibling (e.g. `courses/name-subject/`), swap in a new syllabus PDF + calendar + weekly slots, and re-run the same scripts against the new data. No shared database or shared code to untangle — each course is self-contained.
