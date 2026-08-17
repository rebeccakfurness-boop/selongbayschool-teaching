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
  assets/
    lesson-kit.css          <- shared styling for every lesson.html
    lesson-kit.js            <- shared interactive widgets (flip cards, sort activity,
                                guess-then-reveal, quiz, vocab recap, homework checklist,
                                stepper navigation) — fix a widget once, every lesson gets it
  lessons/
    lesson-01/
      guide.html           <- teaching script for the tutor (talking points, anticipated
                               questions, comprehension-check questions) — prep reading
      lesson.html            <- the interactive single-page tool you actually work through
                               live with the student (step-by-step, click-through activities)
      slides.pptx             <- portable static PowerPoint backup of the same content
      worksheet.html       <- practice questions + answer key (boxed off), quick-view
      worksheet.docx       <- same worksheet as an editable/printable Word doc
      flashcards.txt        <- paste into NotebookLM > Studio > Flashcards
    lesson-02/ ...
  scripts/                 <- regeneration tooling (see below)
```

The dashboard's "Up next" card and Full Sequence table link to `guide.html` (teaching script), `worksheet.html`/`.docx`, `lesson.html`/`slides.pptx` (labelled "Interactive lesson"), and `flashcards.txt`.

## What `lesson.html` actually is

Not a slideshow — a single-page interactive tool with step-by-step navigation (dots + prev/next + arrow keys), built from `assets/lesson-kit.css`/`.js`. Each lesson picks from a set of reusable widgets:

- **Flip cards** — tap a term to reveal its definition (recap warm-ups, rewards, vocab)
- **Sort activity** — classify items into categories with instant right/wrong feedback (e.g. factors of production)
- **Interactive reader** — buttons that swap live computed numbers (e.g. pick a scenario, see PED calculated; pick a time slot, see shortage/surplus)
- **Guess-then-reveal cards** — a question/scenario the tutor asks out loud, tapped open once the student has guessed
- **Shift-grid** — a grid of clickable cards that reveal a short explanation (determinants, "why it matters" categories)
- **Data table** — a computed comparison table, with an optional highlighted "best" row
- **Quiz** — multiple choice with instant right/wrong feedback per option
- **Vocab recap + homework checklist** — closing-step summary

Adding a new lesson's `lesson.html`: copy the shell (topbar/stepper markup + kit `<link>`/`<script>` tags) from any existing lesson, then write that lesson's own `<script>` block with its data arrays and `LessonKit.build*(...)` calls. Keep bespoke one-off widgets (a custom diagram etc.) inline in that lesson's script using the shared CSS classes for visual consistency.

## Asking Claude Code to plan a lesson

Say **"plan lesson 7"** (or whichever number) and Claude Code will generate that lesson's `guide.md`/`worksheet.md` (converted to `.html` + `.docx`), `lesson.html` (against the shared kit) + `slides.pptx`, and `flashcards.txt` in `lessons/lesson-07/`, all tied to that lesson's syllabus reference from `data/lessons.json`, then re-run the scripts below so the dashboard picks it up automatically.

## Regenerating things

- **`scripts/build_data.py`** — rebuilds `data/*.json`. Run after changing the holiday dates, the syllabus known/unknown flags, or the lesson sequence in the script itself. Each lesson's `files` flags (which formats exist — `guide_html`, `worksheet_html`, `worksheet_docx`, `lesson_html`, `slides_pptx`, `flashcards`) are auto-detected from disk, so you never need to hand-edit them.
- **`scripts/inject_data.py`** — rebuilds `dashboard.html` from `scripts/dashboard_template.html` + the current `data/*.json`. Run this after `build_data.py`, or after editing the template itself.
- **`scripts/convert_docs.py <lesson-folder>`** — converts a lesson's `guide.md`/`worksheet.md` into styled, printable `guide.html`/`worksheet.html` (the answer key is auto-boxed off). Has a few markdown formatting gotchas noted at the top of the file — read them before writing new lesson markdown.
- **`scripts/build_worksheet_docx.js`** (`node build_worksheet_docx.js`) — generates `worksheet.docx` for a lesson from a structured content object defined at the top of the file (Word docs don't map cleanly from arbitrary markdown/HTML, so worksheets are authored directly in this shape — add a new lesson by adding an entry).
- **`scripts/build_slides_pptx.js`** (`node build_slides_pptx.js`) — generates `slides.pptx` for a lesson using reusable slide-layout helpers (title slide, factor/icon slide, two-column, calc/reveal card, rule table, demand-supply diagram, closing slide) — add a new lesson by calling these helpers with that lesson's content.

Both `.js` scripts need `npm install` once inside `scripts/` (docx, pptxgenjs — not committed, see `.gitignore`).

Typical flow for a new lesson: write `guide.md` + `worksheet.md` + `lesson.html` + `flashcards.txt` into a new `lessons/lesson-NN/` folder → run `convert_docs.py` on that folder for the guide/worksheet → delete the now-redundant `.md` files → add that lesson's content to `build_worksheet_docx.js` and `build_slides_pptx.js` and run both → run `build_data.py` then `inject_data.py`.

## Adding a second student/subject later

Duplicate this whole `tom-economics/` folder as a sibling (e.g. `courses/name-subject/`), swap in a new syllabus PDF + calendar + weekly slots, and re-run the same scripts against the new data. `assets/lesson-kit.css`/`.js` can be copied as-is (it's course-agnostic). No shared database or shared code to untangle — each course is self-contained.
