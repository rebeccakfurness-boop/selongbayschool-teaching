"""Converts a lesson's guide.md / worksheet.md into styled, printable guide.html / worksheet.html.

Markdown gotchas this pipeline needs (python-markdown, not GFM):
  - Always leave a blank line before a list (numbered or bulleted) or it won't parse as a list.
  - Always leave a blank line before a table.
  - Don't use two consecutive `---` lines as a divider — a lone `---` right after a text line
    gets read as a Setext heading underline, not a horizontal rule. Just use a single blank
    line before the "## Answer key" heading; the tutor-box styling below provides the visual break.
  - The answer-key section (anything from "## Answer key" onwards) is auto-wrapped in a
    highlighted "tutor only" box - keep that exact heading text.

Usage: python3 convert_docs.py ../lessons/lesson-03
(converts guide.md -> guide.html and worksheet.md -> worksheet.html inside that folder,
 for whichever of the two source files exist)
"""
import markdown, re, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "doc_template.html")
with open(TEMPLATE) as f:
    TMPL = f.read()

def convert(md_path, html_path, title):
    with open(md_path) as f:
        src = f.read()
    body = markdown.markdown(src, extensions=['tables', 'nl2br'])
    marker = '<h2>Answer key'
    idx = body.find(marker)
    if idx != -1:
        head = body[:idx]
        tail = body[idx:]
        head = re.sub(r'(<hr\s*/?>\s*){1,2}$', '', head)
        body = head + '<div class="tutor-box">' + tail + '</div>'
    out = TMPL.replace('__TITLE__', title).replace('__BODY__', body)
    with open(html_path, 'w') as f:
        f.write(out)
    print('wrote', html_path)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python3 convert_docs.py <lesson-folder>")
        sys.exit(1)
    folder = os.path.abspath(sys.argv[1])
    lesson_name = os.path.basename(folder).replace('-', ' ').title()
    guide_md = os.path.join(folder, "guide.md")
    worksheet_md = os.path.join(folder, "worksheet.md")
    if os.path.exists(guide_md):
        convert(guide_md, os.path.join(folder, "guide.html"), f"{lesson_name} Guide")
    if os.path.exists(worksheet_md):
        convert(worksheet_md, os.path.join(folder, "worksheet.html"), f"{lesson_name} Worksheet")
