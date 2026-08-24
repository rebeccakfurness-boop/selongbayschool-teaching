"""Builds a self-contained, standalone HTML file from a lesson's lesson.html, by inlining
assets/lesson-kit.css and assets/lesson-kit.js (which lesson.html normally references via
relative <link>/<script src> paths that only resolve when the lesson folder sits inside the
full course folder structure). The output has no external file dependencies (aside from the
Google Fonts stylesheet already linked in lesson-kit.css), so it can be published as a Claude
Artifact or opened directly with no unzip/folder-structure requirements.

Usage: python3 build_artifact.py lesson-05 [output_path]
(defaults output_path to <scratch>/lesson05_artifact.html next to this script if omitted)
"""
import re, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
COURSE = os.path.dirname(HERE)

def _harden_theme_css(css):
    """The source lesson-kit.css only guards dark tokens behind a bare
    @media (prefers-color-scheme) query, which is correct for the plain
    file:// case but not for the Artifact viewer, which can also stamp an
    explicit data-theme="light"/"dark" on <html>. Rewrite the dark-token
    block so an explicit choice always wins over the OS preference."""
    m = re.search(r"@media \(prefers-color-scheme: dark\) \{\s*:root \{(.*?)\}\s*\}", css, re.DOTALL)
    if not m:
        return css
    tokens = m.group(1)
    replacement = (
        '@media (prefers-color-scheme: dark) {\n'
        '  :root:not([data-theme="light"]) {' + tokens + '}\n'
        '}\n'
        ':root[data-theme="dark"] {' + tokens + '}'
    )
    return css[:m.start()] + replacement + css[m.end():]

def build(lesson_folder, out_path, title=None):
    lesson_dir = os.path.join(COURSE, "lessons", lesson_folder)
    with open(os.path.join(lesson_dir, "lesson.html")) as f:
        html = f.read()
    with open(os.path.join(COURSE, "assets", "lesson-kit.css")) as f:
        css = f.read()
    css = _harden_theme_css(css)
    css = re.sub(r"@import url\('https://fonts\.googleapis\.com[^']*'\);\n*", "", css)
    with open(os.path.join(COURSE, "assets", "lesson-kit.js")) as f:
        js = f.read()

    # Pull out the page <title>
    m = re.search(r"<title>(.*?)</title>", html)
    page_title = title or (m.group(1) if m else lesson_folder)

    # Pull out everything inside <body>...</body>
    body_m = re.search(r"<body>(.*)</body>", html, re.DOTALL)
    body = body_m.group(1)

    # Drop the two external tags — CSS and lesson-kit.js get inlined instead
    body = body.replace('<link rel="stylesheet" href="../../assets/lesson-kit.css">', '')
    body = body.replace('<script src="../../assets/lesson-kit.js"></script>', f'<script>\n{js}\n</script>')

    out = f"""<title>{page_title}</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap">
<style>
{css}
/* ===== Wide-content safety + reduced-motion (artifact-hardening additions) ===== */
.data-table, table.cmp {{display:block; overflow-x:auto; white-space:nowrap;}}
@media (prefers-reduced-motion: reduce) {{
  .step{{animation:none;}}
  .btn, .dot, .flip-inner, .bar-fill{{transition:none;}}
}}
</style>
{body}
"""
    with open(out_path, "w") as f:
        f.write(out)
    print("wrote", out_path)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 build_artifact.py <lesson-folder> [output_path]")
        sys.exit(1)
    lesson_folder = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, f"{lesson_folder}_artifact.html")
    build(lesson_folder, out_path)
