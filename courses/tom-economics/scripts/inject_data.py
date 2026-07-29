"""Rebuilds ../dashboard.html from dashboard_template.html + the data/*.json files.

Run this any time data/*.json changes (e.g. after build_data.py) so the
dashboard's embedded data stays in sync.

Usage: python3 inject_data.py
"""
import os

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)  # courses/tom-economics
TEMPLATE = os.path.join(HERE, "dashboard_template.html")
OUT = os.path.join(BASE, "dashboard.html")

with open(os.path.join(BASE, "data", "calendar.json")) as f:
    calendar_json = f.read()
with open(os.path.join(BASE, "data", "syllabus.json")) as f:
    syllabus_json = f.read()
with open(os.path.join(BASE, "data", "lessons.json")) as f:
    lessons_json = f.read()

with open(TEMPLATE) as f:
    html = f.read()

html = html.replace("__CALENDAR_JSON__", calendar_json)
html = html.replace("__SYLLABUS_JSON__", syllabus_json)
html = html.replace("__LESSONS_JSON__", lessons_json)

with open(OUT, "w") as f:
    f.write(html)

print("Wrote", OUT, "-", len(html), "bytes")
