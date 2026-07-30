"""Regenerates data/calendar.json, data/syllabus.json and data/lessons.json.

Run this after changing the holiday ranges, the syllabus known/unknown flags,
or the lesson unit sequence below. `materials_generated` is auto-detected by
checking whether lessons/lesson-NN/guide.html exists on disk, so it never
needs to be edited by hand.

Usage: python3 build_data.py
"""
import datetime, json, os

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)  # courses/tom-economics
OUT = os.path.join(BASE, "data")

# ---------- calendar.json ----------
# Source: Selong Bay School Academic Calendar 26/27 (official PDF).
holiday_ranges = [
    ("2026-07-01", "2026-07-05", "Term break"),
    ("2026-07-20", "2026-07-26", "Term break"),
    ("2026-10-05", "2026-10-09", "Mid-term break"),
    ("2026-12-14", "2026-12-31", "Winter break"),
    ("2027-01-04", "2027-01-10", "Winter break"),
    ("2027-03-08", "2027-03-21", "Spring break"),
    ("2027-06-21", "2027-06-30", "End of year"),
]
public_holidays = [
    ("2026-08-17", "Public holiday"),
    ("2026-08-25", "Public holiday"),
    ("2027-01-01", "New Year's Day"),
    ("2027-03-26", "Public holiday"),
    ("2027-05-01", "Labour Day"),
    ("2027-05-20", "Public holiday"),
    ("2027-06-01", "Public holiday"),
    ("2027-06-06", "Public holiday"),
]

calendar = {
    "academic_year": "26/27",
    "school": "Selong Bay School",
    "school_holidays": [{"start": s, "end": e, "label": l} for s, e, l in holiday_ranges],
    "public_holidays": [{"date": d, "label": l} for d, l in public_holidays],
}
with open(os.path.join(OUT, "calendar.json"), "w") as f:
    json.dump(calendar, f, indent=2)

# ---------- syllabus.json ----------
# "known": True means confirmed already taught, from the Term 1 Week 1 recap worksheet.
syllabus = {
  "code": "0455",
  "title": "Cambridge IGCSE Economics",
  "exam_series": "June 2027",
  "topics": [
    {"id":"1","title":"The basic economic problem","subtopics":[
      {"id":"1.1","title":"The nature of the basic economic problem","known":True},
      {"id":"1.2","title":"Factors of production","known":False},
      {"id":"1.3","title":"Opportunity cost","known":True},
      {"id":"1.4","title":"Production possibility curve (PPC) diagrams","known":True},
    ]},
    {"id":"2","title":"The allocation of resources","subtopics":[
      {"id":"2.1","title":"The role of markets in allocating resources","known":True},
      {"id":"2.2","title":"Demand","known":True},
      {"id":"2.3","title":"Supply","known":True},
      {"id":"2.4","title":"Price determination","known":True,"note":"2.4.1-2.4.2 known; 2.4.3 disequilibrium covered in Lesson 2"},
      {"id":"2.5","title":"Price changes","known":False},
      {"id":"2.6","title":"Price elasticity of demand (PED)","known":False},
      {"id":"2.7","title":"Price elasticity of supply (PES)","known":False},
      {"id":"2.8","title":"Market economic system","known":False},
      {"id":"2.9","title":"Market failure","known":False},
      {"id":"2.10","title":"Mixed economic system","known":False},
    ]},
    {"id":"3","title":"Microeconomic decision-makers","subtopics":[
      {"id":"3.1","title":"Money and banking","known":False},
      {"id":"3.2","title":"Households","known":False},
      {"id":"3.3","title":"Workers","known":False},
      {"id":"3.4","title":"Firms","known":False},
      {"id":"3.5","title":"Firms and production","known":False},
      {"id":"3.6","title":"Firms' costs, revenue and objectives","known":False},
      {"id":"3.7","title":"Types of markets","known":False},
    ]},
    {"id":"4","title":"Government and the macroeconomy","subtopics":[
      {"id":"4.1","title":"Government macroeconomic intervention","known":False},
      {"id":"4.2","title":"Fiscal policy","known":False},
      {"id":"4.3","title":"Monetary policy","known":False},
      {"id":"4.4","title":"Supply-side policy","known":False},
      {"id":"4.5","title":"Economic growth","known":False},
      {"id":"4.6","title":"Employment and unemployment","known":False},
      {"id":"4.7","title":"Inflation","known":False},
    ]},
    {"id":"5","title":"Economic development","subtopics":[
      {"id":"5.1","title":"Living standards","known":False},
      {"id":"5.2","title":"Poverty","known":False},
      {"id":"5.3","title":"Population","known":False},
      {"id":"5.4","title":"Differences in economic development between countries","known":False},
    ]},
    {"id":"6","title":"International trade and globalisation","subtopics":[
      {"id":"6.1","title":"Specialisation and free trade","known":False},
      {"id":"6.2","title":"Globalisation and trade restrictions","known":False},
      {"id":"6.3","title":"Foreign exchange rates","known":False},
      {"id":"6.4","title":"Current account of the balance of payments","known":False},
    ]},
  ]
}
with open(os.path.join(OUT, "syllabus.json"), "w") as f:
    json.dump(syllabus, f, indent=2)

# ---------- lessons.json ----------
def daterange(start, end):
    d = datetime.date.fromisoformat(start); e = datetime.date.fromisoformat(end)
    while d <= e:
        yield d; d += datetime.timedelta(days=1)

non_school = set()
for s, e, _ in holiday_ranges:
    for d in daterange(s, e): non_school.add(d)
for d, _ in public_holidays: non_school.add(datetime.date.fromisoformat(d))

# First lesson slot: the Thursday this course was set up (2026-07-29 was a Wednesday).
today = datetime.date(2026, 7, 29)
end_of_year = datetime.date(2027, 6, 30)
lesson_dates = []
d = today
while d <= end_of_year:
    if d.weekday() in (1, 3) and d not in non_school and d >= today:  # Tue=1, Thu=3
        lesson_dates.append(d)
    d += datetime.timedelta(days=1)

# Teaching units in pedagogical order. Already-known subtopics (per syllabus.json) are skipped.
units = [
 ("content","1.2","Factors of production"),
 ("content","2.4.3 / 2.5","Market disequilibrium & price changes"),
 ("content","2.6","PED: definition, calculation, determinants"),
 ("content","2.6","PED: revenue relationship & significance"),
 ("content","2.7","Price elasticity of supply (PES)"),
 ("content","2.8","Market economic system"),
 ("content","2.9","Market failure: key definitions"),
 ("content","2.9","Market failure: causes & consequences"),
 ("content","2.10","Mixed economy & price controls/tax/subsidies"),
 ("content","2.10","Other government intervention (regulation, privatisation, nationalisation, quotas)"),
 ("review","1-2","Review & consolidation: Topics 1-2"),
 ("content","3.1","Money and banking"),
 ("content","3.2 / 3.3","Households & choice of occupation"),
 ("content","3.3","Wage determination & NMW diagrams"),
 ("content","3.3","Wage differences, mobility, division of labour"),
 ("content","3.4","Types of firms & mergers"),
 ("content","3.4","Economies & diseconomies of scale"),
 ("content","3.5","Firms and production"),
 ("content","3.6","Costs of production (calculations)"),
 ("content","3.6","Revenue & firms' objectives"),
 ("content","3.7","Types of markets: competitive vs monopoly"),
 ("review","3","Review & consolidation: Topic 3"),
 ("content","4.1","Macroeconomic aims & conflicts"),
 ("content","4.2","Government budget & taxation"),
 ("content","4.2","Fiscal policy measures & effects"),
 ("content","4.3","Monetary policy"),
 ("content","4.4","Supply-side policy"),
 ("content","4.5","Economic growth"),
 ("content","4.5","Recession & growth policies"),
 ("content","4.6","Unemployment: definitions, measurement, types"),
 ("content","4.6","Unemployment: consequences & policies"),
 ("content","4.7","Inflation: definitions, measurement, causes"),
 ("content","4.7","Inflation: consequences & policies"),
 ("review","4","Review & consolidation: Topic 4"),
 ("content","5.1","Living standards (GDP/head, HDI)"),
 ("content","5.2","Poverty"),
 ("content","5.3","Population"),
 ("content","5.4","International development differences"),
 ("review","5","Review & consolidation: Topic 5"),
 ("content","6.1","Specialisation & free trade"),
 ("content","6.2","Globalisation & multinational companies"),
 ("content","6.2","Trade restrictions"),
 ("content","6.3","Foreign exchange rates: basics & reasons to trade currency"),
 ("content","6.3","Exchange rate determination & consequences"),
 ("content","6.4","Balance of payments: structure & causes"),
 ("content","6.4","Balance of payments: consequences & policies"),
 ("review","6","Review & consolidation: Topic 6 + full syllabus map"),
 ("revision","1","Revision blitz: Topic 1"),
 ("revision","2","Revision blitz: Topic 2"),
 ("revision","3","Revision blitz: Topic 3"),
 ("revision","4","Revision blitz: Topic 4"),
 ("revision","5","Revision blitz: Topic 5"),
 ("revision","6","Revision blitz: Topic 6"),
 ("exam-skill","P1","Paper 1 command words & MCQ technique + timed set"),
 ("exam-skill","P1","Paper 1 timed practice set 2 + review"),
 ("exam-skill","P2-A","Paper 2 Section A (data response) technique"),
 ("exam-skill","P2-B","Paper 2 Section B part (d) evaluation technique (Level 3 answers)"),
 ("past-paper","PP1","Past paper 1: sit (P1+P2 extract)"),
 ("past-paper","PP1","Past paper 1: review & feedback"),
 ("past-paper","PP2","Past paper 2: sit"),
 ("past-paper","PP2","Past paper 2: review & feedback"),
 ("past-paper","PP3","Past paper 3: sit"),
 ("past-paper","PP3","Past paper 3: review & feedback"),
 ("past-paper","PP4","Past paper 4: sit"),
 ("past-paper","PP4","Past paper 4: review & feedback"),
 ("past-paper","PP5","Past paper 5: sit"),
 ("past-paper","PP5","Past paper 5: review & feedback"),
 ("buffer","-","Flex/re-teach slot (weakest topic, reactive)"),
 ("buffer","-","Flex/re-teach slot (weakest topic, reactive)"),
 ("buffer","-","Flex/re-teach slot (weakest topic, reactive)"),
 ("buffer","-","Final mixed past-paper drill"),
 ("buffer","-","Final mixed past-paper drill"),
 ("buffer","-","Confidence-building light review"),
 ("buffer","-","Confidence-building light review + exam-day logistics chat"),
]

def has_file(folder_name, filename):
    return os.path.exists(os.path.join(BASE, "lessons", folder_name, filename))

lessons = []
for i, ((phase, ref, title), date) in enumerate(zip(units, lesson_dates), 1):
    folder_name = f"lesson-{i:02d}"
    files = {
        "guide_html": has_file(folder_name, "guide.html"),
        "worksheet_html": has_file(folder_name, "worksheet.html"),
        "worksheet_docx": has_file(folder_name, "worksheet.docx"),
        "slides_html": has_file(folder_name, "slides.html"),
        "slides_pptx": has_file(folder_name, "slides.pptx"),
        "flashcards": has_file(folder_name, "flashcards.txt"),
    }
    lessons.append({
        "lesson": i,
        "date": date.isoformat(),
        "weekday": date.strftime("%A"),
        "phase": phase,
        "syllabus_ref": ref,
        "title": title,
        "materials_generated": files["guide_html"],
        "files": files,
        "folder": f"lessons/{folder_name}",
    })

with open(os.path.join(OUT, "lessons.json"), "w") as f:
    json.dump({"lessons": lessons, "total_slots_available": len(lesson_dates)}, f, indent=2)

print("Wrote calendar.json, syllabus.json, lessons.json to", OUT)
print("Total lessons:", len(lessons), "| with materials generated:", sum(l["materials_generated"] for l in lessons))
