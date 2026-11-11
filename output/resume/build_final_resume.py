"""Build Pradeep's final resume as ATS-safe HTML (-> PDF via chromium) and DOCX
from one content definition. Single column, standard headings, no tables,
no icons, no letter-spacing: everything a parser can read cleanly."""
import html, subprocess, pathlib
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

OUT = pathlib.Path(__file__).parent
NAVY = "14304f"

# ------------------------------------------------------------------ content
NAME = "Pradeep Penugonda"
TITLE = "Secondary Science Teacher (Years 7-10) | VCE Biology"
CONTACT = ("Carnegie VIC 3163  |  +61 452 668 984  |  penugonda56@outlook.com  |  penugonda.com.au")
STATUS = ("VIT 727895, registered (provisional; full-registration evidence lodged September 2026)  |  "
          "Australian Permanent Resident  |  Working with Children Check: current")

SUMMARY = (
    "Secondary Science teacher at St Francis Xavier College since 2023, currently teaching seven Years 7 to 9 "
    "Science classes across two campuses. Research scientist by training, with three master's degrees "
    "(Teaching; Biotechnology and Bioinformatics with High Distinction; Microbiology) and three peer-reviewed "
    "publications. Data-informed practitioner: a documented Year 7 cycle of inquiry lifted the class median from "
    "51% to 70% and moved students scoring 80% or above from 0 to 11 of 28. VCE Biology content-ready after three "
    "years tutoring Units 1 to 4. Seeking an ongoing Science position with a pathway into senior Biology and "
    "curriculum leadership."
)

TEACHING_AREAS = [
    "Science, Years 7 to 10: current, and the specialisation of my Master of Teaching.",
    "Biology, VCE Units 1 to 4: content expertise and three years of private tutoring; seeking a first classroom allocation.",
    "Chemistry, Years 7 to 10: grounded in laboratory research and industry quality-control experience.",
    "Open to additional junior allocations if useful to the timetable: Years 7 to 8 Mathematics and junior Humanities "
    "(not yet taught; willing to take on with support).",
]

SKILLS = [
    "Victorian Curriculum 2.0 (Science)", "VCE Biology Units 1 to 4", "Inquiry-based and practical science",
    "Explicit instruction and formative assessment", "Differentiation, PLP and ILP adjustments, inclusive practice",
    "Data-informed teaching (NAPLAN, PAT, diagnostic and summative data)", "Literacy and numeracy in Science",
    "EAL vocabulary scaffolding", "Laboratory safety and practical management",
    "Learning technologies, LMS, custom animation and digital resource design",
    "Assessment moderation, feedback and reporting",
    "Child safety: Child Safe Standards, Ministerial Order 1359, PROTECT, Mandatory Reporting",
    "Pastoral care and student wellbeing", "Collaborative and cross-campus curriculum planning",
    "Accreditation to Teach in a Catholic School (in progress)",
]

EXPERIENCE = [
    {
        "org": "St Francis Xavier College, Berwick and Officer campuses",
        "role": "Secondary Science Teacher, Years 7 to 9",
        "dates": "January 2023 to present",
        "bullets": [
            "Teach seven Science classes across Years 7, 8 and 9 (full-time load) on two campuses, planning against "
            "Victorian Curriculum 2.0 following STAVCON 2025 training.",
            "Ran a formal cycle of inquiry with a Year 7 class of 28 students (6 on Personalised Learning Plans): "
            "profiled the class with NAPLAN, PAT and diagnostic data, set success criteria in advance and measured "
            "growth with three independent instruments. Class median rose from 51% to 70%; students at 80% or above "
            "rose from 0 to 11; one extension student moved from 27% to 90%.",
            "Rebuilt lesson structure around explicit instruction and same-lesson feedback: visible learning "
            "intention, chunked instruction, structured practical, then immediate low-stakes assessment.",
            "Introduced dissection practicals at Year 7 (a first for the campus) and designed a pluck dissection "
            "(respiratory and digestive systems), shared across both campuses with full instructions. A Year 7 "
            "water-rocket investigation I designed was adopted by a colleague's STEM class.",
            "Authored and submitted proposals for two new 2027 electives: Biomedical Science, and Motion, Energy and Materials.",
            "Produce Standard and Personalised Learning Plan versions of assessments and practicals, used across both "
            "campuses; pre-teach subject vocabulary for EAL learners in a College where about a third of students "
            "speak a language other than English at home.",
            "Run a voluntary lunchtime Science support service, approved by the Head of Education and open to any "
            "student in the College; students from other teachers' classes seek me out and book sessions.",
            "Design custom animations and digital resources for every unit, and build educational apps and "
            "interactive web resources (penugonda.com.au).",
            "Proposed and ran an approved Year 9 Astronomy film study with classification check, parent notification "
            "and an opt-out alternative; send a written Science update to College leadership each term.",
            "Member of Bosco House; take my turn preparing and reading the Acknowledgement of Country and a prayer "
            "before the whole House.",
            "Completed the Bill Rogers behaviour management series and applied it in practice; acted on peer "
            "observation feedback on higher-order questioning, confirmed as consistent at follow-up observation.",
        ],
    },
    {
        "org": "Private tuition (self-employed)",
        "role": "VCE Biology Tutor, Units 1 to 4",
        "dates": "2023 to present",
        "bullets": [
            "One-to-one tuition covering course content, SAC preparation and exam technique; design of revision "
            "resources and practice assessments for individual students.",
        ],
    },
    {
        "org": "Teaching placements (55 days of supervised teaching practice)",
        "role": "",
        "dates": "2019 and 2022",
        "bullets": [
            "Mullauna College, 2022 (40 days), Year 9 Science: designed and delivered a four-week student-led research "
            "unit from scratch, including the course booklet, research-process guide, animations and a research-poster showcase.",
            "East Preston Islamic College, 2019 (15 days), Years 7 to 9 Science.",
        ],
    },
    {
        "org": "Eurofins, Melbourne",
        "role": "Laboratory Assistant (internship)",
        "dates": "2021",
        "bullets": [
            "pH meter and pipette calibration, media preparation, colony counting, quality-control growth-promotion "
            "testing and Preservative Efficacy Testing under Good Laboratory Practice.",
        ],
    },
]

EDUCATION = [
    ("Master of Teaching (Secondary Education)", "Victoria University, Melbourne", "2023",
     "Specialisation: Years 7 to 10 Science"),
    ("Master of Science (Biotechnology and Bioinformatics)", "La Trobe University, Melbourne", "2017",
     "Graduated with High Distinction"),
    ("Master of Science (Microbiology)", "Bangalore University, India", "2013", ""),
    ("Bachelor of Science (Microbiology)", "Osmania University, India", "2010", ""),
]

REGISTRATION = [
    "Victorian Institute of Teaching: registration 727895, provisional; full-registration evidence lodged September 2026.",
    "Accreditation to Teach in a Catholic School: in progress, 23 of 25 hours complete (Diocese of Sale).",
    "LANTITE: satisfied (embedded in the Victorian initial teacher education program).",
    "STAVCON 2025 (Victorian Curriculum 2.0 Science); Bill Rogers Behaviour Management series; Behavioural Leadership (2025).",
    "Child safety and wellbeing training, current: Mandatory Reporting, PROTECT, Child Safe Standards, NCCD, "
    "Disability Standards for Education, Anaphylaxis, CPR.",
]

PUBLICATIONS = [
    "Ganoderma applanatum (Persoon) Patouillard as a source of anticancer and antioxidant agent. "
    "International Journal of Pharma and Bio Sciences, 5(2): (B) 276-282, April 2014.",
    "Antioxidant and anticancer activity of Tricholoma giganteum Massee, an edible wild mushroom. "
    "Academic Journal of Cancer Research, January 2014.",
    "Evaluation of antimicrobial activity of some selected Basidiomycetous fungi. "
    "International Journal of Pharma and Bio Sciences, 4(4): (B) 964-971, October 2013.",
]

REFEREES = [
    ("Rosalie Harding", "Pedagogical Coach, St Peter's College (former mentor and supervisor, St Francis Xavier College)",
     "0434 578 465", "rharding@stpeters.vic.edu.au"),
    ("Melinda Davison", "Teacher, Doveton College (former Professional Practice Leader, St Francis Xavier College)",
     "0468 357 576", "melinda.davison@education.vic.gov.au"),
]

# ------------------------------------------------------------------ HTML / PDF
def e(s): return html.escape(s)

def build_html():
    h = []
    h.append(f"""<html><head><meta charset="utf-8"><title>{e(NAME)} - Resume</title><style>
@page {{ size: A4; margin: 13mm 16mm 12mm 16mm; }}
body {{ font-family: Arial, Helvetica, sans-serif; font-size: 9.9pt; line-height: 1.32; color:#1a1a1a; margin:0; }}
h1 {{ font-size: 19pt; margin:0; color:#{NAVY}; }}
.title {{ font-size: 11pt; font-weight:700; color:#{NAVY}; margin:1px 0 3px; }}
.contact, .status {{ font-size: 8.8pt; color:#333; margin:0; }}
h2 {{ font-size: 10.3pt; color:#{NAVY}; border-bottom: 1.2px solid #{NAVY}; margin: 9px 0 4px; padding-bottom:1px; text-transform: uppercase; letter-spacing: 0; }}
p {{ margin: 0 0 4px; }}
ul {{ margin: 2px 0 4px 16px; padding:0; }}
li {{ margin: 0 0 2.5px; }}
.job {{ margin-top: 5px; }}
.job b {{ color:#{NAVY}; }}
.dates {{ float:right; color:#333; }}
.skills {{ margin:0; }}
</style></head><body>""")
    h.append(f"<h1>{e(NAME)}</h1><p class='title'>{e(TITLE)}</p>")
    h.append(f"<p class='contact'>{e(CONTACT)}</p><p class='status'>{e(STATUS)}</p>")

    h.append("<h2>Professional Summary</h2>")
    h.append(f"<p>{e(SUMMARY)}</p>")

    h.append("<h2>Teaching Areas and Availability</h2><ul>")
    h += [f"<li>{e(t)}</li>" for t in TEACHING_AREAS]
    h.append("</ul>")

    h.append("<h2>Key Skills</h2>")
    h.append("<p class='skills'>" + e("  |  ".join(SKILLS)) + "</p>")

    h.append("<h2>Teaching Experience</h2>")
    for j in EXPERIENCE:
        role = f"<br>{e(j['role'])}" if j["role"] else ""
        h.append(f"<p class='job'><span class='dates'>{e(j['dates'])}</span><b>{e(j['org'])}</b>{role}</p><ul>")
        h += [f"<li>{e(b)}</li>" for b in j["bullets"]]
        h.append("</ul>")

    h.append("<h2>Education</h2><ul>")
    for deg, inst, yr, note in EDUCATION:
        extra = f". {e(note)}" if note else ""
        h.append(f"<li><b>{e(deg)}</b>, {e(inst)}, {e(yr)}{extra}</li>")
    h.append("</ul>")

    h.append("<h2>Registration, Accreditation and Professional Learning</h2><ul>")
    h += [f"<li>{e(r)}</li>" for r in REGISTRATION]
    h.append("</ul>")

    h.append("<h2>Publications</h2><ol style='margin:2px 0 4px 18px;padding:0'>")
    h += [f"<li>{e(p)}</li>" for p in PUBLICATIONS]
    h.append("</ol>")

    h.append("<h2>Referees</h2><ul>")
    for n, r, ph, em in REFEREES:
        h.append(f"<li><b>{e(n)}</b>, {e(r)}. {e(ph)}, {e(em)}</li>")
    h.append("</ul></body></html>")
    return "\n".join(h)

# ------------------------------------------------------------------ DOCX
def build_docx(path):
    d = Document()
    for s in d.sections:
        s.top_margin = s.bottom_margin = Cm(1.3); s.left_margin = s.right_margin = Cm(1.6)
    st = d.styles["Normal"]; st.font.name = "Arial"; st.font.size = Pt(10)
    navy = RGBColor(0x14, 0x30, 0x4f)

    def para(text, size=10, bold=False, color=None, space_after=2, align=None):
        p = d.add_paragraph(); r = p.add_run(text); r.font.size = Pt(size); r.bold = bold
        if color: r.font.color.rgb = color
        p.paragraph_format.space_after = Pt(space_after); p.paragraph_format.space_before = Pt(0)
        if align: p.alignment = align
        return p

    def heading(text):
        p = para(text.upper(), size=10.5, bold=True, color=navy, space_after=2)
        p.paragraph_format.space_before = Pt(8)
        # bottom border
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        pPr = p._p.get_or_add_pPr(); pbdr = OxmlElement("w:pBdr"); b = OxmlElement("w:bottom")
        for k, v in (("w:val", "single"), ("w:sz", "6"), ("w:space", "1"), ("w:color", NAVY)): b.set(qn(k), v)
        pbdr.append(b); pPr.append(pbdr)

    def bullets(items):
        for it in items:
            p = d.add_paragraph(it, style="List Bullet"); p.paragraph_format.space_after = Pt(1.5)

    para(NAME, size=19, bold=True, color=navy, space_after=0)
    para(TITLE, size=11, bold=True, color=navy, space_after=2)
    para(CONTACT, size=9, space_after=0); para(STATUS, size=9, space_after=2)

    heading("Professional Summary"); para(SUMMARY, space_after=2)
    heading("Teaching Areas and Availability"); bullets(TEACHING_AREAS)
    heading("Key Skills"); para("  |  ".join(SKILLS), space_after=2)
    heading("Teaching Experience")
    for j in EXPERIENCE:
        p = d.add_paragraph(); p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(1)
        r = p.add_run(j["org"]); r.bold = True; r.font.color.rgb = navy
        if j["role"]: p.add_run(f"  |  {j['role']}")
        p.add_run(f"    {j['dates']}").italic = True
        bullets(j["bullets"])
    heading("Education")
    for deg, inst, yr, note in EDUCATION:
        p = d.add_paragraph(style="List Bullet"); p.paragraph_format.space_after = Pt(1.5)
        p.add_run(deg).bold = True; p.add_run(f", {inst}, {yr}" + (f". {note}" if note else ""))
    heading("Registration, Accreditation and Professional Learning"); bullets(REGISTRATION)
    heading("Publications")
    for pub in PUBLICATIONS:
        p = d.add_paragraph(pub, style="List Number"); p.paragraph_format.space_after = Pt(1.5)
    heading("Referees")
    for n, r_, ph, em in REFEREES:
        p = d.add_paragraph(style="List Bullet"); p.paragraph_format.space_after = Pt(1.5)
        p.add_run(n).bold = True; p.add_run(f", {r_}. {ph}, {em}")
    d.save(path)

if __name__ == "__main__":
    base = OUT / "Pradeep_Penugonda_Resume_Science_Teacher"
    (base.with_suffix(".html")).write_text(build_html(), encoding="utf-8")
    subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu",
                    f"--print-to-pdf={base.with_suffix('.pdf')}", "--no-pdf-header-footer",
                    f"file://{base.with_suffix('.html')}"], capture_output=True)
    build_docx(base.with_suffix(".docx"))
    print("built:", base.with_suffix(".pdf").name, base.with_suffix(".docx").name)
