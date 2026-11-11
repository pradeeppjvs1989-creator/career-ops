"""Build Pradeep's final resume from one content source.

  python3 build_final_resume.py            -> science (default) + vce variants
  python3 build_final_resume.py science    -> Years 7-10 Science-first version
  python3 build_final_resume.py vce        -> VCE-Biology-forward version
  flags: --mixed    adds the Maths / Humanities open-to line (mixed-load ads only)
         --catholic adds the House prayer bullet (Catholic schools only)

ATS-safe: single column, standard headings, no tables, no icons. Each variant
is written as HTML -> PDF (chromium) and DOCX.
"""
import html, subprocess, pathlib, sys
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = pathlib.Path(__file__).parent
NAVY = "14304f"
CHROME = "/opt/pw-browsers/chromium"

# ------------------------------------------------------------------ content
NAME = "Pradeep Penugonda"
CONTACT = "Carnegie VIC 3163  |  +61 452 668 984  |  penugonda56@outlook.com  |  penugonda.com.au"
STATUS = "VIT 727895 (provisional registration)  |  Australian Permanent Resident  |  Working with Children Check: current"

PROFILE_SCIENCE = (
    "Full-time Science teacher at St Francis Xavier College since 2023, currently teaching seven Years 7 to 9 "
    "classes across the Berwick and Officer campuses. Develops practical investigations and differentiated "
    "assessments used across both campuses, with documented student growth. Seeking an ongoing Years 7 to 10 "
    "Science role; Year 10 and senior Biology allocations welcome."
)
PROFILE_VCE = (
    "Science teacher with a postgraduate biology research background: Master of Science in Biotechnology and "
    "Bioinformatics (High Distinction), Master of Science in Microbiology and three peer-reviewed publications. "
    "Full-time at St Francis Xavier College since 2023, teaching seven Years 7 to 9 classes across two campuses, "
    "with three years of VCE Biology tutoring at Units 1 to 4. Seeking an ongoing Science role that includes a "
    "first school-based VCE Biology allocation."
)
VCE_LINE = ("VCE Biology tutoring, Units 1 to 4: three years supporting students with course content, "
            "SAC preparation and exam technique.")

AREAS_SCIENCE = "Science, Years 7 to 10  |  Biology to VCE Units 1 to 4  |  STEM, Years 7 to 10  |  Chemistry, Years 7 to 10"
AREAS_VCE = "Biology, Years 10 to 12 (VCE Units 1 to 4)  |  Science, Years 7 to 10  |  STEM, Years 7 to 10  |  Chemistry, Years 7 to 10"
AREAS_MIXED = "  |  Open to Years 7 to 8 Mathematics and junior Humanities (not yet taught)"

SFX_BULLETS = [
    "Teach seven Science classes across Years 7, 8 and 9 (full-time load) on two campuses, with units planned "
    "to Victorian Curriculum 2.0 following STAVCON 2025 training.",
    "Ran a Year 7 cycle of inquiry (class of 28, six on Personalised Learning Plans): the class median on the "
    "Separation Techniques unit test rose from 51% at diagnostic to 70%, and students scoring 80% or above rose "
    "from 0 to 11 of 28, measured against three instruments including a formal lab report on the shared faculty rubric.",
    "Build Standard and Personalised Learning Plan versions of assessments and practicals, used across both "
    "campuses, and pre-teach subject vocabulary for EAL learners.",
    "Introduced dissection practicals at Year 7, a first for the campus, and designed a pluck (respiratory and "
    "digestive) dissection shared across both campuses with full instructions; a Year 7 water-rocket investigation "
    "I designed was adopted by a colleague's STEM class.",
    "Authored and submitted proposals for two 2027 electives: Biomedical Science, and Motion, Energy and Materials.",
    "Run a voluntary lunchtime Science support service, approved by the Head of Education and open to any student "
    "in the College; students from other teachers' classes book sessions. Design custom animations and digital "
    "resources for every unit.",
]
CATHOLIC_BULLET = ("Member of Bosco House; take my turn preparing and reading the Acknowledgement of Country and "
                   "a prayer before the whole House, and attend College Masses and liturgies.")

OTHER_EXPERIENCE = [
    ("Private VCE Biology Tutor (self-employed)", "2023 to present",
     "One-to-one tuition in Units 1 to 4: course content, SAC preparation, exam technique and revision resources."),
    ("Teaching placements, 55 days supervised practice", "2019 and 2022",
     "Mullauna College, 2022 (40 days), Year 9 Science: designed and delivered a four-week student-led research unit. "
     "East Preston Islamic College, 2019 (15 days), Years 7 to 9 Science."),
    ("Eurofins, Melbourne: Laboratory Assistant (internship)", "2021",
     "Microbiology quality control under Good Laboratory Practice."),
]

EDUCATION = [
    ("Master of Teaching (Secondary Education)", "Victoria University, Melbourne, 2023. Specialisation: Years 7 to 10 Science"),
    ("Master of Science (Biotechnology and Bioinformatics)", "La Trobe University, Melbourne, 2017. High Distinction"),
    ("Master of Science (Microbiology)", "Bangalore University, India, 2013"),
    ("Bachelor of Science (Microbiology)", "Osmania University, India, 2010"),
]

REGISTRATION = [
    "Victorian Institute of Teaching: registration 727895, provisional. Full-registration evidence complete; "
    "recommendation panel at the end of Term 4 2026.",
    "Accreditation to Teach in a Catholic School: in progress, 23 of 25 hours complete (Diocese of Sale).",
    "LANTITE satisfied (embedded in the Victorian initial teacher education program).",
    "Professional learning: STAVCON 2025 (Victorian Curriculum 2.0 Science); Bill Rogers Behaviour Management "
    "series, applied in practice; peer observation cycle on higher-order questioning.",
    "Child safety and wellbeing training, current: Mandatory Reporting, PROTECT, Child Safe Standards, NCCD, "
    "Disability Standards for Education, Anaphylaxis, CPR.",
]

PUBLICATIONS = ("Three peer-reviewed publications in microbiology and mycology (International Journal of Pharma "
                "and Bio Sciences; Academic Journal of Cancer Research), 2013 to 2014. Full citations on request.")

REFEREES = [
    ("Rosalie Harding", "Pedagogical Coach, St Peter's College (former mentor and supervisor, St Francis Xavier College)",
     "0434 578 465", "rharding@stpeters.vic.edu.au"),
    ("Melinda Davison", "Teacher, Doveton College (former Professional Practice Leader, St Francis Xavier College)",
     "0468 357 576", "melinda.davison@education.vic.gov.au"),
]

def assemble(variant, mixed=False, catholic=False):
    vce = variant == "vce"
    return {
        "headline": "Secondary Science Teacher | VCE Biology" if vce else "Secondary Science Teacher",
        "profile": PROFILE_VCE if vce else PROFILE_SCIENCE,
        "vce_line": None if vce else VCE_LINE,
        "areas": (AREAS_VCE if vce else AREAS_SCIENCE) + (AREAS_MIXED if mixed else ""),
        "sfx": SFX_BULLETS + ([CATHOLIC_BULLET] if catholic else []),
    }

# ------------------------------------------------------------------ HTML
def e(s): return html.escape(s)

def build_html(c):
    h = [f"""<html><head><meta charset="utf-8"><title>{e(NAME)} - Resume</title><style>
@page {{ size: A4; margin: 14mm 17mm 13mm 17mm; }}
body {{ font-family: Arial, Helvetica, sans-serif; font-size: 10pt; line-height: 1.33; color:#1a1a1a; margin:0; }}
h1 {{ font-size: 19pt; margin:0; color:#{NAVY}; }}
.headline {{ font-size: 11.5pt; font-weight:700; color:#{NAVY}; margin:1px 0 4px; }}
.small {{ font-size: 8.9pt; color:#333; margin:0; }}
h2 {{ font-size: 10.4pt; color:#{NAVY}; border-bottom: 1.2px solid #{NAVY}; margin: 10px 0 4px; padding-bottom:1px; text-transform: uppercase; }}
p {{ margin: 0 0 4px; }}
ul {{ margin: 2px 0 4px 16px; padding:0; }}  li {{ margin: 0 0 3px; }}
.job {{ margin-top: 5px; }}  .job b {{ color:#{NAVY}; }}  .dates {{ float:right; color:#333; }}
</style></head><body>"""]
    h.append(f"<h1>{e(NAME)}</h1><p class='headline'>{e(c['headline'])}</p>")
    h.append(f"<p class='small'>{e(CONTACT)}</p><p class='small'>{e(STATUS)}</p>")
    h.append("<h2>Profile</h2>")
    h.append(f"<p>{e(c['profile'])}</p>")
    if c["vce_line"]: h.append(f"<p>{e(c['vce_line'])}</p>")
    h.append(f"<p><b>Teaching areas:</b> {e(c['areas'])}</p>")

    h.append("<h2>Teaching Experience</h2>")
    h.append("<p class='job'><span class='dates'>January 2023 to present</span>"
             "<b>St Francis Xavier College, Berwick and Officer campuses</b><br>Secondary Science Teacher, Years 7 to 9</p><ul>")
    h += [f"<li>{e(b)}</li>" for b in c["sfx"]]
    h.append("</ul>")
    for org, dates, line in OTHER_EXPERIENCE:
        h.append(f"<p class='job'><span class='dates'>{e(dates)}</span><b>{e(org)}</b></p><ul><li>{e(line)}</li></ul>")

    h.append("<h2>Education</h2><ul>")
    h += [f"<li><b>{e(d)}</b>, {e(rest)}</li>" for d, rest in EDUCATION]
    h.append("</ul>")

    h.append("<h2>Registration and Professional Learning</h2><ul>")
    h += [f"<li>{e(r)}</li>" for r in REGISTRATION]
    h.append("</ul>")

    h.append(f"<h2>Publications</h2><p>{e(PUBLICATIONS)}</p>")

    h.append("<h2>Referees</h2><ul>")
    h += [f"<li><b>{e(n)}</b>, {e(r)}. {e(ph)}, {e(em)}</li>" for n, r, ph, em in REFEREES]
    h.append("</ul></body></html>")
    return "\n".join(h)

# ------------------------------------------------------------------ DOCX
def build_docx(c, path):
    d = Document()
    for s in d.sections:
        s.top_margin = s.bottom_margin = Cm(1.4); s.left_margin = s.right_margin = Cm(1.7)
    d.styles["Normal"].font.name = "Arial"; d.styles["Normal"].font.size = Pt(10)
    navy = RGBColor(0x14, 0x30, 0x4f)

    def para(text, size=10, bold=False, color=None, after=2, before=0):
        p = d.add_paragraph(); r = p.add_run(text); r.font.size = Pt(size); r.bold = bold
        if color: r.font.color.rgb = color
        p.paragraph_format.space_after = Pt(after); p.paragraph_format.space_before = Pt(before)
        return p

    def heading(text):
        p = para(text.upper(), size=10.5, bold=True, color=navy, after=2, before=8)
        pPr = p._p.get_or_add_pPr(); pbdr = OxmlElement("w:pBdr"); b = OxmlElement("w:bottom")
        for k, v in (("w:val", "single"), ("w:sz", "6"), ("w:space", "1"), ("w:color", NAVY)): b.set(qn(k), v)
        pbdr.append(b); pPr.append(pbdr)

    def bullets(items):
        for it in items:
            p = d.add_paragraph(it, style="List Bullet"); p.paragraph_format.space_after = Pt(1.5)

    def job(org, role, dates):
        p = d.add_paragraph(); p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(1)
        r = p.add_run(org); r.bold = True; r.font.color.rgb = navy
        p.add_run(f"    {dates}").italic = True
        if role: p.add_run(f"\n{role}")

    para(NAME, size=19, bold=True, color=navy, after=0)
    para(c["headline"], size=11.5, bold=True, color=navy, after=2)
    para(CONTACT, size=9, after=0); para(STATUS, size=9, after=2)
    heading("Profile"); para(c["profile"])
    if c["vce_line"]: para(c["vce_line"])
    p = para("", after=2); p.add_run("Teaching areas: ").bold = True; p.add_run(c["areas"])
    heading("Teaching Experience")
    job("St Francis Xavier College, Berwick and Officer campuses", "Secondary Science Teacher, Years 7 to 9", "January 2023 to present")
    bullets(c["sfx"])
    for org, dates, line in OTHER_EXPERIENCE:
        job(org, "", dates); bullets([line])
    heading("Education")
    for deg, rest in EDUCATION:
        p = d.add_paragraph(style="List Bullet"); p.paragraph_format.space_after = Pt(1.5)
        p.add_run(deg).bold = True; p.add_run(f", {rest}")
    heading("Registration and Professional Learning"); bullets(REGISTRATION)
    heading("Publications"); para(PUBLICATIONS)
    heading("Referees")
    for n, r_, ph, em in REFEREES:
        p = d.add_paragraph(style="List Bullet"); p.paragraph_format.space_after = Pt(1.5)
        p.add_run(n).bold = True; p.add_run(f", {r_}. {ph}, {em}")
    d.save(path)

# ------------------------------------------------------------------ main
if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    mixed = "--mixed" in sys.argv; catholic = "--catholic" in sys.argv
    variants = args or ["science", "vce"]
    for v in variants:
        c = assemble(v, mixed, catholic)
        suffix = {"science": "Science_Teacher", "vce": "Science_VCE_Biology"}[v]
        suffix += ("_mixed" if mixed else "") + ("_catholic" if catholic else "")
        base = OUT / f"Pradeep_Penugonda_Resume_{suffix}"
        base.with_suffix(".html").write_text(build_html(c), encoding="utf-8")
        subprocess.run([CHROME, "--headless", "--no-sandbox", "--disable-gpu",
                        f"--print-to-pdf={base.with_suffix('.pdf')}", "--no-pdf-header-footer",
                        f"file://{base.with_suffix('.html')}"], capture_output=True)
        build_docx(c, base.with_suffix(".docx"))
        print("built:", base.with_suffix(".pdf").name)
