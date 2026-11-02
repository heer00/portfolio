from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable, KeepTogether

ROOT = Path(__file__).parent
OUT = ROOT / "assets"
OUT.mkdir(exist_ok=True)

name = "Heer Mehta"
contact = "mehta.heer910@gmail.com  |  linkedin.com/in/heer-mehta-14833928b  |  github.com/heer00"
summary = ("M.Sc. Information Technology student with hands-on experience exploring enterprise AI workflows and building "
           "backend APIs, full-stack applications, and machine learning projects. Open to software development roles and "
           "freelance projects, with a focus on backend and full-stack engineering.")
education = [
    ("Dhirubhai Ambani University", "M.Sc. Information Technology | Gandhinagar, Gujarat | 2025–Present", "CPI: 6.38"),
    ("Shreyarth University", "B.Sc. Information Technology | Ahmedabad, Gujarat | 2022–2025", "CPI: 8.88"),
]
experience = [
    ("Tata Motors - Digital.ai Labs", "Intern | Thane, Mumbai | May 2026–July 2026", [
        "Researched LLM, RAG, vector database, embedding, prompt engineering, and agent workflow approaches for enterprise AI use cases.",
        "Supported model experimentation and technical documentation using AI-assisted development tools.",
        "Worked in a Linux-based Oracle Cloud Infrastructure environment supporting AI model deployment workflows.",
    ])
]
projects = [
    ("Book-E-Pedia | React, Django, REST APIs, SQLite", "Contributed to a full-stack bookstore with catalog and product pages, cart functionality, and admin management. Integrated a React frontend with Django APIs and relational data models."),
    ("InsightEdge | Python, FastAPI, PyMuPDF, Firebase", "Collaborated on a hackathon career guidance platform. Contributed resume parsing, structured skill extraction, role matching, and backend endpoints for resume analysis and chatbot integration."),
    ("Customer Segmentation | Python, Scikit-learn, Streamlit", "Applied K-Means clustering to customer behavior data; used feature scaling and the elbow method to assess clusters, then built an interactive Streamlit application."),
    ("Forest Fire Detection | Python, TensorFlow/Keras, Scikit-learn", "Built and evaluated a CNN-based wildfire detection project, documenting data preprocessing, visualization, and model analysis in Jupyter notebooks."),
]
skills = [
    ("Languages", "Python, JavaScript, TypeScript, C++, SQL"),
    ("Frontend", "React, HTML5, CSS3, Bootstrap, responsive design"),
    ("Backend & APIs", "Django, Django REST Framework, FastAPI, REST APIs"),
    ("Databases", "PostgreSQL, MySQL, SQLite, Firebase"),
    ("Data & AI", "Pandas, NumPy, Scikit-learn, LLMs, RAG, embeddings"),
    ("Tools", "Git, GitHub, Linux, Postman, VS Code, Oracle Cloud Infrastructure"),
    ("Foundations", "Data structures, OOP, operating systems, DBMS, computer networks"),
]
achievements = [
    "GenAI Powered Data Analytics Job Simulation, Forage | April 2026",
    "Deloitte Australia Data Analytics Virtual Experience, Forage | March 2026",
    "Participant, Ingenius Hackathon 2026 and InnovAItion Hackathon 2026",
]

# Word version
doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(.45); sec.bottom_margin = Inches(.45)
sec.left_margin = Inches(.62); sec.right_margin = Inches(.62)
normal = doc.styles['Normal']; normal.font.name = 'Aptos'; normal.font.size = Pt(9.6); normal.font.color.rgb = RGBColor(35, 45, 40)
normal.paragraph_format.space_after = Pt(1.5); normal.paragraph_format.line_spacing = 1.0
title = doc.add_paragraph(style='Title'); title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.paragraph_format.space_after = Pt(1)
r = title.add_run(name); r.font.name = 'Aptos Display'; r.font.size = Pt(23); r.font.bold = True; r.font.color.rgb = RGBColor(32, 6, 17)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after = Pt(1)
r = p.add_run('SOFTWARE DEVELOPER  |  M.Sc. INFORMATION TECHNOLOGY'); r.bold=True; r.font.size=Pt(9)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(6)
r=p.add_run(contact); r.font.size=Pt(8.3); r.font.color.rgb=RGBColor(82,96,87)

def heading(text):
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(5); p.paragraph_format.space_after=Pt(2)
    r=p.add_run(text.upper()); r.bold=True; r.font.size=Pt(10); r.font.color.rgb=RGBColor(92,93,22)
    pPr=p._p.get_or_add_pPr(); borders=OxmlElement('w:pBdr'); bottom=OxmlElement('w:bottom'); bottom.set(qn('w:val'),'single'); bottom.set(qn('w:sz'),'5'); bottom.set(qn('w:space'),'1'); bottom.set(qn('w:color'),'4D1022'); borders.append(bottom); pPr.append(borders)
def bullet(text):
    p=doc.add_paragraph(style='List Bullet'); p.paragraph_format.left_indent=Inches(.17); p.paragraph_format.first_line_indent=Inches(-.12); p.paragraph_format.space_after=Pt(1)
    p.add_run(text)
heading('Profile'); doc.add_paragraph(summary)
heading('Education')
for school, meta, result in education:
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(1)
    r=p.add_run(school); r.bold=True
    p.add_run('  |  '+meta+'  |  '+result)
heading('Experience')
for employer, meta, items in experience:
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(1)
    r=p.add_run(employer); r.bold=True; p.add_run('  |  '+meta)
    for item in items: bullet(item)
heading('Selected Projects')
for label, detail in projects:
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(2)
    r=p.add_run(label+' — '); r.bold=True; p.add_run(detail)
heading('Technical Skills')
for label, value in skills:
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(0)
    r=p.add_run(label+': '); r.bold=True; p.add_run(value)
heading('Achievements')
for item in achievements: bullet(item)
doc.core_properties.title='Heer Mehta - Software Developer Resume'
doc.core_properties.subject='Software development resume'
doc.core_properties.author='Heer Mehta'
doc.save(OUT/'Heer_Mehta_Resume.docx')

# Matching PDF version
green=colors.HexColor('#5C5D16'); muted=colors.HexColor('#5C5D16'); ink=colors.HexColor('#200611'); rule=colors.HexColor('#4D1022')
styles=getSampleStyleSheet()
styles.add(ParagraphStyle(name='Name',parent=styles['Normal'],fontName='Helvetica-Bold',fontSize=21,leading=23,textColor=green,alignment=TA_CENTER,spaceAfter=1))
styles.add(ParagraphStyle(name='Role',parent=styles['Normal'],fontName='Helvetica-Bold',fontSize=8.5,leading=10,textColor=ink,alignment=TA_CENTER,spaceAfter=2))
styles.add(ParagraphStyle(name='Contact',parent=styles['Normal'],fontName='Helvetica',fontSize=7.7,leading=9,textColor=muted,alignment=TA_CENTER,spaceAfter=5))
styles.add(ParagraphStyle(name='SectionX',parent=styles['Normal'],fontName='Helvetica-Bold',fontSize=10,leading=12,textColor=green,spaceBefore=5,spaceAfter=1.5))
styles.add(ParagraphStyle(name='BodyX',parent=styles['Normal'],fontName='Helvetica',fontSize=9.4,leading=11.6,textColor=ink,spaceAfter=1.5))
styles.add(ParagraphStyle(name='SmallX',parent=styles['Normal'],fontName='Helvetica',fontSize=9,leading=11,textColor=ink,spaceAfter=1.5))
story=[Paragraph(name,styles['Name']),Paragraph('SOFTWARE DEVELOPER  |  M.Sc. INFORMATION TECHNOLOGY',styles['Role']),Paragraph('mehta.heer910@gmail.com  |  linkedin.com/in/heer-mehta-14833928b  |  github.com/heer00',styles['Contact'])]
def sect(t): story.extend([Paragraph(t.upper(),styles['SectionX']),HRFlowable(width='100%',thickness=.5,color=rule,spaceAfter=2)])
sect('Profile'); story.append(Paragraph(summary,styles['BodyX']))
sect('Education')
for school,meta,result in education: story.append(Paragraph(f'<b>{school}</b>  |  {meta}  |  {result}',styles['SmallX']))
sect('Experience')
for employer,meta,items in experience:
    story.append(Paragraph(f'<b>{employer}</b>  |  {meta}',styles['SmallX']))
    for item in items: story.append(Paragraph('•  '+item,styles['SmallX']))
sect('Selected Projects')
for label,detail in projects: story.append(Paragraph(f'<b>{label}</b> — {detail}',styles['SmallX']))
sect('Technical Skills')
for label,value in skills: story.append(Paragraph(f'<b>{label}:</b> {value}',styles['SmallX']))
sect('Achievements')
for item in achievements: story.append(Paragraph('•  '+item,styles['SmallX']))
pdf=SimpleDocTemplate(str(OUT/'Heer_Mehta_Resume.pdf'),pagesize=letter,rightMargin=.62*inch,leftMargin=.62*inch,topMargin=.38*inch,bottomMargin=.38*inch,title='Heer Mehta - Software Developer Resume',author='Heer Mehta')
pdf.build(story)
