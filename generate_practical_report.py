import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import os

doc = docx.Document()

# Set standard 1-inch margins
for section in doc.sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    
    # Header & Footer setup
    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hrun = hp.add_run('MIT AOE | Dept of Software Engineering | WT Lab Practical Exam')
    hrun.font.name = 'Times New Roman'
    hrun.font.size = Pt(8.5)
    hrun.font.color.rgb = RGBColor(148, 163, 184)
    
    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    frun = fp.add_run('JobConnect — Full-Stack Job Portal | Yash Mali (PRN: 202401040251) | Division B')
    frun.font.name = 'Times New Roman'
    frun.font.size = Pt(8.5)
    frun.font.color.rgb = RGBColor(148, 163, 184)

# Color Constants
PRIMARY_COLOR = RGBColor(15, 23, 42)      # Dark Slate #0F172A
ACCENT_BLUE = RGBColor(37, 99, 235)       # Royal Blue #2563EB
TEXT_MUTED = RGBColor(100, 116, 139)      # Gray #64748B
WHITE = RGBColor(255, 255, 255)

w_ns = nsdecls('w')

def set_cell_background(cell, hex_color):
    xml_str = f'<w:shd {w_ns} w:fill="{hex_color}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(xml_str))

def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders_xml = (
        f'<w:tblBorders {w_ns}>'
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>'
        f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="1E293B"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
        f'<w:insideV w:val="none"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(parse_xml(borders_xml))

def style_standard_table(table, col_widths=None):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)
    
    # Style Header Row
    for idx, cell in enumerate(table.rows[0].cells):
        set_cell_background(cell, "1E293B")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            for r in p.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(10)
                r.font.bold = True
                r.font.color.rgb = WHITE
                
    # Style Data Rows
    for r_idx, row in enumerate(table.rows[1:]):
        bg = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for cell in row.cells:
            if bg != "FFFFFF":
                set_cell_background(cell, bg)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.line_spacing = 1.15
                for r in p.runs:
                    r.font.name = 'Times New Roman'
                    r.font.size = Pt(9.5)
                    
    if col_widths:
        for row in table.rows:
            for idx, w in enumerate(col_widths):
                if idx < len(row.cells):
                    row.cells[idx].width = Inches(w)

def add_title(text, size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12, color=PRIMARY_COLOR):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return p

def add_heading(text, level=1):
    h = doc.add_heading(text, level=level)
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(6)
    for run in h.runs:
        run.font.name = 'Times New Roman'
        run.font.color.rgb = PRIMARY_COLOR
    return h

def add_p(text, bold_prefix='', space_after=6, italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_b = p.add_run(bold_prefix)
        r_b.font.name = 'Times New Roman'
        r_b.font.size = Pt(11)
        r_b.font.bold = True
    r = p.add_run(text)
    r.font.name = 'Times New Roman'
    r.font.size = Pt(11)
    r.font.italic = italic
    return p

def add_code_block(code_text):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    set_cell_background(cell, "0F172A")
    cell.width = Inches(6.5)
    
    tblPr = table._tbl.tblPr
    tblPr.append(parse_xml(f'<w:tblBorders {w_ns}><w:top w:val="none"/><w:left w:val="single" w:sz="18" w:space="0" w:color="3B82F6"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>'))
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(code_text)
    run.font.name = 'Courier New'
    run.font.size = Pt(8.5)
    run.font.color.rgb = RGBColor(226, 232, 240)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

# ═══════════════════════════════════════════════════════════════════════════════
# 1. COVER PAGE
# ═══════════════════════════════════════════════════════════════════════════════

add_title('MIT ACADEMY OF ENGINEERING, ALANDI, PUNE', size=16, bold=True, space_after=3)
add_title('(An Autonomous Institute Affiliated to Savitribai Phule Pune University)', size=10, bold=False, space_after=18, color=TEXT_MUTED)

add_title('DEPARTMENT OF SOFTWARE ENGINEERING', size=13, bold=True, space_after=4)
add_title('WEB TECHNOLOGY LABORATORY (WT LAB)', size=14, bold=True, space_after=4)
add_title('WEB TECHNOLOGY LABORATORY — PRACTICAL EXAMINATION', size=14, bold=True, space_after=28)

add_title('Course Project Report on', size=11, bold=False, space_after=4, color=TEXT_MUTED)
add_title('JOBCONNECT', size=22, bold=True, space_after=4, color=ACCENT_BLUE)
add_title('Full-Stack Job Portal & Recruitment Management System', size=13, bold=True, space_after=32)

# Student Details Table
t_meta = doc.add_table(rows=5, cols=2)
t_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(t_meta)
meta_rows = [
    ('Candidate Name:', 'Yash Mali'),
    ('Permanent Registration No. (PRN):', '202401040251'),
    ('Branch & Program:', 'B.Tech in Software Engineering'),
    ('Class & Division:', 'Third Year (TY) — Division B'),
    ('Academic Year:', '2026 – 2027')
]
for idx, (lbl, val) in enumerate(meta_rows):
    row = t_meta.rows[idx]
    set_cell_background(row.cells[0], "F1F5F9")
    r0 = row.cells[0].paragraphs[0].add_run(lbl)
    r0.font.bold = True
    r0.font.name = 'Times New Roman'
    r0.font.size = Pt(10.5)
    r1 = row.cells[1].paragraphs[0].add_run(val)
    r1.font.name = 'Times New Roman'
    r1.font.size = Pt(10.5)
    row.cells[0].width = Inches(3.0)
    row.cells[1].width = Inches(3.5)

doc.add_paragraph().paragraph_format.space_before = Pt(20)

# Teacher Evaluation Rubric Sheet
add_title('FACULTY EVALUATION & ASSESSMENT RUBRICS', size=11, bold=True, space_after=6)
t_rubric = doc.add_table(rows=1, cols=4)
r_hdr = t_rubric.rows[0].cells
r_hdr[0].paragraphs[0].add_run('Evaluation Parameter')
r_hdr[1].paragraphs[0].add_run('Max Marks')
r_hdr[2].paragraphs[0].add_run('Marks Obtained')
r_hdr[3].paragraphs[0].add_run('Remarks & Sign')
rubric_data = [
    ('1. Problem Identification, UI/UX Design & Responsiveness', '05', '', ''),
    ('2. RESTful API Architecture & Server Route Implementation', '05', '', ''),
    ('3. Database Connectivity, Persistence & Dual-Layer Validation', '05', '', ''),
    ('4. Individual Demonstration, Viva-Voce & Code Explanation', '05', '', ''),
    ('TOTAL MARKS', '20', '', '')
]
for p1, p2, p3, p4 in rubric_data:
    row = t_rubric.add_row().cells
    row[0].paragraphs[0].add_run(p1)
    row[1].paragraphs[0].add_run(p2)
    row[2].paragraphs[0].add_run(p3)
    row[3].paragraphs[0].add_run(p4)
    if 'TOTAL' in p1:
        for c in row:
            for run in c.paragraphs[0].runs:
                run.font.bold = True
style_standard_table(t_rubric, [3.2, 1.0, 1.1, 1.2])

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 2. STUDENT DECLARATION & TABLE OF CONTENTS
# ═══════════════════════════════════════════════════════════════════════════════

add_heading('STUDENT DECLARATION', level=1)
add_p(
    'I, Yash Mali (PRN: 202401040251), hereby declare that the project entitled "JobConnect — Full-Stack Job Portal & Recruitment Management System" '
    'is a bona fide submission carried out by me as an individual Practical Examination project for the Web Technology Laboratory course. '
    'The front-end design, back-end Node.js/Express.js RESTful API, schema-based JSON persistence, and dual-layer validation mechanisms described in this '
    'report represent my original work and have been individually tested and verified on the local host execution environment.'
)

doc.add_paragraph().paragraph_format.space_before = Pt(12)
p_sig = doc.add_paragraph()
p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
r_sig1 = p_sig.add_run('Yash Mali\n')
r_sig1.font.bold = True
r_sig1.font.name = 'Times New Roman'
r_sig1.font.size = Pt(11)
r_sig2 = p_sig.add_run('PRN: 202401040251 | TY B.Tech Software Engineering (Div B)\nMIT Academy of Engineering, Pune')
r_sig2.font.name = 'Times New Roman'
r_sig2.font.size = Pt(10)
r_sig2.font.color.rgb = TEXT_MUTED

doc.add_paragraph().paragraph_format.space_after = Pt(14)

add_heading('TABLE OF CONTENTS', level=1)
t_index = doc.add_table(rows=1, cols=3)
i_hdr = t_index.rows[0].cells
i_hdr[0].paragraphs[0].add_run('S.No.')
i_hdr[1].paragraphs[0].add_run('Chapter / Section Title')
i_hdr[2].paragraphs[0].add_run('Topic Coverage')
index_data = [
    ('1', 'Project Aim & Problem Statement', 'Real-world recruitment challenges, project motivation & objectives'),
    ('2', 'Course Outcomes (CO) Mapping', 'Mapping project deliverables to WT Lab course educational objectives'),
    ('3', 'System Architecture & Data Flow', 'Decoupled client-server model, asynchronous fetch() flow diagram'),
    ('4', 'Hardware & Software Specifications', 'Execution prerequisites, runtimes, package dependencies'),
    ('5', 'Database Schema & Persistence Design', 'Collections schema for jobs and applications, JSON storage engine'),
    ('6', 'RESTful API Specification', 'Comprehensive HTTP endpoints, request bodies, query params & response codes'),
    ('7', 'Dual-Layer Validation Matrix', 'Client-side on-blur regex verification vs server-side HTTP 400 enforcement'),
    ('8', 'Core Implementation & Code Snippets', 'Server routes, validation logic, dynamic client-side rendering script'),
    ('9', 'Application Screenshots & Demonstration', 'Visual walkthrough of Homepage, Job Specs, Apply Form, & Dashboard'),
    ('10', 'Teacher Demonstration & Viva-Voce Guide', 'Comprehensive Q&A covering architectural and technical decisions'),
    ('11', 'Conclusion & Future Enhancements', 'Summary of achievements, modular extensibility & conclusion')
]
for s, title, desc in index_data:
    row = t_index.add_row().cells
    row[0].paragraphs[0].add_run(s)
    row[1].paragraphs[0].add_run(title)
    row[2].paragraphs[0].add_run(desc)
style_standard_table(t_index, [0.6, 2.5, 3.4])

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 3. AIM, PROBLEM STATEMENT & COURSE OUTCOMES
# ═══════════════════════════════════════════════════════════════════════════════

add_heading('1. PROJECT AIM & PROBLEM STATEMENT', level=1)
add_p(
    'To design, implement, and demonstrate an individual, complete full-stack web application named "JobConnect — Full-Stack Job Portal & Recruitment Management System" '
    'addressing modern digital recruitment workflows. The project showcases responsive front-end user experience, asynchronous RESTful API design using Node.js and Express.js, '
    'persistent file-backed JSON database management, rigorous dual-layer client/server input validation, and recruiter candidate pipeline tracking.',
    bold_prefix='1.1 Project Aim: '
)

add_p(
    'In contemporary employment ecosystems, candidate application experiences suffer from high friction, opaque progress tracking, and cumbersome paper resume exchanges. '
    'Simultaneously, hiring managers and corporate recruiters struggle to organize applicant pools, track candidates through selection stages, and publish dynamic openings without complex database overhead. '
    'JobConnect eliminates these challenges by providing a unified, decoupled digital platform that bridges student applicants directly with corporate hiring teams.',
    bold_prefix='1.2 Problem Statement: '
)

add_heading('2. COURSE OUTCOMES (CO) MAPPING', level=1)
add_p('This Practical Examination project directly fulfills the prescribed learning outcomes of the Web Technology Laboratory course:')

t_co = doc.add_table(rows=1, cols=3)
t_co.rows[0].cells[0].paragraphs[0].add_run('Course Outcome')
t_co.rows[0].cells[1].paragraphs[0].add_run('Statement of Competency')
t_co.rows[0].cells[2].paragraphs[0].add_run('Project Implementation in JobConnect')
co_data = [
    ('CO1', 'Design responsive and interactive client interfaces using HTML5, CSS3, and modern DOM manipulation.', 'Responsive UI featuring CSS Grid, custom tokens, chips, card feeds, and dynamic modal dialogs.'),
    ('CO2', 'Develop asynchronous, client-side event-driven applications using JavaScript and Web APIs.', 'Asynchronous fetch() calls, real-time input event listeners, on-blur validation, and character counters.'),
    ('CO3', 'Build modular back-end RESTful microservices using Node.js and Express framework.', 'Structured Express REST API listening on port 3000 with 9 endpoints adhering to standard HTTP verbs and status codes.'),
    ('CO4', 'Implement data persistence, error handling, and robust input validation security mechanisms.', 'Atomic read/write JSON database (db.json) with dual-layer regex validation preventing malformed or invalid inputs.')
]
for co, stmt, impl in co_data:
    row = t_co.add_row().cells
    row[0].paragraphs[0].add_run(co)
    row[1].paragraphs[0].add_run(stmt)
    row[2].paragraphs[0].add_run(impl)
style_standard_table(t_co, [0.8, 2.7, 3.0])

add_heading('3. SYSTEM ARCHITECTURE & DATA FLOW', level=1)
add_p(
    'JobConnect employs a decoupled full-stack architecture adhering to the Client-Server architectural pattern. '
    'The presentation tier is completely decoupled from the back-end application server, communicating exclusively through serialized JSON HTTP request/response cycles.'
)

add_code_block(
    "+-------------------------------------------------------------------------+\n"
    "|                       CLIENT TIER (Front-End)                           |\n"
    "|  - index.html        (Live Job Feed & Multi-faceted Category Filters)   |\n"
    "|  - job-detail.html   (Detailed Specs & Validated Direct Application)    |\n"
    "|  - dashboard.html    (Recruiter Operations & 5-Stage Status Tracking)   |\n"
    "|  - post-job.html     (Vacancies Publishing Portal)                      |\n"
    "+------------------------------------+------------------------------------+\n"
    "                                     | Asynchronous JSON HTTP Requests\n"
    "                                     | (GET, POST, PATCH, DELETE via fetch)\n"
    "                                     v\n"
    "+-------------------------------------------------------------------------+\n"
    "|                    APPLICATION SERVER TIER (Node.js)                    |\n"
    "|  - Express.js v5 REST API Controller                                    |\n"
    "|  - Request Validation Middleware (Indian Phone Regex, Email RFC, URLs)   |\n"
    "|  - CORS & URL-Encoded / JSON Body Parsing                               |\n"
    "|  - HTTP Status Code Dispatcher (200 OK, 201 Created, 400 Bad, 404 N/F)  |\n"
    "+------------------------------------+------------------------------------+\n"
    "                                     | Atomic File I/O Serialization\n"
    "                                     | (fs.readFileSync / fs.writeFileSync)\n"
    "                                     v\n"
    "+-------------------------------------------------------------------------+\n"
    "|                   PERSISTENT STORAGE TIER (File-Backed)                 |\n"
    "|  - db.json (Schema collections for jobs[] and candidate applications[]) |\n"
    "|  - db.example.json (Sanitized initial seed fallback)                    |\n"
    "+-------------------------------------------------------------------------+"
)

# ═══════════════════════════════════════════════════════════════════════════════
# 4. SPECIFICATIONS & DATABASE DESIGN
# ═══════════════════════════════════════════════════════════════════════════════

add_heading('4. HARDWARE & SOFTWARE SPECIFICATIONS', level=1)

t_spec = doc.add_table(rows=1, cols=3)
t_spec.rows[0].cells[0].paragraphs[0].add_run('Category')
t_spec.rows[0].cells[1].paragraphs[0].add_run('Component / Tool')
t_spec.rows[0].cells[2].paragraphs[0].add_run('Minimum Requirement / Description')
spec_data = [
    ('Hardware', 'Processor', 'Intel Core i3 / AMD Ryzen 3 / Apple Silicon (M1/M2/M3) or higher'),
    ('Hardware', 'Primary Memory (RAM)', '4 GB RAM minimum (8 GB recommended for multitasking)'),
    ('Hardware', 'Storage Space', 'Minimum 500 MB free hard disk / SSD space'),
    ('Software', 'Operating System', 'Cross-platform: macOS 14+, Windows 10/11, or Ubuntu Linux 20.04+'),
    ('Software', 'Runtime Environment', 'Node.js (v18.x, v20.x, or v22.x LTS installed)'),
    ('Software', 'Package Manager', 'npm (v9.x or higher)'),
    ('Software', 'Application Framework', 'Express.js v5 (Fast, unopinionated minimalist web framework)'),
    ('Software', 'Web Client', 'Google Chrome, Microsoft Edge, Mozilla Firefox, or Safari'),
    ('Software', 'Code Editor', 'Visual Studio Code (VS Code) with JavaScript ES6 extensions')
]
for c, comp, req in spec_data:
    row = t_spec.add_row().cells
    row[0].paragraphs[0].add_run(c)
    row[1].paragraphs[0].add_run(comp)
    row[2].paragraphs[0].add_run(req)
style_standard_table(t_spec, [1.1, 2.0, 3.4])

add_heading('5. DATABASE SCHEMA & PERSISTENCE DESIGN', level=1)
add_p(
    'The project utilizes a persistent, schema-driven file-backed JSON database (`db.json`). '
    'This lightweight architecture avoids external database daemon installations while guaranteeing atomic data persistence across server reboots. '
    'The database encapsulates two core entity collections:'
)

add_p('1. jobs Collection: Stores published vacancies with fields `id` (String primary key), `title`, `company`, `category`, `location`, `type`, `salary`, `experience`, `skills` (Array of Strings), `description`, `requirements`, and `postedAt` (ISO Date).')
add_p('2. applications Collection: Stores candidate submissions with fields `id` (String primary key), `jobId` (Foreign key referencing jobs.id), `jobTitle`, `company`, `name`, `email`, `phone`, `experience`, `qualification`, `resumeUrl`, `coverLetter`, `status` (Enum: Pending, Under Review, Shortlisted, Rejected, Offered), and `appliedAt` (ISO Date).')

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 5. REST API & VALIDATION MATRIX
# ═══════════════════════════════════════════════════════════════════════════════

add_heading('6. RESTFUL API SPECIFICATION', level=1)
add_p('All application operations are powered by a RESTful HTTP service adhering to standard architectural constraints:')

t_api = doc.add_table(rows=1, cols=4)
t_api.rows[0].cells[0].paragraphs[0].add_run('Endpoint')
t_api.rows[0].cells[1].paragraphs[0].add_run('Method')
t_api.rows[0].cells[2].paragraphs[0].add_run('Payload / Query')
t_api.rows[0].cells[3].paragraphs[0].add_run('Status Code & Action')
api_data = [
    ('/api/stats', 'GET', 'None', '200 OK — Aggregated platform metrics (Total jobs, applications, shortlisted).'),
    ('/api/jobs', 'GET', '?search=&type=&category=&location=', '200 OK — Array of jobs filtered by multi-criteria search parameters.'),
    ('/api/jobs/:id', 'GET', 'URL param id', '200 OK — Detailed job specs; returns 404 if identifier is invalid.'),
    ('/api/jobs', 'POST', 'JSON: {title, company, location, ...}', '201 Created — Validates and creates a new job posting.'),
    ('/api/jobs/:id', 'DELETE', 'URL param id', '200 OK — Deletes job posting and cascading candidate applications.'),
    ('/api/apply', 'POST', 'JSON: {jobId, name, email, phone, ...}', '201 Created / 400 Bad Request — Strict dual validation candidate application.'),
    ('/api/applications', 'GET', '?jobId=&status=&search=', '200 OK — Candidate submissions filtered for recruiter dashboard.'),
    ('/api/applications/:id/status', 'PATCH', 'JSON: {status: "Offered"}', '200 OK — Updates applicant lifecycle status in database.'),
    ('/api/applications/:id', 'DELETE', 'URL param id', '200 OK — Deletes specific applicant profile from database.')
]
for ep, m, py, st in api_data:
    row = t_api.add_row().cells
    row[0].paragraphs[0].add_run(ep)
    row[1].paragraphs[0].add_run(m)
    row[2].paragraphs[0].add_run(py)
    row[3].paragraphs[0].add_run(st)
style_standard_table(t_api, [1.5, 0.7, 2.0, 2.3])

add_heading('7. DUAL-LAYER VALIDATION MATRIX', level=1)
add_p('To safeguard database integrity and provide real-time user feedback, validation is implemented symmetrically on both client and server tiers:')

t_val = doc.add_table(rows=1, cols=4)
t_val.rows[0].cells[0].paragraphs[0].add_run('Field Name')
t_val.rows[0].cells[1].paragraphs[0].add_run('Client-Side Check')
t_val.rows[0].cells[2].paragraphs[0].add_run('Server-Side Enforcement')
t_val.rows[0].cells[3].paragraphs[0].add_run('Validation Rule / Error Trigger')
val_data = [
    ('Candidate Name', 'on-blur length check', 'name.trim().length >= 3', 'Must contain at least 3 characters.'),
    ('Email Address', 'on-blur Regex test', r'/^[^\s@]+@[^\s@]+\.[^\s@]+$/', 'Must match valid email structure (e.g. user@domain.com).'),
    ('Mobile Number', '10-digit Indian Regex', r'/^[6-9]\d{9}$/', 'Must be 10 digits starting with 6, 7, 8, or 9.'),
    ('Qualification', 'Dropdown non-empty selection', 'qualification.trim().length >= 2', 'Must select an accredited degree/diploma.'),
    ('Experience', 'Finite numeric bounds', '0 <= Number(exp) <= 50', 'Numeric value between 0 (Fresher) and 50 years.'),
    ('Resume URL', 'Protocol validation', 'new URL(url).protocol in [http, https]', 'Optional; if provided must start with http:// or https://'),
    ('Job Title', 'Minimum length check', 'title.trim().length >= 3', 'Must be at least 3 characters long.'),
    ('Job Description', 'Minimum length check', 'description.trim().length >= 15', 'Must provide meaningful description (min 15 chars).')
]
for f, c_chk, s_chk, err in val_data:
    row = t_val.add_row().cells
    row[0].paragraphs[0].add_run(f)
    row[1].paragraphs[0].add_run(c_chk)
    row[2].paragraphs[0].add_run(s_chk)
    row[3].paragraphs[0].add_run(err)
style_standard_table(t_val, [1.1, 1.6, 1.8, 2.0])

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 6. CORE CODE IMPLEMENTATION
# ═══════════════════════════════════════════════════════════════════════════════

add_heading('8. CORE IMPLEMENTATION & CODE SNIPPETS', level=1)

add_p('8.1 Server Route: Submitting Candidate Application with Validation (server.js):', bold_prefix='')
add_code_block(
    "app.post('/api/apply', (req, res) => {\n"
    "  const { jobId, name, email, phone, experience, qualification, resumeUrl, coverLetter } = req.body;\n"
    "  const errors = {};\n"
    "  if (!name || name.trim().length < 3) errors.name = 'Full name must contain at least 3 characters.';\n"
    "  if (!email || !isValidEmail(email)) errors.email = 'Please provide a valid email address.';\n"
    "  if (!phone || !isValidIndianPhone(phone)) errors.phone = 'Enter a valid 10-digit Indian phone number.';\n"
    "  if (Object.keys(errors).length > 0) return res.status(400).json({ success: false, errors });\n\n"
    "  const db = readDB();\n"
    "  const newApp = { id: 'app_' + Date.now(), jobId, name: name.trim(), email: email.trim(), status: 'Pending', ... };\n"
    "  db.applications.unshift(newApp);\n"
    "  writeDB(db);\n"
    "  res.status(201).json({ success: true, message: 'Application submitted successfully!', application: newApp });\n"
    "});"
)

add_p('8.2 Server Route: Recruiter Status Transition & Atomic Persistence (server.js):', bold_prefix='')
add_code_block(
    "app.patch('/api/applications/:id/status', (req, res) => {\n"
    "  const { status } = req.body;\n"
    "  const allowedStatuses = ['Pending', 'Under Review', 'Shortlisted', 'Rejected', 'Offered'];\n"
    "  if (!allowedStatuses.includes(status)) return res.status(400).json({ error: 'Invalid status' });\n\n"
    "  const db = readDB();\n"
    "  const idx = db.applications.findIndex(a => String(a.id) === String(req.params.id));\n"
    "  if (idx === -1) return res.status(404).json({ error: 'Application record not found' });\n\n"
    "  db.applications[idx].status = status;\n"
    "  db.applications[idx].updatedAt = new Date().toISOString().split('T')[0];\n"
    "  writeDB(db);\n"
    "  res.json({ success: true, message: `Status updated to \"${status}\"`, application: db.applications[idx] });\n"
    "});"
)

add_p('8.3 Client Asynchronous Data Fetching & Dynamic Card Grid (index.html):', bold_prefix='')
add_code_block(
    "async function loadJobs() {\n"
    "  const params = new URLSearchParams({ search: activeSearch, category: activeCategory, type: activeType });\n"
    "  const res = await fetch(`/api/jobs?${params}`);\n"
    "  const jobs = await res.json();\n"
    "  document.getElementById('jobGrid').innerHTML = jobs.map(j => `\n"
    "    <div class=\"job-card\">\n"
    "      <div class=\"card-top\">\n"
    "        <div class=\"company-avatar\">${j.company[0]}</div>\n"
    "        <a href=\"job-detail.html?id=${j.id}\" class=\"job-role-title\">${j.title}</a>\n"
    "        <span class=\"badge-tag\">${j.type}</span>\n"
    "      </div>\n"
    "      <div class=\"card-meta-pills\"><span>📍 ${j.location}</span><span>💰 ${j.salary}</span></div>\n"
    "      <a href=\"job-detail.html?id=${j.id}\" class=\"btn btn-primary btn-sm\">View Details & Apply →</a>\n"
    "    </div>\n"
    "  `).join('');\n"
    "}"
)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 7. APPLICATION DEMONSTRATION & SCREENSHOTS
# ═══════════════════════════════════════════════════════════════════════════════

add_heading('9. SCREENSHOTS AND APPLICATION DEMONSTRATION', level=1)
img_folder = os.path.join(os.path.dirname(__file__), 'screenshots')
screenshots = [
    ('Figure_1_Homepage.png', 'Figure 1: JobConnect Homepage displaying live metric counters, search bar, category filter chips, and dynamic job card feed.'),
    ('Figure_2_Job_Details.png', 'Figure 2: Job Details page showing full role specifications, requirements, compensation, and experience criteria.'),
    ('Figure_3_Candidate_Validation.png', 'Figure 3: Interactive Candidate Direct Application Form populated with student data and client-side validation rules.'),
    ('Figure_4_Submission_Success.png', 'Figure 4: Confirmed application submission displaying instantaneous success notification banner and toast message.'),
    ('Figure_5_Recruiter_Dashboard.png', 'Figure 5: Recruiter & Operations Dashboard showing candidate table, inline status dropdowns, and live aggregate metrics.'),
    ('Figure_6_Post_Job.png', 'Figure 6: Recruiter Job Posting Portal with category selection, compensation ranges, and skill tag publishing.')
]
for fname, caption in screenshots:
    fpath = os.path.join(img_folder, fname)
    if os.path.exists(fpath):
        doc.add_paragraph().paragraph_format.space_before = Pt(8)
        doc.add_picture(fpath, width=Inches(6.2))
        cp = doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_after = Pt(14)
        cpr = cp.add_run(caption)
        cpr.font.name = 'Times New Roman'
        cpr.font.size = Pt(9.5)
        cpr.font.italic = True

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 8. TEACHER DEMONSTRATION & VIVA-VOCE GUIDE
# ═══════════════════════════════════════════════════════════════════════════════

add_heading('10. TEACHER DEMONSTRATION & VIVA-VOCE GUIDE', level=1)
add_p('During the evaluation presentation to the course examiner, the following key technical concepts can be defended:')

viva_qa = [
    ('Q1: What real-world problem does JobConnect solve?',
     'A1: JobConnect provides an end-to-end digital employment platform connecting candidates with vetted tech positions, eliminating paper resumes through structured digital submissions and providing recruiters with live applicant tracking pipelines.'),
    ('Q2: Explain the client-server interaction and architectural design in this project.',
     'A2: The front-end is decoupled from the server. When a user interacts with the UI (e.g. searching, applying, updating status), client JavaScript sends asynchronous fetch() HTTP requests with JSON payloads to the Express REST API. The server validates the request, updates db.json, and returns appropriate HTTP status codes without full page reloads.'),
    ('Q3: How is data stored and managed across server restarts?',
     'A3: Data is stored in db.json as schema-compliant JSON documents for jobs and applications. Node.js fs module reads and atomically writes serialized records, ensuring data survives server restarts without external database engine dependencies.'),
    ('Q4: What is dual-layer validation and why is it necessary?',
     'A4: Dual-layer validation validates data at both tiers. Client-side validation uses on-blur events and regex checks to provide immediate visual feedback (red borders, clear error text). Server-side validation inspects incoming payloads and rejects malformed requests with HTTP 400 Bad Request, guaranteeing database integrity even if client scripts are bypassed.'),
    ('Q5: What are the primary recruiter operations available in the dashboard?',
     'A5: Recruiters have access to the Dashboard (dashboard.html) displaying live statistics, an applicant table with inline status dropdowns (Pending, Under Review, Shortlisted, Rejected, Offered), candidate cover letter modal viewers, and a job creation hub (post-job.html).'),
    ('Q6: Why did you use PATCH instead of PUT for candidate status updates?',
     'A6: A PUT request replaces an entire entity, whereas PATCH applies a partial update. Since updating candidate pipeline status only modifies the status and updatedAt attributes while keeping qualifications and personal details intact, PATCH /api/applications/:id/status is semantically correct.')
]
for q, a in viva_qa:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    rq = p.add_run(q + '\n')
    rq.font.bold = True
    rq.font.size = Pt(10.5)
    rq.font.name = 'Times New Roman'
    ra = p.add_run(a)
    ra.font.size = Pt(10.5)
    ra.font.name = 'Times New Roman'

add_heading('11. CONCLUSION & FUTURE ENHANCEMENTS', level=1)
add_p(
    'The JobConnect Full-Stack Mini-Project successfully fulfills all requirements of the Web Technology Laboratory Practical Examination. '
    'It demonstrates an individual, production-ready web application integrating modern front-end engineering, asynchronous Node.js and Express RESTful API backend, '
    'schema-driven persistent JSON database, robust dual-layer validation, and an interactive recruiter workflow.',
    bold_prefix='11.1 Conclusion: '
)

add_p(
    '1. Authentication & Role-Based Access Control (RBAC): Integrating JWT (JSON Web Tokens) or OAuth2 to distinguish recruiter logins from candidate accounts.\n'
    '2. Cloud File Storage: Integrating AWS S3 or Cloudinary for direct PDF resume file uploads and previews.\n'
    '3. Automated Email Alerts: Sending automated status change notifications (e.g. "You have been Shortlisted!") to candidate emails via Nodemailer.',
    bold_prefix='11.2 Future Scope & Enhancements: '
)

# Save report
primary_path = '202401040251_YashMali_JobConnect.docx'
doc.save(primary_path)
print('Successfully saved report to', primary_path)
