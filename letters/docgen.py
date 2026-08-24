# -*- coding: utf-8 -*-
"""
docgen.py
Generates .docx internship letters from a data row, replicating the exact layout
and formatting of official Mahidol University ICT letters:
  - Confirm_Letter_V1        -> used for "Student Referral Letter"
  - Request_Intern_Letter_V1 -> used for "Request Letter for Internship Placement"

Font standard: TH Sarabun New 14pt
Paragraph Indent: 1 tab (1.27 cm)
"""
import io
import os
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import subprocess
import tempfile

LOGO_PATH = os.path.join(os.path.dirname(__file__), 'static', 'letters', 'images', 'mu_logo.jpeg')

FACULTY_TH = "คณะเทคโนโลยีสารสนเทศและการสื่อสาร"
FACULTY_TH_ICT = "คณะเทคโนโลยีสารสนเทศและการสื่อสาร (ICT)"
UNIV_TH = "มหาวิทยาลัยมหิดล"
ADDRESS_LINE1_TH = "เลขที่ 999 ถนนพุทธมณฑล สาย 4 ตำบลศาลายา"
ADDRESS_LINE2_TH = "อำเภอพุทธมณฑล จังหวัดนครปฐม 73170"
TEL_TH = "โทร. 0-2441-0909 โทรสาร 0-2441-0808"
DEAN_TH = "(ดร.พัฒนศักดิ์ มงคลวัฒน์)"
DEAN_TITLE_TH = "คณบดีคณะเทคโนโลยีสารสนเทศและการสื่อสาร"
COURSE_CONTACT_TH = (
    "อาจารย์ผู้ดูแลรายวิชา อาจารย์ผศ. ดร.ศิริเพ็ญ พงษ์ไพเชฐ "
    "อีเมล์ siripen.pon@mahidol.ac.th หรือ โทรศัพท์ 083-0610496"
)

SUPPORTED_LETTER_TYPES = {
    "Student Referral Letter": "confirm",
    "Request Letter for Internship Placement": "request",
}

# Standard 1 tab indent = 0.5 inches = 1.27 cm
TAB_INDENT_CM = 1.27

# ---------------------------------------------------------------- helpers

def _set_font(run, size=14, bold=False, color=None, name="TH Sarabun New"):
    """Applies font styles with full Complex Script (CS) XML attributes required for Thai text."""
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = name

    rpr = run._element.get_or_add_rPr()

    # ASCII and Complex Script fonts
    rFonts = rpr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rpr.makeelement(qn('w:rFonts'), {})
        rpr.append(rFonts)
    rFonts.set(qn('w:ascii'), name)
    rFonts.set(qn('w:hAnsi'), name)
    rFonts.set(qn('w:cs'), name)
    rFonts.set(qn('w:eastAsia'), name)

    # Bold setting for complex script
    if bold:
        bCs = rpr.find(qn('w:bCs'))
        if bCs is None:
            bCs = rpr.makeelement(qn('w:bCs'), {})
            rpr.append(bCs)

    # Size setting for complex script (14pt = 28 half-points)
    szCs = rpr.find(qn('w:szCs'))
    if szCs is None:
        szCs = rpr.makeelement(qn('w:szCs'), {})
        rpr.append(szCs)
    szCs.set(qn('w:val'), str(int(size * 2)))

    if color:
        run.font.color.rgb = RGBColor(*color)


def _para(doc, text="", align=WD_ALIGN_PARAGRAPH.LEFT, size=14, bold=False,
          space_after=3, space_before=0, line_spacing=1.15, indent_first=None):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.line_spacing = line_spacing
    if indent_first is not None:
        p.paragraph_format.first_line_indent = Cm(indent_first)
    if text:
        r = p.add_run(text)
        _set_font(r, size=size, bold=bold)
    return p


def _base_doc():
    doc = Document()
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(1.2)    # Shifted upper
    section.bottom_margin = Cm(1.2) # Shifted upper
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.0)
    
    style = doc.styles['Normal']
    style.font.name = 'TH Sarabun New'
    style.font.size = Pt(14)
    return doc


def _letterhead(doc, faculty_name):
    """Top-centered logo with tight right-aligned faculty address positioned upper on the page."""
    if os.path.exists(LOGO_PATH):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_before = Pt(0)
        p_logo.paragraph_format.space_after = Pt(0)
        p_logo.paragraph_format.line_spacing = 1.0
        run_logo = p_logo.add_run()
        run_logo.add_picture(LOGO_PATH, width=Cm(2.2))

    # All header lines set to line_spacing=1.0 and space_after=0 for tight placement
    _para(doc, faculty_name, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, size=14, space_before=0, space_after=0, line_spacing=1.0)
    _para(doc, UNIV_TH, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, size=14, space_before=0, space_after=0, line_spacing=1.0)
    _para(doc, ADDRESS_LINE1_TH, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, size=14, space_before=0, space_after=0, line_spacing=1.0)
    _para(doc, ADDRESS_LINE2_TH, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, size=14, space_before=0, space_after=0, line_spacing=1.0)
    _para(doc, TEL_TH, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, size=14, space_before=0, space_after=2, line_spacing=1.0)


def convert_docx_bytes_to_pdf(docx_bytes):
    """Converts DOCX binary data into PDF binary data using headless LibreOffice."""
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "temp_letter.docx")
        pdf_path = os.path.join(tmpdir, "temp_letter.pdf")
        
        # Save DOCX buffer to temporary file
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)
        
        # Run LibreOffice CLI conversion
        cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, docx_path]
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        # Read converted PDF bytes
        with open(pdf_path, "rb") as f:
            return f.read()


def generate_pdf_for_row(row, ref_no="___"):
    """Generates PDF bytes and filename for a given row."""
    docx_buf, fname = generate_for_row(row, ref_no)
    if docx_buf is None:
        return None, fname  # Return error message
    
    pdf_bytes = convert_docx_bytes_to_pdf(docx_buf.getvalue())
    pdf_buf = io.BytesIO(pdf_bytes)
    pdf_fname = fname.replace('.docx', '.pdf')
    
    return pdf_buf, pdf_fname


def _ref_and_date(doc, ref_no):
    today = datetime.now()
    thai_be_year = today.year + 543
    thai_months = ["", "มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                   "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]
    date_th = f"{today.day} {thai_months[today.month]} {thai_be_year}"
    
    _para(doc, f"ที่ อว 78.363 / {ref_no}", size=14, space_after=0, line_spacing=1.0)
    _para(doc, f"วันที่ {date_th}", size=14, space_after=4, line_spacing=1.0)


def _student_table(doc, rows):
    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    table.autofit = False  # Prevents Word/LibreOffice from overriding widths
    
    headers = ["ลำดับ", "รหัสนักศึกษา", "ชื่อ - นามสกุล", "ฝ่ายงาน / ตำแหน่ง"]
    # Total printable width = 16.5 cm
    widths = [1.2, 2.8, 6.0, 6.5]
    
    # Set width on column objects
    for i, w in enumerate(widths):
        table.columns[i].width = Cm(w)

    # Set header row
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.width = Cm(widths[i])
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.space_before = Pt(1)
        r = p.add_run(h)
        _set_font(r, bold=True, size=14)
        
    # Set data rows
    for idx, row in enumerate(rows, start=1):
        cells = table.add_row().cells
        values = [str(idx), row['student_id'], row['full_name'], row['dept_position']]
        for i, v in enumerate(values):
            cells[i].width = Cm(widths[i])
            p = cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i in (0, 1) else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.space_before = Pt(1)
            r = p.add_run(v)
            _set_font(r, size=14)
            
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _signature_block(doc):
    _para(doc, "ขอแสดงความนับถือ", align=WD_ALIGN_PARAGRAPH.CENTER, size=14, space_before=6, space_after=36)
    _para(doc, DEAN_TH, align=WD_ALIGN_PARAGRAPH.CENTER, size=14, space_after=0, line_spacing=1.0)
    _para(doc, DEAN_TITLE_TH, align=WD_ALIGN_PARAGRAPH.CENTER, size=14, space_after=0, line_spacing=1.0)
    _para(doc, UNIV_TH, align=WD_ALIGN_PARAGRAPH.CENTER, size=14, space_after=0, line_spacing=1.0)


def _full_name(row):
    return f"{row.get('Prefix','').strip()}{row.get('First Name','').strip()} {row.get('Last Name','').strip()}".strip()


def _course_label(row):
    """Returns cleaned course name with non-breaking spaces (\u00A0) to prevent
    Word/LibreOffice from over-expanding space characters around English course codes."""
    val = row.get('Please select the course for this internship', '').strip()
    cleaned_val = ' '.join(val.split())
    return cleaned_val.replace(' ', '\u00A0')


# ---------------------------------------------------------------- builders

def build_confirm_letter(row, ref_no="___"):
    """'Student Referral Letter' -> sent once the company has already agreed."""
    doc = _base_doc()
    _letterhead(doc, FACULTY_TH)
    _ref_and_date(doc, ref_no)
    
    company_name = row.get('Company/Organization Name for Internship', 'XXXX').strip()
    contact_person = row.get('Letter Addressed To (Company Contact Person)', 'XXXX').strip()
    position = row.get('Position', 'XXXX').strip()
    year_level = row.get('Year Level', 'X').strip()
    period = row.get('Internship Period', 'XX – XX').strip()
    course = _course_label(row)

    _para(doc, "เรื่อง  ขอส่งนักศึกษาเข้าฝึกงาน", bold=True, size=14, space_after=3)
    _para(doc, f"เรียน  คุณ {contact_person}", size=14, space_after=0, line_spacing=1.0)
    _para(doc, f"       ตำแหน่ง {position}", size=14, space_after=0, line_spacing=1.0)
    _para(doc, f"       บริษัท {company_name}", size=14, space_after=4, line_spacing=1.0)

    body = (
        f"ตามที่ บริษัท {company_name} ได้แจ้งความประสงค์ยินดีที่จะรับนักศึกษา"
        f"ระดับปริญญาตรี ชั้นปีที่ {year_level} หลักสูตรวิทยาศาสตรบัณฑิต สาขาวิชาวิทยาการและเทคโนโลยีดิจิทัล "
        f"ของคณะเทคโนโลยีสารสนเทศและการสื่อสาร มหาวิทยาลัยมหิดล เข้าร่วมฝึกงานกับหน่วยงานของท่าน ซึ่งเป็นส่วนหนึ่งของรายวิชา\u00A0"
        f"{course}\u00A0เพื่อเพิ่มพูนทักษะและนำความรู้ที่ได้ศึกษามาใช้ในการปฏิบัติงานจริง โดยมีระยะเวลาการฝึกงาน"
        f"ระหว่างวันที่ {period}"
    )
    _para(doc, body, align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=4)
    _para(doc, "ในการนี้ คณะฯ จึงใคร่ขอส่งนักศึกษาเข้ารับการฝึกงาน จำนวน 1 คน คือ",
          align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=4)

    _student_table(doc, [{
        'student_id': row.get('Student ID (e.g. 6587999)', '').strip(),
        'full_name': _full_name(row),
        'dept_position': f"{row.get('Department/Division for Internship','').strip()} / {row.get('Internship Position','').strip()}".strip(" /"),
    }])

    closing = (
        "ทั้งนี้ ในช่วงระหว่างการฝึกงาน อาจารย์นิเทศก์จากทางคณะฯ จะประสานเพื่อขออนุญาตเข้านิเทศนักศึกษา"
        "และเยี่ยมชมการฝึกงานของนักศึกษากับบริษัท ภายหลังสิ้นสุดการฝึกงาน ทางคณะฯ ขอความอนุเคราะห์ท่านทำแบบประเมินผล"
        "การฝึกงานนักศึกษา เพื่อนำข้อมูลไปพัฒนาปรับปรุงหลักสูตรของคณะฯ ต่อไป หากต้องการสอบถามข้อมูลเพิ่มเติม สามารถติดต่อ"
        f"{COURSE_CONTACT_TH}"
    )
    _para(doc, closing, align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=4)
    _para(doc, "จึงเรียนมาเพื่อโปรดทราบและขอขอบคุณมา ณ โอกาสนี้", indent_first=TAB_INDENT_CM, size=14, space_after=0)
    _signature_block(doc)
    return doc


def build_request_letter(row, ref_no="___"):
    """'Request Letter for Internship Placement' -> initial outreach to a company."""
    doc = _base_doc()
    _letterhead(doc, FACULTY_TH_ICT)
    _ref_and_date(doc, ref_no)
    
    company_name = row.get('Company/Organization Name for Internship', 'XXXXXX').strip()
    contact_person = row.get('Letter Addressed To (Company Contact Person)', 'XXXXX').strip()
    year_level = row.get('Year Level', '2/3').strip()
    period = row.get('Internship Period', '2 มิถุนายน - 25 กรกฎาคม 2568').strip()
    course = _course_label(row)

    _para(doc, "เรื่อง  ขอความอนุเคราะห์รับนักศึกษาเข้าฝึกงาน", bold=True, size=14, space_after=3)
    _para(doc, f"เรียน  คุณ {contact_person}", size=14, space_after=0, line_spacing=1.0)
    _para(doc, "       ผู้จัดการ", size=14, space_after=0, line_spacing=1.0)
    _para(doc, f"       บริษัท {company_name}", size=14, space_after=4, line_spacing=1.0)

    body = (
        f"เนื่องด้วยนักศึกษาระดับปริญญาตรี หลักสูตรวิทยาศาสตรบัณฑิต สาขาวิชาวิทยาการและเทคโนโลยีดิจิทัล (DST) "
        f"ชั้นปีที่ {year_level} ของคณะเทคโนโลยีสารสนเทศและการสื่อสาร (ICT) มหาวิทยาลัยมหิดล "
        f"มีความประสงค์จะขอเข้ารับการฝึกงาน ณ บริษัท {company_name} "
        f"ซึ่งเป็นส่วนหนึ่งของรายวิชา\u00A0{course}\u00A0เพื่อเพิ่มพูนทักษะและนำความรู้ที่ได้ศึกษามาใช้ในการปฏิบัติงานจริง"
    )
    _para(doc, body, align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=4)

    body2 = (
        "ในการนี้ คณะฯ ได้พิจารณาแล้วเห็นว่าหน่วยงานของท่านสามารถสร้างองค์ความรู้และประสบการณ์"
        "อันเป็นประโยชน์และให้หลักการการทำงานที่ดีแก่นักศึกษา คณะฯ จึงใคร่ขอความอนุเคราะห์จากท่านพิจารณา"
        f"รับนักศึกษาของคณะฯ เข้ารับการฝึกงาน โดยมีระยะเวลาการฝึกงานระหว่างวันที่ {period} "
        "จำนวน 1 คน คือ"
    )
    _para(doc, body2, align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=4)

    _student_table(doc, [{
        'student_id': row.get('Student ID (e.g. 6587999)', '').strip(),
        'full_name': _full_name(row),
        'dept_position': row.get('Internship Position', '').strip(),
    }])

    closing = f"ทั้งนี้ หากต้องการสอบถามข้อมูลเพิ่มเติม สามารถติดต่อ{COURSE_CONTACT_TH}"
    _para(doc, closing, align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=4)
    _para(doc, "จึงเรียนมาเพื่อโปรดพิจารณาให้ความอนุเคราะห์ จะขอบคุณยิ่ง", indent_first=TAB_INDENT_CM, size=14, space_after=0)
    _signature_block(doc)
    return doc


def generate_for_row(row, ref_no="___"):
    """Returns (BytesIO, filename) or (None, reason) if unsupported letter type."""
    letter_type = row.get('Letter Type', '').strip()
    kind = SUPPORTED_LETTER_TYPES.get(letter_type)
    if kind is None:
        return None, f"No template yet for letter type: '{letter_type}'"
    doc = build_confirm_letter(row, ref_no) if kind == "confirm" else build_request_letter(row, ref_no)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    safe_name = _full_name(row).replace(' ', '_') or row.get('Student ID (e.g. 6587999)', 'student').strip()
    fname = f"{kind}_{row.get('Student ID (e.g. 6587999)','').strip()}_{safe_name}.docx"
    return buf, fname


def extract_fields_from_docx(file_obj):
    """Reads a generated letter back and pulls out the fields worth comparing
    against the source Excel row."""
    import re
    doc = Document(file_obj)
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    fields = {'recipient_name': '', 'recipient_position': '', 'company': '',
              'student_id': '', 'student_name': '', 'dept_position': '', 'period': ''}

    for p in paras:
        if p.startswith('เรียน'):
            fields['recipient_name'] = p.replace('เรียน', '').replace('คุณ', '').strip()
        elif p.startswith('ตำแหน่ง'):
            fields['recipient_position'] = p.replace('ตำแหน่ง', '').strip()
        elif p.startswith('บริษัท'):
            fields['company'] = p.replace('บริษัท', '').strip()

    if doc.tables:
        table = doc.tables[0]
        if len(table.rows) > 1:
            cells = [c.text.strip() for c in table.rows[1].cells]
            if len(cells) > 1:
                fields['student_id'] = cells[1]
            if len(cells) > 2:
                fields['student_name'] = cells[2]
            if len(cells) > 3:
                fields['dept_position'] = cells[3]

    for p in paras:
        m = re.search(r'ระหว่างวันที่\s*(.+?)(?:\s*จำนวน|$)', p)
        if m:
            fields['period'] = m.group(1).strip()
            break

    return fields