# -*- coding: utf-8 -*-
"""
docgen.py
สร้างเอกสารหนังสือขอความอนุเคราะห์และหนังสือส่งตัวฝึกงาน (.docx) จากข้อมูลในแต่ละแถว
โดยจำลองเค้าโครงและรูปแบบของเอกสารทางการของ คณะ ICT มหาวิทยาลัยมหิดล:
  - Confirm_Letter_V1        -> ใช้สำหรับ "หนังสือส่งตัวนักศึกษาฝึกงาน"
  - Request_Intern_Letter_V1 -> ใช้สำหรับ "หนังสือขอความอนุเคราะห์รับนักศึกษาฝึกงาน"

มาตรฐานฟอนต์: TH Sarabun New 14pt
ระยะย่อหน้า: 1 แท็บ (1.27 ซม.)
"""
import io
import os
import re
import shutil
import subprocess
import tempfile
from datetime import datetime
from io import BytesIO
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx2pdf import convert
from docxtpl import DocxTemplate
from django.conf import settings
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from docx.oxml import OxmlElement

FONT_PATH = os.path.join(settings.BASE_DIR, 'letters', 'static', 'fonts', 'THSarabunNew.ttf')
pdfmetrics.registerFont(TTFont("THSarabun", FONT_PATH))
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

# ---------------------------------------------------------------- หัวข้อฟิลด์ฟอร์ม (ภาษาไทย)
H_LETTER_TYPE = ('ประเภทหนังสือที่ขอ', 'ประเภทหนังสือ', 'Letter Type')
H_COURSE = ('โปรดเลือกรายวิชาของการไปฝึกงาน', 'รายวิชา')
H_STUDENT_ID = ('รหัสนักศึกษา (เช่น 6587999)', 'รหัสนักศึกษา')
H_PREFIX = ('คำนำหน้าชื่อ',)
H_FIRST_NAME = ('ชื่อ (ภาษาไทย)', 'ชื่อ')
H_LAST_NAME = ('นามสกุล (ภาษาไทย)', 'นามสกุล')
H_YEAR_LEVEL = ('ชั้นปีที่',)
H_YEAR = H_YEAR_LEVEL  # ตัวแปร Alias เพื่อรองรับการนำเข้าจาก views.py
H_COMPANY = (
    'ชื่อหน่วยงาน/บริษัท ที่ต้องการฝึกงาน', 
    'Company/Organization Name for Internship',
    'Company',
    'Company Name',
    'บริษัท',
    'ชื่อบริษัท',
    'หน่วยงาน'
)
H_CONTACT_PERSON = ('โดยทำหนังสือถึงคุณ (สอบถามผู้ติดต่อฝั่งบริษัท)',)
H_CONTACT_POSITION = ('ตำแหน่ง')
H_DEPARTMENT = ('นักศึกษาเข้าฝึกงานในส่วนของฝ่ายงาน', )
H_INTERN_POSITION = ('ตำแหน่งของการฝึกงาน')
H_PERIOD = ('ช่วงเวลาของการฝึกงาน',)

# จับคู่ประเภทหนังสือภายใน docgen ('confirm'/'request') ไปยังค่าของ LetterTemplate.letter_type ('confirmation'/'request')
KIND_TO_TEMPLATE_TYPE = {'confirm': 'confirmation', 'request': 'request', 'extension': 'extension'}

SUPPORTED_LETTER_TYPES = {
    "หนังสือส่งตัวนักศึกษาฝึกงาน": "confirm",
    "หนังสือขอความอนุเคราะห์รับนักศึกษาฝึกงาน": "request",
    "หนังสือขอขยายระยะเวลาฝึกงาน": "extension",
    "Student Referral Letter": "confirm",
    "Request Letter for Internship Placement": "request",
    "หนังสือส่งตัว": "confirm",
    "หนังสือขอความอนุเคราะห์": "request",
    "ขอขยายระยะเวลาการฝึกงานของนักศึกษา": "extension",
}

# ฟิลด์ข้อมูลทั้งหมดที่สามารถใช้ใน Custom Template (Jinja2)
AVAILABLE_SOURCE_FIELDS = [
    ('ref_no', 'เลขที่หนังสือ (เช่น อว 78.363 / 001)'),
    ('date_th', 'วันที่ปัจจุบัน (รูปแบบภาษาไทย)'),
    ('student_id', 'รหัสนักศึกษา'),
    ('student_name', 'ชื่อ-นามสกุล นักศึกษา'),
    ('prefix', 'คำนำหน้าชื่อ'),
    ('first_name', 'ชื่อจริง'),
    ('last_name', 'นามสกุล'),
    ('year_level', 'ชั้นปีที่'),
    ('company', 'ชื่อหน่วยงาน/บริษัท'),
    ('contact_person', 'ชื่อผู้ติดต่อ'),
    ('contact_position', 'ตำแหน่งผู้ติดต่อ'),
    ('department', 'ฝ่ายงาน/แผนก'),
    ('internship_position', 'ตำแหน่งงานฝึกงาน'),
    ('dept_position', 'ฝ่ายงาน / ตำแหน่ง (รวมกัน)'),
    ('period', 'ช่วงเวลาการฝึกงาน'),
    ('course', 'รายวิชา'),
    ('faculty_name', 'ชื่อคณะ'),
    ('university_name', 'ชื่อมหาวิทยาลัย'),
    ('address_line1', 'ที่อยู่ บรรทัดที่ 1'),
    ('address_line2', 'ที่อยู่ บรรทัดที่ 2'),
    ('tel', 'เบอร์โทรศัพท์ / โทรสาร'),
    ('dean_name', 'ชื่อคณบดี'),
    ('dean_title', 'ตำแหน่งคณบดี'),
    ('course_contact', 'ข้อมูลติดต่ออาจารย์ผู้ดูแลรายวิชา'),
]


def apply_thai_font_fix(doc: Document, font_name: str = "TH Sarabun New") -> Document:
    """
    Forces TH Sarabun New + th-TH language tags across all document elements
    (paragraphs, tables, headers, footers, SDTs, styles) and removes Word Theme font
    overrides that trigger broken Thai character rendering in PDF engines.
    """
    theme_attrs = ["asciiTheme", "cstheme", "hAnsiTheme", "eastAsiaTheme"]

    xml_roots = [doc._element]
    for section in doc.sections:
        for container in [
            section.header, section.first_page_header,
            section.footer, section.first_page_footer
        ]:
            if container and getattr(container, "_element", None) is not None:
                xml_roots.append(container._element)

    # 1. Target every single <w:r> across all XML document trees
    for root in xml_roots:
        for r in root.xpath('.//w:r'):
            rPr = r.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                r.insert(0, rPr)

            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.append(rFonts)

            # Strip Theme Attributes
            for attr in theme_attrs:
                key = qn(f"w:{attr}")
                if key in rFonts.attrib:
                    del rFonts.attrib[key]

            # Assign explicit font family across all script engines
            rFonts.set(qn("w:ascii"), font_name)
            rFonts.set(qn("w:hAnsi"), font_name)
            rFonts.set(qn("w:cs"), font_name)
            rFonts.set(qn("w:eastAsia"), font_name)

            # Force Thai Language & Bidi tagging
            lang = rPr.find(qn("w:lang"))
            if lang is None:
                lang = OxmlElement("w:lang")
                rPr.append(lang)
            lang.set(qn("w:val"), "th-TH")
            lang.set(qn("w:bidi"), "th-TH")
            lang.set(qn("w:eastAsia"), "th-TH")

    # 2. Update styles.xml and default document fonts
    try:
        styles_root = doc.styles.element
        for rFonts in styles_root.xpath('.//w:rFonts'):
            for attr in theme_attrs:
                key = qn(f"w:{attr}")
                if key in rFonts.attrib:
                    del rFonts.attrib[key]
            rFonts.set(qn("w:ascii"), font_name)
            rFonts.set(qn("w:hAnsi"), font_name)
            rFonts.set(qn("w:cs"), font_name)
            rFonts.set(qn("w:eastAsia"), font_name)
    except Exception:
        pass

    return doc


def _get_libreoffice_bin():
    """ค้นหาไฟล์ Executable ของ LibreOffice บน Linux, macOS หรือ Windows"""
    soffice_bin = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_bin:
        mac_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        win_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
        if os.path.exists(mac_path):
            soffice_bin = mac_path
        elif os.path.exists(win_path):
            soffice_bin = win_path
    return soffice_bin


def _get_val(row, *keys):
    """ดึงข้อมูลและทำความสะอาดข้อความจาก dict ของข้อมูล โดยลองใช้ คีย์ ตามลำดับ"""
    for key in keys:
        if isinstance(keys[0], (list, tuple)) and not isinstance(key, str):
            for sub_k in key:
                res = _get_val(row, sub_k)
                if res:
                    return res
            continue

        val = row.get(key) if isinstance(row, dict) else getattr(row, key, None)
        if val is None:
            continue
        try:
            import pandas as pd
            if pd.isna(val):
                continue
        except (ImportError, TypeError, ValueError):
            pass

        if isinstance(val, float):
            if val.is_integer():
                val = int(val)

        val_str = str(val).strip()
        if val_str.endswith('.0') and val_str[:-2].isdigit():
            val_str = val_str[:-2]

        if val_str.lower() in ('nan', 'none', 'nat'):
            continue

        if val_str:
            return val_str
    return ''


def normalize_letter_type(raw_val):
    """แปลงข้อความประเภทหนังสือให้อยู่ในรูปแบบมาตรฐาน (confirm, request, extension)"""
    if not raw_val:
        return "request"

    val_str = str(raw_val).strip()
    if val_str in SUPPORTED_LETTER_TYPES:
        return SUPPORTED_LETTER_TYPES[val_str]

    val_lower = val_str.lower()
    if any(k in val_lower for k in ['referral', 'confirm', 'ส่งตัว', 'dispatch']):
        return "confirm"
    if any(k in val_lower for k in ['request', 'placement', 'ขอความอนุเคราะห์', 'อนุเคราะห์']):
        return "request"
    if any(k in val_lower for k in ['extension', 'ขยายระยะเวลา', 'ขยาย']):
        return "extension"

    return val_str


TAB_INDENT_CM = 1.27
FONT_CANDIDATES = ["TH Sarabun New"]


def _available_font_families():
    """คืนค่ารายการชื่อฟอนต์ที่ติดตั้งอยู่ในระบบผ่าน fc-list"""
    try:
        result = subprocess.run(
            ["fc-list", ":family"],
            capture_output=True, text=True, timeout=5,
        )
        return {line.strip() for line in result.stdout.splitlines() if line.strip()}
    except Exception:
        return set()


def _resolve_thai_font():
    """เลือกฟอนต์ภาษาไทยแรกที่พบจาก FONT_CANDIDATES หากไม่พบจะใช้ TH Sarabun New เป็นค่าเริ่มต้น"""
    families = _available_font_families()
    if not families:
        return FONT_CANDIDATES[0]
    for candidate in FONT_CANDIDATES:
        if any(candidate in fam for fam in families):
            return candidate
    return FONT_CANDIDATES[0]


THAI_FONT = _resolve_thai_font()

# TH Sarabun New is the visually "correct"/expected font for these official
# letters, but it appears to trigger a LibreOffice-specific PDF export bug:
# documents render perfectly on-screen in Writer, but headless PDF export
# corrupts complex-script glyph shaping (distinct Thai syllables collapse
# into repeated wrong characters). This happens even with an isolated
# LibreOffice profile, which rules out profile/cache corruption as the cause.
# Sarabun (the Google Fonts rebuild) is visually very close but doesn't share
# TH Sarabun New's legacy OpenType tables, so it sidesteps the bug. Used only
# for the PDF-bound document — DOCX output is unaffected and keeps THAI_FONT.
_families = _available_font_families()
PDF_SAFE_FONT = "Sarabun" if any(fam.split(':')[0] == "Sarabun" for fam in _families) else THAI_FONT


def _set_font(run, size=14, bold=False, color=None, name=None):
    """กำหนดรูปแบบฟอนต์พร้อมคุณสมบัติ Complex Script (CS) สำหรับข้อความภาษาไทย"""
    if name is None:
        name = THAI_FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = name

    rpr = run._element.get_or_add_rPr()

    rFonts = rpr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rpr.makeelement(qn('w:rFonts'), {})
        rpr.append(rFonts)
    rFonts.set(qn('w:ascii'), name)
    rFonts.set(qn('w:hAnsi'), name)
    rFonts.set(qn('w:cs'), name)
    rFonts.set(qn('w:eastAsia'), name)

    if bold:
        bCs = rpr.find(qn('w:bCs'))
        if bCs is None:
            bCs = rpr.makeelement(qn('w:bCs'), {})
            rpr.append(bCs)
        bCs.set(qn('w:val'), 'true')

    szCs = rpr.find(qn('w:szCs'))
    if szCs is None:
        szCs = rpr.makeelement(qn('w:szCs'), {})
        rpr.append(szCs)
    szCs.set(qn('w:val'), str(int(size * 2)))

    if color:
        run.font.color.rgb = RGBColor(*color)


def _para(doc, text="", align=WD_ALIGN_PARAGRAPH.LEFT, size=14, bold=False,
          space_after=3, space_before=0, line_spacing=1, indent_first=None):
    """สร้างย่อหน้าใหม่พร้อมตั้งค่ารูปแบบและฟอนต์ (ปรับ line_spacing เป็น 1.2 เพื่อป้องกันวรรณยุกต์ซ้อนโดนตัด)"""
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
    """สร้างโครงสร้างเอกสาร Word พื้นฐาน พร้อมตั้งค่าระยะขอบกระดาษ"""
    doc = Document()
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(1.2)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.0)
    
    style = doc.styles['Normal']
    style.font.name = THAI_FONT
    style.font.size = Pt(14)
    return doc


def _letterhead(doc, faculty_name):
    """สร้างส่วนหัวกระดาษของคณะและมหาวิทยาลัย"""
    if os.path.exists(LOGO_PATH):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_before = Pt(0)
        p_logo.paragraph_format.space_after = Pt(0)
        p_logo.paragraph_format.line_spacing = 1.0
        run_logo = p_logo.add_run()
        run_logo.add_picture(LOGO_PATH, width=Cm(2.2))

    _para(doc, faculty_name, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, size=14, space_before=0, space_after=0, line_spacing=1.15)
    _para(doc, UNIV_TH, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, size=14, space_before=0, space_after=0, line_spacing=1.15)
    _para(doc, ADDRESS_LINE1_TH, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, size=14, space_before=0, space_after=0, line_spacing=1.15)
    _para(doc, ADDRESS_LINE2_TH, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, size=14, space_before=0, space_after=0, line_spacing=1.15)
    _para(doc, TEL_TH, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, size=14, space_before=0, space_after=2, line_spacing=1.15)


def _ref_and_date(doc, ref_no):
    """สร้างส่วนเลขที่หนังสือและวันที่ (รูปแบบ พ.ศ.)"""
    today = datetime.now()
    thai_be_year = today.year + 543
    thai_months = ["", "มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                   "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]
    date_th = f"{today.day} {thai_months[today.month]} {thai_be_year}"
    
    _para(doc, f"ที่ อว 78.363 / {ref_no}", size=14, space_after=0, line_spacing=1.15)
    _para(doc, f"วันที่ {date_th}", size=14, space_after=3, line_spacing=1.15)


def _student_table(doc, rows):
    """สร้างตารางแสดงรายชื่อนักศึกษา"""
    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    table.autofit = False
    
    headers = ["ลำดับ", "รหัสนักศึกษา", "ชื่อ - นามสกุล", "ฝ่ายงาน / ตำแหน่ง"]
    widths = [1.2, 2.8, 6.0, 6.5]
    
    for i, w in enumerate(widths):
        table.columns[i].width = Cm(w)

    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.width = Cm(widths[i])
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.space_before = Pt(1)
        r = p.add_run(h)
        _set_font(r, bold=True, size=14)
        
    for idx, row in enumerate(rows, start=1):
        cells = table.add_row().cells
        values = [
            str(idx),
            row.get('student_id', ''),
            row.get('full_name', ''),
            row.get('dept_position', '')
        ]
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
    """สร้างส่วนลงนามท้ายเอกสาร"""
    _para(doc, "ขอแสดงความนับถือ", align=WD_ALIGN_PARAGRAPH.CENTER, size=14, space_before=4, space_after=36)
    _para(doc, DEAN_TH, align=WD_ALIGN_PARAGRAPH.CENTER, size=14, space_after=0, line_spacing=1.15)
    _para(doc, DEAN_TITLE_TH, align=WD_ALIGN_PARAGRAPH.CENTER, size=14, space_after=0, line_spacing=1.15)
    _para(doc, UNIV_TH, align=WD_ALIGN_PARAGRAPH.CENTER, size=14, space_after=0, line_spacing=1.15)


def _full_name(row):
    """ประกอบชื่อและนามสกุลเต็มของนักศึกษา"""
    if isinstance(row, dict):
        prefix = _get_val(row, *H_PREFIX)
        first_name = _get_val(row, *H_FIRST_NAME)
        last_name = _get_val(row, *H_LAST_NAME)
        return f"{prefix}{first_name} {last_name}".strip()
    return getattr(row, 'display_name', '')


def _course_label(row):
    """ดึงรหัสและชื่อรายวิชา"""
    if isinstance(row, dict):
        val = _get_val(row, *H_COURSE)
    else:
        val = getattr(row, 'course_id', '')
    cleaned_val = ' '.join(str(val).split())
    return cleaned_val.replace(' ', '\u00A0')


def build_confirm_letter(row, ref_no="___"):
    """สร้างเอกสารหนังสือส่งตัวนักศึกษาเข้าฝึกงาน"""
    doc = _base_doc()
    _letterhead(doc, FACULTY_TH)
    _ref_and_date(doc, ref_no)
    
    r_dict = row if isinstance(row, dict) else (getattr(row, 'data', {}) or {})
    company_name = _get_val(r_dict, *H_COMPANY) or 'XXXX'
    contact_person = _get_val(r_dict, *H_CONTACT_PERSON) or 'XXXX'
    position = _get_val(r_dict, *H_CONTACT_POSITION) or 'XXXX'
    year_level = _get_val(r_dict, *H_YEAR_LEVEL) or 'X'
    period = _get_val(r_dict, *H_PERIOD) or 'XX – XX'
    course = _course_label(r_dict)

    _para(doc, "เรื่อง  ขอส่งนักศึกษาเข้าฝึกงาน", bold=True, size=14, space_after=3)
    _para(doc, f"เรียน  คุณ {contact_person}", size=14, space_after=0, line_spacing=1.15)
    _para(doc, f"       ตำแหน่ง {position}", size=14, space_after=0, line_spacing=1.15)
    _para(doc, f"       บริษัท {company_name}", size=14, space_after=3, line_spacing=1.15)

    body = (
        f"ตามที่ บริษัท {company_name} ได้แจ้งความประสงค์ยินดีที่จะรับนักศึกษา"
        f"ระดับปริญญาตรี ชั้นปีที่ {year_level} หลักสูตรวิทยาศาสตรบัณฑิต สาขาวิชาวิทยาการและเทคโนโลยีดิจิทัล "
        f"ของคณะเทคโนโลยีสารสนเทศและการสื่อสาร มหาวิทยาลัยมหิดล เข้าร่วมฝึกงานกับหน่วยงานของท่าน ซึ่งเป็นส่วนหนึ่งของรายวิชา\u00A0"
        f"{course}\u00A0เพื่อเพิ่มพูนทักษะและนำความรู้ที่ได้ศึกษามาใช้ในการปฏิบัติงานจริง โดยมีระยะเวลาการฝึกงาน"
        f"ระหว่างวันที่ {period}"
    )
    _para(doc, body, align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=3)
    _para(doc, "ในการนี้ คณะฯ จึงใคร่ขอส่งนักศึกษาเข้ารับการฝึกงาน จำนวน 1 คน คือ",
          align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=3)

    _student_table(doc, [{
        'student_id': _get_val(r_dict, *H_STUDENT_ID) or getattr(row, 'student_id', ''),
        'full_name': _full_name(r_dict) or getattr(row, 'display_name', ''),
        'dept_position': f"{_get_val(r_dict, *H_DEPARTMENT)} / {_get_val(r_dict, *H_INTERN_POSITION)}".strip(" /"),
    }])

    closing = (
        "ทั้งนี้ ในช่วงระหว่างการฝึกงาน อาจารย์นิเทศก์จากทางคณะฯ จะประสานเพื่อขออนุญาตเข้านิเทศนักศึกษา"
        "และเยี่ยมชมการฝึกงานของนักศึกษากับบริษัท ภายหลังสิ้นสุดการฝึกงาน ทางคณะฯ ขอความอนุเคราะห์ท่านทำแบบประเมินผล"
        "การฝึกงานนักศึกษา เพื่อนำข้อมูลไปพัฒนาปรับปรุงหลักสูตรของคณะฯ ต่อไป หากต้องการสอบถามข้อมูลเพิ่มเติม สามารถติดต่อ"
        f"{COURSE_CONTACT_TH}"
    )
    _para(doc, closing, align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=3)
    _para(doc, "จึงเรียนมาเพื่อโปรดทราบและขอขอบคุณมา ณ โอกาสนี้", indent_first=TAB_INDENT_CM, size=14, space_after=0)
    _signature_block(doc)
    return doc


def build_request_letter(row, ref_no="___"):
    """สร้างเอกสารหนังสือขอความอนุเคราะห์รับนักศึกษาเข้าฝึกงาน"""
    doc = _base_doc()
    _letterhead(doc, FACULTY_TH_ICT)
    _ref_and_date(doc, ref_no)
    
    r_dict = row if isinstance(row, dict) else (getattr(row, 'data', {}) or {})
    company_name = _get_val(r_dict, *H_COMPANY) or 'XXXXXX'
    contact_person = _get_val(r_dict, *H_CONTACT_PERSON) or 'XXXXX'
    year_level = _get_val(r_dict, *H_YEAR_LEVEL) or '2/3'
    period = _get_val(r_dict, *H_PERIOD) or '2 มิถุนายน - 25 กรกฎาคม 2568'
    course = _course_label(r_dict)

    _para(doc, "เรื่อง  ขอความอนุเคราะห์รับนักศึกษาเข้าฝึกงาน", bold=True, size=14, space_after=3)
    _para(doc, f"เรียน  คุณ {contact_person}", size=14, space_after=0, line_spacing=1.15)
    _para(doc, "       ผู้จัดการ", size=14, space_after=0, line_spacing=1.15)
    _para(doc, f"       บริษัท {company_name}", size=14, space_after=3, line_spacing=1.15)

    body = (
        f"เนื่องด้วยนักศึกษาระดับปริญญาตรี หลักสูตรวิทยาศาสตรบัณฑิต สาขาวิชาวิทยาการและเทคโนโลยีดิจิทัล (DST) "
        f"ชั้นปีที่ {year_level} ของคณะเทคโนโลยีสารสนเทศและการสื่อสาร (ICT) มหาวิทยาลัยมหิดล "
        f"มีความประสงค์จะขอเข้ารับการฝึกงาน ณ บริษัท {company_name} "
        f"ซึ่งเป็นส่วนหนึ่งของรายวิชา\u00A0{course}\u00A0เพื่อเพิ่มพูนทักษะและนำความรู้ที่ได้ศึกษามาใช้ในการปฏิบัติงานจริง"
    )
    _para(doc, body, align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=3)

    body2 = (
        "ในการนี้ คณะฯ ได้พิจารณาแล้วเห็นว่าหน่วยงานของท่านสามารถสร้างองค์ความรู้และประสบการณ์"
        "อันเป็นประโยชน์และให้หลักการการทำงานที่ดีแก่นักศึกษา คณะฯ จึงใคร่ขอความอนุเคราะห์จากท่านพิจารณา"
        f"รับนักศึกษาของคณะฯ เข้ารับการฝึกงาน โดยมีระยะเวลาการฝึกงานระหว่างวันที่ {period} "
        "จำนวน 1 คน คือ"
    )
    _para(doc, body2, align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=3)

    _student_table(doc, [{
        'student_id': _get_val(r_dict, *H_STUDENT_ID) or getattr(row, 'student_id', ''),
        'full_name': _full_name(r_dict) or getattr(row, 'display_name', ''),
        'dept_position': _get_val(r_dict, *H_INTERN_POSITION),
    }])

    closing = f"ทั้งนี้ หากต้องการสอบถามข้อมูลเพิ่มเติม สามารถติดต่อ{COURSE_CONTACT_TH}"
    _para(doc, closing, align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=3)
    _para(doc, "จึงเรียนมาเพื่อโปรดพิจารณาให้ความอนุเคราะห์ จะขอบคุณยิ่ง", indent_first=TAB_INDENT_CM, size=14, space_after=0)
    _signature_block(doc)
    return doc


def build_extension_letter(row, ref_no="___"):
    """สร้างเอกสารหนังสือขอขยายระยะเวลาการฝึกงานของนักศึกษา"""
    doc = _base_doc()
    _letterhead(doc, FACULTY_TH_ICT)
    _ref_and_date(doc, ref_no)
    
    r_dict = row if isinstance(row, dict) else (getattr(row, 'data', {}) or {})
    company_name = _get_val(r_dict, *H_COMPANY) or 'XXXXXX'
    contact_person = _get_val(r_dict, *H_CONTACT_PERSON) or 'XXXXX'
    position = _get_val(r_dict, *H_CONTACT_POSITION) or 'ผู้จัดการ'
    year_level = _get_val(r_dict, *H_YEAR_LEVEL) or '2/3'
    period = _get_val(r_dict, *H_PERIOD) or 'XX – XX'
    course = _course_label(r_dict)

    _para(doc, "เรื่อง  ขอขยายระยะเวลาการฝึกงานของนักศึกษา", bold=True, size=14, space_after=3)
    _para(doc, f"เรียน  คุณ {contact_person}", size=14, space_after=0, line_spacing=1.15)
    _para(doc, f"       ตำแหน่ง {position}", size=14, space_after=0, line_spacing=1.15)
    _para(doc, f"       บริษัท {company_name}", size=14, space_after=3, line_spacing=1.15)

    body = (
        f"ตามที่ บริษัท {company_name} ได้ให้ความอนุเคราะห์รับนักศึกษาระดับปริญญาตรี หลักสูตรวิทยาศาสตรบัณฑิต "
        f"สาขาวิชาวิทยาการและเทคโนโลยีดิจิทัล (DST) ชั้นปีที่ {year_level} ของคณะเทคโนโลยีสารสนเทศและการสื่อสาร (ICT) "
        f"มหาวิทยาลัยมหิดล เข้ารับการฝึกงานในรายวิชา\u00A0{course}\u00A0นั้น"
    )
    _para(doc, body, align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=3)

    body2 = (
        f"เนื่องด้วยการปฏิบัติงานฝึกงานดังกล่าวจำเป็นต้องใช้ระยะเวลาในการดำเนินงานเพิ่มเติม เพื่อให้นักศึกษาได้เรียนรู้"
        f"และเพิ่มพูนทักษะการทำงานได้อย่างสมบูรณ์ คณะฯ จึงใคร่ขอความอนุเคราะห์จากท่านพิจารณาขยายระยะเวลาการฝึกงาน"
        f"ของนักศึกษาออกไป จนถึงวันที่ {period} จำนวน 1 คน คือ"
    )
    _para(doc, body2, align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=3)

    _student_table(doc, [{
        'student_id': _get_val(r_dict, *H_STUDENT_ID) or getattr(row, 'student_id', ''),
        'full_name': _full_name(r_dict) or getattr(row, 'display_name', ''),
        'dept_position': _get_val(r_dict, *H_INTERN_POSITION),
    }])

    closing = f"ทั้งนี้ หากต้องการสอบถามข้อมูลเพิ่มเติม สามารถติดต่อ{COURSE_CONTACT_TH}"
    _para(doc, closing, align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=3)
    _para(doc, "จึงเรียนมาเพื่อโปรดพิจารณาให้ความอนุเคราะห์ จะขอบคุณยิ่ง", indent_first=TAB_INDENT_CM, size=14, space_after=0)
    _signature_block(doc)
    return doc


def _build_template_context(row_or_record, kind, ref_no="___"):
    """สร้าง Context Data สำหรับการนำไปใส่ใน Template docxtpl"""
    r_dict = row_or_record if isinstance(row_or_record, dict) else (getattr(row_or_record, 'data', {}) or {})

    today = datetime.now()
    thai_be_year = today.year + 543
    thai_months = ["", "มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                   "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]
    date_th = f"{today.day} {thai_months[today.month]} {thai_be_year}"

    department = _get_val(r_dict, *H_DEPARTMENT)
    internship_position = _get_val(r_dict, *H_INTERN_POSITION)
    student_id = _get_val(r_dict, *H_STUDENT_ID) or getattr(row_or_record, 'student_id', '')
    student_name = _full_name(r_dict) or getattr(row_or_record, 'display_name', '')
    company = _get_val(r_dict, *H_COMPANY) or getattr(row_or_record, 'company', '')
    course = _course_label(r_dict)

    return {
        'ref_no': f"อว 78.363 / {ref_no}",
        'date_th': date_th,
        'faculty_name': FACULTY_TH if kind == "confirm" else FACULTY_TH_ICT,
        'university_name': UNIV_TH,
        'address_line1': ADDRESS_LINE1_TH,
        'address_line2': ADDRESS_LINE2_TH,
        'tel': TEL_TH,
        'dean_name': DEAN_TH,
        'dean_title': DEAN_TITLE_TH,
        'course_contact': COURSE_CONTACT_TH,
        'student_id': student_id,
        'student_name': student_name,
        'full_name': student_name,
        'prefix': _get_val(r_dict, *H_PREFIX),
        'first_name': _get_val(r_dict, *H_FIRST_NAME),
        'last_name': _get_val(r_dict, *H_LAST_NAME),
        'year_level': _get_val(r_dict, *H_YEAR_LEVEL),
        'company': company,
        'contact_person': _get_val(r_dict, *H_CONTACT_PERSON),
        'contact_position': _get_val(r_dict, *H_CONTACT_POSITION),
        'department': department,
        'internship_position': internship_position,
        'dept_position': f"{department} / {internship_position}".strip(" /"),
        'period': _get_val(r_dict, *H_PERIOD),
        'course': course,
        'course_id': course,
        'excel': r_dict,
    }


def _resolve_mapped_context(row_or_record, kind, ref_no, mappings):
    """สร้าง Context โดยอ้างอิงตามค่า Field Mapping ที่ตั้งค่าไว้"""
    base_context = _build_template_context(row_or_record, kind, ref_no)
    context = {}
    for mapping in mappings:
        if mapping.source_type == mapping.SOURCE_STATIC:
            context[mapping.placeholder] = mapping.source_value
        elif mapping.source_type == mapping.SOURCE_FIELD:
            context[mapping.placeholder] = base_context.get(mapping.source_value, '')
    return context


def render_with_template(template_obj, row_or_record, kind, ref_no="___"):
    """ประมวลผลไฟล์ LetterTemplate รวมเข้ากับข้อมูล Context"""
    mappings = list(template_obj.field_mappings.all()) if hasattr(template_obj, 'field_mappings') else []
    if mappings:
        context = _resolve_mapped_context(row_or_record, kind, ref_no, mappings)
    else:
        context = _build_template_context(row_or_record, kind, ref_no)

    template_obj.file.open('rb')
    try:
        doc = DocxTemplate(template_obj.file)
        doc.render(context)
        # Apply complex script Thai font fix to the rendered DocxTemplate
        apply_thai_font_fix(doc.docx, font_name=THAI_FONT)
    finally:
        template_obj.file.close()

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def _try_user_template(row_or_record, kind, ref_no="___"):
    """พยายามเรนเดอร์เอกสารจาก Custom Template ที่ผู้ใช้เคยอัปโหลดไว้ หากไม่มีจะคืนค่า None"""
    template_type = KIND_TO_TEMPLATE_TYPE.get(kind, kind)
    try:
        from .models import LetterTemplate
        template_obj = LetterTemplate.objects.get(letter_type=template_type, is_active=True)
    except Exception:
        return None

    try:
        return render_with_template(template_obj, row_or_record, kind, ref_no)
    except Exception:
        return None


def generate_for_row(row, ref_no="___"):
    """
    สร้างเอกสาร Word Buffer สำหรับข้อมูลนักศึกษา 1 แถว
    คืนค่าเป็น: (BytesIO_buffer, filename) เมื่อสำเร็จ หรือ (None, error_message) เมื่อล้มเหลว
    """
    raw_type = _get_val(row, *H_LETTER_TYPE) if isinstance(row, dict) else getattr(row, 'letter_type', '')
    kind = normalize_letter_type(raw_type)

    if not kind:
        return None, f"ไม่พบหรือประเภทหนังสือไม่ถูกต้อง ('{raw_type}')"

    r_dict = row if isinstance(row, dict) else (getattr(row, 'data', {}) or {})
    student_id = _get_val(r_dict, *H_STUDENT_ID) or getattr(row, 'student_id', '')
    safe_id = re.sub(r'[^A-Za-z0-9]', '', str(student_id)) or "noid"

    student_name = _full_name(row)
    safe_name = student_name.strip().replace(' ', '_') or "student"

    custom_buf = _try_user_template(row, kind, ref_no)
    if custom_buf:
        return custom_buf, f"{kind}_{safe_id}_{safe_name}.docx"

    if kind == "confirm":
        doc = build_confirm_letter(row, ref_no)
    elif kind == "extension":
        doc = build_extension_letter(row, ref_no)
    elif kind == "request":
        doc = build_request_letter(row, ref_no)
    else:
        return None, f"ไม่รองรับประเภทหนังสือ '{kind}'"

    # Force Thai font properties and CTL language tags across the generated document
    apply_thai_font_fix(doc, font_name=THAI_FONT)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf, f"{kind}_{safe_id}_{safe_name}.docx"


def generate_letter_buffer(student_record, letter_type='request'):
    """ฟังก์ชันช่วยสำหรับเรียกใช้งานใน Views ที่ต้องการรับเฉพาะไฟล์ Buffer"""
    buf, _ = generate_for_row(student_record)
    return buf


def convert_docx_bytes_to_pdf(docx_bytes: bytes) -> bytes:
    """Converts DOCX binary data to PDF binary data using headless LibreOffice on Linux."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        input_docx_path = os.path.join(tmp_dir, "temp_doc.docx")
        
        # Write input DOCX to temp file
        with open(input_docx_path, "wb") as f:
            f.write(docx_bytes)

        # Execute headless LibreOffice CLI conversion
        command = [
            "libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            tmp_dir,
            input_docx_path,
        ]
        
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)
        
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice error: {result.stderr.decode('utf-8')}")

        output_pdf_path = os.path.join(tmp_dir, "temp_doc.pdf")
        if not os.path.exists(output_pdf_path):
            raise FileNotFoundError("LibreOffice completed without throwing an error, but no PDF was generated.")

        # Read converted PDF bytes
        with open(output_pdf_path, "rb") as f:
            return f.read()


def convert_docx_to_pdf(docx_path: str, output_dir: str, timeout: int = 30) -> str:
    """แปลงไฟล์ DOCX ตามที่อยู่ path ที่ระบุไปเป็น PDF"""
    docx_path = Path(docx_path).resolve()
    output_dir = Path(output_dir).resolve()

    soffice_bin = _get_libreoffice_bin()
    if not soffice_bin:
        raise RuntimeError("ไม่พบโปรแกรม LibreOffice ในระบบ")

    env = os.environ.copy()
    env["LANG"] = "th_TH.UTF-8"
    env["LC_ALL"] = "th_TH.UTF-8"

    # Use a fresh, isolated LibreOffice profile for this conversion. Sharing the
    # default profile with another running/stale LibreOffice instance can break
    # Complex Text Layout shaping (which Thai needs) even though the same file
    # renders correctly when opened normally in LibreOffice Writer.
    with tempfile.TemporaryDirectory() as lo_profile_dir:
        profile_uri = Path(lo_profile_dir).resolve().as_uri()

        cmd = [
            soffice_bin,
            "--headless",
            "--invisible",
            "--nologo",
            f"-env:UserInstallation={profile_uri}",
            "--convert-to", "pdf:writer_pdf_Export",
            "--outdir", str(output_dir),
            str(docx_path)
        ]

        try:
            subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                env=env,
                timeout=timeout,
                check=True
            )
        except subprocess.TimeoutExpired:
            raise TimeoutError(f"การแปลง PDF หมดเวลาหลังจาก {timeout} วินาที")
        except subprocess.CalledProcessError as e:
            raise RuntimeError(f"การแปลง PDF ล้มเหลว: {e.stderr.decode('utf-8', errors='ignore')}")

    pdf_path = output_dir / f"{docx_path.stem}.pdf"
    if not pdf_path.exists():
        raise FileNotFoundError(f"ไม่พบไฟล์ PDF ที่ปลายทาง {pdf_path}")

    return str(pdf_path)


def generate_pdf_for_row(row, ref_no="___"):
    """สร้างเอกสาร PDF จากข้อมูลนักศึกษา 1 แถว คืนค่าเป็น (pdf_buffer, filename)"""
    res = generate_for_row(row, ref_no)
    if not res or res[0] is None:
        raise ValueError(f"การสร้างไฟล์ DOCX ล้มเหลว: {res[1] if res else 'ไม่ทราบสาเหตุ'}")

    docx_buf, fname = res
    pdf_fname = fname.rsplit('.', 1)[0] + '.pdf'

    # Re-open the generated docx and swap its font to PDF_SAFE_FONT just for
    # this conversion. TH Sarabun New (THAI_FONT / the docx default) appears
    # to trigger a LibreOffice PDF-export bug that corrupts Thai glyph shaping
    # even though the docx itself renders fine — see PDF_SAFE_FONT comment.
    # The original docx_buf (used for direct .docx download) is left untouched.
    docx_buf.seek(0)
    pdf_input_doc = Document(docx_buf)
    apply_thai_font_fix(pdf_input_doc, font_name=PDF_SAFE_FONT)
    pdf_input_buf = BytesIO()
    pdf_input_doc.save(pdf_input_buf)
    pdf_input_buf.seek(0)

    soffice_bin = _get_libreoffice_bin()
    if not soffice_bin:
        raise RuntimeError("ไม่พบโปรแกรม LibreOffice ในระบบ (โปรดติดตั้ง LibreOffice เพื่อใช้ฟังก์ชันแปลง PDF)")

    env = os.environ.copy()
    env["LANG"] = "th_TH.UTF-8"
    env["LC_ALL"] = "th_TH.UTF-8"

    with tempfile.TemporaryDirectory() as tmpdir:
        input_docx = os.path.join(tmpdir, "document.docx")

        with open(input_docx, "wb") as f:
            f.write(pdf_input_buf.getvalue())

        # Isolated profile per conversion (see convert_docx_to_pdf for why this
        # matters for Thai text specifically) — nested inside tmpdir so it's
        # cleaned up automatically along with everything else here.
        lo_profile_dir = os.path.join(tmpdir, "lo_profile")
        os.makedirs(lo_profile_dir, exist_ok=True)
        profile_uri = Path(lo_profile_dir).resolve().as_uri()

        cmd = [
            soffice_bin,
            "--headless",
            "--invisible",
            "--nologo",
            f"-env:UserInstallation={profile_uri}",
            "--convert-to", "pdf:writer_pdf_Export",
            "--outdir", tmpdir,
            input_docx
        ]

        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env)
        if result.returncode != 0:
            raise RuntimeError(f"การแปลงไฟล์ PDF ล้มเหลว: {result.stderr.decode('utf-8', errors='ignore')}")

        output_pdf = os.path.join(tmpdir, "document.pdf")
        if not os.path.exists(output_pdf):
            raise FileNotFoundError(f"ไม่พบไฟล์ PDF ที่แปลงสำเร็จ Log: {result.stdout.decode('utf-8', errors='ignore')}")

        pdf_buf = BytesIO()
        with open(output_pdf, "rb") as f:
            pdf_buf.write(f.read())
        pdf_buf.seek(0)

    return pdf_buf, pdf_fname


def extract_fields_from_docx(file_obj):
    """สกัดกวาดหาตัวแปร Jinja2 หรือข้อความฟิลด์ในไฟล์ docx"""
    try:
        doc = DocxTemplate(file_obj)
        vars_list = list(doc.get_undeclared_template_variables())
        if vars_list:
            return {'template_variables': vars_list}
    except Exception:
        pass

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