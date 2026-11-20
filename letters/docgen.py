# -*- coding: utf-8 -*-
"""
docgen.py
สร้างเอกสารหนังสือขอความอนุเคราะห์และหนังสือส่งตัวฝึกงาน (.docx / .pdf) จากข้อมูลในแต่ละแถว
โดยจำลองเค้าโครงและรูปแบบของเอกสารทางการของ คณะ ICT มหาวิทยาลัยมหิดล:
  - Confirm_Letter_V1        -> ใช้สำหรับ "หนังสือส่งตัวนักศึกษาฝึกงาน"
  - Request_Intern_Letter_V1 -> ใช้สำหรับ "หนังสือขอความอนุเคราะห์รับนักศึกษาฝึกงาน"
  - Extension_Letter         -> ใช้สำหรับ "หนังสือขอขยายระยะเวลาฝึกงาน"

มาตรฐานฟอนต์: TH Sarabun New 14pt
ระยะย่อหน้า: 1 แท็บ (1.27 ซม.)
"""
import contextvars
import io
from django.core.files.base import ContentFile
import logging
import os
import re
import shutil
import subprocess
import tempfile
from datetime import datetime
from io import BytesIO
from pathlib import Path

from django.conf import settings
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx2pdf import convert
from docxtpl import DocxTemplate
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

logger = logging.getLogger(__name__)

try:
    FONT_PATH = os.path.join(settings.BASE_DIR, 'letters', 'static', 'fonts', 'THSarabunNew.ttf')
    if os.path.exists(FONT_PATH):
        pdfmetrics.registerFont(TTFont("THSarabun", FONT_PATH))
except Exception as e:
    logger.warning("Could not register THSarabun TTFont: %s", e)

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
H_CONTACT_POSITION = ('ตำแหน่ง',)
H_DEPARTMENT = ('นักศึกษาเข้าฝึกงานในส่วนของฝ่ายงาน', )
H_INTERN_POSITION = ('ตำแหน่งของการฝึกงาน',)
H_PERIOD = ('ช่วงเวลาของการฝึกงาน',)

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
    ('instructor_name', 'ชื่ออาจารย์ผู้ดูแลรายวิชา'),
    ('instructor_email', 'อีเมล์อาจารย์ผู้ดูแลรายวิชา'),
    ('instructor_phone', 'โทรศัพท์อาจารย์ผู้ดูแลรายวิชา'),
]


def normalize_letter_type(raw_val):
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


def resolve_letter_type(raw_val):
    raw = str(raw_val or '').strip()
    if raw:
        try:
            from .models import LetterTypeMapping
            m = LetterTypeMapping.objects.select_related('letter_type').filter(raw_text=raw).first()
            if m:
                return m.letter_type.slug
        except Exception:
            logger.exception("LetterTypeMapping lookup failed")
    return normalize_letter_type(raw)


TAB_INDENT_CM = 1.27

THAI_MONTHS = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
               "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]


def _compose_contact(name, email, phone):
    line = f"อาจารย์ผู้ดูแลรายวิชา {name}".strip()
    if email:
        line += f" อีเมล์ {email}"
    if phone:
        line += (" หรือ" if email else "") + f" โทรศัพท์ {phone}"
    return line


DEFAULT_INSTRUCTOR = ("อาจารย์ผศ. ดร.ศิริเพ็ญ พงษ์ไพเชฐ", "siripen.pon@mahidol.ac.th", "083-0610496")


def _contact_parts(r_dict):
    """(name, email, phone) of the instructor of the subject named in this row (from the Subjects page).

    A subject without instructor details uses the default instructor.
    """
    if isinstance(r_dict, dict) and r_dict.get('__tag_mode__'):          # built-in letter -> template with tags
        return '{{ instructor_name }}', '{{ instructor_email }}', '{{ instructor_phone }}'
    try:
        from .models import Course
        raw = re.sub(r'[^A-Za-z0-9]', '', _get_val(r_dict, *H_COURSE)).upper()
        if raw:
            for c in Course.objects.exclude(instructor_name='', instructor_email='', instructor_phone=''):
                code = re.sub(r'[^A-Za-z0-9]', '', c.code).upper()
                if code and code in raw:
                    return c.instructor_name, c.instructor_email, c.instructor_phone
    except Exception:
        pass
    return DEFAULT_INSTRUCTOR


def _contact_line(r_dict):
    """The sentence printed at the end of the built-in letters, built from the instructor's name, email and phone."""
    return _compose_contact(*_contact_parts(r_dict))


_CFG_SCOPE = contextvars.ContextVar('letter_cfg', default=None)


def _cfg():
    """The letter settings. While one letter is being built they are read from the database once
    (building a letter asks for them once per paragraph, which meant about 20 identical queries per letter)."""
    scoped = _CFG_SCOPE.get()
    return scoped if scoped is not None else _load_cfg()


def _load_cfg():
    """Editable letter settings (reference prefix, date, dean, supervisor, spacing).

    Falls back to the built-in values when the database is not available, so
    docgen keeps working outside Django (tests, scripts).
    """
    cfg = {
        'ref_prefix': 'อว 78.363',
        'letter_date': None,
        'dean_name': DEAN_TH,
        'dean_title': DEAN_TITLE_TH,
        'line_spacing': 1.15,
    }
    try:
        from .models import LetterSettings
        obj = LetterSettings.load()
        cfg.update({
            'ref_prefix': obj.ref_prefix or cfg['ref_prefix'],
            'letter_date': obj.letter_date,
            'dean_name': obj.dean_name or cfg['dean_name'],
            'dean_title': obj.dean_title or cfg['dean_title'],
            'line_spacing': float(obj.line_spacing or cfg['line_spacing']),
        })
    except Exception:
        pass
    return cfg


def _letter_date_th(cfg=None):
    cfg = cfg or _cfg()
    d = cfg['letter_date'] or datetime.now()
    return f"{d.day} {THAI_MONTHS[d.month]} {d.year + 543}"
FONT_CANDIDATES = ["TH Sarabun New"]


def _available_font_families():
    try:
        result = subprocess.run(
            ["fc-list", ":family"],
            capture_output=True, text=True, timeout=5,
        )
        return {line.strip() for line in result.stdout.splitlines() if line.strip()}
    except Exception:
        return set()


def _resolve_thai_font():
    families = _available_font_families()
    if not families:
        return FONT_CANDIDATES[0]
    for candidate in FONT_CANDIDATES:
        if any(candidate in fam for fam in families):
            return candidate
    return FONT_CANDIDATES[0]


THAI_FONT = _resolve_thai_font()

_families = _available_font_families()
PDF_SAFE_FONT = "Sarabun" if any(fam.split(':')[0] == "Sarabun" for fam in _families) else THAI_FONT


_RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
              'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden',
              'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
              'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout',
              'specVanish', 'oMath']


def _rpr_child(rpr, name):
    """Find or create <w:name> inside <w:rPr> at its schema position.

    Word is strict about child order; elements appended in the wrong place can
    make it open the file as 'unreadable' or ignore the formatting.
    """
    el = rpr.find(qn('w:' + name))
    if el is not None:
        return el
    el = OxmlElement('w:' + name)
    idx = _RPR_ORDER.index(name)
    later = set(_RPR_ORDER[idx + 1:])
    for child in rpr:
        if child.tag.split('}')[-1] in later:
            child.addprevious(el)
            return el
    rpr.append(el)
    return el


def apply_thai_font_fix(doc: Document, font_name: str = "TH Sarabun New") -> Document:
    theme_attrs = ["asciiTheme", "cstheme", "hAnsiTheme", "eastAsiaTheme"]

    xml_roots = [doc._element]
    for section in doc.sections:
        for container in [
            section.header, section.first_page_header,
            section.footer, section.first_page_footer
        ]:
            if container and getattr(container, "_element", None) is not None:
                xml_roots.append(container._element)

    for root in xml_roots:
        for r in root.xpath('.//w:r'):
            rPr = r.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                r.insert(0, rPr)

            rFonts = _rpr_child(rPr, 'rFonts')

            for attr in theme_attrs:
                key = qn(f"w:{attr}")
                if key in rFonts.attrib:
                    del rFonts.attrib[key]

            rFonts.set(qn("w:ascii"), font_name)
            rFonts.set(qn("w:hAnsi"), font_name)
            rFonts.set(qn("w:cs"), font_name)
            rFonts.set(qn("w:eastAsia"), font_name)

            lang = _rpr_child(rPr, 'lang')
            lang.set(qn("w:val"), "th-TH")
            lang.set(qn("w:bidi"), "th-TH")
            lang.set(qn("w:eastAsia"), "th-TH")

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


def extract_fields_from_docx(file_or_path_or_bytes):
    if isinstance(file_or_path_or_bytes, bytes):
        file_obj = io.BytesIO(file_or_path_or_bytes)
    elif isinstance(file_or_path_or_bytes, (str, Path)):
        file_obj = str(file_or_path_or_bytes)
    elif hasattr(file_or_path_or_bytes, 'read'):
        content = file_or_path_or_bytes.read()
        if hasattr(file_or_path_or_bytes, 'seek'):
            file_or_path_or_bytes.seek(0)
        file_obj = io.BytesIO(content)
    else:
        file_obj = file_or_path_or_bytes

    try:
        doc = DocxTemplate(file_obj)
        try:
            vars_set = doc.get_undeclared_variables()
        except AttributeError:
            vars_set = doc.get_undeclared_template_variables()
        return {'template_variables': sorted(list(vars_set))}
    except Exception as e:
        logger.exception("Error extracting fields from docx template: %s", e)
        return {'template_variables': []}


def _get_libreoffice_bin():
    soffice_bin = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_bin:
        mac_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        win_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
        if os.path.exists(mac_path):
            soffice_bin = mac_path
        elif os.path.exists(win_path):
            soffice_bin = win_path
    return soffice_bin


def _norm_header(h):
    """Ignore stray spaces and pandas' duplicate-column suffix (.1, .2)."""
    h = re.sub(r'\.\d+$', '', str(h).strip())
    return re.sub(r'\s+', '', h).lower()


def _fuzzy_row_value(row, key):
    """Find a value by header ignoring spaces / '.N' duplicate suffix.
    Returns the first non-empty value among matching columns."""
    target = _norm_header(key)
    for k, v in row.items():
        if _norm_header(k) == target:
            if v is not None and str(v).strip() not in ('', 'nan', 'None', 'NaT'):
                return v
    return None


def _get_val(row, *keys):
    for key in keys:
        if isinstance(keys[0], (list, tuple)) and not isinstance(key, str):
            for sub_k in key:
                res = _get_val(row, sub_k)
                if res:
                    return res
            continue

        val = row.get(key) if isinstance(row, dict) else getattr(row, key, None)
        if (val is None or str(val).strip() == '') and isinstance(row, dict):
            val = _fuzzy_row_value(row, key)
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


def _set_font(run, size=14, bold=False, color=None, name=None):
    if name is None:
        name = THAI_FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = name

    rpr = run._element.get_or_add_rPr()

    rFonts = _rpr_child(rpr, 'rFonts')
    rFonts.set(qn('w:ascii'), name)
    rFonts.set(qn('w:hAnsi'), name)
    rFonts.set(qn('w:cs'), name)
    rFonts.set(qn('w:eastAsia'), name)

    if bold:
        _rpr_child(rpr, 'bCs').set(qn('w:val'), 'true')

    _rpr_child(rpr, 'szCs').set(qn('w:val'), str(int(size * 2)))

    if color:
        run.font.color.rgb = RGBColor(*color)


def _para(doc, text="", align=WD_ALIGN_PARAGRAPH.LEFT, size=14, bold=False,
          space_after=3, space_before=0, line_spacing=1.2, indent_first=None):
    p = doc.add_paragraph()
    if align == WD_ALIGN_PARAGRAPH.LEFT and indent_first is not None:
        align = WD_ALIGN_PARAGRAPH.THAI_JUSTIFY      # body text: even left and right margins
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.line_spacing = _cfg()['line_spacing']     # one setting for the whole letter
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
    section.top_margin = Cm(1.2)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.0)
    
    style = doc.styles['Normal']
    style.font.name = THAI_FONT
    style.font.size = Pt(14)
    rpr = style.element.get_or_add_rPr()
    fonts = _rpr_child(rpr, 'rFonts')
    for attr in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
        fonts.set(qn('w:' + attr), THAI_FONT)
    _rpr_child(rpr, 'szCs').set(qn('w:val'), '28')
    _rpr_child(rpr, 'lang').set(qn('w:bidi'), 'th-TH')
    style.paragraph_format.line_spacing = _cfg()['line_spacing']
    return doc


def _letterhead(doc, faculty_name):
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
    cfg = _cfg()
    _para(doc, f"ที่ {cfg['ref_prefix']} / {ref_no}", size=14, space_after=0)
    _para(doc, f"วันที่ {_letter_date_th(cfg)}", size=14, space_after=3)


def _student_table(doc, rows):
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
    _para(doc, "ขอแสดงความนับถือ", align=WD_ALIGN_PARAGRAPH.CENTER, size=14, space_before=4, space_after=40)
    cfg = _cfg()
    _para(doc, cfg['dean_name'], align=WD_ALIGN_PARAGRAPH.CENTER, size=14, space_after=0, line_spacing=1.15)
    _para(doc, cfg['dean_title'], align=WD_ALIGN_PARAGRAPH.CENTER, size=14, space_after=0, line_spacing=1.15)
    _para(doc, UNIV_TH, align=WD_ALIGN_PARAGRAPH.CENTER, size=14, space_after=0, line_spacing=1.15)


def _full_name(row):
    if isinstance(row, dict):
        prefix = _get_val(row, *H_PREFIX)
        first_name = _get_val(row, *H_FIRST_NAME)
        last_name = _get_val(row, *H_LAST_NAME)
        return f"{prefix}{first_name} {last_name}".strip()
    return getattr(row, 'display_name', '')


def _course_label(row):
    if isinstance(row, dict):
        val = _get_val(row, *H_COURSE)
    else:
        val = getattr(row, 'course_id', '')
    cleaned_val = ' '.join(str(val).split())
    return cleaned_val


def build_confirm_letter(row, ref_no="___"):
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
        f"ของคณะเทคโนโลยีสารสนเทศและการสื่อสาร มหาวิทยาลัยมหิดล เข้าร่วมฝึกงานกับหน่วยงานของท่าน ซึ่งเป็นส่วนหนึ่งของรายวิชา "
        f"{course} เพื่อเพิ่มพูนทักษะและนำความรู้ที่ได้ศึกษามาใช้ในการปฏิบัติงานจริง โดยมีระยะเวลาการฝึกงาน"
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
        f"{_contact_line(r_dict)}"
    )
    _para(doc, closing, align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=3)
    _para(doc, "จึงเรียนมาเพื่อโปรดทราบและขอขอบคุณมา ณ โอกาสนี้", indent_first=TAB_INDENT_CM, size=14, space_after=0)
    _signature_block(doc)
    return doc


def build_request_letter(row, ref_no="___"):
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
        f"ซึ่งเป็นส่วนหนึ่งของรายวิชา {course} เพื่อเพิ่มพูนทักษะและนำความรู้ที่ได้ศึกษามาใช้ในการปฏิบัติงานจริง"
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

    closing = f"ทั้งนี้ หากต้องการสอบถามข้อมูลเพิ่มเติม สามารถติดต่อ{_contact_line(r_dict)}"
    _para(doc, closing, align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=3)
    _para(doc, "จึงเรียนมาเพื่อโปรดพิจารณาให้ความอนุเคราะห์ จะขอบคุณยิ่ง", indent_first=TAB_INDENT_CM, size=14, space_after=0)
    _signature_block(doc)
    return doc


def build_extension_letter(row, ref_no="___"):
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
        f"มหาวิทยาลัยมหิดล เข้ารับการฝึกงานในรายวิชา {course} นั้น"
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

    closing = f"ทั้งนี้ หากต้องการสอบถามข้อมูลเพิ่มเติม สามารถติดต่อ{_contact_line(r_dict)}"
    _para(doc, closing, align=WD_ALIGN_PARAGRAPH.LEFT, indent_first=TAB_INDENT_CM, size=14, space_after=3)
    _para(doc, "จึงเรียนมาเพื่อโปรดพิจารณาให้ความอนุเคราะห์ จะขอบคุณยิ่ง", indent_first=TAB_INDENT_CM, size=14, space_after=0)
    _signature_block(doc)
    return doc


BUILTIN_TITLES = {                      # keyed by the template's letter type (see KIND_TO_TEMPLATE_TYPE)
    'confirmation': 'หนังสือขอส่งตัวฝึกงาน (กรณีที่บริษัทรับเข้าฝึกงานแล้ว)',
    'request': 'หนังสือขอความอนุเคราะห์รับนักศึกษาเข้าฝึกงาน (กรณียังไม่ได้รับเข้าฝึกงาน)',
    'extension': 'หนังสือขอขยายระยะเวลาการฝึกงาน',
}


def build_tagged_letter(kind):
    """The built-in letter with {{ tags }} where the student's data goes, as .docx bytes.

    It is the same layout the system draws by itself, so it can be opened in the
    editor (or Word), changed, and saved as an ordinary template version.
    """
    row = {
        H_COMPANY[0]: '{{ company }}', H_CONTACT_PERSON[0]: '{{ contact_person }}',
        H_CONTACT_POSITION[0]: '{{ contact_position }}', H_YEAR_LEVEL[0]: '{{ year_level }}',
        H_PERIOD[0]: '{{ period }}', H_COURSE[0]: '{{ course }}', H_STUDENT_ID[0]: '{{ student_id }}',
        H_FIRST_NAME[0]: '{{ student_name }}', H_PREFIX[0]: '', H_LAST_NAME[0]: '',
        H_DEPARTMENT[0]: '{{ department }}', H_INTERN_POSITION[0]: '{{ internship_position }}',
        '__tag_mode__': True,
    }
    kind = {v: k for k, v in KIND_TO_TEMPLATE_TYPE.items()}.get(kind, kind)
    builder = {'confirm': build_confirm_letter, 'request': build_request_letter,
               'extension': build_extension_letter}[kind]
    doc = builder(row, ref_no='___')
    for para in doc.paragraphs:                     # reference number and date are filled per letter
        if para.text.startswith('ที่ ') and para.runs:
            para.runs[0].text = 'ที่ {{ ref_no }}'
            for r in para.runs[1:]:
                r.text = ''
        elif para.text.startswith('วันที่ ') and para.runs:
            para.runs[0].text = 'วันที่ {{ date_th }}'
            for r in para.runs[1:]:
                r.text = ''
    apply_thai_font_fix(doc, font_name=THAI_FONT)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _build_template_context(row_or_record, kind, ref_no="___"):
    """
    Builds context dictionary for Jinja2 template rendering.
    Combines pre-defined system fields and automatically merges all raw uploaded Excel/CSV columns.
    """
    r_dict = row_or_record if isinstance(row_or_record, dict) else (getattr(row_or_record, 'data', {}) or {})

    cfg = _cfg()
    date_th = _letter_date_th(cfg)

    department = _get_val(r_dict, *H_DEPARTMENT)
    internship_position = _get_val(r_dict, *H_INTERN_POSITION)
    student_id = _get_val(r_dict, *H_STUDENT_ID) or getattr(row_or_record, 'student_id', '')
    student_name = _full_name(r_dict) or getattr(row_or_record, 'display_name', '')
    company = _get_val(r_dict, *H_COMPANY) or getattr(row_or_record, 'company', '')
    course = _course_label(r_dict)

    context = {
        'ref_no': f"{cfg['ref_prefix']} / {ref_no}",
        'date_th': date_th,
        'faculty_name': FACULTY_TH if kind == "confirm" else FACULTY_TH_ICT,
        'university_name': UNIV_TH,
        'address_line1': ADDRESS_LINE1_TH,
        'address_line2': ADDRESS_LINE2_TH,
        'tel': TEL_TH,
        'dean_name': cfg['dean_name'],
        'dean_title': cfg['dean_title'],
        'instructor_name': _contact_parts(r_dict)[0],
        'instructor_email': _contact_parts(r_dict)[1],
        'instructor_phone': _contact_parts(r_dict)[2],
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

    # Automatically merge all custom raw uploaded columns from Excel/CSV
    if isinstance(r_dict, dict):
        for raw_k, raw_v in r_dict.items():
            if raw_k and raw_k not in context and raw_v is not None:
                context[raw_k] = str(raw_v).strip()

    return context


def missing_placeholders(mappings, base_context):
    """Placeholders of a template that cannot be filled for one record."""
    missing = []
    for m in mappings:
        if m.source_type == m.SOURCE_UNMAPPED:
            missing.append(m.placeholder)
        elif m.source_type == m.SOURCE_FIELD:
            val = base_context.get(m.source_value)
            if not str(val or '').strip():
                missing.append(m.placeholder)
    return missing


def _resolve_mapped_context(row_or_record, kind, ref_no, mappings):
    """Resolves mapped context, including new user-defined Excel variables like 'nickname'."""
    base_context = _build_template_context(row_or_record, kind, ref_no)
    context = {}
    r_dict = row_or_record if isinstance(row_or_record, dict) else (getattr(row_or_record, 'data', {}) or {})

    for mapping in mappings:
        if mapping.source_type == mapping.SOURCE_STATIC:
            context[mapping.placeholder] = mapping.source_value
        elif mapping.source_type == mapping.SOURCE_FIELD:
            mapped_key = mapping.source_value.strip()

            # 1. Look in base_context
            val = base_context.get(mapped_key)

            # 2. Direct lookup in raw Excel row dictionary
            if val is None or not str(val).strip():
                val = r_dict.get(mapped_key, '')

            # 3. Case-insensitive fallback (e.g., 'nickname' matches 'Nickname' or 'NICKNAME' in Excel)
            if not str(val or '').strip():
                for excel_header, excel_val in r_dict.items():
                    if str(excel_header).strip().lower() == mapped_key.lower():
                        val = excel_val
                        break

            context[mapping.placeholder] = str(val or '').strip()

    return context

def apply_logo(doc, logo_bytes):
    """Swap the first picture found (header first, then body) for logo_bytes.

    Keeps the picture's width and position and re-computes its height from
    the new image, so a different logo does not look stretched.
    Returns True when a picture was replaced.
    """
    parts = []
    for section in doc.sections:
        if not section.header.is_linked_to_previous:
            parts.append(section.header.part)
    parts.append(doc.part)

    for part in parts:
        blips = part.element.xpath('.//a:blip')
        if not blips:
            continue
        blip = blips[0]
        rid, img = part.get_or_add_image(BytesIO(logo_bytes))
        blip.set(qn('r:embed'), rid)

        if img.px_width:
            ratio = img.px_height / img.px_width
            for drawing in blip.xpath('ancestor::w:drawing[1]'):
                for ext in drawing.xpath('.//wp:extent | .//a:ext[@cx]'):
                    cx = int(ext.get('cx'))
                    ext.set('cy', str(int(cx * ratio)))
        return True
    return False


def bake_logo_into_template(template_obj, logo):
    """Put the chosen logo into the template file itself.

    The logo is part of the template, not something added to every letter.
    Done once when the logo is chosen, so generating a letter is only
    "copy the template, fill in the values".
    Returns True when a picture was replaced.
    """
    if not (logo and logo.image):
        return False
    logo.image.open('rb')
    try:
        logo_bytes = logo.image.read()
    finally:
        logo.image.close()
    template_obj.file.open('rb')
    try:
        doc = Document(BytesIO(template_obj.file.read()))
    finally:
        template_obj.file.close()
    if not apply_logo(doc, logo_bytes):
        return False
    buf = io.BytesIO()
    doc.save(buf)
    name = os.path.basename(template_obj.file.name)
    template_obj.file.save(name, ContentFile(buf.getvalue()), save=False)
    return True


def render_with_template(template_obj, row_or_record, kind, ref_no="___"):
    """Copy the template and fill in the values with docxtpl - nothing else.

    Fonts, justification, line spacing and the logo all live in the template
    file, so the letter comes out exactly as the template looks in Word.
    """
    mappings = list(template_obj.field_mappings.all()) if hasattr(template_obj, 'field_mappings') else []
    if mappings:
        context = _resolve_mapped_context(row_or_record, kind, ref_no, mappings)
    else:
        context = _build_template_context(row_or_record, kind, ref_no)

    template_obj.file.open('rb')
    try:
        doc = DocxTemplate(BytesIO(template_obj.file.read()))   # in-memory copy; the stored template is never touched
    finally:
        template_obj.file.close()
    doc.render(context)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def year_of(semester):
    """'2026/1' -> '2026'"""
    return re.split(r'[/\s]', str(semester or '').strip())[0]


def pick_template(kind, semester=''):
    """The active template of this letter type (None = use the built-in layout)."""
    from .models import LetterTemplate
    template_type = KIND_TO_TEMPLATE_TYPE.get(kind, kind)
    return (LetterTemplate.objects.filter(letter_type__in={template_type, kind}, is_active=True)
            .order_by('-version').first())


def _try_user_template(row_or_record, kind, ref_no="___", semester=''):
    try:
        template_obj = pick_template(kind, semester)
        if template_obj is None:
            return None
        return render_with_template(template_obj, row_or_record, kind, ref_no)
    except Exception:
        logger.exception("Template rendering failed")
        return None


def generate_for_row(*args, **kwargs):
    """Build one letter. The letter settings are loaded once for the whole letter."""
    token = _CFG_SCOPE.set(_load_cfg())
    try:
        return _generate_for_row(*args, **kwargs)
    finally:
        _CFG_SCOPE.reset(token)


def _generate_for_row(row, ref_no="___", kind_override=None, template_obj=None, force_builtin=False, semester=''):
    if kind_override:
        kind = normalize_letter_type(kind_override)
    else:
        raw_type = _get_val(row, *H_LETTER_TYPE) if isinstance(row, dict) else getattr(row, 'letter_type', '')
        kind = resolve_letter_type(raw_type)

    if not kind:
        return None, f"ไม่พบหรือประเภทหนังสือไม่ถูกต้อง"

    r_dict = row if isinstance(row, dict) else (getattr(row, 'data', {}) or {})
    student_id = _get_val(r_dict, *H_STUDENT_ID) or getattr(row, 'student_id', '')
    safe_id = re.sub(r'[^A-Za-z0-9]', '', str(student_id)) or "noid"

    student_name = _full_name(row)
    safe_name = student_name.strip().replace(' ', '_') or "student"

    if template_obj is not None:
        try:
            custom_buf = render_with_template(template_obj, row, kind, ref_no)
        except Exception as e:
            logger.exception("Template #%s failed to render", template_obj.id)
            return None, f"Template '{template_obj.name}' failed to render: {e}"
    elif force_builtin:
        custom_buf = None
    else:
        custom_buf = _try_user_template(row, kind, ref_no, semester)

    if custom_buf:
        return custom_buf, f"{kind}_{safe_id}_{safe_name}.docx"

    # Built-in layout fallback
    if kind == "confirm":
        doc = build_confirm_letter(row, ref_no=ref_no)
    elif kind == "request":
        doc = build_request_letter(row, ref_no=ref_no)
    elif kind == "extension":
        doc = build_extension_letter(row, ref_no=ref_no)
    else:
        return None, f"ไม่รองรับประเภทหนังสือ: '{kind}'"

    apply_thai_font_fix(doc, font_name=THAI_FONT)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    filename = f"{kind}_{safe_id}_{safe_name}.docx"
    return buf, filename


def convert_docx_bytes_to_pdf(docx_bytes: bytes) -> bytes:
    soffice_bin = _get_libreoffice_bin()
    if not soffice_bin:
        raise RuntimeError("ไม่พบโปรแกรม LibreOffice ในระบบ")

    env = os.environ.copy()
    env["LANG"] = "th_TH.UTF-8"
    env["LC_ALL"] = "th_TH.UTF-8"

    with tempfile.TemporaryDirectory() as tmp_dir:
        input_docx_path = os.path.join(tmp_dir, "temp_doc.docx")
        with open(input_docx_path, "wb") as f:
            f.write(docx_bytes)

        profile_uri = Path(os.path.join(tmp_dir, "lo_profile")).resolve().as_uri()
        command = [
            soffice_bin, "--headless", "--invisible", "--nologo", "--norestore",
            f"-env:UserInstallation={profile_uri}",
            "--convert-to", "pdf:writer_pdf_Export",
            "--outdir", tmp_dir, input_docx_path,
        ]
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                env=env, timeout=60, check=False)
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice error: {result.stderr.decode('utf-8', errors='ignore')}")

        output_pdf_path = os.path.join(tmp_dir, "temp_doc.pdf")
        if not os.path.exists(output_pdf_path):
            raise FileNotFoundError("LibreOffice finished but no PDF was generated.")
        with open(output_pdf_path, "rb") as f:
            return f.read()


def convert_docx_to_pdf(input_path_or_bytes, output_path=None):
    if isinstance(input_path_or_bytes, (bytes, io.BytesIO)):
        pdf_bytes = convert_docx_bytes_to_pdf(input_path_or_bytes)
        if output_path:
            with open(output_path, "wb") as f:
                f.write(pdf_bytes)
            return output_path
        return pdf_bytes
    else:
        with open(input_path_or_bytes, "rb") as f:
            docx_bytes = f.read()
        pdf_bytes = convert_docx_bytes_to_pdf(docx_bytes)
        if output_path:
            with open(output_path, "wb") as f:
                f.write(pdf_bytes)
            return output_path
        return pdf_bytes


def generate_pdf_for_row(row, ref_no="___"):
    doc_buf, docx_filename = generate_for_row(row, ref_no=ref_no)
    if not doc_buf:
        return None, docx_filename

    try:
        pdf_bytes = convert_docx_bytes_to_pdf(doc_buf)
        pdf_filename = os.path.splitext(docx_filename)[0] + ".pdf"
        return io.BytesIO(pdf_bytes), pdf_filename
    except Exception as e:
        logger.exception("Failed to generate PDF for row")
        return None, f"เกิดข้อผิดพลาดในการแปลงไฟล์ PDF: {str(e)}"