# -*- coding: utf-8 -*-
import io
import logging
import os
import re
import subprocess
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path

from docx import Document
import pandas as pd

from django.conf import settings
from django.contrib import messages
from django.contrib.auth import login, logout, update_session_auth_hash
from django.contrib.auth.forms import AuthenticationForm, PasswordChangeForm
from django.contrib.auth.models import User
from django.core.files.base import ContentFile
from django.core.mail import EmailMessage
from django.http import FileResponse, Http404, HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils import timezone
from django.views.decorators.clickjacking import xframe_options_sameorigin
from django.views.decorators.http import require_POST

from .decorators import admin_required, get_role, staff_required
from datetime import datetime
from .docgen import (
    bake_logo_into_template,
    pick_template,
    AVAILABLE_SOURCE_FIELDS,
    H_COMPANY,
    H_CONTACT_PERSON,
    H_CONTACT_POSITION,
    H_COURSE,
    H_FIRST_NAME,
    H_INTERN_POSITION,
    H_LAST_NAME,
    H_LETTER_TYPE,
    H_PERIOD,
    H_PREFIX,
    H_STUDENT_ID,
    H_YEAR,
    PDF_SAFE_FONT,
    SUPPORTED_LETTER_TYPES,
    _build_template_context,
    _get_libreoffice_bin,
    apply_thai_font_fix,
    convert_docx_bytes_to_pdf,
    extract_fields_from_docx,
    generate_for_row,
    generate_pdf_for_row,
    missing_placeholders,
    normalize_letter_type,
    render_with_template,
    resolve_letter_type,
)
from .models import (
    AuditLog,
    EmailTemplate,
    GeneratedDocument,
    GenerationBatch,
    LetterTemplate,
    LetterType,
    Course,
    LetterTypeMapping,
    Logo,
    StudentLetter,
    StudentRecord,
    TemplateFieldMapping,
    UploadedBatch,
)

logger = logging.getLogger(__name__)

IMPORTANT_FIELDS = ['student_id', 'display_name', 'company', 'contractor', 'internship_position', 'period', 'year']

raw_priority_headers = [
    H_STUDENT_ID, H_LETTER_TYPE, H_COURSE, H_COMPANY,
    H_CONTACT_PERSON, H_CONTACT_POSITION, H_INTERN_POSITION,
    H_PERIOD, H_PREFIX, H_FIRST_NAME, H_LAST_NAME, H_YEAR
]

PRIORITY_HEADERS = set()
for item in raw_priority_headers:
    if isinstance(item, (tuple, list, set)):
        PRIORITY_HEADERS.update(item)
    else:
        PRIORITY_HEADERS.add(item)

ENGLISH_PRIORITY_KEYS = {
    'Letter Type',
    'Student ID (e.g. 6587999)',
    'Prefix',
    'First Name',
    'Last Name',
    'Company/Organization Name for Internship',
    'Letter Addressed To (Company Contact Person)',
    'Position',
    'Internship Position',
    'Internship Period',
    'Year',
    'ชั้นปีที่',
}

ALL_PRIORITY_KEYS = PRIORITY_HEADERS.union(ENGLISH_PRIORITY_KEYS)
EXCLUDE_KEYS = {'id', 'batch', 'row_index', 'supported', 'missing', 'batch_id'}

THAI_MONTHS = {
    'มกราคม': 1, 'ม.ค': 1, 'กุมภาพันธ์': 2, 'ก.พ': 2, 'มีนาคม': 3, 'มี.ค': 3,
    'เมษายน': 4, 'เม.ย': 4, 'พฤษภาคม': 5, 'พ.ค': 5, 'มิถุนายน': 6, 'มิ.ย': 6,
    'กรกฎาคม': 7, 'ก.ค': 7, 'สิงหาคม': 8, 'ส.ค': 8, 'กันยายน': 9, 'ก.ย': 9,
    'ตุลาคม': 10, 'ต.ค': 10, 'พฤศจิกายน': 11, 'พ.ย': 11, 'ธันวาคม': 12, 'ธ.ค': 12,
}

BUILTIN_KINDS = {'confirm', 'request', 'extension'}


# ==========================================
# HELPER FUNCTIONS
# ==========================================

def _pick(data, keys):
    for k in keys:
        v = data.get(k)
        if v and str(v).strip():
            return str(v).strip()
    return ''


def _field_ready(field):
    if not field:
        return False
    try:
        return field.storage.exists(field.name)
    except Exception:
        return False


def log_activity(request, action, details=""):
    if not request.user or not request.user.is_authenticated:
        return

    x_forwarded_for = request.META.get('HTTP_X_FORWARDED_FOR')
    ip = x_forwarded_for.split(',')[0].strip() if x_forwarded_for else request.META.get('REMOTE_ADDR')

    request._activity_logged = True            # the middleware then skips its generic line
    AuditLog.objects.create(
        user=request.user,
        action=action,
        details=details,
        ip_address=ip
    )


def _load_table(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith('.csv'):
        return pd.read_csv(uploaded_file, dtype=str).fillna('')
    return pd.read_excel(uploaded_file, dtype=str).fillna('')


def _clean_str(s):
    if not s:
        return ''
    s = str(s).lower()
    prefixes = ['mr.', 'mr', 'mrs.', 'mrs', 'ms.', 'ms', 'miss', 'dr.', 'dr', 'นาย', 'นาง', 'นางสาว']
    for p in prefixes:
        if s.startswith(p):
            s = s[len(p):]
    return re.sub(r'[^a-zA-Z0-9\u0e00-\u0e7f]', '', s)


def get_student_email(record):
    data = record.data or {}

    def norm(s):
        return re.sub(r'[^a-zA-Z0-9\u0e00-\u0e7f]', '', str(s)).lower()

    student_keys = ['studentprimaryemail', 'studentemail', 'studentmail', 'อีเมลนักศึกษา', 'student']
    company_keys = ['company', 'contractor', 'contact', 'hr', 'supervisor', 'บริษัท', 'ผู้ติดต่อ', 'ผู้ประสานงาน']

    for k, v in data.items():
        if not v or '@' not in str(v):
            continue
        nk = norm(k)
        if any(sk in nk for sk in student_keys) and not any(ck in nk for ck in company_keys):
            return str(v).strip()

    for k, v in data.items():
        if not v or '@' not in str(v):
            continue
        val_str = str(v).strip().lower()
        nk = norm(k)
        if any(dom in val_str for dom in ['@student', '@u.mahidol', '@mail.mahidol', '.ac.th', '.edu']):
            if not any(ck in nk for ck in company_keys):
                return str(v).strip()

    for k, v in data.items():
        if not v or '@' not in str(v):
            continue
        nk = norm(k)
        if not any(ck in nk for ck in company_keys):
            return str(v).strip()

    return ''


def _get_letter_type_keywords(letter_type_raw):
    category = normalize_letter_type(letter_type_raw)
    keywords = set()

    if category == "request":
        keywords.update([
            _clean_str('Request Letter for Internship Placement'),
            _clean_str('หนังสือขอความอนุเคราะห์'),
            _clean_str('ขอความอนุเคราะห์'),
            'request', 'req'
        ])
    elif category == "confirm":
        keywords.update([
            _clean_str('Student Referral Letter'),
            _clean_str('หนังสือส่งตัวนักศึกษาเข้าฝึกงาน'),
            _clean_str('หนังสือส่งตัว'),
            'referral', 'confirm', 'send', 'dispatch'
        ])
    elif category == "extension":
        keywords.update([
            _clean_str('หนังสือขอขยายระยะเวลาฝึกงาน'),
            _clean_str('ขยายระยะเวลา'),
            'extension', 'ext'
        ])
    else:
        keywords.add(_clean_str(letter_type_raw))

    return {k for k in keywords if k}


def _match_and_save_pdf(records, filename, file_content):
    clean_filename = _clean_str(filename)
    candidates = []

    for record in records:
        clean_id = _clean_str(record.student_id)
        clean_name = _clean_str(record.display_name)

        matched = (clean_id and clean_id in clean_filename) or \
                  (clean_name and len(clean_name) >= 3 and clean_name in clean_filename)
        if matched:
            candidates.append(record)

    if not candidates:
        return False

    if len(candidates) == 1:
        record = candidates[0]
    else:
        type_matches = [
            r for r in candidates
            if _get_letter_type_keywords(r.letter_type) & {
                clean_filename[i:j]
                for i in range(len(clean_filename)) for j in range(i + 1, len(clean_filename) + 1)
            }
        ]
        if len(type_matches) == 1:
            record = type_matches[0]
        else:
            return False

    doc = GeneratedDocument.objects.filter(record=record).last()
    if not doc:
        ref_no = str(2600 + record.row_index)
        doc = GeneratedDocument.objects.create(
            record=record,
            ref_no=ref_no,
            status=GeneratedDocument.STATUS_GENERATED,
            file_name=filename,
        )
    else:
        doc.status = GeneratedDocument.STATUS_GENERATED

    doc.signed_at = timezone.now()
    doc.signed_file.save(os.path.basename(filename), file_content, save=True)
    return True


def _norm_header(h):
    h = re.sub(r'\.\d+$', '', str(h).strip())
    return re.sub(r'\s+', '', h).lower()


def _get_val(row, *keys):
    for key in keys:
        val = row.get(key)
        if (val is None or str(val).strip() in ('', 'nan')) and hasattr(row, 'items'):
            t = _norm_header(key)
            for k, v in row.items():
                if _norm_header(k) == t and v is not None and str(v).strip() not in ('', 'nan', 'None', 'NaT'):
                    val = v
                    break
        if val is None or pd.isna(val):
            continue
        if isinstance(val, float) and val.is_integer():
            val = int(val)

        val_str = str(val).strip()
        if val_str.endswith('.0') and val_str[:-2].isdigit():
            val_str = val_str[:-2]
        if val_str.lower() in ('nan', 'none', 'nat'):
            continue
        if val_str:
            return val_str
    return ''


def _parse_thai_date(date_str):
    """'1 มิ.ย. 2569', '1มิถุนายน 69', 'กรกฏาคม' spelling variants. Returns (day, month, year|None)."""
    date_str = (date_str or '').strip()
    m = re.match(r'(\d{1,2})\s*([^\d]+?)\.?\s*(\d{2,4})?$', date_str)
    if not m:
        return None

    day = int(m.group(1))
    month_raw = m.group(2).strip().rstrip('.').replace('ฏ', 'ฎ')   # กรกฏาคม is a common misspelling
    month_raw = re.sub(r'\s+', '', month_raw)
    month = THAI_MONTHS.get(month_raw)
    if not month:
        return None

    year_raw = m.group(3)
    if year_raw:
        y = int(year_raw)
        if y < 100:            # 69 -> 2569
            y += 2500
        year = y - 543 if y > 2400 else y
    else:
        year = None
    return day, month, year


def _get_semester_from_period(period_str, default='2026/1'):
    if not period_str:
        return default

    range_match = re.search(r'\((.*?)\)', period_str)
    range_text = range_match.group(1) if range_match else period_str

    parts = range_text.split('-')
    if len(parts) != 2:
        return default

    end = _parse_thai_date(parts[1])
    if not end or end[2] is None:
        return default
    end_year = end[2]

    start = _parse_thai_date(parts[0])
    if start and start[1] is not None:
        start_month = start[1]
    else:
        start_month = end[1]

    if start_month in (8, 9, 10, 11, 12):
        sem = 1
    elif start_month in (1, 2, 3, 4, 5):
        sem = 2
    else:
        sem = 3

    return f"{end_year}/{sem}"


def _check_year_eligibility(record):
    if not record.period:
        return False
    m = re.search(r'ปี\s*(\d+)\s*(?:-\s*(\d+))?\s*เท่านั้น', record.period)
    if not m:
        return False

    start = int(m.group(1))
    end = int(m.group(2)) if m.group(2) else start
    allowed = set(range(start, end + 1))

    digits = re.sub(r'\D', '', record.year or '')
    if not digits:
        return None

    return int(digits) not in allowed


def _chosen_template(request, record, templates_by_id):
    """-> (template_obj, force_builtin). No choice submitted = automatic (active template)."""
    key = f'template_{record.id}'
    if key not in request.POST:
        return None, False
    val = request.POST.get(key, '').strip()
    if not val:
        return None, True
    return templates_by_id.get(val), False


# ==========================================
# AUTHENTICATION & PROFILE VIEWS
# ==========================================

def login_view(request):
    if request.user.is_authenticated:
        return redirect('letters:index')

    if request.method == 'POST':
        form = AuthenticationForm(request, data=request.POST)
        if form.is_valid():
            user = form.get_user()
            if not user.is_staff:
                messages.error(request, "This account does not have staff permissions.")
                return render(request, 'letters/login.html', {'form': form})
            profile = getattr(user, 'profile', None) if hasattr(user, 'profile') else None
            if profile is not None and profile.expired:
                messages.error(request, "This account has expired. Ask an admin to extend it.")
                return render(request, 'letters/login.html', {'form': form})
            login(request, user)
            log_activity(request, 'LOGIN', f"{user.username} signed in")
            request.session.set_expiry(1800)
            next_url = request.GET.get('next')
            return redirect(next_url if next_url else 'letters:index')
        else:
            messages.error(request, "Invalid username or password.")
    else:
        form = AuthenticationForm()
    return render(request, 'letters/login.html', {'form': form})


def logout_view(request):
    if request.user.is_authenticated:
        log_activity(request, 'LOGOUT', f"{request.user.username} signed out")
    logout(request)
    return redirect('letters:login')


@staff_required
def profile_view(request):
    password_form = PasswordChangeForm(request.user)

    if request.method == 'POST':
        action = request.POST.get('action')

        if action == 'update_info':
            request.user.email = request.POST.get('email', '').strip()
            request.user.first_name = request.POST.get('first_name', '').strip()
            request.user.last_name = request.POST.get('last_name', '').strip()
            request.user.save()
            messages.success(request, "Profile updated successfully.")
            return redirect('letters:profile')

        elif action == 'change_password':
            password_form = PasswordChangeForm(request.user, request.POST)
            if password_form.is_valid():
                user = password_form.save()
                update_session_auth_hash(request, user)
                messages.success(request, "Password updated successfully.")
                return redirect('letters:profile')
            else:
                messages.error(request, "Please fix the password errors below.")

    return render(request, 'letters/profile.html', {'password_form': password_form})


# ==========================================
# CORE SYSTEM VIEWS
# ==========================================

@staff_required
def index(request):
    recent = UploadedBatch.objects.order_by('-uploaded_at')[:5]
    return render(request, 'letters/index.html', {'recent': recent})


@staff_required
def upload(request):
    # status_views is imported here (not at the top) because it imports this module
    from .status_views import (_bring_in_students, _check_records, _drop_stash, _load_stash,
                               _stash_upload, _unlisted_students)

    token = request.POST.get('token', '') if request.method == 'POST' else ''

    if token:                                    # second step: answer to "are you sure?"
        if request.POST.get('action') == 'cancel':
            _drop_stash(request, token)
            messages.info(request, "Upload cancelled. Nothing was saved.")
            return redirect('letters:index')
        f, info = _load_stash(request, token, 'requests')
        if f is None:
            messages.error(request, "The pending upload expired. Please choose the file again.")
            return redirect('letters:index')
        file_name = info['name']
    elif request.method == 'POST' and 'datafile' in request.FILES:
        f = request.FILES['datafile']
        file_name = f.name
    else:
        messages.error(request, "Please select a valid file to upload.")
        return redirect('letters:index')

    try:
        df = _load_table(f)
    except Exception as e:
        messages.error(request, f"Failed to parse file: {e}")
        if token:
            _drop_stash(request, token)
        return redirect('letters:index')

    records = []
    in_memory_counts = {}
    skipped_rows = []          # Excel row numbers that have no student ID
    skipped_info = []          # the same rows with what can be shown about them

    def _skipped(row, n):
        skipped_rows.append(n)
        skipped_info.append({
            'n': n, 'student_id': '',
            'name': f"{_get_val(row, *H_PREFIX)} {_get_val(row, *H_FIRST_NAME)} {_get_val(row, *H_LAST_NAME)}".strip(),
            'year': _get_val(row, *H_YEAR), 'subject': _get_val(row, *H_COURSE),
            'company': _get_val(row, *H_COMPANY), 'letter_type': _get_val(row, *H_LETTER_TYPE)})

    for i, row in enumerate(df.to_dict(orient='records')):
        s_id_raw = _get_val(row, *H_STUDENT_ID)
        if not s_id_raw:
            _skipped(row, i + 2)
            continue

        s_id = str(s_id_raw).replace('.0', '').strip()
        if not s_id:
            _skipped(row, i + 2)
            continue

        l_type = _get_val(row, *H_LETTER_TYPE)
        course = _get_val(row, *H_COURSE)
        year_val = _get_val(row, *H_YEAR)
        company_name = _get_val(row, *H_COMPANY)
        contractor_name = _get_val(row, *H_CONTACT_PERSON)
        contractor_pos = _get_val(row, *H_CONTACT_POSITION)
        intern_pos = _get_val(row, *H_INTERN_POSITION)
        period_val = _get_val(row, *H_PERIOD)
        sem_val = _get_semester_from_period(period_val)

        prefix = _get_val(row, *H_PREFIX)
        first_name = _get_val(row, *H_FIRST_NAME)
        last_name = _get_val(row, *H_LAST_NAME)
        full_name = f"{prefix} {first_name} {last_name}".strip()

        combo_key = (s_id, l_type, course, sem_val)

        if combo_key not in in_memory_counts:
            db_count = StudentRecord.objects.filter(
                student_id=s_id,
                letter_type=l_type,
                course_id=course,
                semester=sem_val
            ).count()
            in_memory_counts[combo_key] = db_count

        in_memory_counts[combo_key] += 1
        times_val = in_memory_counts[combo_key]

        records.append(StudentRecord(
            batch=None,                          # set once the upload is accepted
            row_index=i,
            student_id=s_id,
            display_name=full_name,
            letter_type=l_type,
            course_id=course,
            semester=sem_val,
            times=times_val,
            company=company_name,
            contractor=contractor_name,
            contractor_position=contractor_pos,
            internship_position=intern_pos,
            year=year_val,
            period=period_val,
            data=row,
        ))

    if not records:
        if token:
            _drop_stash(request, token)
        messages.error(request, "No valid student records found. Please check column header names.")
        return redirect('letters:index')

    # ---- first step: show every row and let the staff choose which ones to import ----
    if not token:
        unlisted = _unlisted_students(records)
        new_ids = set(unlisted)
        rows = []
        for x in _check_records(records):
            r = x['record']
            reasons = [{'level': i['level'], 'text': i['text']} for i in x['issues'] if i['code'] != 'no_student']
            is_new = (r.student_id or '').strip() in new_ids
            if is_new:
                reasons.insert(0, {'level': 'warn', 'text': "Not in the student list (for example an older-year student). "
                                                            "Tick to add this student to the system."})
            rows.append({'n': r.row_index + 2, 'pick': r.row_index, 'student_id': r.student_id,
                         'name': r.display_name, 'year': r.year, 'subject': r.course_id, 'company': r.company,
                         'letter_type': r.letter_type, 'status': 'check' if reasons else 'ready',
                         'reasons': reasons, 'new_student': is_new})
        for sk in skipped_info:
            rows.append({**sk, 'pick': '', 'status': 'err', 'new_student': False,
                         'reasons': [{'level': 'error', 'text': "Student ID is empty, so this row cannot be imported."}]})
        rows.sort(key=lambda row: row['n'])
        token = _stash_upload(request, f, 'requests')
        count = lambda st: sum(1 for row in rows if row['status'] == st)
        return render(request, 'letters/import_select.html', {
            'file_name': file_name, 'token': token, 'rows': rows, 'action_url': reverse('letters:upload'),
            'n_rows': len(rows), 'n_ready': count('ready'), 'n_check': count('check'), 'n_err': count('err'),
        })

    # ---- second step: import only the rows that were ticked ----
    n_total = len(records)
    picked = set(request.POST.getlist('pick'))
    records = [r for r in records if str(r.row_index) in picked]
    n_unselected = n_total - len(records)
    if not records:
        _drop_stash(request, token)
        messages.warning(request, "Nothing was imported because no rows were selected.")
        return redirect('letters:index')
    unlisted = _unlisted_students(records)          # ticking such a row confirms adding that student
    _, brought = _bring_in_students(request, unlisted, set(unlisted))
    batch = UploadedBatch.objects.create(file_name=file_name, uploaded_by=request.user)
    for r in records:
        r.batch = batch
    StudentRecord.objects.bulk_create(records)
    if brought:
        messages.info(request, f"{brought} student(s) who were not in the student list were added to the system.")
    if n_unselected:
        messages.warning(request, f"{n_unselected} row(s) were not selected and were not imported.")
    _drop_stash(request, token)
    messages.success(request, f"Imported {len(records)} of {n_total} record(s).")
    return redirect('letters:preview', batch_id=batch.id)


@staff_required
def preview(request, batch_id):
    batch = get_object_or_404(UploadedBatch, id=batch_id)
    records = list(batch.records.all())

    # Fetch ALL templates (active & inactive) so staff can preview any option
    templates = list(LetterTemplate.objects.prefetch_related('field_mappings').order_by('letter_type', 'name'))
    for t in templates:
        t.mapping_list = list(t.field_mappings.all())

    fit_data = {}
    counts = {}
    for r in records:
        kind = resolve_letter_type(r.letter_type)
        base = _build_template_context(r, kind)

        _t = pick_template(kind, r.semester)
        r.default_template_id = _t.id if _t else ''
        r.missing = [f for f in IMPORTANT_FIELDS if not getattr(r, f).strip()]
        r.year_flag = _check_year_eligibility(r)
        builtin_ok = bool(r.letter_type.strip()) and kind in BUILTIN_KINDS
        r.supported = builtin_ok or _t is not None

        clean_type = r.letter_type.strip() if r.letter_type else ""
        r.builtin_label = f"{clean_type} (built-in)" if clean_type else "Built-in layout"

        # Calculate missing placeholders for built-in and EVERY template
        fit_data[r.id] = {
            'builtin': {'supported': builtin_ok, 'missing': r.missing},
            'templates': {t.id: missing_placeholders(t.mapping_list, base) for t in templates},
        }
        counts[r.letter_type] = counts.get(r.letter_type, 0) + 1

    return render(request, 'letters/preview.html', {
        'batch': batch,
        'records': records,
        'counts': counts,
        'templates': templates,
        'fit_data': fit_data,
    })

@staff_required
@require_POST
def update_semester(request, record_id):
    record = get_object_or_404(StudentRecord, id=record_id)
    new_semester = request.POST.get('semester', '').strip()

    if not new_semester:
        return JsonResponse({'ok': False, 'error': 'Semester cannot be empty.'}, status=400)

    record.semester = new_semester
    record.save(update_fields=['semester'])
    log_activity(request, 'EDIT_RECORD', f"Manually set semester to '{new_semester}' for record #{record.id} ({record.display_name})")

    return JsonResponse({'ok': True, 'semester': record.semester})


@staff_required
def record_detail(request, record_id):
    record = get_object_or_404(StudentRecord, id=record_id)
    document = GeneratedDocument.objects.filter(record=record).last()

    if request.method == 'POST':
        updated_data = dict(record.data or {})
        for field_key in list(updated_data.keys()):
            if field_key in request.POST:
                updated_data[field_key] = request.POST.get(field_key)

        # Company and internship period are only changed by an admin, and every change is recorded.
        from .status_views import log_changes
        new_company = _pick(updated_data, H_COMPANY) or record.company
        new_period = _pick(updated_data, H_PERIOD) or record.period
        if (new_company != record.company or new_period != record.period):
            if get_role(request.user) != 'admin':
                messages.error(request, "Only an admin can change the company or the internship period. "
                                        "Ask an admin after the change is approved on paper.")
                return redirect('letters:record_detail', record_id=record.id)
            log_changes(request, record, new_company, new_period, request.POST.get('change_reason', ''))

        record.data = updated_data
        record.company = _pick(updated_data, H_COMPANY) or record.company
        record.contractor = _pick(updated_data, H_CONTACT_PERSON) or record.contractor
        record.contractor_position = _pick(updated_data, H_CONTACT_POSITION) or record.contractor_position
        record.internship_position = _pick(updated_data, H_INTERN_POSITION) or record.internship_position
        record.period = _pick(updated_data, H_PERIOD) or record.period
        record.year = _pick(updated_data, H_YEAR) or record.year
        record.save()

        ref_no = document.ref_no if (document and document.ref_no) else str(2600 + record.row_index)
        buf, fname_or_reason = generate_for_row(record.data, ref_no=ref_no, semester=record.semester)
        if buf:
            if not document:
                document = GeneratedDocument(record=record, ref_no=ref_no)
            document.status = GeneratedDocument.STATUS_GENERATED
            document.file_name = fname_or_reason
            document.file.save(fname_or_reason, ContentFile(buf.getvalue()), save=True)

        log_activity(request, 'EDIT_RECORD', f"Updated student record #{record.id} ({record.display_name})")
        return redirect('letters:record_detail', record_id=record.id)

    if not document or not document.file:
        ref_no = str(2600 + record.row_index)
        buf, fname_or_reason = generate_for_row(record.data, ref_no=ref_no, semester=record.semester)
        if buf:
            if not document:
                document = GeneratedDocument.objects.create(
                    record=record,
                    ref_no=ref_no,
                    status=GeneratedDocument.STATUS_GENERATED,
                    file_name=fname_or_reason,
                )
            document.file.save(fname_or_reason, ContentFile(buf.getvalue()), save=True)

    raw_data = record.data or {}
    valid_keys = [
        k for k in raw_data.keys()
        if k and not k.startswith('Unnamed') and k.lower() not in EXCLUDE_KEYS
    ]

    primary_fields = [(k, raw_data[k]) for k in valid_keys if k in ALL_PRIORITY_KEYS]
    secondary_fields = [(k, raw_data[k]) for k in valid_keys if k not in ALL_PRIORITY_KEYS]

    return render(request, 'letters/record_detail.html', {
        'record': record,
        'primary_fields': primary_fields,
        'secondary_fields': secondary_fields,
        'document': document,
    })


@staff_required
def add_record(request):
    if request.method == 'POST':
        student_id = request.POST.get('student_id', '').strip()
        display_name = request.POST.get('display_name', '').strip()
        letter_type = request.POST.get('letter_type', '').strip()
        year = request.POST.get('year', '').strip()
        course_id = request.POST.get('course_id', '').strip()
        semester = request.POST.get('semester', '').strip()
        company = request.POST.get('company', '').strip()
        contractor = request.POST.get('contractor', '').strip()
        internship_position = request.POST.get('internship_position', '').strip()
        period = request.POST.get('period', '').strip()

        existing_attempts = StudentRecord.objects.filter(
            student_id=student_id,
            letter_type=letter_type,
            course_id=course_id,
            semester=semester
        ).count()

        batch = UploadedBatch.objects.order_by('-uploaded_at').first()
        if not batch:
            batch = UploadedBatch.objects.create(file_name="Manual Entry", uploaded_by=request.user)

        new_record = StudentRecord.objects.create(
            batch=batch,
            student_id=student_id,
            display_name=display_name,
            letter_type=letter_type,
            course_id=course_id,
            semester=semester,
            year=year,
            times=existing_attempts + 1,
            company=company,
            contractor=contractor,
            internship_position=internship_position,
            period=period,
            data={
                'Student ID (e.g. 6587999)': student_id,
                'First Name': display_name,
                'ชั้นปีที่': year,
                'Letter Type': letter_type,
                'Course ID': course_id,
                'Semester': semester,
                'Company/Organization Name for Internship': company,
                'Letter Addressed To (Company Contact Person)': contractor,
                'Internship Position': internship_position,
                'Internship Period': period,
            }
        )

        log_activity(request, 'ADD_RECORD', f"Manually added record for {student_id} ({course_id} {semester})")
        messages.success(request, f"Added letter request #{new_record.times} for {display_name}.")
        return redirect('letters:record_detail', record_id=new_record.id)

    letter_types = sorted(set(LetterTypeMapping.objects.values_list('raw_text', flat=True))
                          | {v.strip() for v in StudentRecord.objects.exclude(letter_type='').values_list('letter_type', flat=True).distinct()})
    return render(request, 'letters/add_record.html', {
        'letter_types': letter_types,
        'courses': Course.objects.all(),
    })


@staff_required
def audit_logs_view(request):
    logs = AuditLog.objects.select_related('user')
    who = request.GET.get('user', '').strip()
    action = request.GET.get('action', '').strip()
    if who.isdigit():
        logs = logs.filter(user_id=int(who))
    if action:
        logs = logs.filter(action=action)
    return render(request, 'letters/audit_logs.html', {
        'logs': logs[:300],
        'users': User.objects.order_by('username'),
        'actions': AuditLog.objects.order_by().values_list('action', flat=True).distinct(),
        'who': who, 'action': action,
    })


@staff_required
def generate(request, batch_id):
    batch = get_object_or_404(UploadedBatch, id=batch_id)

    if request.method == "POST":
        selected_ids = request.POST.getlist("record_id")
        fmt = request.POST.get("format", "docx").lower()

        if not selected_ids:
            messages.error(request, "No records selected.")
            return redirect("letters:preview", batch_id=batch_id)

        records = StudentRecord.objects.filter(id__in=selected_ids, batch=batch)
        templates_by_id = {str(t.id): t for t in LetterTemplate.objects.all()}

        zip_buffer = io.BytesIO()
        added_count = 0
        errors = []

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            if fmt == "docx":
                for record in records:
                    try:
                        tobj, fb = _chosen_template(request, record, templates_by_id)
                        buf, filename = generate_for_row(record.data, ref_no="___", template_obj=tobj, force_builtin=fb, semester=record.semester)
                        if buf:
                            from .extra_views import save_document_version
                            save_document_version(record, buf.getvalue(), ref_no=str(record.student_id), file_name=filename)
                            zf.writestr(filename, buf.getvalue())
                            added_count += 1
                        else:
                            errors.append(f"Record #{record.id}: {filename}")
                    except Exception as e:
                        errors.append(f"Record #{record.id}: {str(e)}")
            else:
                with tempfile.TemporaryDirectory() as tmpdir:
                    file_map = {}
                    docx_paths = []

                    for record in records:
                        try:
                            ref = str(record.student_id or record.id or "___")
                            tobj, fb = _chosen_template(request, record, templates_by_id)
                            res = generate_for_row(record.data, ref_no=ref, template_obj=tobj, force_builtin=fb, semester=record.semester)
                            if not res or res[0] is None:
                                err_msg = res[1] if res else "Unknown error"
                                errors.append(f"Record #{record.id}: {err_msg}")
                                continue

                            docx_buf, orig_fname = res
                            from .extra_views import save_document_version
                            save_document_version(record, docx_buf.getvalue(), ref_no=str(record.student_id),
                                                  file_name=orig_fname)
                            pdf_fname = orig_fname.rsplit(".", 1)[0] + ".pdf"

                            docx_buf.seek(0)
                            pdf_input_doc = Document(docx_buf)
                            apply_thai_font_fix(pdf_input_doc, font_name=PDF_SAFE_FONT)

                            doc_filename = f"doc_{record.id}.docx"
                            input_docx_path = os.path.join(tmpdir, doc_filename)
                            pdf_input_doc.save(input_docx_path)

                            docx_paths.append(input_docx_path)
                            file_map[doc_filename] = (
                                pdf_fname,
                                os.path.join(tmpdir, f"doc_{record.id}.pdf"),
                            )
                        except Exception as e:
                            errors.append(f"Record #{record.id}: {str(e)}")

                    if docx_paths:
                        soffice_bin = _get_libreoffice_bin()
                        if not soffice_bin:
                            messages.error(request, "ไม่พบโปรแกรม LibreOffice ในระบบ")
                            return redirect("letters:preview", batch_id=batch_id)

                        env = os.environ.copy()
                        env["LANG"] = "th_TH.UTF-8"
                        env["LC_ALL"] = "th_TH.UTF-8"

                        lo_profile_dir = os.path.join(tmpdir, "lo_profile")
                        os.makedirs(lo_profile_dir, exist_ok=True)
                        profile_uri = Path(lo_profile_dir).resolve().as_uri()

                        cmd = [
                            soffice_bin,
                            "--headless",
                            "--invisible",
                            "--nologo",
                            "--norestore",
                            "--nofirststartwizard",
                            f"-env:UserInstallation={profile_uri}",
                            "--convert-to",
                            "pdf:writer_pdf_Export",
                            "--outdir",
                            tmpdir,
                            *docx_paths,
                        ]

                        try:
                            subprocess.run(
                                cmd,
                                stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE,
                                env=env,
                                timeout=90,
                            )
                        except subprocess.TimeoutExpired:
                            logger.error("Batch LibreOffice conversion timed out.")

                    for doc_filename, (out_zip_fname, target_pdf_path) in file_map.items():
                        if os.path.exists(target_pdf_path):
                            with open(target_pdf_path, "rb") as f:
                                zf.writestr(out_zip_fname, f.read())
                                added_count += 1
                        else:
                            errors.append(f"PDF generation failed for {out_zip_fname}")

        if added_count == 0:
            first_err = errors[0] if errors else "PDF conversion failed."
            messages.error(request, f"Failed to generate documents. Reason: {first_err}")
            return redirect("letters:preview", batch_id=batch_id)

        zip_buffer.seek(0)
        response = HttpResponse(zip_buffer.getvalue(), content_type="application/zip")
        response["Content-Disposition"] = (
            f'attachment; filename="letters_batch_{batch.id}_{fmt}.zip"'
        )
        return response

    return redirect("letters:preview", batch_id=batch_id)


@staff_required
def download(request, run_id):
    gen_batch = get_object_or_404(GenerationBatch, id=run_id)
    generated = gen_batch.documents.filter(status=GeneratedDocument.STATUS_GENERATED)
    if not generated.exists():
        raise Http404("No documents in this batch")

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for doc in generated:
            doc.file.open('rb')
            zf.writestr(doc.file_name, doc.file.read())
            doc.file.close()
    zip_buf.seek(0)
    return FileResponse(zip_buf, as_attachment=True, filename='internship_letters.zip')


@staff_required
def results(request, run_id):
    gen_batch = get_object_or_404(GenerationBatch, id=run_id)
    documents = gen_batch.documents.select_related('record').order_by('id')

    generated_count = documents.filter(status=GeneratedDocument.STATUS_GENERATED).count()
    total_count = documents.count()
    failed_count = total_count - generated_count

    return render(request, 'letters/results.html', {
        'batch': gen_batch,
        'documents': documents,
        'generated_count': generated_count,
        'failed_count': failed_count,
        'total_count': total_count,
    })


# ==========================================
# FILE PREVIEW & DOWNLOAD VIEWS
# ==========================================

@staff_required
def download_record_pdf(request, record_id):
    record = get_object_or_404(StudentRecord, id=record_id)
    document = GeneratedDocument.objects.filter(record=record).last()

    if document and _field_ready(document.signed_file):
        document.signed_file.open('rb')
        response = FileResponse(document.signed_file, content_type='application/pdf')
        filename = os.path.basename(document.signed_file.name)
    else:
        ref_no = request.GET.get('ref_no', '___')
        pdf_buf, filename = generate_pdf_for_row(record.data, ref_no=ref_no, semester=record.semester)
        if not pdf_buf:         # filename holds the reason when nothing could be made
            messages.error(request, f"Could not create the PDF for this record: {filename}")
            return redirect(request.META.get('HTTP_REFERER') or reverse('letters:index'))
        response = HttpResponse(pdf_buf.getvalue(), content_type='application/pdf')

    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response


@xframe_options_sameorigin
@staff_required
def preview_record_pdf(request, record_id):
    record = get_object_or_404(StudentRecord, id=record_id)
    doc = GeneratedDocument.objects.filter(record=record).last()

    # 1) signed PDF, if uploaded
    if doc and _field_ready(doc.signed_file):
        try:
            with doc.signed_file.open('rb') as f:
                data = f.read()
        except Exception as e:
            return HttpResponse(f"Unable to read signed file from storage: {e}", status=404)
        response = HttpResponse(data, content_type='application/pdf')
        response['Content-Disposition'] = f'inline; filename="{os.path.basename(doc.signed_file.name)}"'
        return response

    # 2) generated file, if saved
    if doc and _field_ready(doc.file):
        try:
            with doc.file.open('rb') as f:
                file_bytes = f.read()
        except Exception as e:
            return HttpResponse(f"Unable to read file from storage: {e}", status=404)

        if doc.file.name.lower().endswith('.pdf'):
            pdf_bytes = file_bytes
        else:
            try:
                pdf_bytes = convert_docx_bytes_to_pdf(file_bytes)
            except Exception as e:
                return HttpResponse(f"PDF Conversion Failed: {e}", status=500)
        response = HttpResponse(pdf_bytes, content_type='application/pdf')
        response['Content-Disposition'] = f'inline; filename="student_{record.student_id}.pdf"'
        return response

    # 3) nothing saved yet: generate on the fly
    ref_no = doc.ref_no if (doc and doc.ref_no) else str(2600 + record.row_index)
    try:
        pdf_buf, _name = generate_pdf_for_row(record.data, ref_no=ref_no, semester=record.semester)
    except Exception as e:
        return HttpResponse(f"Unable to generate preview: {e}", status=500)
    if not pdf_buf:
        return HttpResponse(f"Unable to generate preview: {_name}", status=422)
    response = HttpResponse(pdf_buf.getvalue(), content_type='application/pdf')
    response['Content-Disposition'] = f'inline; filename="student_{record.student_id}.pdf"'
    return response


@staff_required
def download_letter_view(request, record_id, letter_type):
    record = get_object_or_404(StudentRecord, id=record_id)
    if letter_type not in ('confirm', 'request', 'extension'):
        raise Http404(f"Unknown letter type: '{letter_type}'")

    ref_no = str(2600 + record.row_index)
    buf, fname_or_reason = generate_for_row(record.data, ref_no=ref_no, kind_override=letter_type, semester=record.semester)
    if buf is None:
        raise Http404(fname_or_reason)

    buf.seek(0)
    return FileResponse(
        buf, as_attachment=True, filename=fname_or_reason,
        content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )


# ==========================================
# BATCH DISPATCH & SIGNED FILE MANAGEMENT
# ==========================================

@staff_required
def upload_signed(request, batch_id=None):
    selected_batch_id = request.GET.get('batch_id') or (str(batch_id) if batch_id else 'all')

    if request.method == 'POST':
        uploaded_files = request.FILES.getlist('signed_files') or request.FILES.getlist('files')
        matched_count = 0

        if selected_batch_id != 'all' and selected_batch_id.isdigit():
            batch = get_object_or_404(UploadedBatch, id=int(selected_batch_id))
            records = batch.records.all()
        else:
            all_records = StudentRecord.objects.order_by('-id')
            seen = set()
            records = []
            for r in all_records:
                key = (r.student_id.strip(), r.letter_type.strip(), r.times)
                if key not in seen:
                    seen.add(key)
                    records.append(r)

        for uploaded_file in uploaded_files:
            fname = uploaded_file.name.lower()
            if fname.endswith('.zip'):
                with zipfile.ZipFile(uploaded_file, 'r') as z:
                    for inner_name in z.namelist():
                        if inner_name.lower().endswith('.pdf') and not inner_name.startswith('__MACOSX'):
                            pdf_bytes = z.read(inner_name)
                            if _match_and_save_pdf(records, inner_name, ContentFile(pdf_bytes, name=inner_name)):
                                matched_count += 1
            elif fname.endswith('.pdf'):
                if _match_and_save_pdf(records, uploaded_file.name, uploaded_file):
                    matched_count += 1

        log_activity(
            request,
            'UPLOAD_SIGNED',
            f"Uploaded signed documents ({selected_batch_id}). Matched {matched_count} file(s)."
        )

        if matched_count > 0:
            messages.success(request, f"Successfully matched {matched_count} signed document(s).")
        else:
            messages.warning(request, "No matching students found. Ensure PDF filenames include Student IDs or Names.")

    return redirect(f"{reverse('letters:signed_preview')}?batch_id={selected_batch_id}")


@staff_required
def signed_preview(request, batch_id=None):
    batches = UploadedBatch.objects.order_by('-uploaded_at')
    selected_batch_id = request.GET.get('batch_id') or (str(batch_id) if batch_id else 'all')

    if selected_batch_id != 'all' and selected_batch_id.isdigit():
        records = StudentRecord.objects.filter(batch_id=int(selected_batch_id))
    else:
        records = StudentRecord.objects.select_related('batch').order_by('-id')

    saved_email = EmailTemplate.load()

    docs = GeneratedDocument.objects.filter(record__in=records).order_by('record_id', 'id')
    latest_doc_map = {doc.record_id: doc for doc in docs}

    record_summary = []
    for r in records:
        doc = latest_doc_map.get(r.id)

        try:
            student_email = get_student_email(r)
        except Exception:
            student_email = ""

        has_signed = False
        if doc and getattr(doc, 'signed_file', None):
            try:
                has_signed = bool(_field_ready(doc.signed_file))
            except Exception:
                has_signed = False

        record_summary.append({
            'record': r,
            'doc': doc,
            'email': student_email,
            'has_signed': has_signed,
        })

    return render(request, 'letters/signed_preview.html', {
        'batches': batches,
        'selected_batch_id': str(selected_batch_id),
        'summary': record_summary,
        'saved_email': saved_email,
    })


def _fill_placeholders(text, record, staff_name, staff_email):
    """Replace {name}, {student_id}, {company}, {staff_name}, {staff_email}.
    Plain str.replace is used so a stray '{' typed by staff cannot crash it."""
    values = {
        '{name}': record.display_name or '',
        '{student_id}': record.student_id or '',
        '{company}': record.company or '',
        '{staff_name}': staff_name,
        '{staff_email}': staff_email,
    }
    for k, v in values.items():
        text = text.replace(k, str(v))
    return text


@staff_required
def send_signed_emails(request):
    if request.method != 'POST':
        return redirect('letters:signed_preview')

    selected_ids = request.POST.getlist('selected_records')
    selected_batch = request.POST.get('batch_id', 'all')

    records = StudentRecord.objects.filter(id__in=selected_ids)

    if not records.exists():
        messages.warning(request, "No student records were selected. Please check at least one box.")
        return redirect(f"{reverse('letters:signed_preview')}?batch_id={selected_batch}")

    staff_email = request.user.email if request.user.email else settings.DEFAULT_FROM_EMAIL
    staff_name = request.user.get_full_name() or request.user.username

    # the message comes from Configuration > Email templates: the letter type's own one, else the default
    from .models import LetterType as _LT
    _types = {t.slug: t for t in _LT.objects.all()}
    _tpl_cache = {}

    def _template_for(rec):
        slug = resolve_letter_type(rec.letter_type)
        if slug not in _tpl_cache:
            _tpl_cache[slug] = EmailTemplate.load(_types.get(slug))
        return _tpl_cache[slug]

    sent_count = 0
    failed = []

    for r in records:
        student_email = request.POST.get(f'email_{r.id}', '').strip() or get_student_email(r)
        doc = GeneratedDocument.objects.filter(record=r).last()

        if doc and _field_ready(doc.signed_file) and student_email:
            _t = _template_for(r)
            subject_tpl, body_tpl = _t.subject, _t.body
            subject = _fill_placeholders(subject_tpl, r, staff_name, staff_email)
            subject = ' '.join(subject.splitlines())   # header must be one line
            body = _fill_placeholders(body_tpl, r, staff_name, staff_email)

            email = EmailMessage(
                subject=subject,
                body=body,
                from_email=staff_email,
                to=[student_email],
                reply_to=[staff_email],
            )

            with doc.signed_file.open('rb') as f:
                email.attach(os.path.basename(doc.signed_file.name), f.read(), 'application/pdf')

            if _field_ready(doc.file):
                with doc.file.open('rb') as f:
                    is_pdf = doc.file.name.lower().endswith('.pdf')
                    c_type = 'application/pdf' if is_pdf else 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    email.attach(os.path.basename(doc.file.name), f.read(), c_type)

            try:
                email.send(fail_silently=False)
            except Exception as e:
                failed.append(f"{r.display_name} ({student_email}): {e}")
                continue

            doc.status = GeneratedDocument.STATUS_SENT
            doc.email_sent_at = timezone.now()
            doc.save()
            sent_count += 1

    log_activity(
        request,
        'SEND_EMAIL',
        f"Sent {sent_count} signed internship document email(s)." + (f" {len(failed)} failed." if failed else "")
    )

    if sent_count > 0:
        messages.success(request, f"Successfully sent {sent_count} email(s) to student(s).")
    if failed:
        messages.warning(request, f"{len(failed)} email(s) failed: " + "; ".join(failed[:5]) + (" …" if len(failed) > 5 else ""))
    if sent_count == 0 and not failed:
        messages.warning(request, "No emails sent. Verify students have attached signed documents and valid email addresses.")

    return redirect(f"{reverse('letters:signed_preview')}?batch_id={selected_batch}")


@staff_required
def view_signed_pdf(request, document_id):
    document = get_object_or_404(GeneratedDocument, id=document_id)
    if not _field_ready(document.signed_file):
        raise Http404("No signed file uploaded for this document yet.")

    document.signed_file.open('rb')
    response = FileResponse(document.signed_file, content_type='application/pdf')
    response['Content-Disposition'] = f'inline; filename="{os.path.basename(document.signed_file.name)}"'
    return response


@staff_required
def upload_signed_by_id(request, letter_id):
    if request.method == 'POST' and 'signed_file' in request.FILES:
        letter = get_object_or_404(StudentLetter, pk=letter_id)
        letter.signed_file = request.FILES['signed_file']
        letter.status = 'signed'
        letter.save()
        messages.success(request, f"Signed file uploaded for student {letter.student_id}.")
    else:
        messages.error(request, "Invalid file upload attempt.")

    return redirect('letters:index')


# ==========================================
# TEMPLATE & LETTER TYPE MANAGEMENT
# ==========================================

@staff_required
def template_list(request):
    from .docgen import BUILTIN_TITLES, build_tagged_letter
    from .template_editor import mini_preview_html
    templates = list(LetterTemplate.objects.order_by('letter_type', '-version'))
    for t in templates:
        mappings = list(t.field_mappings.all())
        t.total_fields = len(mappings)
        t.mapped_fields = len([m for m in mappings if m.source_type != TemplateFieldMapping.SOURCE_UNMAPPED])

    letter_types = LetterType.objects.order_by('display_name')
    names = {lt.slug: lt.display_name for lt in letter_types}
    cards = []
    slugs = list(dict.fromkeys([t.letter_type for t in templates] + [lt.slug for lt in letter_types] + list(BUILTIN_TITLES)))
    for slug in slugs:
        mine = [t for t in templates if t.letter_type == slug]
        shown = next((t for t in mine if t.is_active), mine[0] if mine else None)
        if shown is None and slug not in BUILTIN_TITLES:
            continue                                    # a letter type nobody has made a template for yet
        card = {'slug': slug, 'name': names.get(slug) or BUILTIN_TITLES.get(slug, slug), 'shown': shown,
                'n_versions': len(mine), 'builtin_only': shown is None}
        try:
            if shown is not None:
                shown.file.open('rb')
                try:
                    card['mini'] = mini_preview_html(shown.file.read())
                finally:
                    shown.file.close()
            else:
                card['mini'] = mini_preview_html(build_tagged_letter(slug))
        except Exception:
            logger.exception("Could not draw the template thumbnail")
            card['mini'] = ''
        cards.append(card)
    return render(request, 'letters/template_list.html', {
        'templates': templates,
        'cards': cards,
        'letter_types': letter_types,
        'update_types': _update_types(letter_types),
    })


@staff_required
def template_upload(request):
    if request.method != 'POST' or 'template_file' not in request.FILES:
        messages.error(request, "Please select a template file and letter type.")
        return redirect('letters:template_list')

    letter_type = request.POST.get('letter_type', '').strip()
    if not LetterType.objects.filter(slug=letter_type).exists():
        messages.error(request, "Please select a valid, existing letter type.")
        return redirect('letters:template_list')

    f = request.FILES['template_file']
    last = (LetterTemplate.objects.filter(letter_type=letter_type)
            .order_by('-version').first())
    template_obj = LetterTemplate.objects.create(
        letter_type=letter_type, name=f.name, file=f,
        version=(last.version + 1) if last else 1, is_active=False,
    )
    created = last is None

    template_obj.file.open('rb')
    try:
        detected = extract_fields_from_docx(template_obj.file)
    finally:
        template_obj.file.close()

    placeholders = detected.get('template_variables') or []
    if not placeholders:
        messages.warning(
            request,
            f"Uploaded '{f.name}', but no {{{{ tag }}}}-style placeholders were found in it."
        )
        return redirect('letters:template_list')

    for tag in placeholders:
        TemplateFieldMapping.objects.get_or_create(template=template_obj, placeholder=tag)

    action = 'Uploaded new version of' if not created else 'Uploaded'
    log_activity(request, 'UPLOAD_TEMPLATE', f"{action} template '{f.name}' for '{letter_type}' v{template_obj.version}")
    messages.success(request, f"Saved as version {template_obj.version}. Map its {len(placeholders)} field(s) before activating. Older versions are kept.")
    return redirect('letters:template_mapping', template_id=template_obj.id)


@staff_required
def template_from_builtin(request):
    """Turn one of the built-in letters into an editable template version and open it in the editor."""
    from .docgen import AVAILABLE_SOURCE_FIELDS, BUILTIN_TITLES, build_tagged_letter
    kind = request.POST.get('kind', '').strip()
    if request.method != 'POST' or kind not in BUILTIN_TITLES:
        messages.error(request, "Please choose which built-in letter to edit.")
        return redirect('letters:template_list')
    lt, _ = LetterType.objects.get_or_create(slug=kind, defaults={'display_name': BUILTIN_TITLES[kind]})
    last = (LetterTemplate.objects.filter(letter_type=kind)
            .order_by('-version').first())
    ver = (last.version if last else 0) + 1
    name = f"{kind}_builtin_v{ver}.docx"
    t = LetterTemplate(letter_type=kind, name=name, version=ver, is_active=False)
    t.file.save(name, ContentFile(build_tagged_letter(kind)), save=True)
    _premapped_tags(t)
    log_activity(request, 'UPLOAD_TEMPLATE', f"Made the built-in '{kind}' letter editable as template v{ver}")
    messages.success(request, "The built-in letter is now an editable template (all tags are already mapped). Change it below, then save, preview and activate.")
    return redirect('letters:template_editor', template_id=t.id)


def _premapped_tags(t):
    """Map every tag of a template that has the same name as a system field; returns the tag list."""
    from .docgen import AVAILABLE_SOURCE_FIELDS
    known = {tag for tag, _label in AVAILABLE_SOURCE_FIELDS}
    t.file.open('rb')
    try:
        tags = extract_fields_from_docx(t.file).get('template_variables') or []
    finally:
        t.file.close()
    for tag in tags:
        TemplateFieldMapping.objects.get_or_create(
            template=t, placeholder=tag,
            defaults={'source_type': TemplateFieldMapping.SOURCE_FIELD if tag in known else TemplateFieldMapping.SOURCE_UNMAPPED,
                      'source_value': tag if tag in known else ''})
    return tags


def _update_types(letter_types):
    """Letter types that can be updated: every saved type plus the built-in letters."""
    from .docgen import BUILTIN_TITLES
    out = [{'slug': lt.slug, 'name': lt.display_name, 'builtin': lt.slug in BUILTIN_TITLES} for lt in letter_types]
    have = {o['slug'] for o in out}
    out += [{'slug': k, 'name': title, 'builtin': True} for k, title in BUILTIN_TITLES.items() if k not in have]
    return out


@staff_required
def template_create(request):
    """Create a template for a NEW letter type, from the system's letter or from an uploaded Word file."""
    from django.utils.text import slugify
    from .docgen import build_tagged_letter
    if request.method != 'POST':
        return redirect('letters:template_list')
    name = request.POST.get('display_name', '').strip()[:150]
    if not name:
        messages.error(request, "Please give the new letter type a name.")
        return redirect('letters:template_list')
    if LetterType.objects.filter(display_name__iexact=name).exists():
        messages.error(request, f"A letter type named '{name}' already exists. Use Update template to edit it.")
        return redirect('letters:template_list')
    up = request.FILES.get('template_file')
    if up and not up.name.lower().endswith('.docx'):
        messages.error(request, "The template file must be a Word .docx file.")
        return redirect('letters:template_list')

    base = slugify(name)[:40] or 'letter'
    slug, n = base, 1
    while LetterType.objects.filter(slug=slug).exists():
        n += 1
        slug = f"{base}-{n}"
    lt = LetterType.objects.create(slug=slug, display_name=name)
    LetterTypeMapping.objects.get_or_create(raw_text=name, defaults={'letter_type': lt})
    if up:
        t = LetterTemplate.objects.create(letter_type=slug, name=up.name, file=up,
                                          version=1, is_active=False)
        tags = _premapped_tags(t)
        log_activity(request, 'UPLOAD_TEMPLATE', f"Created letter type '{name}' with uploaded template '{up.name}'")
        if not tags:
            messages.warning(request, f"Letter type '{name}' was created, but no {{{{ tag }}}} placeholders were found in the file.")
            return redirect('letters:template_list')
        messages.success(request, f"Letter type '{name}' created. Check the tag mapping, then preview and activate.")
        return redirect('letters:template_mapping', template_id=t.id)
    fname = f"{slug}_v1.docx"
    t = LetterTemplate(letter_type=slug, name=fname, version=1, is_active=False)
    t.file.save(fname, ContentFile(build_tagged_letter('request')), save=True)
    _premapped_tags(t)
    log_activity(request, 'UPLOAD_TEMPLATE', f"Created letter type '{name}' from the system letter")
    messages.success(request, f"Letter type '{name}' created from the system's letter. Change the wording below, then save, preview and activate.")
    return redirect('letters:template_editor', template_id=t.id)


@admin_required
def letter_settings(request):
    """Who signs the letters and the few numbers every built-in letter shares."""
    from decimal import Decimal, InvalidOperation
    from .models import LetterSettings
    obj = LetterSettings.load()
    if request.method == 'POST':
        name = request.POST.get('approver_name', '').strip()
        title = request.POST.get('approver_title', '').strip()
        prefix = request.POST.get('ref_prefix', '').strip()
        if not name or not title or not prefix:
            messages.error(request, "Approver name, approver title and reference prefix cannot be empty.")
            return redirect('letters:letter_settings')
        try:
            spacing = Decimal(request.POST.get('line_spacing', '1.0') or '1.0')
            if not Decimal('0.8') <= spacing <= Decimal('2.0'):
                raise InvalidOperation
        except InvalidOperation:
            messages.error(request, "Line spacing must be a number between 0.8 and 2.0.")
            return redirect('letters:letter_settings')
        fixed = request.POST.get('letter_date', '').strip()
        try:
            letter_date = datetime.strptime(fixed, '%Y-%m-%d').date() if fixed else None
        except ValueError:
            messages.error(request, "The fixed letter date is not a valid date.")
            return redirect('letters:letter_settings')
        obj.dean_name, obj.dean_title, obj.ref_prefix = name, title, prefix
        obj.line_spacing, obj.letter_date = spacing, letter_date
        obj.save()
        messages.success(request, "Letter settings saved. New letters use them straight away.")
        return redirect('letters:letter_settings')
    return render(request, 'letters/letter_settings.html', {'s': obj})


@staff_required
def lettertype_list(request):
    types = LetterType.objects.prefetch_related('mappings').order_by('display_name')
    unmapped_raw_values = (
        StudentRecord.objects.exclude(letter_type='')
        .values_list('letter_type', flat=True).distinct()
    )
    known_raw = set(LetterTypeMapping.objects.values_list('raw_text', flat=True))
    unmapped_raw_values = sorted({v.strip() for v in unmapped_raw_values if v.strip() not in known_raw})
    return render(request, 'letters/lettertype_list.html', {
        'types': types,
        'unmapped_raw_values': unmapped_raw_values,
    })


@staff_required
def lettertype_create(request):
    if request.method == 'POST':
        slug = request.POST.get('slug', '').strip()
        display_name = request.POST.get('display_name', '').strip()
        if not slug or not display_name:
            messages.error(request, "Both slug and display name are required.")
        else:
            LetterType.objects.get_or_create(slug=slug, defaults={'display_name': display_name})
            log_activity(request, 'CREATE_LETTER_TYPE', f"Created letter type '{slug}'")
            messages.success(request, f"Created letter type '{display_name}'.")
    return redirect('letters:lettertype_list')


@staff_required
def preview_record_docx(request, record_id):
    record = get_object_or_404(StudentRecord, id=record_id)
    doc = GeneratedDocument.objects.filter(record=record).last()

    if not doc or not _field_ready(doc.file):
        # no generated Word file saved yet: build it now
        ref_no = doc.ref_no if (doc and doc.ref_no) else str(2600 + record.row_index)
        buf, fname = generate_for_row(record.data, ref_no=ref_no, semester=record.semester)
        if not buf:
            return HttpResponse("No generated document found for this record.", status=404)
        response = HttpResponse(
            buf.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        response['Content-Disposition'] = f'inline; filename="{fname}"'
        return response

    try:
        with doc.file.open('rb') as f:
            file_bytes = f.read()
        response = HttpResponse(
            file_bytes,
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        filename = os.path.basename(doc.file.name)
        response['Content-Disposition'] = f'inline; filename="{filename}"'
        return response
    except Exception as e:
        return HttpResponse(f"Unable to read file: {str(e)}", status=404)


@staff_required
def mapping_create(request):
    if request.method == 'POST':
        raw_text = request.POST.get('raw_text', '').strip()
        letter_type_id = request.POST.get('letter_type_id')
        if not raw_text or not letter_type_id:
            messages.error(request, "Both the raw text and a letter type are required.")
        else:
            letter_type = get_object_or_404(LetterType, id=letter_type_id)
            LetterTypeMapping.objects.update_or_create(raw_text=raw_text, defaults={'letter_type': letter_type})
            log_activity(request, 'CREATE_TYPE_MAPPING', f"Mapped '{raw_text}' -> '{letter_type.slug}'")
            messages.success(request, f"Mapped '{raw_text}' to '{letter_type.display_name}'.")
    return redirect('letters:lettertype_list')


@staff_required
def template_mapping(request, template_id):
    template_obj = get_object_or_404(LetterTemplate, id=template_id)
    mappings = template_obj.field_mappings.order_by('placeholder')

    if request.method == 'POST':
        for mapping in mappings:
            src_type = request.POST.get(f'source_type_{mapping.id}', TemplateFieldMapping.SOURCE_UNMAPPED)
            src_val = request.POST.get(f'source_value_{mapping.id}', '').strip()
            custom_val = request.POST.get(f'custom_value_{mapping.id}', '').strip()

            mapping.source_type = src_type

            if src_type == TemplateFieldMapping.SOURCE_FIELD:
                # Use typed custom column name (e.g. 'nickname') if '__custom__' or text is provided
                if src_val == '__custom__' or custom_val:
                    mapping.source_value = custom_val
                else:
                    mapping.source_value = src_val
            elif src_type == TemplateFieldMapping.SOURCE_STATIC:
                mapping.source_value = src_val
            else:
                mapping.source_value = ''

            mapping.save()

        log_activity(request, 'MAP_TEMPLATE_FIELDS', f"Updated mapping for template #{template_obj.id}")
        messages.success(request, "Field mapping saved.")
        return redirect('letters:template_mapping', template_id=template_obj.id)

    # Scan existing records to extract uploaded Excel header columns
    dynamic_columns = set()
    recent_records = StudentRecord.objects.order_by('-id')[:100]
    for rec in recent_records:
        if rec.data and isinstance(rec.data, dict):
            for k in rec.data.keys():
                if k and not k.startswith('Unnamed') and k.lower() not in EXCLUDE_KEYS:
                    dynamic_columns.add(k)

    standard_keys = {item[0] for item in AVAILABLE_SOURCE_FIELDS}
    standard_fields = AVAILABLE_SOURCE_FIELDS

    custom_fields = [
        (col, col) for col in sorted(dynamic_columns) if col not in standard_keys
    ]

    return render(request, 'letters/template_mapping.html', {
        'template': template_obj,
        'mappings': mappings,
        'standard_fields': standard_fields,
        'custom_fields': custom_fields,
    })



@staff_required
def template_preview(request, template_id):
    template_obj = get_object_or_404(LetterTemplate, id=template_id)
    sample_record = StudentRecord.objects.order_by('-id').first()
    if not sample_record:
        messages.error(request, "No student records exist yet to preview against.")
        return redirect('letters:template_mapping', template_id=template_obj.id)

    kind = normalize_letter_type(sample_record.letter_type)
    ref_no = str(2600 + sample_record.row_index)

    docx_buf = render_with_template(template_obj, sample_record, kind, ref_no)
    pdf_bytes = convert_docx_bytes_to_pdf(docx_buf.getvalue())

    response = HttpResponse(pdf_bytes, content_type='application/pdf')
    response['Content-Disposition'] = f'inline; filename="preview_{template_obj.name}.pdf"'
    return response


@staff_required
def template_activate(request, template_id):
    template_obj = get_object_or_404(LetterTemplate, id=template_id)
    unmapped = template_obj.field_mappings.filter(source_type=TemplateFieldMapping.SOURCE_UNMAPPED).count()
    if unmapped:
        messages.error(request, f"Can't activate — {unmapped} placeholder(s) still unmapped.")
        return redirect('letters:template_mapping', template_id=template_obj.id)

    # one active version per letter type
    LetterTemplate.objects.filter(letter_type=template_obj.letter_type).exclude(pk=template_obj.pk).update(is_active=False)
    template_obj.is_active = True
    template_obj.save()
    log_activity(request, 'ACTIVATE_TEMPLATE', f"Activated template #{template_obj.id} for '{template_obj.letter_type}'")
    messages.success(request, f"'{template_obj.name}' is now active for '{template_obj.letter_type}'.")
    return redirect('letters:template_list')

# Add these views to letters/views.py

@staff_required
def lettertype_delete(request, type_id):
    if request.method == 'POST':
        letter_type = get_object_or_404(LetterType, id=type_id)
        name = letter_type.display_name
        letter_type.delete()
        log_activity(request, 'DELETE_LETTER_TYPE', f"Deleted letter type '{name}'")
        messages.success(request, f"Deleted letter type '{name}'.")
    return redirect('letters:lettertype_list')


@staff_required
def mapping_update(request, mapping_id):
    if request.method == 'POST':
        mapping = get_object_or_404(LetterTypeMapping, id=mapping_id)
        new_type_id = request.POST.get('letter_type_id')
        if new_type_id:
            new_type = get_object_or_404(LetterType, id=new_type_id)
            old_name = mapping.letter_type.display_name
            mapping.letter_type = new_type
            mapping.save()
            log_activity(request, 'UPDATE_TYPE_MAPPING', f"Reassigned '{mapping.raw_text}' from '{old_name}' to '{new_type.display_name}'")
            messages.success(request, f"Reassigned '{mapping.raw_text}' to '{new_type.display_name}'.")
    return redirect('letters:lettertype_list')


@staff_required
def mapping_delete(request, mapping_id):
    if request.method == 'POST':
        mapping = get_object_or_404(LetterTypeMapping, id=mapping_id)
        raw_text = mapping.raw_text
        mapping.delete()
        log_activity(request, 'DELETE_TYPE_MAPPING', f"Deleted mapping for '{raw_text}'")
        messages.success(request, f"Unmapped '{raw_text}'.")
    return redirect('letters:lettertype_list')


@staff_required
def template_delete(request, template_id):
    if request.method == 'POST':
        template_obj = get_object_or_404(LetterTemplate, id=template_id)
        name = template_obj.name
        template_obj.delete()
        log_activity(request, 'DELETE_TEMPLATE', f"Deleted template '{name}'")
        messages.success(request, f"Deleted template '{name}'.")
    return redirect('letters:template_list')



# ----------------------------------------------------------------------
# Template editing: choose or upload the letter logo
# ----------------------------------------------------------------------
@staff_required
def template_edit(request, template_id):
    template_obj = get_object_or_404(LetterTemplate, id=template_id)
    back = (reverse('letters:template_editor', args=[template_obj.id]) if request.POST.get('next') == 'editor'
            else reverse('letters:template_edit', args=[template_obj.id]))

    if request.method == 'POST':
        new_image = request.FILES.get('new_logo')
        if new_image:
            if not (new_image.content_type or '').startswith('image/'):
                messages.error(request, "The logo must be a PNG or JPG image.")
                return redirect(back)
            logo = Logo.objects.create(
                name=request.POST.get('new_logo_name', '').strip() or new_image.name,
                image=new_image,
            )
        else:
            logo_id = request.POST.get('logo_id', '')
            logo = Logo.objects.filter(id=logo_id).first() if logo_id else None

        template_obj.logo = logo
        baked = False
        if logo:
            try:
                baked = bake_logo_into_template(template_obj, logo)
            except Exception:
                logger.exception("Could not put logo into template %s", template_obj.pk)
        template_obj.save()
        if logo and not baked:
            messages.warning(request, "No picture was found in this template to replace, so the logo was not applied. Add a placeholder picture to the template header in Word first.")
        log_activity(request, 'EDIT_TEMPLATE', f"Set logo of '{template_obj.name}' to '{logo.name if logo else 'file default'}'")
        messages.success(
            request,
            "Logo saved. Letters generated from now on use it; earlier letters keep the old logo.",
        )
        return redirect(back)

    return render(request, 'letters/template_edit.html', {
        'template': template_obj,
        'logos': Logo.objects.all(),
    })
