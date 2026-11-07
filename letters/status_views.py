# -*- coding: utf-8 -*-
"""
Request Status feature: subjects, students (registrations), request status
per subject, and a stand-alone document-status page for every request.

Copy this file next to views.py (letters/status_views.py) and add the URLs
listed in urls_snippet.txt.
"""
import re
import uuid
from collections import Counter, defaultdict
from datetime import date

from django.contrib import messages
from django.core.exceptions import ValidationError
from django.core.files.storage import default_storage
from django.core.validators import validate_email
from django.core.paginator import Paginator
from django.db import transaction
from django.db.models import Count, Q
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST

from .decorators import staff_required
from .docgen import normalize_letter_type
from .models import (
    Course,
    Enrollment,
    GeneratedDocument,
    LetterType,
    LetterTypeMapping,
    Student,
    StudentRecord,
    UploadedBatch,
)
from .views import (
    _field_ready,
    _load_table,
    _parse_thai_date,
    get_student_email,
    log_activity,
)

PAGE_SIZE = 50
SEMESTER_RE = re.compile(r'^\d{4}/\d$')          # 2026/1
STUDENT_ID_RE = re.compile(r'^\d{7,10}$')          # digits only
SUBJECT_RE = re.compile(r'^\s*([A-Za-z]{2,6}\s*-?\s*\d{3,4}[A-Za-z]?)\s*[-:]?\s*(.*)$')   # "ITE497 สหกิจศึกษา"
VALID_YEARS = {'1', '2', '3', '4'}


def _valid_email(value):
    try:
        validate_email(value)
        return True
    except ValidationError:
        return False


DEFAULT_MIN_YEAR = 2          # year 1 students may not register internship subjects


def _year_block(year, code, min_year):
    """Reason text if a student of this study year may not register the subject, else None."""
    digits = re.sub(r'\D', '', str(year or ''))
    if digits and int(digits) < int(min_year):
        return f"Year {digits} students cannot register {code} (minimum year {min_year})"
    return None


def _one_per_semester_block(student, course, semester):
    """Reason text if the student already has a different subject in this semester, else None."""
    other = (Enrollment.objects.filter(student=student, semester=semester)
             .exclude(course=course).select_related('course').first())
    if other:
        return (f"Already registered in {other.course.code} for {semester} - "
                f"a student can register only one subject per semester")
    return None


def _split_subject(text):
    """'ITE497 สหกิจศึกษา' -> ('ITE497', 'สหกิจศึกษา'); returns (None, '') if it is not a subject code."""
    m = SUBJECT_RE.match(text or '')
    if not m:
        return None, ''
    return _norm_code(m.group(1)), m.group(2).strip()

# ----------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------
STAGE_RECEIVED, STAGE_DRAFT, STAGE_SIGNED, STAGE_SENT = 0, 1, 2, 3
STAGE_INFO = {
    STAGE_RECEIVED: ('Received', 'bg-gray-100 text-gray-700'),
    STAGE_DRAFT: ('Draft ready', 'bg-yellow-100 text-yellow-800'),
    STAGE_SIGNED: ('Signed', 'bg-blue-100 text-blue-700'),
    STAGE_SENT: ('Emailed', 'bg-green-100 text-green-700'),
}

# "refer" in the UI = the 'confirm' (referral/confirmation) letter type
BASE_KINDS = [
    ('request', 'Request'),
    ('confirm', 'Confirm / Referral'),
    ('extension', 'Extension'),
]


def _norm_code(value):
    """ITCS 499 / itcs-499 -> ITCS499"""
    return re.sub(r'[\s\-_/]+', '', str(value or '')).upper()


def _has_file(field):
    """Cheap check (no storage call): is a file name stored on the field?"""
    return bool(field and getattr(field, 'name', ''))


def _list_stage(doc):
    """Stage used on list pages. Does not touch storage, so it stays fast."""
    if doc is None:
        return STAGE_RECEIVED
    if doc.email_sent_at or doc.status == GeneratedDocument.STATUS_SENT:
        return STAGE_SENT
    if _has_file(doc.signed_file):
        return STAGE_SIGNED
    if _has_file(doc.file) or doc.status == GeneratedDocument.STATUS_GENERATED:
        return STAGE_DRAFT
    return STAGE_RECEIVED


def _kind_resolver():
    """Same logic as docgen.resolve_letter_type, but loads the mapping once."""
    mapping = {
        m.raw_text: m.letter_type.slug
        for m in LetterTypeMapping.objects.select_related('letter_type')
    }

    def resolve(raw):
        raw = str(raw or '').strip()
        return mapping.get(raw) or normalize_letter_type(raw)

    return resolve


def _kind_columns():
    """Letter-type columns: the 3 built-in ones + any extra LetterType rows."""
    cols = list(BASE_KINDS)
    known = {slug for slug, _ in cols}
    for lt in LetterType.objects.order_by('id'):
        if lt.slug not in known:
            cols.append((lt.slug, lt.display_name))
            known.add(lt.slug)
    return cols


def _course_matches(code_norm, record_course):
    rec = _norm_code(record_course)
    return bool(code_norm) and bool(rec) and (rec == code_norm or code_norm in rec)


def _build_rows(enrollments):
    """
    For every enrollment find the student's request records for that subject
    and work out the status of each letter type.

    Matching rule: same student ID and the record's course_id equals (or
    contains) the subject code, ignoring spaces/case. The semester is not part
    of the match because the record's semester is derived from the internship
    period; it is shown in the cell tooltip instead.
    """
    enrollments = list(enrollments)
    columns = _kind_columns()
    col_slugs = {slug for slug, _ in columns}
    resolve = _kind_resolver()

    student_ids = {e.student.student_id.strip() for e in enrollments}
    records = list(
        StudentRecord.objects.filter(student_id__in=student_ids).order_by('id')
    )
    by_student = defaultdict(list)
    for r in records:
        by_student[r.student_id.strip()].append(r)

    latest_doc = {}
    for d in GeneratedDocument.objects.filter(record__in=records).order_by('id'):
        latest_doc[d.record_id] = d            # later id wins

    used_other = False
    rows = []
    for e in enrollments:
        code_norm = _norm_code(e.course.code)
        cells = {}
        for r in by_student.get(e.student.student_id.strip(), []):
            if not _course_matches(code_norm, r.course_id):
                continue
            kind = resolve(r.letter_type)
            if kind not in col_slugs:
                kind = 'other'
                used_other = True
            doc = latest_doc.get(r.id)
            cells[kind] = {                    # newest record per letter type
                'record': r,
                'doc': doc,
                'stage': _list_stage(doc),
                'label': STAGE_INFO[_list_stage(doc)][0],
                'css': STAGE_INFO[_list_stage(doc)][1],
            }
        rows.append({
            'enrollment': e,
            'student': e.student,
            'course': e.course,
            'cells': cells,
            'requested': bool(cells),
            'best_stage': max((c['stage'] for c in cells.values()), default=None),
        })

    if used_other:
        columns = columns + [('other', 'Other')]
    for row in rows:
        row['cell_list'] = [row['cells'].get(slug) for slug, _ in columns]
    return rows, columns


# ----------------------------------------------------------------------
# 1. Request status (main page)
# ----------------------------------------------------------------------
@staff_required
def request_status(request):
    q = request.GET.get('q', '').strip()                 # subject text
    subject_id = request.GET.get('subject', '').strip()
    semester = request.GET.get('semester', '').strip()
    student_q = request.GET.get('student', '').strip()
    state = request.GET.get('state', 'all')

    enrollments = Enrollment.objects.select_related('student', 'course')
    if q:
        enrollments = enrollments.filter(
            Q(course__code__icontains=q) | Q(course__name__icontains=q) | Q(course__name_th__icontains=q)
        )
    if subject_id.isdigit():
        enrollments = enrollments.filter(course_id=int(subject_id))
    if semester:
        enrollments = enrollments.filter(semester=semester)
    if student_q:
        enrollments = enrollments.filter(
            Q(student__student_id__icontains=student_q) | Q(student__display_name__icontains=student_q)
        )
    enrollments = enrollments.order_by('course__code', 'student__student_id')

    rows, columns = _build_rows(enrollments)
    total = len(rows)
    requested = sum(1 for r in rows if r['requested'])
    emailed = sum(1 for r in rows if r['best_stage'] == STAGE_SENT)

    if state == 'requested':
        rows = [r for r in rows if r['requested']]
    elif state == 'not_requested':
        rows = [r for r in rows if not r['requested']]

    page = Paginator(rows, PAGE_SIZE).get_page(request.GET.get('page'))
    params = request.GET.copy()
    params.pop('page', None)

    return render(request, 'letters/request_status.html', {
        'page': page,
        'columns': columns,
        'total': total,
        'requested': requested,
        'not_requested': total - requested,
        'emailed': emailed,
        'courses': Course.objects.all(),
        'semesters': Enrollment.objects.order_by('-semester').values_list('semester', flat=True).distinct(),
        'f': {'q': q, 'subject': subject_id, 'semester': semester, 'student': student_q, 'state': state},
        'querystring': params.urlencode(),
        'has_roster': Enrollment.objects.exists(),
    })


# ----------------------------------------------------------------------
# 2. Document status (stand-alone page for one request)
# ----------------------------------------------------------------------
def _first(data, keys):
    for k in keys:
        v = (data or {}).get(k)
        if v and str(v).strip():
            return str(v).strip()
    return ''


@staff_required
def document_status(request, record_id):
    record = get_object_or_404(StudentRecord.objects.select_related('batch'), id=record_id)
    doc = GeneratedDocument.objects.filter(record=record).last()
    kind = _kind_resolver()(record.letter_type)
    kind_label = dict(_kind_columns()).get(kind, record.letter_type or '—')

    draft_ready = bool(doc and (_field_ready(doc.file) or doc.status in
                                (GeneratedDocument.STATUS_GENERATED, GeneratedDocument.STATUS_SENT)))
    skipped = bool(doc and doc.status == GeneratedDocument.STATUS_SKIPPED and not draft_ready)
    signed_ready = bool(doc and _field_ready(doc.signed_file))
    sent = bool(doc and (doc.email_sent_at or doc.status == GeneratedDocument.STATUS_SENT))

    try:
        email_to = get_student_email(record)
    except Exception:
        email_to = ''

    form_time = _first(record.data, ['Timestamp', 'timestamp', 'ประทับเวลา'])
    batch = record.batch

    steps = [
        {
            'title': 'Request received',
            'state': 'done',
            'time': batch.uploaded_at if batch else None,
            'detail': (f"Uploaded in batch #{batch.id} — {batch.file_name}" if batch else ''),
            'extra': (f"Form submitted: {form_time}" if form_time else ''),
            'links': [],
        },
        {
            'title': 'Draft letter generated',
            'state': 'done' if draft_ready else ('skipped' if skipped else 'todo'),
            'time': doc.created_at if (doc and draft_ready) else None,
            'detail': (f"Reference no. {doc.ref_no}" if doc and doc.ref_no and draft_ready else
                       (f"Skipped: {doc.skip_reason}" if skipped and doc.skip_reason else '')),
            'extra': '',
            'links': ([{'label': 'Open draft (.docx)', 'url': reverse('letters:preview_record_docx', args=[record.id])},
                       {'label': 'Preview PDF', 'url': reverse('letters:preview_record_pdf', args=[record.id])}]
                      if draft_ready else []),
        },
        {
            'title': 'Signed letter uploaded',
            'state': 'done' if signed_ready else 'todo',
            'time': None,
            'detail': 'Signed PDF is attached to this request.' if signed_ready else 'Waiting for the signed PDF.',
            'extra': '',
            'links': ([{'label': 'Open signed PDF', 'url': reverse('letters:view_signed_pdf', args=[doc.id])}]
                      if signed_ready else []),
        },
        {
            'title': 'Emailed to student',
            'state': 'done' if sent else 'todo',
            'time': doc.email_sent_at if (doc and sent) else None,
            'detail': (f"Sent to {email_to}" if sent and email_to else
                       (f"Will be sent to {email_to}" if email_to else 'No student email found in the data.')),
            'extra': '',
            'links': [],
        },
    ]

    # progress + "what to do next"
    done_count = sum(1 for s in steps if s['state'] == 'done')
    if done_count == 4:
        current = STAGE_SENT
        next_step = {'text': 'All steps are complete.', 'url': '', 'label': ''}
    elif skipped:
        current = STAGE_RECEIVED
        next_step = {'text': 'The draft was skipped. Check the reason and generate it again.',
                     'url': reverse('letters:preview', args=[batch.id]) if batch else '',
                     'label': 'Open batch preview'}
    elif not draft_ready:
        current = STAGE_RECEIVED
        next_step = {'text': 'The draft letter has not been generated yet.',
                     'url': reverse('letters:preview', args=[batch.id]) if batch else '',
                     'label': 'Open batch preview'}
    elif not signed_ready:
        current = STAGE_DRAFT
        next_step = {'text': 'Waiting for the signed PDF to be uploaded.',
                     'url': reverse('letters:signed_preview') + (f"?batch_id={batch.id}" if batch else ''),
                     'label': 'Upload signed papers'}
    else:
        current = STAGE_SIGNED
        next_step = {'text': 'Signed letter is ready to be emailed to the student.',
                     'url': reverse('letters:signed_preview') + (f"?batch_id={batch.id}" if batch else ''),
                     'label': 'Send emails'}

    student = Student.objects.filter(student_id=record.student_id.strip()).first()
    others = []
    docs_by_record = {}
    siblings = list(StudentRecord.objects.filter(student_id=record.student_id).exclude(id=record.id).order_by('-id'))
    for d in GeneratedDocument.objects.filter(record__in=siblings).order_by('id'):
        docs_by_record[d.record_id] = d
    resolve = _kind_resolver()
    for r in siblings:
        st = _list_stage(docs_by_record.get(r.id))
        others.append({'record': r, 'kind': resolve(r.letter_type), 'label': STAGE_INFO[st][0], 'css': STAGE_INFO[st][1]})

    return render(request, 'letters/document_status.html', {
        'record': record,
        'doc': doc,
        'student': student,
        'kind_label': kind_label,
        'steps': steps,
        'stage_label': STAGE_INFO[current][0],
        'stage_css': STAGE_INFO[current][1],
        'progress': int(done_count / 4 * 100),
        'next_step': next_step,
        'others': others,
    })


# ----------------------------------------------------------------------
# 3. Subjects
# ----------------------------------------------------------------------
@staff_required
def subject_list(request):
    q = request.GET.get('q', '').strip()
    courses = Course.objects.annotate(n_students=Count('enrollments'))
    if q:
        courses = courses.filter(Q(code__icontains=q) | Q(name__icontains=q) | Q(name_th__icontains=q))
    edit = None
    edit_id = request.GET.get('edit', '')
    if edit_id.isdigit():
        edit = Course.objects.filter(id=int(edit_id)).first()
    return render(request, 'letters/subject_list.html', {'courses': courses, 'edit': edit, 'q': q})


@staff_required
@require_POST
def subject_save(request):
    code = _norm_code(request.POST.get('code'))
    name = request.POST.get('name', '').strip()
    name_th = request.POST.get('name_th', '').strip()
    course_id = request.POST.get('id', '')
    min_year = request.POST.get('min_year', str(DEFAULT_MIN_YEAR)).strip()

    if not code:
        messages.error(request, "Subject code is required.")
        return redirect('letters:subject_list')
    if min_year not in VALID_YEARS:
        messages.error(request, "Minimum year must be 1-4.")
        return redirect('letters:subject_list')
    min_year = int(min_year)

    clash = Course.objects.filter(code=code)
    if course_id.isdigit():
        clash = clash.exclude(id=int(course_id))
    if clash.exists():
        messages.error(request, f"Subject {code} already exists.")
        return redirect('letters:subject_list')

    if course_id.isdigit():
        course = get_object_or_404(Course, id=int(course_id))
        course.code, course.name, course.name_th, course.min_year = code, name, name_th, min_year
        course.save()
        messages.success(request, f"Subject {code} updated.")
    else:
        Course.objects.create(code=code, name=name, name_th=name_th, min_year=min_year)
        messages.success(request, f"Subject {code} added.")
    log_activity(request, 'EDIT_SUBJECT', f"Saved subject {code}")
    return redirect('letters:subject_list')


@staff_required
@require_POST
def subject_delete(request, pk):
    course = get_object_or_404(Course, pk=pk)
    code = course.code
    course.delete()                                      # also removes its registrations
    log_activity(request, 'DELETE_SUBJECT', f"Deleted subject {code}")
    messages.success(request, f"Subject {code} deleted.")
    return redirect('letters:subject_list')


SUBJECT_IMPORT_ALIASES = {
    'code': ['code', 'subject', 'subjectcode', 'courseid', 'รหัสวิชา', 'รหัส'],
    'name': ['name', 'nameen', 'englishname', 'subjectname', 'ชื่อวิชา', 'ชื่อภาษาอังกฤษ'],
    'name_th': ['nameth', 'thainame', 'ชื่อไทย', 'ชื่อวิชาภาษาไทย'],
    'min_year': ['minyear', 'minimumyear', 'ชั้นปีขั้นต่ำ'],
}


@staff_required
@require_POST
def subject_import(request):
    """CSV/Excel with columns: code, name, name_th (code may also be 'ITE497 สหกิจศึกษา')."""
    f = request.FILES.get('file')
    if not f:
        messages.error(request, "Choose a CSV or Excel file first.")
        return redirect('letters:subject_list')
    try:
        df = _load_table(f)
    except Exception as e:
        messages.error(request, f"Could not read the file: {e}")
        return redirect('letters:subject_list')

    col = {}
    for header in df.columns:
        key = _header_key(header)
        for field, aliases in SUBJECT_IMPORT_ALIASES.items():
            if field not in col and key in aliases:
                col[field] = header
    if 'code' not in col:
        messages.error(request, "The file needs a subject code column (code).")
        return redirect('letters:subject_list')

    def cell(row, field):
        return str(row.get(col[field], '') if field in col else '').strip()

    created = updated = 0
    problems = []
    with transaction.atomic():
        for idx, row in enumerate(df.to_dict(orient='records')):
            line = idx + 2
            code, name_from_code = _split_subject(cell(row, 'code'))
            if not code:
                problems.append((line, f"'{cell(row, 'code')}' is not a valid subject code (e.g. ITE497) - row skipped"))
                continue
            name = cell(row, 'name')
            name_th = cell(row, 'name_th') or name_from_code
            min_year_text = cell(row, 'min_year')
            if min_year_text and min_year_text not in VALID_YEARS:
                problems.append((line, f"minimum year '{min_year_text}' must be 1-4 - {DEFAULT_MIN_YEAR} used"))
                min_year_text = ''
            min_year = int(min_year_text) if min_year_text else None
            course, was_created = Course.objects.get_or_create(
                code=code, defaults={'name': name, 'name_th': name_th, 'min_year': min_year or DEFAULT_MIN_YEAR})
            if was_created:
                created += 1
            else:
                changed = False
                if name and course.name != name:
                    course.name, changed = name, True
                if name_th and course.name_th != name_th:
                    course.name_th, changed = name_th, True
                if min_year and course.min_year != min_year:
                    course.min_year, changed = min_year, True
                if changed:
                    course.save()
                    updated += 1

    log_activity(request, 'IMPORT_SUBJECTS', f"Imported subjects: {created} new, {updated} updated, {len(problems)} problem(s)")
    messages.success(request, f"Imported {created} new and updated {updated} existing subject(s).")
    if problems:
        shown = "; ".join(f"row {n}: {t}" for n, t in problems[:10])
        messages.warning(request, f"{len(problems)} problem(s) found in the file. {shown}")
    return redirect('letters:subject_list')


# ----------------------------------------------------------------------
# 4. Students and registrations
# ----------------------------------------------------------------------
@staff_required
def student_list(request):
    q = request.GET.get('q', '').strip()
    year = request.GET.get('year', '').strip()
    students = Student.objects.prefetch_related('enrollments__course')
    if q:
        students = students.filter(Q(student_id__icontains=q) | Q(display_name__icontains=q))
    if year:
        students = students.filter(year=year)
    page = Paginator(students, PAGE_SIZE).get_page(request.GET.get('page'))
    params = request.GET.copy()
    params.pop('page', None)
    return render(request, 'letters/student_list.html', {
        'page': page,
        'q': q,
        'year': year,
        'years': Student.objects.exclude(year='').order_by('year').values_list('year', flat=True).distinct(),
        'querystring': params.urlencode(),
    })


@staff_required
def student_form(request, pk=None):
    student = get_object_or_404(Student, pk=pk) if pk else None
    courses = Course.objects.all()
    form = {
        'student_id': student.student_id if student else '',
        'display_name': student.display_name if student else '',
        'email': student.email if student else '',
        'year': student.year if student else '',
        'semester': '',
        'courses': [],
    }
    errors = []

    if request.method == 'POST':
        form.update({
            'student_id': request.POST.get('student_id', '').strip(),
            'display_name': request.POST.get('display_name', '').strip(),
            'email': request.POST.get('email', '').strip(),
            'year': request.POST.get('year', '').strip(),
            'semester': request.POST.get('semester', '').strip(),
            'courses': request.POST.getlist('courses'),
        })
        if not form['student_id']:
            errors.append("Student ID is required.")
        elif not STUDENT_ID_RE.match(form['student_id']):
            errors.append("Student ID must be 7-10 digits.")
        if form['email'] and not _valid_email(form['email']):
            errors.append("Email address is not valid.")
        if form['year'] and form['year'] not in VALID_YEARS:
            errors.append("Year must be 1-4.")
        if not form['display_name']:
            errors.append("Student name is required.")
        clash = Student.objects.filter(student_id=form['student_id'])
        if student:
            clash = clash.exclude(pk=student.pk)
        if form['student_id'] and clash.exists():
            errors.append(f"Student ID {form['student_id']} already exists.")
        picked = [int(c) for c in form['courses'] if str(c).isdigit()]
        if picked and not SEMESTER_RE.match(form['semester']):
            errors.append("Enter the semester as year/term, e.g. 2026/1, when you select subjects.")
        if len(set(picked)) > 1:
            errors.append("A student can register only one subject per semester - choose one subject.")
        for course in Course.objects.filter(id__in=set(picked)):
            reason = _year_block(form['year'], course.code, course.min_year)
            if reason:
                errors.append(reason + ".")
            elif not form['year'] and picked:
                errors.append("Select the student's year before registering a subject.")
            if student is not None and SEMESTER_RE.match(form['semester']):
                reason = _one_per_semester_block(student, course, form['semester'])
                if reason:
                    errors.append(reason + ".")

        if not errors:
            with transaction.atomic():
                if student is None:
                    student = Student()
                student.student_id = form['student_id']
                student.display_name = form['display_name']
                student.email = form['email']
                student.year = form['year']
                student.save()
                added = 0
                for cid in (c for c in form['courses'] if str(c).isdigit()):
                    if cid.isdigit():
                        _, created = Enrollment.objects.get_or_create(
                            student=student, course_id=int(cid), semester=form['semester'])
                        added += created
            log_activity(request, 'EDIT_STUDENT', f"Saved student {student.student_id} (+{added} registrations)")
            messages.success(request, f"Saved {student.student_id}" + (f" with {added} new registration(s)." if added else "."))
            return redirect('letters:student_detail', pk=student.pk)

    return render(request, 'letters/student_form.html', {
        'student': student,
        'form': form,
        'errors': errors,
        'courses': courses,
        'selected_courses': {int(c) for c in form['courses'] if str(c).isdigit()},
    })


@staff_required
def student_detail(request, pk):
    student = get_object_or_404(Student, pk=pk)
    enrollments = student.enrollments.select_related('student', 'course')
    rows, columns = _build_rows(enrollments)

    # requests from this student that match none of the registered subjects
    matched_ids = {c['record'].id for r in rows for c in r['cells'].values()}
    resolve = _kind_resolver()
    unmatched = []
    recs = list(StudentRecord.objects.filter(student_id=student.student_id).exclude(id__in=matched_ids))
    docs = {}
    for d in GeneratedDocument.objects.filter(record__in=recs).order_by('id'):
        docs[d.record_id] = d
    for r in recs:
        st = _list_stage(docs.get(r.id))
        unmatched.append({'record': r, 'kind': resolve(r.letter_type),
                          'label': STAGE_INFO[st][0], 'css': STAGE_INFO[st][1]})

    return render(request, 'letters/student_detail.html', {
        'student': student, 'rows': rows, 'columns': columns, 'unmatched': unmatched,
    })


@staff_required
@require_POST
def student_delete(request, pk):
    student = get_object_or_404(Student, pk=pk)
    sid = student.student_id
    student.delete()
    log_activity(request, 'DELETE_STUDENT', f"Deleted student {sid}")
    messages.success(request, f"Student {sid} deleted. Their request records are kept.")
    return redirect('letters:student_list')


@staff_required
@require_POST
def enrollment_delete(request, pk):
    enrollment = get_object_or_404(Enrollment.objects.select_related('student', 'course'), pk=pk)
    student_pk = enrollment.student_id
    enrollment.delete()
    messages.success(request, "Registration removed.")
    return redirect('letters:student_detail', pk=student_pk)


# ----------------------------------------------------------------------
# 5. Import students + registrations from CSV / Excel
# ----------------------------------------------------------------------
IMPORT_ALIASES = {
    'student_id': ['studentid', 'id', 'รหัสนักศึกษา', 'รหัส'],
    'name': ['name', 'fullname', 'displayname', 'ชื่อ', 'ชื่อนามสกุล', 'ชื่อสกุล'],
    'year': ['year', 'ชั้นปี', 'ปี'],
    'email': ['email', 'อีเมล'],
    'semester': ['semester', 'ภาคเรียน', 'เทอม'],
    'subjects': ['subject', 'subjects', 'course', 'courses', 'courseid', 'รหัสวิชา', 'วิชา'],
}


def _header_key(h):
    return re.sub(r'[\s_\-\.]+', '', str(h or '').lower())


# ---- "are you sure?" step: hold the file between the check and the import ----
def _stash_upload(request, f, kind, **extra):
    """Save the uploaded file temporarily and remember it in the session. Returns a token."""
    token = uuid.uuid4().hex
    f.seek(0)
    path = default_storage.save(f"pending_imports/{token}/{f.name}", f)
    pending = request.session.get('pending_imports', {})
    pending[token] = {'path': path, 'name': f.name, 'kind': kind, **extra}
    request.session['pending_imports'] = pending
    request.session.modified = True
    return token


def _load_stash(request, token, kind):
    """-> (open file, info) or (None, None) if the token is unknown/expired."""
    info = request.session.get('pending_imports', {}).get(token)
    if not info or info.get('kind') != kind:
        return None, None
    try:
        return default_storage.open(info['path']), info
    except Exception:
        return None, None


def _drop_stash(request, token):
    pending = request.session.get('pending_imports', {})
    info = pending.pop(token, None)
    if info:
        try:
            default_storage.delete(info['path'])
        except Exception:
            pass
    request.session['pending_imports'] = pending
    request.session.modified = True


def _problem(line, level, who, text):
    """level: 'skip' = this data will not be imported, 'warn' = imported but please check."""
    return {'where': f"Row {line}", 'level': level, 'who': who, 'text': text}


def _import_students_table(df, default_semester, commit):
    """
    Read the table and create/update students, subjects and registrations.
    commit=False is a dry run: everything is rolled back, only the report is returned.
    """
    col = {}
    for header in df.columns:
        key = _header_key(header)
        for field, aliases in IMPORT_ALIASES.items():
            if field not in col and key in aliases:
                col[field] = header
    if 'student_id' not in col or 'name' not in col:
        return {'error': "The file needs at least a student ID column and a name column."}

    def cell(row, field):
        return str(row.get(col[field], '') if field in col else '').strip()

    regs = rows_ok = 0
    new_ids, updated_ids = set(), set()      # distinct students (a student can appear on several rows)
    new_courses = set()
    problems = []
    seen = {}                     # student_id -> (row number, name) for in-file duplicates
    seen_regs = set()

    with transaction.atomic():
        for idx, row in enumerate(df.to_dict(orient='records')):
            line = idx + 2        # row number as shown in Excel (header = row 1)
            sid, name = cell(row, 'student_id'), cell(row, 'name')

            if not sid:
                problems.append(_problem(line, 'skip', name or '(no ID)', "Student ID is empty - row not imported"))
                continue
            if not STUDENT_ID_RE.match(sid):
                problems.append(_problem(line, 'skip', sid, f"Student ID '{sid}' must be 7-10 digits - row not imported"))
                continue
            if not name:
                problems.append(_problem(line, 'skip', sid, "Name is empty - row not imported"))
                continue

            if sid in seen and seen[sid][1] != name:
                problems.append(_problem(line, 'warn', sid,
                    f"Duplicate ID with a different name: '{name}' here vs '{seen[sid][1]}' on row {seen[sid][0]} - first name kept"))
                name = seen[sid][1]
            seen.setdefault(sid, (line, name))

            student, created = Student.objects.get_or_create(student_id=sid, defaults={'display_name': name})
            if created:
                new_ids.add(sid)
            else:
                if sid not in new_ids:
                    updated_ids.add(sid)
                if student.display_name != name:
                    problems.append(_problem(line, 'warn', sid,
                        f"Already in the system as '{student.display_name}', file says '{name}' - saved name kept"))

            year = re.sub(r'\D', '', cell(row, 'year'))
            if cell(row, 'year'):
                if year in VALID_YEARS:
                    student.year = year
                else:
                    problems.append(_problem(line, 'warn', sid, f"Year '{cell(row, 'year')}' must be 1-4 - year not saved"))

            email = cell(row, 'email')
            if email:
                if _valid_email(email):
                    student.email = email
                else:
                    problems.append(_problem(line, 'warn', sid, f"Email '{email}' is not valid - email not saved"))
            student.save()
            rows_ok += 1

            semester = cell(row, 'semester') or default_semester
            for part in filter(None, (p.strip() for p in re.split(r'[,;\n]+', cell(row, 'subjects')))):
                code, course_name = _split_subject(part)
                if not code:
                    problems.append(_problem(line, 'skip', sid, f"Subject '{part}' is not a valid code (e.g. ITE497) - not registered"))
                    continue
                if not SEMESTER_RE.match(semester):
                    problems.append(_problem(line, 'skip', sid, f"Semester '{semester}' must look like 2026/1 - {code} not registered"))
                    continue
                if (sid, code, semester) in seen_regs:
                    problems.append(_problem(line, 'warn', sid, f"{code} is listed again for {semester} - counted once"))
                    continue
                seen_regs.add((sid, code, semester))

                # rule 1: minimum study year   (a subject that does not exist yet gets the default)
                existing = Course.objects.filter(code=code).first()
                min_year = existing.min_year if existing else DEFAULT_MIN_YEAR
                eff_year = student.year
                reason = _year_block(eff_year, code, min_year)
                if reason:
                    problems.append(_problem(line, 'skip', sid, reason + " - not registered"))
                    continue
                if not re.sub(r'\D', '', eff_year or ''):
                    problems.append(_problem(line, 'warn', sid, f"Year is not known, so the minimum year for {code} was not checked"))

                # rule 2: one subject per semester
                if existing:
                    reason = _one_per_semester_block(student, existing, semester)
                    if reason:
                        problems.append(_problem(line, 'skip', sid, f"{code}: {reason} - not registered"))
                        continue

                course, c_created = Course.objects.get_or_create(code=code, defaults={'name': course_name})
                if c_created:
                    new_courses.add(code)
                    problems.append(_problem(line, 'warn', sid, f"Subject {code} is not in the Subjects list - it will be created"))
                elif course_name and not course.name:
                    course.name = course_name
                    course.save()
                _, e_created = Enrollment.objects.get_or_create(student=student, course=course, semester=semester)
                regs += e_created

        if not commit:
            transaction.set_rollback(True)          # dry run: keep nothing

    return {'new_students': len(new_ids), 'updated': len(updated_ids), 'regs': regs, 'rows_ok': rows_ok,
            'total_rows': len(df), 'new_courses': new_courses, 'problems': problems}


@staff_required
@require_POST
def student_import(request):
    token = request.POST.get('token', '')
    default_semester = request.POST.get('default_semester', '').strip()

    if token:                                                   # second step: answer to "are you sure?"
        if request.POST.get('action') == 'cancel':
            _drop_stash(request, token)
            messages.info(request, "Import cancelled. Nothing was saved.")
            return redirect('letters:student_list')
        f, info = _load_stash(request, token, 'students')
        if f is None:
            messages.error(request, "The pending import expired. Please choose the file again.")
            return redirect('letters:student_list')
        default_semester = info.get('default_semester', '')
        file_name = info['name']
    else:                                                       # first step: new file
        f = request.FILES.get('file')
        if not f:
            messages.error(request, "Choose a CSV or Excel file first.")
            return redirect('letters:student_list')
        if default_semester and not SEMESTER_RE.match(default_semester):
            messages.error(request, "Default semester must look like 2026/1.")
            return redirect('letters:student_list')
        file_name = f.name

    try:
        df = _load_table(f)
    except Exception as e:
        messages.error(request, f"Could not read the file: {e}")
        if token:
            _drop_stash(request, token)
        return redirect('letters:student_list')

    if not token:
        dry = _import_students_table(df, default_semester, commit=False)
        if dry.get('error'):
            messages.error(request, dry['error'])
            return redirect('letters:student_list')
        if dry['problems']:                                     # ask before importing
            token = _stash_upload(request, f, 'students', default_semester=default_semester)
            n_skip = sum(1 for p in dry['problems'] if p['level'] == 'skip')
            return render(request, 'letters/import_confirm.html', {
                'title': 'Check the student file before importing',
                'n_problems': len(dry['problems']),
                'file_name': file_name,
                'token': token,
                'action_url': reverse('letters:student_import'),
                'cancel_url': reverse('letters:student_list'),
                'summary': [
                    ('Rows in file', dry['total_rows']),
                    ('New students', dry['new_students']),
                    ('Already in the system (will be updated)', dry['updated']),
                    ('New registrations', dry['regs']),
                    ('Problems found', len(dry['problems'])),
                ],
                'problems': dry['problems'][:200],
                'more': max(0, len(dry['problems']) - 200),
                'confirm_note': (f"“Import anyway” saves everything that is valid. "
                                 f"{n_skip} item(s) marked “Not imported” are left out; items marked “Check” are saved as described."),
            })

    result = _import_students_table(df, default_semester, commit=True)
    if token:
        _drop_stash(request, token)
    problems = result['problems']
    log_activity(request, 'IMPORT_STUDENTS',
                 f"Imported {result['new_students']} new / {result['updated']} existing students, "
                 f"{result['regs']} registrations, {len(problems)} problem(s) confirmed")
    msg = f"Imported {result['new_students']} new and {result['updated']} existing student(s), {result['regs']} new registration(s)."
    if result['new_courses']:
        msg += f" New subjects created: {', '.join(sorted(result['new_courses']))}."
    messages.success(request, msg)
    if problems:
        shown = "; ".join(f"{p['where'].lower()}: {p['text']}" for p in problems[:10])
        more = f" (+{len(problems) - 10} more)" if len(problems) > 10 else ""
        messages.warning(request, f"{len(problems)} problem(s) were in the file. {shown}{more}")
    return redirect('letters:student_list')


# ----------------------------------------------------------------------
# 6. Data check: compare every request with the student / subject lists
# ----------------------------------------------------------------------
ERROR, WARN = 'error', 'warn'

ISSUE_LABELS = {
    'bad_id': 'Invalid student ID',
    'no_student': 'Student not in list',
    'no_letter_type': 'Letter type empty',
    'bad_letter_type': 'Letter type unknown',
    'no_subject': 'Subject empty',
    'bad_subject': 'Subject unknown',
    'not_registered': 'Not registered in subject',
    'bad_year': 'Invalid year',
    'year_not_allowed': 'Year not allowed for subject',
    'two_subjects': 'Two subjects in one semester',
    'year_mismatch': 'Year differs from list',
    'name_mismatch': 'Name differs from list',
    'missing_field': 'Required field empty',
    'bad_period': 'Period unreadable',
    'bad_email': 'Invalid email',
    'duplicate': 'Duplicate request',
}

_NAME_PREFIX_RE = re.compile(r'^(นางสาว|นาง|นาย|น\.ส\.?|mrs\.?|mr\.?|ms\.?|miss|dr\.?)', re.I)


def _name_key(value):
    """Name without title, spaces or punctuation, for comparing two spellings."""
    s = re.sub(r'\s+', '', str(value or '')).lower()
    s = _NAME_PREFIX_RE.sub('', s)
    return re.sub(r'[^a-z0-9฀-๿]', '', s)


def _period_problem(period):
    """None if the period reads as 'D M YYYY - D M YYYY' with end >= start, else a short reason."""
    text = (period or '').strip()
    if not text:
        return None                                   # reported as an empty field instead
    m = re.search(r'\((.*?)\)', text)
    parts = (m.group(1) if m else text).split('-')
    if len(parts) != 2:
        return 'is not a date range such as "1 พ.ย. 2568 - 31 ม.ค. 2569"'
    start, end = _parse_thai_date(parts[0]), _parse_thai_date(parts[1])
    if not start or not end or end[2] is None:
        return 'dates could not be read'
    try:
        end_d = date(end[2], end[1], end[0])
        start_year = start[2] or end[2]
        start_d = date(start_year, start[1], start[0])
        if start[2] is None and start_d > end_d:      # "1 พ.ย. - 31 ม.ค. 2569" crosses a year
            start_d = date(start_year - 1, start[1], start[0])
    except ValueError:
        return 'contains a date that does not exist'
    if end_d < start_d:
        return 'ends before it starts'
    return None


def _check_records(records):
    """-> list of {'record', 'issues': [(level, code, message)], 'worst'} in record order."""
    students = {s.student_id.strip(): s for s in Student.objects.all()}
    enrolled = defaultdict(set)
    for e in Enrollment.objects.all():
        enrolled[e.student_id].add(e.course_id)
    courses = [(_norm_code(c.code), c) for c in Course.objects.all()]
    # The comparisons below only make sense once the lists are in use.
    use_students, use_courses = bool(students), bool(courses)
    known_kinds = {slug for slug, _ in _kind_columns()}
    resolve = _kind_resolver()

    # which subject does each request belong to? (needed for the "one subject per semester" check)
    def subject_of(r):
        norm = _norm_code(r.course_id)
        return next((c for code, c in courses if code and code in norm), None) if norm else None

    subjects_by_semester = defaultdict(set)
    for r in records:
        c = subject_of(r)
        if c is not None and (r.student_id or '').strip() and r.semester:
            subjects_by_semester[((r.student_id or '').strip(), r.semester)].add(c.code)

    results = []
    for r in records:
        issues = []
        sid = (r.student_id or '').strip()

        # student
        stu = None
        if not STUDENT_ID_RE.match(sid):
            issues.append((ERROR, 'bad_id', f"Student ID '{sid}' must be 7-10 digits"))
        else:
            stu = students.get(sid)
            if stu is None and use_students:
                issues.append((ERROR, 'no_student', f"Student {sid} is not in the student list"))

        # letter type
        raw_type = (r.letter_type or '').strip()
        if not raw_type:
            issues.append((ERROR, 'no_letter_type', "Letter type is empty (the system would treat it as a Request letter)"))
        elif resolve(raw_type) not in known_kinds:
            issues.append((ERROR, 'bad_letter_type', f"Letter type '{raw_type}' is not recognised"))

        # subject
        subject_text = (r.course_id or '').strip()
        if not subject_text:
            issues.append((ERROR, 'no_subject', "Subject is empty"))
        else:
            norm = _norm_code(subject_text)
            course = next((c for code, c in courses if code and code in norm), None)
            if course is None and use_courses:
                issues.append((ERROR, 'bad_subject', f"Subject '{subject_text}' is not in the Subjects list"))
            elif course is not None:
                form_year = re.sub(r'\D', '', r.year or '')
                known_year = stu.year if (stu is not None and stu.year) else form_year
                reason = _year_block(known_year, course.code, course.min_year)
                if reason:
                    issues.append((ERROR, 'year_not_allowed', reason))
                elif stu is not None and course.id not in enrolled[stu.id]:
                    issues.append((ERROR, 'not_registered', f"{stu.student_id} is not registered in {course.code}"))
                taken = subjects_by_semester.get((sid, r.semester), set())
                if len(taken) > 1:
                    issues.append((WARN, 'two_subjects',
                                   f"Requests for {', '.join(sorted(taken))} are in the same semester {r.semester} - "
                                   f"a student can register only one subject per semester"))

        # year
        year = re.sub(r'\D', '', r.year or '')
        if year and year not in VALID_YEARS:
            issues.append((ERROR, 'bad_year', f"Year '{r.year}' must be 1-4"))
        elif year and stu is not None and stu.year and stu.year != year:
            issues.append((WARN, 'year_mismatch', f"Year is {year} but the student list says {stu.year}"))

        # name
        if stu is not None and r.display_name.strip() and _name_key(stu.display_name) != _name_key(r.display_name):
            issues.append((WARN, 'name_mismatch', f"Name '{r.display_name.strip()}' differs from the list ('{stu.display_name}')"))

        # required fields
        missing = [label for label, value in (
            ('student name', r.display_name), ('year', r.year), ('company', r.company),
            ('contact person', r.contractor), ('internship position', r.internship_position),
            ('period', r.period)) if not str(value or '').strip()]
        if missing:
            issues.append((ERROR, 'missing_field', "Empty: " + ", ".join(missing)))

        # period
        problem = _period_problem(r.period)
        if problem:
            issues.append((ERROR, 'bad_period', f"Period {problem}"))

        # emails anywhere in the form row
        for key, value in (r.data or {}).items():
            text = str(value or '').strip()
            if text and ('mail' in str(key).lower() or 'อีเมล' in str(key)) and not _valid_email(text):
                label = str(key).split('\n')[0][:40]
                issues.append((WARN, 'bad_email', f"Email '{text}' in '{label}' is not valid"))

        # repeat submission
        if (r.times or 1) > 1:
            issues.append((WARN, 'duplicate', f"Same request was submitted before (time #{r.times})"))

        worst = ERROR if any(i[0] == ERROR for i in issues) else (WARN if issues else None)
        results.append({'record': r, 'issues': [{'level': l, 'code': c, 'text': t} for l, c, t in issues], 'worst': worst})
    return results


@staff_required
def data_check(request):
    batches = list(UploadedBatch.objects.order_by('-uploaded_at'))
    batch_id = request.GET.get('batch_id') or (str(batches[0].id) if batches else '')
    show = request.GET.get('show', 'issues')              # issues | ok | all

    records = StudentRecord.objects.all()
    if batch_id.isdigit():
        records = records.filter(batch_id=int(batch_id))
    results = _check_records(list(records.order_by('batch_id', 'row_index', 'id')))

    n_error = sum(1 for x in results if x['worst'] == ERROR)
    n_warn = sum(1 for x in results if x['worst'] == WARN)
    n_ok = sum(1 for x in results if x['worst'] is None)
    kinds = Counter(i['code'] for x in results for i in x['issues'])
    kind_summary = [(code, ISSUE_LABELS[code], n) for code, n in kinds.most_common()]

    if show == 'issues':
        results = [x for x in results if x['worst']]
    elif show == 'ok':
        results = [x for x in results if not x['worst']]

    page = Paginator(results, PAGE_SIZE).get_page(request.GET.get('page'))
    params = request.GET.copy()
    params.pop('page', None)
    return render(request, 'letters/data_check.html', {
        'page': page,
        'batches': batches,
        'batch_id': batch_id,
        'show': show,
        'total': n_error + n_warn + n_ok,
        'n_error': n_error,
        'n_warn': n_warn,
        'n_ok': n_ok,
        'kind_summary': kind_summary,
        'querystring': params.urlencode(),
        'has_roster': Student.objects.exists(),
    })
