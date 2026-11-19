# -*- coding: utf-8 -*-
"""Document versions, signed upload by hand, letter settings, accounts, template guidance."""
import io
import os
import re

from django.contrib import messages
from django.contrib.auth.models import User
from django.core.files.base import ContentFile
from django.db import transaction
from django.http import FileResponse, Http404, HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST
from docx import Document

from .decorators import admin_required, get_role, staff_required
from .docgen import convert_docx_bytes_to_pdf, generate_for_row, resolve_letter_type
from .models import GeneratedDocument, LetterSettings, StaffProfile, StudentRecord
from .views import _field_ready, log_activity

DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'


# ----------------------------------------------------------------------
# Document versions: only the latest version is kept
# ----------------------------------------------------------------------
def save_document_version(record, docx_bytes, ref_no='', file_name=''):
    """Store docx_bytes as the newest version of this request's letter.

    The row is updated in place, so the previous file is deleted and only the
    latest version remains. A signed file and the email date stay attached.
    """
    kind = resolve_letter_type(record.letter_type) or 'letter'
    name = file_name or f"{record.student_id}_{kind}.docx"
    with transaction.atomic():
        doc = GeneratedDocument.objects.select_for_update().filter(record=record).order_by('-id').first()
        if doc is None:
            doc = GeneratedDocument(record=record, ref_no=ref_no, version=0)
        old_file = doc.file if (doc.pk and doc.file) else None
        doc.version = (doc.version or 0) + 1
        doc.status = GeneratedDocument.STATUS_GENERATED
        doc.file_name = name
        if ref_no and not doc.ref_no:
            doc.ref_no = ref_no
        doc.save()
        if old_file:
            old_file.delete(save=False)              # keep only the latest version
        doc.file.save(os.path.basename(name), ContentFile(docx_bytes), save=True)
    return doc


@staff_required
def record_docx_download(request, record_id):
    """Download the current Word file so staff can fix it; creates version 1 if none exists yet."""
    record = get_object_or_404(StudentRecord, id=record_id)
    doc = GeneratedDocument.objects.filter(record=record).order_by('-id').first()
    if not doc or not _field_ready(doc.file):
        buf, fname = generate_for_row(record.data, ref_no=str(record.student_id), semester=record.semester)
        if not buf:
            raise Http404(fname)
        doc = save_document_version(record, buf.getvalue(), ref_no=str(record.student_id))
    doc.file.open('rb')
    return FileResponse(doc.file, as_attachment=True,
                        filename=f"{record.student_id}_v{doc.version}.docx", content_type=DOCX_TYPE)


@staff_required
@require_POST
def record_docx_upload(request, record_id):
    record = get_object_or_404(StudentRecord, id=record_id)
    back = (redirect('letters:record_detail', record_id=record.id)
            if request.POST.get('next') == 'detail' else redirect('letters:document_status', record_id=record.id))
    f = request.FILES.get('edited_file')
    if not f or not f.name.lower().endswith('.docx'):
        messages.error(request, "Choose the edited Word file (.docx).")
        return back
    data = f.read()
    try:
        Document(io.BytesIO(data))                    # must open as a Word file
    except Exception:
        messages.error(request, "That file is not a valid .docx and was not saved.")
        return back
    doc = save_document_version(record, data)
    log_activity(request, 'EDIT_DOCUMENT', f"Uploaded edited letter for {record.student_id}, now version {doc.version}")
    messages.success(request, f"Saved as version {doc.version}. The earlier version was replaced.")
    return back


@staff_required
def record_pages(request, record_id):
    """Number of pages of the current letter (the preview shows it)."""
    record = get_object_or_404(StudentRecord, id=record_id)
    doc = GeneratedDocument.objects.filter(record=record).order_by('-id').first()
    try:
        if doc and _field_ready(doc.file):
            with doc.file.open('rb') as fh:
                data = fh.read()
        else:
            buf, _ = generate_for_row(record.data, ref_no=str(record.student_id), semester=record.semester)
            data = buf.getvalue()
        pdf = convert_docx_bytes_to_pdf(data)
        return JsonResponse({'pages': count_pdf_pages(pdf)})
    except Exception as e:
        return JsonResponse({'pages': None, 'error': str(e)}, status=500)


def count_pdf_pages(pdf_bytes):
    try:
        from pypdf import PdfReader
        return len(PdfReader(io.BytesIO(pdf_bytes)).pages)
    except ImportError:
        return len(re.findall(rb'/Type\s*/Page[^s]', pdf_bytes))


# ----------------------------------------------------------------------
# Signed letter: choose the student, the system renames the file
# ----------------------------------------------------------------------
@staff_required
@require_POST
def signed_upload_manual(request):
    f = request.FILES.get('signed_file')
    record = StudentRecord.objects.filter(id=request.POST.get('record_id') or 0).first()
    if not f or not f.name.lower().endswith('.pdf') or record is None:
        messages.error(request, "Choose a PDF and the student it belongs to.")
        return redirect('letters:signed_preview')
    kind = resolve_letter_type(record.letter_type) or 'letter'
    new_name = f"{record.student_id}_{kind}_signed.pdf"
    doc = GeneratedDocument.objects.filter(record=record).order_by('-id').first()
    if doc is None:
        doc = GeneratedDocument.objects.create(record=record, ref_no=str(record.student_id),
                                               status=GeneratedDocument.STATUS_GENERATED, file_name=new_name)
    if doc.signed_file:
        doc.signed_file.delete(save=False)           # replace the earlier signed file
    doc.signed_at = timezone.now()
    doc.signed_file.save(new_name, f, save=True)
    log_activity(request, 'UPLOAD_SIGNED', f"Signed file for {record.student_id} saved as {new_name}")
    messages.success(request, f"Saved for {record.display_name} as {new_name}.")
    return redirect('letters:signed_preview')


# ----------------------------------------------------------------------
# Configuration: Email templates (one default + optional one per letter type)
# ----------------------------------------------------------------------
@admin_required
def email_templates(request):
    from .models import EmailTemplate, LetterType
    from .views import _fill_placeholders
    types = list(LetterType.objects.order_by('display_name'))
    own = {t.letter_type_id: t for t in EmailTemplate.objects.exclude(letter_type__isnull=True)}
    key = request.GET.get('type') or request.POST.get('type') or 'default'
    lt = next((t for t in types if t.slug == key), None)
    if key != 'default' and lt is None:
        key = 'default'
    row = own.get(lt.id) if lt else EmailTemplate.load()

    if request.method == 'POST':
        action = request.POST.get('action', 'save')
        if action == 'remove' and lt and row:
            row.delete()
            log_activity(request, 'EDIT_EMAIL_MESSAGE', f"Removed the own email message of {lt.display_name}")
            messages.success(request, f"{lt.display_name} now uses the default email message.")
            return redirect(f"{request.path}?type={key}")
        subject = ' '.join(request.POST.get('email_subject', '').splitlines()).strip()
        body = request.POST.get('email_body', '').replace('\r\n', '\n').strip()
        if action == 'reset':
            subject, body = EmailTemplate.DEFAULT_SUBJECT, EmailTemplate.DEFAULT_BODY
        if not subject or not body:
            messages.error(request, "Subject and message cannot be empty.")
            return redirect(f"{request.path}?type={key}")
        if lt and not row:
            row = EmailTemplate(letter_type=lt)
        row.subject, row.body, row.updated_by = subject, body, request.user
        row.save()
        log_activity(request, 'EDIT_EMAIL_MESSAGE', f"Saved the email message for {lt.display_name if lt else 'all letter types'}")
        messages.success(request, "Email message saved. It is used the next time emails are sent.")
        return redirect(f"{request.path}?type={key}")

    default_row = EmailTemplate.load()
    items = [{'slug': 'default', 'name': 'Default (all letter types)', 'custom': True}]
    items += [{'slug': t.slug, 'name': t.display_name, 'custom': t.id in own} for t in types]

    class _Sample:      # fictional student, so the preview shows what a student will read
        display_name, student_id, company = 'Somchai Rakthai', '6687001', 'Siam Digital Co., Ltd.'
    shown = row or default_row
    sample_subject = _fill_placeholders(shown.subject, _Sample, request.user.get_full_name() or request.user.username, request.user.email or 'staff@example.com')
    sample_body = _fill_placeholders(shown.body, _Sample, request.user.get_full_name() or request.user.username, request.user.email or 'staff@example.com')
    return render(request, 'letters/email_templates.html', {
        'items': items, 'key': key, 'lt': lt, 'row': row, 'shown': shown, 'uses_default': bool(lt and not row),
        'default_subject': EmailTemplate.DEFAULT_SUBJECT, 'default_body': EmailTemplate.DEFAULT_BODY,
        'sample_subject': sample_subject, 'sample_body': sample_body,
    })


# ----------------------------------------------------------------------
# Accounts (admin only)
# ----------------------------------------------------------------------
@admin_required
def user_list(request):
    users = User.objects.select_related('profile').order_by('username')
    return render(request, 'letters/user_list.html', {'users': users})


@admin_required
def user_form(request, pk=None):
    user = get_object_or_404(User, pk=pk) if pk else None
    profile = getattr(user, 'profile', None) if user else None
    if request.method == 'POST':
        p = request.POST
        username = p.get('username', '').strip()
        password = p.get('password', '')
        role = p.get('role', StaffProfile.ROLE_STAFF)
        if role not in dict(StaffProfile.ROLE_CHOICES):
            role = StaffProfile.ROLE_STAFF
        if not username or (user is None and len(password) < 8):
            messages.error(request, "Enter a username and a password of at least 8 characters.")
        elif User.objects.filter(username=username).exclude(pk=pk or 0).exists():
            messages.error(request, "That username is taken.")
        else:
            if user is None:
                user = User(username=username, is_staff=True)
            user.username = username
            user.first_name = p.get('first_name', '').strip()
            user.last_name = p.get('last_name', '').strip()          # surname
            user.email = p.get('email', '').strip()
            user.is_active = p.get('is_active', 'on') == 'on' if pk else True
            if password:
                user.set_password(password)
            user.save()
            StaffProfile.objects.update_or_create(
                user=user, defaults={'role': role, 'expires_on': p.get('expires_on') or None})
            log_activity(request, 'EDIT_ACCOUNT', f"{'Updated' if pk else 'Created'} account {username} ({role})")
            messages.success(request, f"Account {username} saved.")
            return redirect('letters:user_list')
    return render(request, 'letters/user_form.html', {'u': user, 'profile': profile, 'roles': StaffProfile.ROLE_CHOICES})


# ----------------------------------------------------------------------
# Template help: a starter file with every tag in it
# ----------------------------------------------------------------------
@staff_required
def template_starter(request):
    from .docgen import AVAILABLE_SOURCE_FIELDS, THAI_FONT, _set_font, _base_doc
    doc = _base_doc()
    p = doc.add_paragraph()
    _set_font(p.add_run("ที่ {{ ref_no }}"))
    _set_font(doc.add_paragraph().add_run("วันที่ {{ date_th }}"))
    _set_font(doc.add_paragraph().add_run("เรียน {{ contact_person }} บริษัท {{ company }}"))
    _set_font(doc.add_paragraph().add_run(
        "ด้วย {{ student_name }} รหัสนักศึกษา {{ student_id }} ชั้นปีที่ {{ year_level }} "
        "ขอเข้าฝึกงานในช่วง {{ period }} รายวิชา {{ course }}"))
    _set_font(doc.add_paragraph().add_run("{{ dean_name }}"))
    _set_font(doc.add_paragraph().add_run("{{ dean_title }}"))
    doc.add_page_break()
    _set_font(doc.add_paragraph().add_run("Tags you can use (delete this page):"), bold=True)
    for tag, label in AVAILABLE_SOURCE_FIELDS:
        _set_font(doc.add_paragraph().add_run("{{ %s }}  =  %s" % (tag, label)))
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return FileResponse(buf, as_attachment=True, filename='starter_template.docx', content_type=DOCX_TYPE)
