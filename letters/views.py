# -*- coding: utf-8 -*-
import glob
import io
import os
import re
import subprocess
import tempfile
import zipfile

import pandas as pd
from django.contrib import messages
from django.contrib.auth import login, logout, update_session_auth_hash
from django.contrib.auth.forms import AuthenticationForm, PasswordChangeForm
from django.core.files.base import ContentFile
from django.core.mail import EmailMessage
from django.http import FileResponse, Http404
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.urls import reverse

from .decorators import staff_required
from .docgen import (
    SUPPORTED_LETTER_TYPES,
    extract_fields_from_docx,
    generate_for_row,
    generate_pdf_for_row,
)
from .models import AuditLog, GeneratedDocument, GenerationBatch, StudentRecord, UploadedBatch

IMPORTANT_FIELDS = ['student_id', 'display_name', 'company', 'contractor', 'internship_position', 'period']

PRIORITY_FIELDS = [
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
]

EXCLUDE_KEYS = {'id', 'batch', 'row_index', 'supported', 'missing', 'batch_id'}


def log_activity(request, action, details=""):
    """Helper function to record staff actions in AuditLog."""
    if not request.user or not request.user.is_authenticated:
        return

    x_forwarded_for = request.META.get('HTTP_X_FORWARDED_FOR')
    if x_forwarded_for:
        ip = x_forwarded_for.split(',')[0].strip()
    else:
        ip = request.META.get('REMOTE_ADDR')

    AuditLog.objects.create(
        user=request.user,
        action=action,
        details=details,
        ip_address=ip
    )


# ==========================================
# AUTHENTICATION & PROFILE VIEWS
# ==========================================

def login_view(request):
    """Renders login form and handles staff authentication."""
    if request.user.is_authenticated:
        return redirect('letters:index')

    if request.method == 'POST':
        form = AuthenticationForm(request, data=request.POST)
        if form.is_valid():
            user = form.get_user()
            if not user.is_staff:
                messages.error(request, "This account does not have staff permissions.")
                return render(request, 'letters/login.html', {'form': form})
            login(request, user)

            request.session.set_expiry(1800)  # 30-minute inactivity timeout

            next_url = request.GET.get('next')
            return redirect(next_url if next_url else 'letters:index')
        else:
            messages.error(request, "Invalid username or password.")
    else:
        form = AuthenticationForm()
    return render(request, 'letters/login.html', {'form': form})


def logout_view(request):
    """Terminates the user session."""
    logout(request)
    return redirect('letters:login')


@staff_required
def profile_view(request):
    """Allows logged-in staff to manage profile info and update password."""
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
# HELPER FUNCTIONS
# ==========================================

def _load_table(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith('.csv'):
        df = pd.read_csv(uploaded_file, dtype=str).fillna('')
    else:
        df = pd.read_excel(uploaded_file, dtype=str).fillna('')
    return df


def _norm(s):
    return (s or '').strip().lower()


def _clean_str(s):
    """Strips honorifics, spaces, underscores, and punctuation for fuzzy matching."""
    if not s:
        return ''
    s = str(s).lower()
    prefixes = ['mr.', 'mr', 'mrs.', 'mrs', 'ms.', 'ms', 'miss', 'dr.', 'dr', 'นาย', 'นาง', 'นางสาว']
    for p in prefixes:
        if s.startswith(p):
            s = s[len(p):]
    return re.sub(r'[^a-zA-Z0-9\u0e00-\u0e7f]', '', s)


def get_student_email(record):
    """Aggressively normalizes keys to match STUDENT PRIMARY E-MAIL while filtering out company headers."""
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


def _match_and_save_pdf(records, filename, file_content):
    """Matches a PDF file to a Student Record and ensures status is set to STATUS_GENERATED."""
    clean_filename = _clean_str(filename)

    for record in records:
        clean_id = _clean_str(record.student_id)
        clean_name = _clean_str(record.display_name)

        matched = (clean_id and clean_id in clean_filename) or \
                  (clean_name and len(clean_name) >= 3 and clean_name in clean_filename)

        if matched:
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

            doc.signed_file.save(os.path.basename(filename), file_content, save=True)
            return True

    return False


# ==========================================
# CORE SYSTEM VIEWS (STAFF PROTECTED)
# ==========================================

@staff_required
def index(request):
    recent = UploadedBatch.objects.order_by('-uploaded_at')[:5]
    return render(request, 'letters/index.html', {'recent': recent})


@staff_required
def upload(request):
    if request.method != 'POST' or 'datafile' not in request.FILES:
        return redirect('letters:index')

    f = request.FILES['datafile']
    df = _load_table(f)

    batch = UploadedBatch.objects.create(file_name=f.name, uploaded_by=request.user)
    records = []
    skipped_duplicates = 0

    for i, row in enumerate(df.to_dict(orient='records')):
        s_id = str(row.get('Student ID (e.g. 6587999)', '')).strip()
        l_type = str(row.get('Letter Type', '')).strip()

        # Check if record already exists globally across any batch
        if s_id and StudentRecord.objects.filter(student_id=s_id, letter_type=l_type).exists():
            skipped_duplicates += 1
            continue

        name = f"{row.get('Prefix', '')}{row.get('First Name', '')} {row.get('Last Name', '')}".strip()
        records.append(StudentRecord(
            batch=batch,
            row_index=i,
            letter_type=l_type,
            student_id=s_id,
            display_name=name,
            company=row.get('Company/Organization Name for Internship', ''),
            contractor=row.get('Letter Addressed To (Company Contact Person)', ''),
            contractor_position=row.get('Position', ''),
            internship_position=row.get('Internship Position', ''),
            period=row.get('Internship Period', ''),
            data=row,
        ))

    if records:
        StudentRecord.objects.bulk_create(records)

    log_activity(
        request, 
        'UPLOAD_EXCEL', 
        f"Batch #{batch.id} uploaded: {f.name} ({len(records)} added, {skipped_duplicates} duplicates skipped)"
    )

    if skipped_duplicates > 0:
        messages.info(request, f"Uploaded {len(records)} new record(s). Skipped {skipped_duplicates} duplicate record(s).")
    else:
        messages.success(request, f"Successfully uploaded {len(records)} record(s).")

    return redirect('letters:preview', batch_id=batch.id)

@staff_required
def preview(request, batch_id):
    batch = get_object_or_404(UploadedBatch, id=batch_id)
    records = list(batch.records.all())
    for r in records:
        r.supported = r.letter_type.strip() in SUPPORTED_LETTER_TYPES
        r.missing = [f for f in IMPORTANT_FIELDS if not getattr(r, f).strip()]
    counts = {}
    for r in records:
        counts[r.letter_type] = counts.get(r.letter_type, 0) + 1
    return render(request, 'letters/preview.html', {
        'batch': batch,
        'records': records,
        'counts': counts,
    })


@staff_required
def record_detail(request, record_id):
    record = get_object_or_404(StudentRecord, id=record_id)
    document = GeneratedDocument.objects.filter(
        record=record,
        status=GeneratedDocument.STATUS_GENERATED
    ).last()

    if request.method == 'POST':
        updated_data = dict(record.data)
        for field_key in updated_data.keys():
            if field_key in request.POST:
                updated_data[field_key] = request.POST.get(field_key)

        record.data = updated_data
        record.company = request.POST.get('Company/Organization Name for Internship', record.company)
        record.contractor = request.POST.get('Letter Addressed To (Company Contact Person)', record.contractor)
        record.internship_position = request.POST.get('Internship Position', record.internship_position)
        record.period = request.POST.get('Internship Period', record.period)
        record.save()

        ref_no = document.ref_no if document else str(2600 + record.row_index)
        buf, fname_or_reason = generate_for_row(record.data, ref_no=ref_no)

        if buf:
            if not document:
                document = GeneratedDocument(record=record, ref_no=ref_no)
            document.status = GeneratedDocument.STATUS_GENERATED
            document.file_name = fname_or_reason
            document.file.save(fname_or_reason, ContentFile(buf.getvalue()), save=True)

        # Log action
        log_activity(
            request, 
            'EDIT_RECORD', 
            f"Updated student record #{record.id} ({record.display_name})"
        )

        return redirect('letters:record_detail', record_id=record.id)

    if not document:
        ref_no = str(2600 + record.row_index)
        buf, fname_or_reason = generate_for_row(record.data, ref_no=ref_no)
        if buf:
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

    def get_priority(key):
        try:
            return (0, PRIORITY_FIELDS.index(key))
        except ValueError:
            return (1, key)

    sorted_keys = sorted(valid_keys, key=get_priority)

    primary_fields = [(k, raw_data[k]) for k in sorted_keys if k in PRIORITY_FIELDS]
    secondary_fields = [(k, raw_data[k]) for k in sorted_keys if k not in PRIORITY_FIELDS]

    return render(request, 'letters/record_detail.html', {
        'record': record,
        'primary_fields': primary_fields,
        'secondary_fields': secondary_fields,
        'document': document,
    })
@staff_required
def audit_logs_view(request):
    """Staff UI page to view system audit history."""
    logs = AuditLog.objects.select_related('user').all()[:100]
    return render(request, 'letters/audit_logs.html', {'logs': logs})

@staff_required
def generate(request, batch_id):
    batch = get_object_or_404(UploadedBatch, id=batch_id)
    if request.method != 'POST':
        return redirect('letters:preview', batch_id=batch.id)

    file_format = request.POST.get('format', 'docx').lower()
    selected_ids = request.POST.getlist('record_id')
    records = StudentRecord.objects.filter(id__in=selected_ids, batch=batch)

    zip_buf = io.BytesIO()
    gen_batch = GenerationBatch.objects.create(batch=batch)

    if file_format == 'pdf':
        with tempfile.TemporaryDirectory() as tmp_dir:
            record_file_map = {}

            for record in records:
                ref_no = str(2600 + record.row_index)
                buf, fname_or_reason = generate_for_row(record.data, ref_no=ref_no)

                if buf is None:
                    GeneratedDocument.objects.create(
                        record=record,
                        status=GeneratedDocument.STATUS_SKIPPED,
                        skip_reason=fname_or_reason,
                    )
                else:
                    docx_path = os.path.join(tmp_dir, fname_or_reason)
                    with open(docx_path, 'wb') as f:
                        f.write(buf.getvalue())
                    record_file_map[fname_or_reason] = (record, ref_no)

            docx_files = glob.glob(os.path.join(tmp_dir, '*.docx'))
            if docx_files:
                subprocess.run(
                    ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', tmp_dir] + docx_files,
                    check=True
                )

            with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                for docx_name, (record, ref_no) in record_file_map.items():
                    pdf_name = os.path.splitext(docx_name)[0] + '.pdf'
                    pdf_path = os.path.join(tmp_dir, pdf_name)

                    if os.path.exists(pdf_path):
                        with open(pdf_path, 'rb') as pf:
                            pdf_data = pf.read()

                        doc = GeneratedDocument.objects.create(
                            record=record,
                            status=GeneratedDocument.STATUS_GENERATED,
                            file_name=pdf_name,
                            ref_no=ref_no,
                        )
                        doc.file.save(pdf_name, ContentFile(pdf_data), save=True)
                        gen_batch.documents.add(doc)
                        zf.writestr(pdf_name, pdf_data)
                    else:
                        GeneratedDocument.objects.create(
                            record=record,
                            status=GeneratedDocument.STATUS_SKIPPED,
                            skip_reason="PDF conversion failed",
                        )
    else:
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for record in records:
                ref_no = str(2600 + record.row_index)
                buf, fname_or_reason = generate_for_row(record.data, ref_no=ref_no)

                if buf is None:
                    GeneratedDocument.objects.create(
                        record=record,
                        status=GeneratedDocument.STATUS_SKIPPED,
                        skip_reason=fname_or_reason,
                    )
                else:
                    doc = GeneratedDocument.objects.create(
                        record=record,
                        status=GeneratedDocument.STATUS_GENERATED,
                        file_name=fname_or_reason,
                        ref_no=ref_no,
                    )
                    doc.file.save(fname_or_reason, ContentFile(buf.getvalue()), save=True)
                    gen_batch.documents.add(doc)
                    zf.writestr(fname_or_reason, buf.getvalue())

    # Log action
    log_activity(
        request, 
        'GENERATE_DOCS', 
        f"Generated {records.count()} {file_format.upper()} files for Batch #{batch.id}"
    )

    zip_buf.seek(0)
    return FileResponse(
        zip_buf,
        as_attachment=True,
        filename=f'internship_letters_batch_{batch.id}_{file_format}.zip'
    )


@staff_required
def results(request, run_id):
    gen_batch = get_object_or_404(GenerationBatch, id=run_id)
    documents = gen_batch.documents.all()
    generated = documents.filter(status=GeneratedDocument.STATUS_GENERATED)
    skipped = documents.filter(status=GeneratedDocument.STATUS_SKIPPED)
    return render(request, 'letters/results.html', {
        'run_id': gen_batch.id,
        'generated': generated,
        'skipped': skipped,
    })


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
def download_record_pdf(request, record_id):
    record = get_object_or_404(StudentRecord, id=record_id)
    document = GeneratedDocument.objects.filter(
        record=record,
        status=GeneratedDocument.STATUS_GENERATED
    ).last()

    ref_no = document.ref_no if document else str(2600 + record.row_index)
    pdf_buf, filename = generate_pdf_for_row(record.data, ref_no=ref_no)
    if pdf_buf is None:
        raise Http404(filename)

    pdf_buf.seek(0)
    return FileResponse(
        pdf_buf,
        as_attachment=True,
        filename=filename,
        content_type='application/pdf'
    )

@staff_required
def upload_signed(request, batch_id=None):
    """Uploads signed PDF(s) or ZIP file and matches against filtered/all students."""
    selected_batch_id = request.GET.get('batch_id') or (str(batch_id) if batch_id else 'all')

    if request.method == 'POST':
        uploaded_files = request.FILES.getlist('signed_files') or request.FILES.getlist('files')
        matched_count = 0

        if selected_batch_id != 'all':
            batch = get_object_or_404(UploadedBatch, id=selected_batch_id)
            records = batch.records.all()
        else:
            # Match against latest deduplicated records
            all_records = StudentRecord.objects.order_by('-id')
            seen = set()
            records = []
            for r in all_records:
                key = (r.student_id.strip(), r.letter_type.strip())
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
    """Displays student records with batch filtering and deduplication."""
    batches = UploadedBatch.objects.order_by('-uploaded_at')
    selected_batch_id = request.GET.get('batch_id') or (str(batch_id) if batch_id else 'all')

    if selected_batch_id != 'all':
        selected_batch = get_object_or_404(UploadedBatch, id=selected_batch_id)
        records = StudentRecord.objects.filter(batch=selected_batch)
    else:
        # Get all records, deduplicated by student_id and letter_type (keeping latest upload)
        all_records = StudentRecord.objects.select_related('batch').order_by('-id')
        seen = set()
        records = []
        for r in all_records:
            key = (r.student_id.strip(), r.letter_type.strip())
            if key not in seen:
                seen.add(key)
                records.append(r)

    record_summary = []
    for r in records:
        doc = GeneratedDocument.objects.filter(record=r).last()
        student_email = get_student_email(r)
        record_summary.append({
            'record': r,
            'doc': doc,
            'email': student_email,
            'has_signed': bool(doc and doc.signed_file),
        })

    return render(request, 'letters/signed_preview.html', {
        'batches': batches,
        'selected_batch_id': str(selected_batch_id),
        'summary': record_summary,
    })

@staff_required
def send_signed_emails(request, batch_id):
    """Dispatches signed PDFs and draft DOCX files to selected student emails."""
    batch = get_object_or_404(UploadedBatch, id=batch_id)

    if request.method != 'POST':
        return redirect('letters:signed_preview', batch_id=batch.id)

    selected_ids = request.POST.getlist('selected_records')
    records = batch.records.filter(id__in=selected_ids)

    if not records.exists():
        messages.warning(request, "No student records were selected. Please check at least one box.")
        return redirect('letters:signed_preview', batch_id=batch.id)

    sent_count = 0
    for r in records:
        student_email = request.POST.get(f'email_{r.id}', '').strip() or get_student_email(r)
        doc = GeneratedDocument.objects.filter(record=r, status=GeneratedDocument.STATUS_GENERATED).last()

        if doc and doc.signed_file and student_email:
            subject = f"Signed Internship Letter & Draft - {r.display_name}"
            body = (
                f"Dear {r.display_name},\n\n"
                f"Please find attached your official signed internship letter (PDF) "
                f"and draft copy (.docx).\n\n"
                f"Faculty of ICT, Mahidol University"
            )

            email = EmailMessage(
                subject=subject,
                body=body,
                to=[student_email],
            )

            doc.signed_file.open('rb')
            email.attach(
                os.path.basename(doc.signed_file.name),
                doc.signed_file.read(),
                'application/pdf'
            )
            doc.signed_file.close()

            if doc.file:
                doc.file.open('rb')
                email.attach(
                    os.path.basename(doc.file.name),
                    doc.file.read(),
                    'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                )
                doc.file.close()

            email.send(fail_silently=True)

            doc.email_sent_at = timezone.now()
            doc.save()
            sent_count += 1

    # Log action
    log_activity(
        request, 
        'SEND_EMAIL', 
        f"Sent {sent_count} signed letter emails for Batch #{batch.id}"
    )

    messages.success(request, f"Successfully processed {sent_count} email(s) with PDF and DOCX attachments.")
    return redirect('letters:signed_preview', batch_id=batch.id)


@staff_required
def signed_papers_main(request):
    """Navbar destination: opens signed preview page with all batches by default."""
    return redirect('letters:signed_preview')


@staff_required
def view_signed_pdf(request, document_id):
    """Serves the signed PDF file inline for the mini preview iframe."""
    doc = get_object_or_404(GeneratedDocument, id=document_id)
    if not doc.signed_file:
        raise Http404("Signed PDF not found.")

    response = FileResponse(doc.signed_file.open('rb'), content_type='application/pdf')
    response['Content-Disposition'] = f'inline; filename="{os.path.basename(doc.signed_file.name)}"'
    return response


@staff_required
def view_draft_doc(request, document_id):
    """Serves the generated draft file for preview."""
    doc = get_object_or_404(GeneratedDocument, id=document_id)
    if not doc.file:
        raise Http404("Draft document not found.")

    is_pdf = doc.file.name.lower().endswith('.pdf')
    content_type = 'application/pdf' if is_pdf else 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

    response = FileResponse(doc.file.open('rb'), content_type=content_type)
    response['Content-Disposition'] = f'inline; filename="{os.path.basename(doc.file.name)}"'
    return response


@staff_required
def compare(request, document_id):
    doc = get_object_or_404(GeneratedDocument, id=document_id)
    if doc.status != GeneratedDocument.STATUS_GENERATED or not doc.file:
        raise Http404("This document wasn't generated, so there's nothing to compare.")

    record = doc.record
    doc.file.open('rb')
    extracted = extract_fields_from_docx(doc.file)
    doc.file.close()

    raw_rows = [
        ('Student ID', record.student_id, extracted['student_id']),
        ('Student Name', record.display_name, extracted['student_name']),
        ('Company', record.company, extracted['company']),
        ('Contractor / Contact', record.contractor, extracted['recipient_name']),
        ('Internship Position', record.internship_position, extracted['dept_position']),
        ('Period', record.period, extracted['period']),
    ]
    comparison = []
    for i, (label, excel_val, doc_val) in enumerate(raw_rows):
        if _norm(excel_val) == _norm(doc_val):
            status = 'match'
        elif not doc_val.strip():
            status = 'missing'
        elif _norm(excel_val) in _norm(doc_val) or _norm(doc_val) in _norm(excel_val):
            status = 'partial'
        else:
            status = 'mismatch'
        comparison.append({
            'row_key': f'row{i}', 'label': label,
            'excel': excel_val, 'doc': doc_val, 'status': status,
        })

    return render(request, 'letters/compare.html', {
        'record': record,
        'document': doc,
        'comparison': comparison,
        'match_count': sum(1 for c in comparison if c['status'] == 'match'),
        'issue_count': sum(1 for c in comparison if c['status'] != 'match'),
    })