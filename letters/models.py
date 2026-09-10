# -*- coding: utf-8 -*-
from django.db import models
from django.contrib.auth.models import User


class UploadedBatch(models.Model):
    """One uploaded Excel/CSV file."""
    file_name = models.CharField(max_length=255)
    uploaded_at = models.DateTimeField(auto_now_add=True)
    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)

    def __str__(self):
        uploader = self.uploaded_by.username if self.uploaded_by else "System"
        return f"Batch #{self.id} - {self.file_name} (Uploaded by {uploader})"


class BatchUpload(models.Model):
    """Legacy or alternate batch tracking model."""
    filename = models.CharField(max_length=255)
    uploaded_at = models.DateTimeField(auto_now_add=True)
    total_records = models.IntegerField(default=0)

    def __str__(self):
        return f"Batch #{self.id} - {self.filename}"


class StudentRecord(models.Model):
    batch = models.ForeignKey('UploadedBatch', on_delete=models.CASCADE, related_name='records')
    row_index = models.IntegerField(default=0)
    
    # Core Tracking Keys
    student_id = models.CharField(max_length=50)
    display_name = models.CharField(max_length=255)
    letter_type = models.CharField(max_length=100)
    course_id = models.CharField(max_length=50, blank=True, default='')   # e.g., "ITCS499"
    semester = models.CharField(max_length=20, blank=True, default='')    # e.g., "2026/1"
    times = models.PositiveIntegerField(default=1)                       # 1st time, 2nd time, etc.
    
    # Details
    year = models.CharField(max_length=10, blank=True, default='')
    company = models.CharField(max_length=255, blank=True)
    contractor = models.CharField(max_length=255, blank=True)
    contractor_position = models.CharField(max_length=255, blank=True)
    internship_position = models.CharField(max_length=255, blank=True)
    period = models.CharField(max_length=255, blank=True)
    data = models.JSONField(default=dict)

    class Meta:
        ordering = ['-id']

    def __str__(self):
        return f"{self.student_id} - {self.display_name} ({self.course_id} {self.semester} #{self.times})"


class LetterTemplate(models.Model):
    TEMPLATE_TYPES = [
        ('request', 'Internship Request Letter'),
        ('confirm', 'Internship Confirmation/Referral Letter'),
        ('extension', 'Internship Extension Letter'),
    ]

    name = models.CharField(max_length=150)
    letter_type = models.CharField(max_length=50, choices=TEMPLATE_TYPES, unique=True)
    file = models.FileField(upload_to='letter_templates/')
    is_active = models.BooleanField(default=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.name} ({self.get_letter_type_display()})"


class LetterType(models.Model):
    slug = models.SlugField(max_length=50, unique=True)
    display_name = models.CharField(max_length=150)

    def __str__(self):
        return self.display_name


class LetterTypeMapping(models.Model):
    raw_text = models.CharField(max_length=255, unique=True)
    letter_type = models.ForeignKey(LetterType, related_name='mappings', on_delete=models.CASCADE)

    def __str__(self):
        return f"{self.raw_text} → {self.letter_type.slug}"


class GeneratedDocument(models.Model):
    STATUS_GENERATED = 'generated'
    STATUS_SKIPPED = 'skipped'
    STATUS_SENT = 'sent'
    STATUS_CHOICES = [
        (STATUS_GENERATED, 'Generated'),
        (STATUS_SKIPPED, 'Skipped'),
        (STATUS_SENT, 'Sent'),
    ]

    record = models.ForeignKey(StudentRecord, related_name='documents', on_delete=models.CASCADE)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES)
    file_name = models.CharField(max_length=255, blank=True)
    file = models.FileField(upload_to='generated_letters/', blank=True, null=True)
    skip_reason = models.CharField(max_length=255, blank=True)
    ref_no = models.CharField(max_length=20, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    signed_file = models.FileField(upload_to='signed_docs/', null=True, blank=True)
    email_sent_at = models.DateTimeField(null=True, blank=True)

    def __str__(self):
        student_name = self.record.display_name if self.record else "Unknown Student"
        student_id = self.record.student_id if self.record else "N/A"
        letter_type = self.record.letter_type if self.record else "N/A"
        identifier = self.file_name or self.skip_reason or f"Ref: {self.ref_no}"
        status_label = self.get_status_display()
        return f"{student_name} ({student_id}) | {identifier} [{letter_type}] ({status_label})"


class GenerationBatch(models.Model):
    batch = models.ForeignKey(UploadedBatch, related_name='generation_runs', on_delete=models.CASCADE)
    created_at = models.DateTimeField(auto_now_add=True)
    documents = models.ManyToManyField(GeneratedDocument, related_name='generation_batches')


class AuditLog(models.Model):
    ACTION_CHOICES = [
        ('UPLOAD_EXCEL', 'Uploaded Excel Batch'),
        ('GENERATE_DOCS', 'Generated Draft Letters'),
        ('UPLOAD_SIGNED', 'Uploaded Signed PDFs'),
        ('SEND_EMAIL', 'Dispatched Emails'),
        ('EDIT_RECORD', 'Edited Student Record'),
        ('ADD_RECORD', 'Manually Added Student Record'),
    ]

    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    action = models.CharField(max_length=50, choices=ACTION_CHOICES)
    details = models.TextField(blank=True)
    ip_address = models.GenericIPAddressField(null=True, blank=True)
    timestamp = models.DateTimeField(auto_now_add=True)


class TemplateFieldMapping(models.Model):
    SOURCE_FIELD = 'field'
    SOURCE_STATIC = 'static'
    SOURCE_UNMAPPED = 'unmapped'
    SOURCE_CHOICES = [
        (SOURCE_FIELD, 'Data Field'),
        (SOURCE_STATIC, 'Static Text'),
        (SOURCE_UNMAPPED, 'Not Mapped Yet'),
    ]

    template = models.ForeignKey('LetterTemplate', related_name='field_mappings', on_delete=models.CASCADE)
    placeholder = models.CharField(max_length=100)
    source_type = models.CharField(max_length=10, choices=SOURCE_CHOICES, default=SOURCE_UNMAPPED)
    source_value = models.CharField(max_length=255, blank=True, default='')

    class Meta:
        unique_together = ('template', 'placeholder')


class StudentLetter(models.Model):
    student_id = models.CharField(max_length=20)
    display_name = models.CharField(max_length=255, blank=True, null=True)
    letter_type = models.CharField(max_length=50)
    company = models.CharField(max_length=255, blank=True, null=True)
    course_id = models.CharField(max_length=100, blank=True, null=True)
    data = models.JSONField(default=dict, blank=True)
    signed_file = models.FileField(upload_to='signed_letters/', blank=True, null=True)
    status = models.CharField(max_length=20, default='pending')
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.student_id} - {self.display_name or ''} ({self.letter_type})"