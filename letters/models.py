# -*- coding: utf-8 -*-
from django.db import models


class UploadedBatch(models.Model):
    """One uploaded Excel/CSV file."""
    file_name = models.CharField(max_length=255)
    uploaded_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.file_name} ({self.uploaded_at:%Y-%m-%d %H:%M})"


class StudentRecord(models.Model):
    """One row from the uploaded file. Raw columns are kept as JSON so the
    schema doesn't have to change every time the Google Form adds a field —
    docgen.py reads from `data` by the original column names. A few key
    fields are duplicated onto real columns too, purely so the preview page
    can query/sort/filter them without touching the JSON blob."""
    batch = models.ForeignKey(UploadedBatch, related_name='records', on_delete=models.CASCADE)
    row_index = models.IntegerField()
    letter_type = models.CharField(max_length=120, blank=True)
    student_id = models.CharField(max_length=40, blank=True)
    display_name = models.CharField(max_length=200, blank=True)
    company = models.CharField(max_length=200, blank=True)
    contractor = models.CharField(max_length=200, blank=True)  # company contact person
    contractor_position = models.CharField(max_length=200, blank=True)
    internship_position = models.CharField(max_length=200, blank=True)
    period = models.CharField(max_length=120, blank=True)
    data = models.JSONField()  # full original row, used by docgen.py

    class Meta:
        ordering = ['row_index']

    def __str__(self):
        return f"{self.student_id} — {self.display_name}"


class GeneratedDocument(models.Model):
    STATUS_GENERATED = 'generated'
    STATUS_SKIPPED = 'skipped'
    STATUS_CHOICES = [(STATUS_GENERATED, 'Generated'), (STATUS_SKIPPED, 'Skipped')]

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
        return f"{self.file_name or self.skip_reason} ({self.status})"


class GenerationBatch(models.Model):
    """Groups the documents produced by one 'Generate' click, so the
    results/download page can show/zip just that run."""
    batch = models.ForeignKey(UploadedBatch, related_name='generation_runs', on_delete=models.CASCADE)
    created_at = models.DateTimeField(auto_now_add=True)
    documents = models.ManyToManyField(GeneratedDocument, related_name='generation_batches')
