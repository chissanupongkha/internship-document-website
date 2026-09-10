from django.contrib import admin
from django.utils.html import format_html
from .models import StudentRecord, UploadedBatch, GeneratedDocument, GenerationBatch, LetterTemplate, AuditLog


@admin.register(UploadedBatch)
class UploadedBatchAdmin(admin.ModelAdmin):
    list_display = ('id', 'file_name', 'uploaded_at', 'uploaded_by')
    search_fields = ('file_name', 'uploaded_by__username')
    ordering = ('-uploaded_at',)


@admin.register(StudentRecord)
class StudentRecordAdmin(admin.ModelAdmin):
    list_display = ('student_id', 'display_name', 'company', 'letter_type', 'course_id', 'semester', 'times')
    search_fields = ('student_id', 'display_name', 'company')
    list_filter = ('letter_type', 'semester', 'course_id')
    readonly_fields = ('display_all_excel_data',)

    def display_all_excel_data(self, obj):
        if not obj.data or not isinstance(obj.data, dict):
            return "No raw Excel data found."

        rows_html = "".join(
            f"<tr>"
            f"<th style='text-align:left; padding:6px 12px; background:#f8f9fa; border:1px solid #dee2e6; width:30%;'>{key}</th>"
            f"<td style='padding:6px 12px; border:1px solid #dee2e6;'>{value}</td>"
            f"</tr>"
            for key, value in obj.data.items()
        )
        return format_html(
            "<table style='width:100%; border-collapse:collapse; font-size:13px;'>"
            "<thead><tr><th style='text-align:left; padding:8px; background:#e9ecef;'>Column Header</th>"
            "<th style='text-align:left; padding:8px; background:#e9ecef;'>Excel Cell Value</th></tr></thead>"
            "<tbody>{}</tbody></table>",
            format_html(rows_html)
        )

    display_all_excel_data.short_description = "Full Excel Data Row"


@admin.register(LetterTemplate)
class LetterTemplateAdmin(admin.ModelAdmin):
    list_display = ('name', 'letter_type', 'updated_at')
    search_fields = ('name', 'letter_type')


@admin.register(GeneratedDocument)
class GeneratedDocumentAdmin(admin.ModelAdmin):
    list_display = ('record', 'status', 'ref_no', 'file_name', 'created_at')
    list_filter = ('status', 'created_at')
    search_fields = ('record__student_id', 'record__display_name', 'ref_no')


@admin.register(GenerationBatch)
class GenerationBatchAdmin(admin.ModelAdmin):
    list_display = ('id', 'batch', 'created_at')
    ordering = ('-created_at',)


@admin.register(AuditLog)
class AuditLogAdmin(admin.ModelAdmin):
    list_display = ('user', 'action', 'details', 'ip_address', 'timestamp')
    list_filter = ('action', 'timestamp')
    search_fields = ('user__username', 'details', 'ip_address')
    ordering = ('-timestamp',)