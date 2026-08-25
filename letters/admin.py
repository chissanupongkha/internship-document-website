from django.contrib import admin
from .models import UploadedBatch, StudentRecord, GeneratedDocument, GenerationBatch, AuditLog

admin.site.register(UploadedBatch)
admin.site.register(StudentRecord)
admin.site.register(GeneratedDocument)
admin.site.register(GenerationBatch)


@admin.register(AuditLog)
class AuditLogAdmin(admin.ModelAdmin):
    list_display = ('timestamp', 'user', 'action', 'details', 'ip_address')
    list_filter = ('action', 'user', 'timestamp')
    search_fields = ('user__username', 'details', 'ip_address')
    readonly_fields = ('user', 'action', 'details', 'ip_address', 'timestamp')

    def has_add_permission(self, request):
        return False  # Prevents manual tampering of audit logs