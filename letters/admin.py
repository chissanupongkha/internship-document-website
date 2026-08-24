from django.contrib import admin
from .models import UploadedBatch, StudentRecord, GeneratedDocument, GenerationBatch

admin.site.register(UploadedBatch)
admin.site.register(StudentRecord)
admin.site.register(GeneratedDocument)
admin.site.register(GenerationBatch)
