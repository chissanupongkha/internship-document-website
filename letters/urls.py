from django.urls import path
from . import views, template_editor
from . import status_views
from . import report_views
from . import extra_views

app_name = 'letters'

urlpatterns = [
    # Dashboard & Batch Management
    path('', views.index, name='index'),
    path('upload/', views.upload, name='upload'),
    path('preview/<int:batch_id>/', views.preview, name='preview'),
    path('generate/<int:batch_id>/', views.generate, name='generate'),
    path('results/<int:run_id>/', views.results, name='results'),
    path('download/<int:run_id>/', views.download, name='download'),

    # Student Record Management
    path('record/add/', views.add_record, name='add_record'),
    path('record/<int:record_id>/', views.record_detail, name='record_detail'),
    path('record/<int:record_id>/pdf/', views.download_record_pdf, name='download_record_pdf'),
    path('download/<int:record_id>/<str:letter_type>/', views.download_letter_view, name='download_letter'),
    path('document/<int:document_id>/signed-pdf/', views.view_signed_pdf, name='view_signed_pdf'),
    path('upload-signed/', views.upload_signed, name='upload_signed'),
    path('upload-signed/<int:batch_id>/', views.upload_signed, name='upload_signed_batch'),
    path('letters/<int:letter_id>/upload-signed/', views.upload_signed_by_id, name='upload_signed_by_id'),

    # Signed Documents Workflow (Aliases provided for backward compatibility)
    path('signed-papers/', views.signed_preview, name='signed_papers_main'),
    path('signed/', views.signed_preview, name='signed_preview'),
    path('signed/<int:batch_id>/', views.signed_preview, name='signed_preview_batch'),
    path('signed/upload/', views.upload_signed, name='upload_signed_default'),
    path('send-signed-emails/', views.send_signed_emails, name='send_signed_emails'),

    # Auth & Logs
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('profile/', views.profile_view, name='profile'),
    path('audit-logs/', views.audit_logs_view, name='audit_logs'),

    path('record/<int:record_id>/preview-pdf/', views.preview_record_pdf, name='preview_record_pdf'),
    path('record/<int:record_id>/preview-docx/', views.preview_record_docx, name='preview_record_docx'),
    path('record/<int:record_id>/update-semester/', views.update_semester, name='update_semester'),
    path('record/<int:record_id>/update-ref-no/', views.update_ref_no, name='update_ref_no'),

    # Templates & Management
    path('templates/', views.template_list, name='template_list'),
    path('templates/upload/', views.template_upload, name='template_upload'),
    path('templates/from-builtin/', views.template_from_builtin, name='template_from_builtin'),
    path('templates/create/', views.template_create, name='template_create'),
    path('templates/<int:template_id>/edit/', views.template_edit, name='template_edit'),
    path('templates/<int:template_id>/editor/', template_editor.template_editor, name='template_editor'),
    path('templates/<int:template_id>/restore/', template_editor.template_restore, name='template_restore'),
    path('templates/<int:template_id>/mapping/', views.template_mapping, name='template_mapping'),
    path('templates/<int:template_id>/preview/', views.template_preview, name='template_preview'),
    path('templates/<int:template_id>/activate/', views.template_activate, name='template_activate'),
    path('templates/<int:template_id>/delete/', views.template_delete, name='template_delete'),

    # Letter Types & Raw Value Mappings
    path('letter-types/', views.lettertype_list, name='lettertype_list'),
    path('letter-settings/', views.letter_settings, name='letter_settings'),
    path('letter-types/create/', views.lettertype_create, name='lettertype_create'),
    path('letter-types/<int:type_id>/delete/', views.lettertype_delete, name='lettertype_delete'),
    path('letter-types/mapping/', views.mapping_create, name='mapping_create'),
    path('mappings/<int:mapping_id>/update/', views.mapping_update, name='mapping_update'),
    path('mappings/<int:mapping_id>/delete/', views.mapping_delete, name='mapping_delete'),
    
    # --- add to letters/urls.py -------------------------------------------------
# (keep your existing imports; add this one)


# --- add inside urlpatterns = [ ... ] ---------------------------------------
    # Document versions, signed upload, settings, accounts, template help
    path('record/<int:record_id>/docx/', extra_views.record_docx_download, name='record_docx_download'),
    path('record/<int:record_id>/docx/upload/', extra_views.record_docx_upload, name='record_docx_upload'),
    path('record/<int:record_id>/pages/', extra_views.record_pages, name='record_pages'),
    path('signed/upload-manual/', extra_views.signed_upload_manual, name='signed_upload_manual'),
    path('configuration/email-templates/', extra_views.email_templates, name='email_templates'),
    path('accounts/', extra_views.user_list, name='user_list'),
    path('accounts/new/', extra_views.user_form, name='user_new'),
    path('accounts/<int:pk>/', extra_views.user_form, name='user_edit'),
    path('templates/starter/', extra_views.template_starter, name='template_starter'),

    # Reports (under Tracking)
    path('reports/', report_views.report, name='report'),
    path('reports/export/', report_views.report_export, name='report_export'),

    # Request status, students, subjects
    path('status/', status_views.request_status, name='request_status'),
    path('request/<int:record_id>/status/', status_views.document_status, name='document_status'),
    path('request/<int:record_id>/change/', status_views.record_change, name='record_change'),
    path('subjects/', status_views.subject_list, name='subject_list'),
    path('subjects/save/', status_views.subject_save, name='subject_save'),
    path('subjects/<int:pk>/delete/', status_views.subject_delete, name='subject_delete'),
    path('students/', status_views.student_list, name='student_list'),
    path('students/new/', status_views.student_form, name='student_new'),
    path('students/import/', status_views.student_import, name='student_import'),
    path('students/<int:pk>/', status_views.student_detail, name='student_detail'),
    path('students/<int:pk>/edit/', status_views.student_form, name='student_edit'),
    path('students/<int:pk>/delete/', status_views.student_delete, name='student_delete'),
    path('enrollments/<int:pk>/delete/', status_views.enrollment_delete, name='enrollment_delete'),

]