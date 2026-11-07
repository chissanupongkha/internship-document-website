from django.urls import path
from . import views
from . import status_views

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

    # Templates & Management
    path('templates/', views.template_list, name='template_list'),
    path('templates/upload/', views.template_upload, name='template_upload'),
    path('templates/<int:template_id>/mapping/', views.template_mapping, name='template_mapping'),
    path('templates/<int:template_id>/preview/', views.template_preview, name='template_preview'),
    path('templates/<int:template_id>/activate/', views.template_activate, name='template_activate'),
    path('templates/<int:template_id>/delete/', views.template_delete, name='template_delete'),

    # Letter Types & Raw Value Mappings
    path('letter-types/', views.lettertype_list, name='lettertype_list'),
    path('letter-types/create/', views.lettertype_create, name='lettertype_create'),
    path('letter-types/<int:type_id>/delete/', views.lettertype_delete, name='lettertype_delete'),
    path('letter-types/mapping/', views.mapping_create, name='mapping_create'),
    path('mappings/<int:mapping_id>/update/', views.mapping_update, name='mapping_update'),
    path('mappings/<int:mapping_id>/delete/', views.mapping_delete, name='mapping_delete'),
    path('save-email-message/', views.save_email_message, name='save_email_message'),
    
    # --- add to letters/urls.py -------------------------------------------------
# (keep your existing imports; add this one)


# --- add inside urlpatterns = [ ... ] ---------------------------------------
    # Request status, students, subjects
    path('status/', status_views.request_status, name='request_status'),
    path('request/<int:record_id>/status/', status_views.document_status, name='document_status'),
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