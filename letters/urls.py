from django.urls import path
from . import views

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
    # urls.py
    path('letters/<int:letter_id>/upload-signed/', views.upload_signed_by_id, name='upload_signed_by_id'),
    # ...
    # Signed Documents Workflow (Aliases provided for backward compatibility)
    path('signed-papers/', views.signed_preview, name='signed_papers_main'),
    path('signed/', views.signed_preview, name='signed_preview'),
    path('signed/<int:batch_id>/', views.signed_preview, name='signed_preview_batch'),
    path('upload-signed/<int:batch_id>/', views.upload_signed, name='upload_signed'),
    path('signed/upload/', views.upload_signed, name='upload_signed_default'),
    path('send-signed-emails/', views.send_signed_emails, name='send_signed_emails'),

    # Auth & Logs
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('profile/', views.profile_view, name='profile'),
    path('audit-logs/', views.audit_logs_view, name='audit_logs'),
    
    path('record/<int:record_id>/preview-pdf/', views.preview_record_pdf, name='preview_record_pdf'),

]