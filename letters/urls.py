from django.urls import path, include
from django.contrib import admin
from . import views

app_name = 'letters'

urlpatterns = [
    path('', views.index, name='index'),
    path('upload/', views.upload, name='upload'),
    path('preview/<int:batch_id>/', views.preview, name='preview'),
    path('generate/<int:batch_id>/', views.generate, name='generate'),
    path('results/<int:run_id>/', views.results, name='results'),
    path('download/<int:run_id>/', views.download, name='download'),
    path('record/<int:record_id>/', views.record_detail, name='record_detail'),
    path('record/<int:record_id>/pdf/', views.download_record_pdf, name='download_record_pdf'),
    path('upload-signed/<int:batch_id>/', views.upload_signed, name='upload_signed'),
    path('signed-preview/<int:batch_id>/', views.signed_preview, name='signed_preview'),
    path('signed-papers/', views.signed_papers_main, name='signed_papers_main'),
    path('send-signed-emails/<int:batch_id>/', views.send_signed_emails, name='send_signed_emails'),
    path('compare/<int:document_id>/', views.compare, name='compare'),
    # Auth & Account Management
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('profile/', views.profile_view, name='profile'),
    path('audit-logs/', views.audit_logs_view, name='audit_logs'),
    path('signed/', views.signed_preview, name='signed_preview'),
    path('signed/<int:batch_id>/', views.signed_preview, name='signed_preview_batch'),
    path('signed/upload/', views.upload_signed, name='upload_signed'),
    path('signed/upload/<int:batch_id>/', views.upload_signed, name='upload_signed_batch'),
    
    # Add this route:
    path('document/<int:document_id>/signed-pdf/', views.view_signed_pdf, name='view_signed_pdf'),
        
]