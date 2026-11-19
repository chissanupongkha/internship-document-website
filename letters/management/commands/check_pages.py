"""Open every page as an admin and list the ones that fail. Changes nothing.

    python manage.py check_pages
"""
import re

from django.conf import settings
from django.contrib.auth.models import User
from django.core.management.base import BaseCommand
from django.test import Client
from django.urls import URLResolver, get_resolver

SKIP = ('logout', 'delete', 'generate', 'send', 'download', 'export', 'pdf', 'docx', 'starter', 'preview-',
        'save', 'upload', 'update', 'activate', 'change', 'import', 'media', 'admin/')


def _routes(resolver, prefix=''):
    for p in resolver.url_patterns:
        if isinstance(p, URLResolver):
            yield from _routes(p, prefix + str(p.pattern))
        else:
            yield prefix + str(p.pattern)


class Command(BaseCommand):
    help = "Open every page as an admin and report pages that crash. Changes nothing."

    def handle(self, *args, **opts):
        from letters.models import Course, Student, StudentRecord, UploadedBatch, LetterTemplate
        admin = User.objects.filter(is_superuser=True, is_active=True).first()
        if not admin:
            self.stderr.write("No admin account. Run: python manage.py createsuperuser")
            return
        settings.ALLOWED_HOSTS = list(settings.ALLOWED_HOSTS) + ['testserver']
        client = Client(raise_request_exception=False)
        client.force_login(admin)
        first = lambda m: (m.objects.order_by('id').first() or type('x', (), {'id': 0})()).id
        ids = {'record_id': first(StudentRecord), 'batch_id': first(UploadedBatch), 'template_id': first(LetterTemplate),
               'student_id': first(Student), 'course_id': first(Course), 'user_id': admin.id}
        bad = checked = 0
        for route in _routes(get_resolver()):
            url = '/' + route.replace('^', '').replace('$', '')
            if any(k in url for k in SKIP):
                continue
            url = re.sub(r'<int:(\w+)>', lambda m: str(ids.get(m.group(1), 1)), url)
            url = re.sub(r'<(?:str|slug):(\w+)>', 'x', url)
            if '(' in url:
                continue
            checked += 1
            r = client.get(url)
            if r.status_code >= 500:
                bad += 1
                self.stdout.write(self.style.ERROR(f"FAIL {r.status_code}  {url}"))
            else:
                self.stdout.write(f"ok   {r.status_code}  {url}")
        self.stdout.write(f"\n{checked} pages checked, {bad} failed.")
        if bad:
            self.stdout.write("Failed pages print their error in the runserver window / this window above. "
                              "Paste the FAIL lines and the red error text here.")
