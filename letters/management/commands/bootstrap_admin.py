import os
from django.contrib.auth.models import User
from django.core.management.base import BaseCommand


class Command(BaseCommand):
    help = "Creates an initial staff superuser from env vars if none exists yet."

    def handle(self, *args, **options):
        username = os.environ.get('ADMIN_USERNAME')
        password = os.environ.get('ADMIN_PASSWORD')
        email = os.environ.get('ADMIN_EMAIL', '')

        if not username or not password:
            self.stdout.write("ADMIN_USERNAME/ADMIN_PASSWORD not set -- skipping bootstrap.")
            return

        if User.objects.filter(is_superuser=True).exists():
            self.stdout.write("A superuser already exists -- skipping bootstrap.")
            return

        User.objects.create_superuser(username=username, email=email, password=password)
        self.stdout.write(f"Created initial superuser '{username}'.")