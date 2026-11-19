"""Load the official logo images into the logo list that staff choose from."""
import os

from django.core.files import File
from django.core.management.base import BaseCommand

from letters.models import Logo

FRIENDLY = {
    'logo-mu_color-2': 'Mahidol - letterhead mark (colour)',
    'logo-mu_black-white-2': 'Mahidol - letterhead mark (black)',
}
LIBRARY = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'logo_library')
CURRENT = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'static', 'letters', 'images', 'mu_logo.jpeg')


def _name_for(filename):
    stem = os.path.splitext(filename)[0]
    return FRIENDLY.get(stem.lower(), f"Mahidol - {stem}")


class Command(BaseCommand):
    help = "Load logo images from letters/logo_library into the logo list (safe to run again)."

    def handle(self, *args, **opts):
        added = 0
        files = []
        if os.path.exists(CURRENT):
            files.append((CURRENT, 'Mahidol - current letterhead logo'))
        if os.path.isdir(LIBRARY):
            for f in sorted(os.listdir(LIBRARY)):
                if f.lower().endswith(('.png', '.jpg', '.jpeg')):
                    files.append((os.path.join(LIBRARY, f), _name_for(f)))
        for path, name in files:
            if Logo.objects.filter(name=name).exists():
                continue
            with open(path, 'rb') as fh:
                logo = Logo(name=name)
                logo.image.save(os.path.basename(path), File(fh), save=True)
            added += 1
        self.stdout.write(self.style.SUCCESS(f"{added} logo(s) added, {len(files) - added} already there."))
