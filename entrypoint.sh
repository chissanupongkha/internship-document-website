#!/bin/sh
set -e

# Create required media directories on local disk (mount a volume at MEDIA_ROOT to keep them)
MEDIA_DIR="${MEDIA_ROOT:-/app/media}"
mkdir -p "$MEDIA_DIR/signed_docs"
mkdir -p "$MEDIA_DIR/generated_letters"
mkdir -p "$MEDIA_DIR/letter_templates" "$MEDIA_DIR/logos"

# Run database migrations
python manage.py migrate --noinput
python manage.py collectstatic --noinput
python manage.py load_logos || true

# Safely create superuser if environment variables exist and user doesn't
if [ "$DJANGO_SUPERUSER_USERNAME" ]; then
    python manage.py shell -c "
from django.contrib.auth import get_user_model;
User = get_user_model();
if not User.objects.filter(username='$DJANGO_SUPERUSER_USERNAME').exists():
    User.objects.create_superuser('$DJANGO_SUPERUSER_USERNAME', '$DJANGO_SUPERUSER_EMAIL', '$DJANGO_SUPERUSER_PASSWORD')
"
fi

exec "$@"