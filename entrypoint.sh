#!/bin/sh
set -e

# Create required media directories on ephemeral container storage
mkdir -p /app/media/signed_docs
mkdir -p /app/media/generated_letters
mkdir -p /app/media/letter_templates

# Run database migrations
python manage.py migrate --noinput
python manage.py collectstatic --noinput

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