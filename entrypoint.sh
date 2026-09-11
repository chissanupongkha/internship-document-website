#!/bin/sh

# Apply database migrations
python manage.py migrate --noinput

# Create superuser automatically if environment variables are set
if [ "$DJANGO_SUPERUSER_USERNAME" ]; then
    python manage.py createsuperuser --noinput || true
fi

exec "$@"