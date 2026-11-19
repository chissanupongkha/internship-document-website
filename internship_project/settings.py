"""
Django settings for the Internship Letter Generator.

Only PostgreSQL is required. Everything else (media files, static files and
email) runs inside Django itself, with no hosted services.
"""
import os
from pathlib import Path

import dj_database_url
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent.parent


def _env_int(name, default):
    """Integer from the environment; a wrong value falls back to the default instead of crashing."""
    try:
        return int(os.getenv(name, default))
    except (TypeError, ValueError):
        return default


def _env_bool(name, default=False):
    return os.getenv(name, str(default)).strip().lower() in ('1', 'true', 'yes', 'on')


DEBUG = _env_bool('DEBUG', False)
SECRET_KEY = os.getenv('SECRET_KEY', 'dev-only-insecure-key-change-me')
if not DEBUG and SECRET_KEY == 'dev-only-insecure-key-change-me':
    raise RuntimeError('Set SECRET_KEY in the environment when DEBUG is off.')

X_FRAME_OPTIONS = 'SAMEORIGIN'
ALLOWED_HOSTS = [h.strip() for h in os.getenv('ALLOWED_HOSTS', 'localhost,127.0.0.1').split(',') if h.strip()]
# Render (or any host behind an HTTPS proxy): allow its public host name automatically
_render_host = os.getenv('RENDER_EXTERNAL_HOSTNAME')
if _render_host and _render_host not in ALLOWED_HOSTS:
    ALLOWED_HOSTS.append(_render_host)
CSRF_TRUSTED_ORIGINS = [o.strip() for o in os.getenv('CSRF_TRUSTED_ORIGINS', '').split(',') if o.strip()]
CSRF_TRUSTED_ORIGINS += [f'https://{h}' for h in ALLOWED_HOSTS if h not in ('localhost', '127.0.0.1', '*') and not h.startswith('.')
                         and f'https://{h}' not in CSRF_TRUSTED_ORIGINS]
SECURE_PROXY_SSL_HEADER = ('HTTP_X_FORWARDED_PROTO', 'https')   # the proxy tells Django the request was HTTPS

SESSION_COOKIE_AGE = 7200            # 2 hours
SESSION_SAVE_EVERY_REQUEST = True    # active users stay logged in
SESSION_EXPIRE_AT_BROWSER_CLOSE = False

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'letters',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'whitenoise.middleware.WhiteNoiseMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
    'letters.middleware.ActivityMiddleware',
]

ROOT_URLCONF = 'internship_project.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    },
]

WSGI_APPLICATION = 'internship_project.wsgi.application'

# --- Database: plain PostgreSQL ------------------------------------------
# Either set DATABASE_URL=postgresql://user:pass@host:5432/dbname
# or the separate DB_* variables below.
_db_url = os.getenv('DATABASE_URL')
if _db_url:
    DATABASES = {'default': dj_database_url.parse(_db_url, conn_max_age=60)}
else:
    DATABASES = {
        'default': {
            'ENGINE': 'django.db.backends.postgresql',
            'NAME': os.getenv('DB_NAME', 'internship_letters'),
            'USER': os.getenv('DB_USER', 'postgres'),
            'PASSWORD': os.getenv('DB_PASSWORD', 'postgres'),
            'HOST': os.getenv('DB_HOST', 'localhost'),
            'PORT': os.getenv('DB_PORT', '5432'),
        }
    }

AUTH_PASSWORD_VALIDATORS = [
    {'NAME': 'django.contrib.auth.password_validation.UserAttributeSimilarityValidator'},
    {'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator'},
    {'NAME': 'django.contrib.auth.password_validation.CommonPasswordValidator'},
    {'NAME': 'django.contrib.auth.password_validation.NumericPasswordValidator'},
]

LANGUAGE_CODE = 'en-us'
TIME_ZONE = 'UTC'
USE_I18N = True
USE_TZ = True
DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'
LOGIN_URL = 'letters:login'

# --- Files: local disk ----------------------------------------------------
# Uploaded templates, logos, generated and signed letters live under
# MEDIA_ROOT. In Docker, mount a volume there so they survive restarts.
STATIC_URL = '/static/'
STATIC_ROOT = BASE_DIR / 'staticfiles'
MEDIA_URL = '/media/'
MEDIA_ROOT = Path(os.getenv('MEDIA_ROOT', BASE_DIR / 'media'))

STORAGES = {
    'default': {'BACKEND': 'django.core.files.storage.FileSystemStorage'},
    'staticfiles': {'BACKEND': 'whitenoise.storage.CompressedManifestStaticFilesStorage'},
}

# --- Email: standard SMTP, credentials only from the environment ----------
# With no EMAIL_HOST set, messages are printed to the console instead.
EMAIL_HOST = os.getenv('EMAIL_HOST', '')
if EMAIL_HOST:
    EMAIL_BACKEND = 'django.core.mail.backends.smtp.EmailBackend'
    EMAIL_PORT = _env_int('EMAIL_PORT', 587)
    EMAIL_USE_TLS = _env_bool('EMAIL_USE_TLS', True)
    EMAIL_USE_SSL = _env_bool('EMAIL_USE_SSL', False)
    EMAIL_HOST_USER = os.getenv('EMAIL_HOST_USER', '')
    EMAIL_HOST_PASSWORD = os.getenv('EMAIL_HOST_PASSWORD', '')
    EMAIL_TIMEOUT = 60
else:
    EMAIL_BACKEND = 'django.core.mail.backends.console.EmailBackend'
DEFAULT_FROM_EMAIL = os.getenv('DEFAULT_FROM_EMAIL', 'Faculty of ICT <noreply@localhost>')
