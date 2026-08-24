# Internship Letter Generator (Django + PostgreSQL + Tailwind)

Upload internship data (CSV/Excel) → preview → generate the internship
letters as .docx files (with the Mahidol University seal on the letterhead)
→ download as a zip. All uploads and generated documents are stored in
PostgreSQL/media, not sessions — safe for multiple staff using it at once.

## Setup

1. Create a Postgres database (or update the env vars below to point at
   an existing one):
   ```
   createdb internship_letters
   ```
2. Set connection details — either export these or copy `.env.example` to
   `.env` and load it however your project usually does (e.g. `python-dotenv`,
   or your existing Django project's settings if you're merging this in):
   ```
   DB_NAME=internship_letters
   DB_USER=postgres
   DB_PASSWORD=postgres
   DB_HOST=localhost
   DB_PORT=5432
   ```
3. Install and run:
   ```
   pip install -r requirements.txt
   python manage.py migrate
   python manage.py runserver
   ```
4. Open http://localhost:8000

## Structure

- `letters/` — the actual app. If you already have a Django project, copy
  this folder in, add `'letters'` to INSTALLED_APPS, and
  `path('', include('letters.urls'))` to your urls.py. Also copy the
  MEDIA_URL/MEDIA_ROOT setting and the `static(...)` line in urls.py (dev only —
  production should serve media via your web server / object storage).
  - `docgen.py` — the letter-building logic (plain python-docx, no Django
    dependency). Edit this to fix layout/wording, swap the logo, or add a
    new letter type.
  - `models.py` — `UploadedBatch` / `StudentRecord` / `GeneratedDocument` /
    `GenerationBatch`. Row data is kept as JSON on `StudentRecord.data` so
    the schema doesn't break every time the Google Form changes a column.
  - `views.py` — upload -> preview -> generate -> download, all DB-backed.
  - `templates/letters/` — Tailwind (via CDN script tag — swap for a real
    Tailwind build + `django-tailwind` before production, since the CDN
    build isn't meant for production use).
  - `static/letters/images/mu_logo.jpeg` — the university seal, extracted
    from your PDF templates. Swap this file for a higher-resolution version
    if you have one.

## Preview page

The preview table now shows company, contractor (contact person + their
title), internship position, and period directly — plus a ⚠ badge if any
of those key fields are blank in the source row. Click "View all →" on a
row to see every column from the upload for that student.

## Comparing a generated document against the source Excel

From the results page, click "Compare with Excel →" on any generated
document. This opens a split-panel view — Excel data on the left, the
*actual text read back out of the generated .docx* on the right (not just
the Excel values shown twice) — with per-field ✓ Match / ⚠ Partial / ✕
Mismatch badges. Hover a field on either side and its match on the other
panel highlights too.

The extraction logic (`extract_fields_from_docx` in docgen.py) parses the
letter's recipient block and student table back out with plain
string/regex matching. It's tuned to the current letter wording — if you
change the wording in `build_confirm_letter` / `build_request_letter`,
double check the extraction still matches (the period regex in particular
looks for "ระหว่างวันที่ ... จำนวน").

## What it supports right now

- **Student Referral Letter** -> matches "Confirm_Letter_V1" template
- **Request Letter for Internship Placement** -> matches "Request_Intern_Letter_V1" template
- **Internship Extension Request Letter** -> no source template yet, rows of
  this type are flagged and skipped rather than generated incorrectly.

## Known shortcuts taken (fix before real staff use)

- **Tailwind via CDN**: fine for development, but Tailwind's own docs say
  not to use the CDN build in production (no purging, slower, no
  customization pipeline). Set up `django-tailwind` or a proper CLI build
  before deploying.
- **Layout accuracy**: rebuilt from the PDF templates by hand (no original
  font metrics), since we only had PDFs, not the source .docx. If you get
  the real .docx, `docxtpl` + `{{merge_fields}}` against the actual file
  will match exactly rather than approximate it.
- **Logo resolution**: extracted from the PDF at low res (199×196px JPEG).
  Fine for the web header; you may want a higher-res version for the
  printed letter.
- **Reference number** (`ที่ อว 78.363/...`) just increments per row — no
  real numbering/sequencing logic wired in yet.
- **No auth** — anyone who can reach the site can upload/generate. Add
  Django's login_required (or your existing project's auth) before this
  is staff-facing.
# internship-document-website
