FROM python:3.11-slim

# Install LibreOffice + fontconfig + Thai locale support
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice \
        fontconfig \
        locales \
    && sed -i '/th_TH.UTF-8/s/^# //g' /etc/locale.gen \
    && locale-gen \
    && rm -rf /var/lib/apt/lists/*

# Copy fonts from your actual project folder into system font directory
COPY letters/static/fonts/*.ttf /usr/share/fonts/truetype/custom/
RUN fc-cache -f -v

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

# Point to your actual WSGI module and Render's standard port
CMD ["gunicorn", "internship_project.wsgi:application", "--bind", "0.0.0.0:10000"]