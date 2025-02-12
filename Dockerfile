FROM python:3.10-slim

WORKDIR /app

# Installation des dépendances système nécessaires
RUN apt-get update && apt-get install -y \
    build-essential \
    python3-numpy \
    python3-pandas \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
# Modifier requirements.txt pour exclure numpy et pandas car ils sont déjà installés via apt
RUN grep -v "numpy\|pandas" requirements.txt > requirements_filtered.txt && \
    pip install --no-cache-dir -r requirements_filtered.txt

COPY . .

ENV PORT=8080
CMD gunicorn --workers=2 --timeout=120 --bind 0.0.0.0:$PORT scriptMake:app 