# Use an official Python runtime as a parent image
FROM python:3.9-slim

# --- Install System Dependencies ---
# This is the key step: install LibreOffice and poppler-utils (for pdf2image)
RUN apt-get update && apt-get install -y \
    libreoffice \
    poppler-utils \
    --no-install-recommends && \
    rm -rf /var/lib/apt/lists/*

# Set the working directory in the container
WORKDIR /app

# Copy the requirements file and install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of your application code into the container
COPY . .

# Tell Render what command to run when starting the web service
# Use gunicorn to run your Flask app
CMD ["gunicorn", "--bind", "0.0.0.0:10000", "--workers", "2", "--timeout", "120", "app:app"]