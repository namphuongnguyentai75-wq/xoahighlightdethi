FROM python:3.11-slim

# Install LibreOffice for PDF conversion and other dependencies
RUN sed -i -e 's/Components: main/Components: main contrib non-free/g' /etc/apt/sources.list.d/debian.sources || sed -i -e 's/ main/ main contrib non-free/g' /etc/apt/sources.list && \
    apt-get update && apt-get install -y debconf-utils && \
    echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections && \
    apt-get install -y \
    libreoffice \
    libreoffice-core \
    libreoffice-writer \
    default-jre \
    fonts-liberation \
    fonts-noto \
    ttf-mscorefonts-installer \
    fontconfig \
    && fc-cache -f -v \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy requirements and install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Expose port
EXPOSE 8000

# Start FastAPI
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
