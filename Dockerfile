FROM python:3.11-slim

# Set environment variables
ENV PYTHONUNBUFFERED=1 \
    DEBIAN_FRONTEND=noninteractive \
    PORT=5000

WORKDIR /app

# Install minimal base utilities
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Install Playwright browser & all required Linux OS libraries
RUN playwright install --with-deps chromium

# Copy application source code
COPY . .

# Expose dynamic application port
EXPOSE 5000

# Start production server with 1 worker and 4 threads to preserve in-memory batch state
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT:-5000} --workers 1 --threads 4 --timeout 0 pcc_web_app:app"]
