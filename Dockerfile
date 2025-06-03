# Use a lightweight Python image
FROM python:3.12-slim

# Set working directory
WORKDIR /app

# Install system dependencies for pandas and Excel
RUN apt-get update && apt-get install -y \
    build-essential \
    libpq-dev \
    && rm -rf /var/lib/apt/lists/*

# Copy project files into the image
COPY . .

# Install required Python packages
RUN pip install --no-cache-dir -r requirements.txt

# Expose port Flask will run on
EXPOSE 8080

# Start the app using Gunicorn
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "app:app"]