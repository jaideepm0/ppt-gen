# Use an official Python image
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install system dependencies for Pillow (libjpeg, zlib)
RUN apt-get update && apt-get install -y libjpeg-dev zlib1g-dev

# Copy project files into the container
COPY . /app

# Copy the fonts directory into the Docker container

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Expose the Flask port
EXPOSE 7860

# Run the Flask app with gunicorn for production
CMD ["gunicorn", "-b", "0.0.0.0:7860", "app:app"]
