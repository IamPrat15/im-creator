# Use Python 3.11 as base image
FROM python:3.11-slim

# Set working directory inside the container
WORKDIR /app

# Copy everything from your server folder into the container
COPY server/ /app/

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Expose the port Render will use
EXPOSE 8000

# Start FastAPI using Uvicorn
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]