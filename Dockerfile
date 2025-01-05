# Use an official Python runtime as a base image
FROM python:3.9-slim

# Set the working directory inside the container
WORKDIR /app

# Copy the current directory contents into the container at /app
COPY . /app

# Install any required Python dependencies
RUN pip install --no-cache-dir -r requirements.txt


ENV BOT_TOKEN="7694519326:AAEIOqEYtMiQvxk5n6H7is1Nv6-cOAJZq-8"


# Command to run your Python application
CMD ["python3", "bot.py"]

