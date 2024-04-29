# Use the official Python base image
FROM python:3.11

# Copy the entry point script to the working directory
COPY src/ ./src/
COPY Rules/ ./Rules/
COPY main.py .


# Set the entry point to the script
ENTRYPOINT ["python", "main.py"]
