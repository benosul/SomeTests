# Use the official Python base image
FROM 3.11-alpine

# Copy the entry point script to the working directory
COPY src/ .
COPY Rules/ .
COPY main.py .


# Set the entry point to the script
ENTRYPOINT ["python", "main.py"]
