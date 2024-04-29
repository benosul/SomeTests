# Use the official Python base image
FROM python:3.9-alpine

# Copy the entry point script to the working directory
COPY entrypoint.sh /entrypoint.sh

# Grant execute permissions to the entry point script
RUN chmod +x /entrypoint.sh

# Set the entry point to the script
ENTRYPOINT ["/entrypoint.sh"]
