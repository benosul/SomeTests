FROM python:3-slim AS builder
ADD . /app
WORKDIR /app

# We are installing a dependency here directly into our app source dir
# RUN pip install --target=/app os
# RUN pip install --target=/app logging
# RUN pip install --target=/app datetime
# RUN pip install --target=/app re

# A distroless container image with Python and some basics like SSL certificates
# https://github.com/GoogleContainerTools/distroless
FROM gcr.io/distroless/python3-debian10
COPY --from=builder /app /app
COPY testFilesAndDirs/Rules/Rules_Avoid.txt /
COPY testFilesAndDirs/Rules/Rules_Have.txt /
WORKDIR /app
ENV PYTHONPATH /app
CMD ["/app/main.py"]
