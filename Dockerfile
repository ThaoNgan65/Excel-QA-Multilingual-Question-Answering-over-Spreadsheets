FROM python:3.11-slim-bookworm
ENV PYTHONUNBUFFERED=1 PYTHONIOENCODING=utf-8
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY excelqa.py .
USER 0:0
ENTRYPOINT ["python", "excelqa.py"]