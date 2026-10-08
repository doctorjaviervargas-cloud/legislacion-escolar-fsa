FROM python:3.12-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN mkdir -p /app/data/Normativa
ENV PORT=8000 APP_DATA_DIR=/app/data
EXPOSE 8000
CMD ["gunicorn", "-c", "gunicorn.conf.py", "app:app"]
