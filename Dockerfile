FROM python:3.12-slim
WORKDIR /app
COPY webapp/requirements.txt webapp/requirements.txt
RUN pip install --no-cache-dir -r webapp/requirements.txt
COPY tools/sample_data.py tools/msp_mirror.py tools/
COPY webapp webapp
ENV DATA_DIR=/data PYTHONUNBUFFERED=1
WORKDIR /app/webapp
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000} --proxy-headers --forwarded-allow-ips='*'"]
