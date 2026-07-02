# 8008 经销商数据报表服务
FROM python:3.13-slim

WORKDIR /app

COPY config/requirements.txt ./config/requirements.txt
RUN pip install --no-cache-dir -r config/requirements.txt

COPY src/ ./src/

RUN mkdir -p /app/data/output /app/logs

EXPOSE 8008

ENV PYTHONUNBUFFERED=1
ENV OUTPUT_DIR=/app/data/output
ENV LOG_DIR=/app/logs

CMD ["python", "src/dealer_report_http.py", "--port", "8008"]
