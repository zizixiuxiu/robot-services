FROM python:3.12-slim

WORKDIR /app

COPY config/requirements.txt ./config/requirements.txt
RUN pip install --no-cache-dir -r config/requirements.txt

COPY src/ ./src/

RUN mkdir -p /app/logs

ENV PYTHONUNBUFFERED=1
ENV FEISHU_LOG_FILE=/app/logs/feishu_bot_ws.log
ENV HTTP_SERVICE_HOST=host.docker.internal

CMD ["python", "src/feishu_bot_ws.py"]
