# 8007 quote maker service
FROM python:3.12-slim

WORKDIR /app

COPY config/requirements.txt ./config/requirements.txt
RUN pip install --no-cache-dir -r config/requirements.txt

COPY src/ ./src/

RUN mkdir -p /app/data/output /app/logs

EXPOSE 8007

ENV PYTHONUNBUFFERED=1

CMD ["python", "src/quote_maker_http.py", "--port", "8007"]
