FROM python:3.12-slim

WORKDIR /app

COPY config/requirements.txt ./config/requirements.txt
RUN pip install --no-cache-dir -r config/requirements.txt

COPY src/app ./app

RUN mkdir -p /app/data /app/logs

EXPOSE 8090

ENV PYTHONUNBUFFERED=1
ENV DATABASE_URL=sqlite:////app/data/inventory.db

CMD ["python", "-m", "uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8090"]
