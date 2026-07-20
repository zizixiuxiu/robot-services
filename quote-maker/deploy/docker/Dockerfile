# 8007 quote maker service
FROM python:3.12-slim

WORKDIR /app

COPY config/requirements.txt ./config/requirements.txt
RUN pip install --no-cache-dir -r config/requirements.txt

COPY src/ ./src/

# 公共库 factory-common（来自 compose additional_contexts，源在仓库根目录）
COPY --from=factory-common factory-common/factory_common /app/factory_common
ENV PYTHONPATH=/app

RUN mkdir -p /app/data/output /app/logs

EXPOSE 8007

ENV PYTHONUNBUFFERED=1

CMD ["python", "src/quote_maker_http.py", "--port", "8007"]
