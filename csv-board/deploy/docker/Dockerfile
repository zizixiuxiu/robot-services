# 8004 CSV 板件转换服务
FROM python:3.12-slim

WORKDIR /app

COPY src/ ./src/

# 公共库 factory-common
COPY --from=factory-common factory-common/factory_common /app/factory_common
ENV PYTHONPATH=/app

RUN mkdir -p /app/data/output /app/logs

EXPOSE 8004

ENV PYTHONUNBUFFERED=1

CMD ["python", "src/csv_board_http.py", "--port", "8004"]
