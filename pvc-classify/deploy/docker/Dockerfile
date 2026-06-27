# 8005 PVC 自动分类服务
FROM python:3.12-slim

WORKDIR /app

# 先复制依赖文件并安装，利用缓存层
COPY config/requirements.txt ./config/requirements.txt
RUN pip install --no-cache-dir -r config/requirements.txt

# 复制应用代码（运行时通常由 volume 覆盖，用于保证镜像完整）
COPY src/ ./src/

# 输出目录
RUN mkdir -p /app/data/output /app/logs

EXPOSE 8005

ENV PYTHONUNBUFFERED=1

CMD ["python", "src/pvc_classify_http.py", "--port", "8005"]
