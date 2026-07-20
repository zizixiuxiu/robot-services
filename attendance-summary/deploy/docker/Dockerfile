# 8009 考勤汇总服务
FROM docker.1ms.run/python:3.11-slim

WORKDIR /app

# 先复制依赖文件并安装，利用缓存层
COPY config/requirements.txt ./config/requirements.txt
RUN pip install --no-cache-dir -i https://pypi.tuna.tsinghua.edu.cn/simple -r config/requirements.txt

# 复制应用代码（运行时通常由 volume 覆盖，用于保证镜像完整）
COPY src/ ./src/

# 输出目录
RUN mkdir -p /app/data/output

EXPOSE 8009

ENV PYTHONUNBUFFERED=1

CMD ["python", "src/attendance_summary_http.py", "--port", "8009"]
