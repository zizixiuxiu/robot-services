# 8003 经销商销售服务（5月业绩核对表）
FROM python:3.13-slim

# 安装 Node.js 20.x 和 npm（generate_may_sales_report.js 需要）
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    ca-certificates \
    gnupg \
    && mkdir -p /etc/apt/keyrings \
    && curl -fsSL https://deb.nodesource.com/gpgkey/nodesource-repo.gpg.key | gpg --dearmor -o /etc/apt/keyrings/nodesource.gpg \
    && echo "deb [signed-by=/etc/apt/keyrings/nodesource.gpg] https://deb.nodesource.com/node_20.x nodistro main" > /etc/apt/sources.list.d/nodesource.list \
    && apt-get update && apt-get install -y --no-install-recommends nodejs \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 先复制依赖文件并安装，利用缓存层
COPY config/package.json config/package-lock.json* ./
RUN npm install

# 复制应用代码（运行时通常由 volume 覆盖，用于保证镜像完整）
COPY src/dealer_sales_http.py src/generate_may_sales_report.js ./

# 输出目录
RUN mkdir -p /app/output_may_sales

EXPOSE 8003

ENV OUTPUT_DIR=/app/output_may_sales

CMD ["python", "dealer_sales_http.py", "--port", "8003"]
