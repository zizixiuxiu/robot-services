# 8006 workshop order conversion service
FROM python:3.13-slim

WORKDIR /app

# Required for .xls -> .xlsx conversion while preserving workbook formatting.
# Do not replace this with pandas conversion: pandas rewrites cell data only and
# loses styles, merged cells, column widths, images, and page layout.
RUN sed -i \
        -e 's|http://deb.debian.org/debian-security|https://mirrors.aliyun.com/debian-security|g' \
        -e 's|http://deb.debian.org/debian|https://mirrors.aliyun.com/debian|g' \
        /etc/apt/sources.list.d/debian.sources \
    && apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-calc fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

COPY config/requirements.txt ./config/requirements.txt
RUN pip install --no-cache-dir -r config/requirements.txt

COPY src/ ./src/

RUN mkdir -p /app/data/output

EXPOSE 8006

ENV PYTHONUNBUFFERED=1

CMD ["python", "src/workshop_order_http.py", "--port", "8006"]
