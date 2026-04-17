# 1. Bắt đầu với một hệ điều hành Linux có sẵn Python 3.10
FROM python:3.10-slim

# 2. Cài đặt LibreOffice
RUN apt-get update && \
    apt-get install -y libreoffice --no-install-recommends && \
    rm -rf /var/lib/apt/lists/*

# 3. Đặt thư mục làm việc trong máy chủ
WORKDIR /app

# 4. Copy file cấu hình và cài các thư viện Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 5. Copy toàn bộ code của bạn vào máy chủ
COPY . .

# 6. Khởi động máy chủ web bằng Gunicorn
CMD gunicorn app:app --bind 0.0.0.0:${PORT:-10000}