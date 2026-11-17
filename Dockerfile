FROM python:3.10-slim

# Ngăn Python tạo bytecode .pyc và bật unbuffered logging
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=8080

# Cài đặt các gói hệ thống cần thiết (bao gồm LibreOffice headless để hỗ trợ cập nhật số trang mục lục trên Linux)
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    curl \
    libreoffice-writer \
    fonts-liberation \
    fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Sao chép và cài đặt thư viện Python
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Sao chép mã nguồn ứng dụng
COPY . .

# Mở cổng mặc định (Google Cloud Run dùng biến môi trường $PORT, mặc định 8080)
EXPOSE 8080

# Chạy Streamlit lắng nghe trên tất cả interface và hỗ trợ cổng động của Cloud Run
ENTRYPOINT ["sh", "-c", "streamlit run app.py --server.port=${PORT:-8080} --server.address=0.0.0.0 --server.headless=true --browser.gatherUsageStats=false"]
