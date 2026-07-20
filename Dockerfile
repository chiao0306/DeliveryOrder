FROM python:3.10-slim

# 設定工作目錄
WORKDIR /app

# 安裝相依套件
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 複製專案原始碼
COPY . .

# Cloud Run 預設會提供 PORT 環境變數
EXPOSE 8080

# 啟動 FastAPI 伺服器
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-8080}"]