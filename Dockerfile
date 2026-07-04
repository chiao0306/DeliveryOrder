FROM python:3.10-slim
WORKDIR /app
COPY . .
RUN pip install --no-cache-dir -r requirements.txt
EXPOSE 8080
# 💡 加上 --server.headless=true、--client.toolbarMode=minimal 確保雲端順利開機
CMD ["streamlit", "run", "app.py", "--server.port=8080", "--server.address=0.0.0.0", "--server.headless=true", "--client.toolbarMode=minimal"]
