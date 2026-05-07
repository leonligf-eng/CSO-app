# 使用輕量級的 Python 基礎映像檔
FROM python:3.10-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

EXPOSE 8080

# ⚠️ 注意這裡：檔名必須是 streamlit_app.py
CMD ["streamlit", "run", "streamlit_app.py", "--server.port=8080", "--server.address=0.0.0.0", "--server.enableCORS=false"]