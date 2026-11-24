FROM python:3.12-slim
WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py login.html ./

# タスクリストの画面一式
COPY index.html style.css app.js m365-login.js auth.js ./static/

RUN useradd -m -u 10001 appuser && chown -R appuser:appuser /app
USER appuser

ENV STATIC_DIR=/app/static
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8080"]
