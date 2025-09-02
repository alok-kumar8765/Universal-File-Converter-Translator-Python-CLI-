
---

# 🐳 Dockerfile

```dockerfile
FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY unified_converter.py .

CMD ["python", "file_toolkit.py"]
