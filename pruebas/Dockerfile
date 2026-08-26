FROM python:3.9.13-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV TZ=America/Lima

CMD ["python", "reporteetm.py"]
