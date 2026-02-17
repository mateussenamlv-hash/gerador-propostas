FROM python:3.11-slim

RUN apt-get update && apt-get install -y libreoffice

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . .

CMD ["python", "app.py"]
