# ใช้ Python base image ที่เบาแต่เพียงพอ
FROM python:3.9-slim

# ติดตั้ง Java และ dependencies ที่จำเป็น
RUN apt-get update && \
    apt-get install -y default-jre curl && \
    apt-get clean

# ตั้ง environment variable ให้ tabula ใช้ Java ได้
ENV PATH="/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin"

# คัดลอก requirements.txt และติดตั้ง Python packages
COPY requirements.txt .
RUN pip install --upgrade pip && pip install -r requirements.txt

# คัดลอกไฟล์ทั้งหมดในโปรเจกต์เข้ามา
COPY . /app
WORKDIR /app

# รัน streamlit
CMD ["streamlit", "run", "app.py", "--server.port=10000", "--server.enableCORS=false"]