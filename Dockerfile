FROM python:3.11-slim

WORKDIR /app

# Instalar dependencias del sistema necesarias para numpy y pandas
RUN apt-get update && apt-get install -y \
    build-essential \
    libpq-dev \
    && rm -rf /var/lib/apt/lists/*

# Copiar los archivos de requisitos primero para aprovechar la caché de Docker
COPY requirements.txt .

# Instalar dependencias de Python
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el código de la aplicación
COPY . .

# Configurar variables de entorno
ENV PORT=5000
ENV OLLAMA_URL=https://evaenespanol.loca.lt
ENV MODEL_NAME=llama3:8b
ENV FRONTEND_URL=https://contafin-front.vercel.app
ENV ENVIRONMENT=production

# Exponer el puerto de la aplicación
EXPOSE 5000

# Comando para iniciar la aplicación con gunicorn
CMD gunicorn --bind 0.0.0.0:$PORT --workers 4 --timeout 120 app:app