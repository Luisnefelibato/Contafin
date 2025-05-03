FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Establecer variables de entorno
ENV PORT=5000
ENV OLLAMA_URL=https://evaenespanol.loca.lt
ENV MODEL_NAME=llama3:8b

# Puerto en el que se ejecutará la aplicación
EXPOSE $PORT

# Comando para iniciar la aplicación
CMD gunicorn --bind 0.0.0.0:$PORT app:app