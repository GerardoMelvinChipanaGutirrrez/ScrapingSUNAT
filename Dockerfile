FROM python:3.9-slim

# Instalar dependencias del sistema
RUN apt-get update && apt-get install -y \
    wget \
    gnupg \
    unzip \
    chromium \
    chromium-driver \
    && rm -rf /var/lib/apt/lists/*

# Configurar variables de entorno para Chrome
ENV CHROME_BIN=/usr/bin/chromium
ENV CHROMEDRIVER_PATH=/usr/bin/chromedriver
ENV PYTHONUNBUFFERED=1

# Crear y configurar el directorio de la aplicación
WORKDIR /app

# Copiar requirements.txt primero para aprovechar el caché de Docker
COPY requirements.txt .

# Instalar dependencias de Python
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el resto del código
COPY . .

# Crear directorios necesarios
RUN mkdir -p uploads

# Exponer el puerto
EXPOSE 8000

# Comando para ejecutar la aplicación
CMD ["gunicorn", "--bind", "0.0.0.0:8000", "run:app"]