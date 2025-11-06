FROM python:3.9-slim

# Instalar dependencias del sistema
RUN apt-get update && apt-get install -y \
    wget \
    gnupg \
    unzip \
    ca-certificates \
    fonts-liberation \
    libnss3 \
    libgconf-2-4 \
    libxss1 \
    libasound2 \
    libatk1.0-0 \
    libatk-bridge2.0-0 \
    libcups2 \
    libdbus-1-3 \
    libgtk-3-0 \
    libx11-xcb1 \
    libxcomposite1 \
    libxrandr2 \
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
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT:-8000} run:app --workers 1 --timeout 120 --log-level debug"]