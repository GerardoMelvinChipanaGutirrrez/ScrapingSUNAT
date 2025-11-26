# 🕷️ ScrapingSUNAT

Este proyecto es una herramienta de **web scraping automatizado** desarrollada en Python para consultar datos de empresas en el portal de **SUNAT (Perú)**.  
A partir de un archivo Excel que contiene **Razones Sociales**, el programa extrae:

- **RUC**
- **Razón social encontrada**
- **Responsable legal o representante**
- **Estado de la empresa y mensajes de error**

Los resultados se guardan en un nuevo archivo Excel, facilitando procesos de verificación empresarial.

---

## 📌 Características principales

- 📝 Lectura automática de Excel con razones sociales.
- 🌐 Scraping de SUNAT usando Selenium.
- 🔄 Manejo de alertas, tiempos y reintentos.
- 🧹 Limpieza y validación de cadenas con expresiones regulares.
- 📤 Exportación final a Excel con los datos recopilados.
- ⚠️ Detección de empresas no encontradas o con errores de búsqueda.

---

## 📦 Instalación

### 1️⃣ Requisitos previos

- Python **3.8 o superior**
- Google Chrome
- Pip actualizado

Para actualizar pip:

```bash
pip install --upgrade pip

Razón social encontrada

Responsable legal o representante

Estado de la empresa y mensajes de error

Los resultados se guardan en un nuevo archivo Excel, facilitando procesos de verificación empresarial.

📌 Características principales

📝 Lectura automática de Excel con razones sociales.

🌐 Scraping de SUNAT usando Selenium.

🔄 Manejo de alertas, tiempos y reintentos.

🧹 Limpieza y validación de cadenas con expresiones regulares.

📤 Exportación final a Excel con los datos recopilados.

⚠️ Detección de empresas no encontradas o con errores de búsqueda.

📦 Instalación
1️⃣ Requisitos previos

Python 3.8 o superior

Google Chrome

Pip actualizado

Para actualizar pip:

pip install --upgrade pip

2️⃣ Instalación de librerías necesarias

Ejecuta estos comandos en tu terminal o PowerShell:

pip install pandas selenium webdriver-manager openpyxl


Esto instalará:

pandas → Manejo de archivos Excel

selenium → Automatiza el navegador

webdriver-manager → Instala y actualiza automáticamente el driver de Chrome

openpyxl → Permite leer/escribir archivos Excel


📂 Cómo usar el scraper

Coloca tu archivo Excel en la carpeta del proyecto.

Debe tener una columna con la Razón Social de las empresas.

Ejecuta el script:

python scraper.py


El programa:

Detectará el Excel automáticamente.

Buscará cada razón social en la web de SUNAT.

Extraerá RUC y responsable.

Guardará todo en un nuevo archivo Excel con fecha y hora.

🔧 Flujo de funcionamiento

Leer Excel usando pandas.

Abrir el navegador con Selenium.

Ingresar la razón social en la web de SUNAT.

Extraer:

RUC

Razón social normalizada

Responsable legal

Manejar errores como:

Timeout

Alertas de “Solo letras y números”

Empresa no encontrada

Guardar resultados en un archivo Excel final.

🎯 Objetivo del proyecto

Este scraper fue desarrollado para automatizar la verificación masiva de empresas, ideal para:

Estudios contables

Áreas administrativas

Validación de proveedores

Auditorías y control interno

Integración con sistemas ERP