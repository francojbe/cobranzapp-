# Imagen base ligera oficial de Python
FROM python:3.12-slim

# Evitar escritura de archivos .pyc y habilitar logs en tiempo real
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=80

# Directorio de trabajo
WORKDIR /app

# Instalar dependencias mínimas del sistema
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Instalar dependencias de Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el código y los archivos necesarios
COPY . .

# Exponer puertos 80 y 8000 para máxima compatibilidad con Easypanel
EXPOSE 80 8000

# Arrancar uvicorn escuchando en 0.0.0.0 y usando la variable PORT (por defecto 80)
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-80}"]
