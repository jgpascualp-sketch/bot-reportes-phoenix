#!/usr/bin/env bash
# Salir si hay error
set -o errexit

# Actualizar paquetes e instalar LibreOffice sin entorno gráfico
apt-get update && apt-get install -y libreoffice --no-install-recommends

# Instalar librerías de Python
pip install --upgrade pip
pip install -r requirements.txt
