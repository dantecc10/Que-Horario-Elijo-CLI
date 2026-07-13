#!/bin/bash

PROJECT_DIR="/media/darkseid/DATA/Repos/Que-Horario-Elijo-CLI"
LOCAL_PORT=5000
REMOTE_PORT=2026
SERVER="root@castelancarpinteyro.com"

cd "$PROJECT_DIR" || exit 1

echo "Activando entorno virtual..."
source .venv/bin/activate

echo "🚀 Iniciando gunicorn en puerto $LOCAL_PORT (1 worker, 4 threads)..."
gunicorn app_web:app --bind 127.0.0.1:$LOCAL_PORT --workers 1 --threads 4 --timeout 120 --access-logfile - &
FLASK_PID=$!

sleep 2

echo "🔗 Conectando túnel: Tu PC ($LOCAL_PORT) <---> Servidor ($REMOTE_PORT)"
echo "🌍 URL pública: https://preview.castelancarpinteyro.com"
echo "🛑 Ctrl+C para detener"

cleanup() {
    echo ""
    echo "Deteniendo servidor y túnel..."
    kill $FLASK_PID 2>/dev/null
    exit 0
}
trap cleanup INT TERM

ssh -R $REMOTE_PORT:localhost:$LOCAL_PORT $SERVER -N
