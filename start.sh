#!/bin/bash
# Start DataPack Platform

cd "$(dirname "$0")"

# Activate venv if exists
if [ -d "venv" ]; then
    source venv/bin/activate
elif [ -d ".venv" ]; then
    source .venv/bin/activate
fi

# Load .env if exists
if [ -f ".env" ]; then
    export $(cat .env | xargs)
fi

# Start server
uvicorn backend.main:app --host 0.0.0.0 --port 8000 --reload
