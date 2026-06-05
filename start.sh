#!/usr/bin/env bash
# Production startup script for the Planilhas service.
# Uses Gunicorn instead of Flask's built-in development server.
set -euo pipefail

PORT="${PORT:-5000}"
exec gunicorn wsgi:app --config gunicorn.conf.py --bind "0.0.0.0:${PORT}"
