#!/usr/bin/env bash
# Production startup script for the Planilhas service.
# Uses Gunicorn instead of Flask's built-in development server.
set -euo pipefail

exec gunicorn wsgi:app --workers 4 --worker-class sync --timeout 120 --max-requests 200
