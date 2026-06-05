#!/usr/bin/env bash
# Production startup script for the Planilhas service.
# Uses Gunicorn instead of Flask's built-in development server.
set -euo pipefail

exec gunicorn "app:app" --config gunicorn.conf.py
