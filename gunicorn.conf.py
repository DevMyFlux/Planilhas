# Gunicorn production configuration for the Planilhas service.
# https://docs.gunicorn.org/en/stable/settings.html

import multiprocessing

# ── Binding ───────────────────────────────────────────────────────────────────
bind = "0.0.0.0:5000"

# ── Workers ───────────────────────────────────────────────────────────────────
# Spreadsheet processing is CPU-bound, but 1 GB memory limit allows only 1-2 workers.
# Use 2 workers to handle concurrent requests without OOM.
workers = 2
worker_class = "sync"

# ── Timeouts ──────────────────────────────────────────────────────────────────
# Large spreadsheets can take a while to process; give each request 120 s.
timeout = 120
graceful_timeout = 30
keepalive = 5

# ── Logging ───────────────────────────────────────────────────────────────────
accesslog = "-"   # stdout
errorlog = "-"    # stderr
loglevel = "info"

# ── Memory safety ─────────────────────────────────────────────────────────────
# Restart a worker after it has served this many requests to reclaim any
# memory that leaked during spreadsheet processing. Aggressive limit ensures
# long-lived cache is cleared.
max_requests = 100
max_requests_jitter = 20

# ── Request size limits ────────────────────────────────────────────────────────
# By default Gunicorn enforces conservative limits that cause HTTP 413 errors
# for large file uploads before the request even reaches Flask.  Set all three
# to 0 (unlimited) so that Flask's own MAX_CONTENT_LENGTH = 50 MB is the sole
# gatekeeper.
limit_request_line = 0
limit_request_fields = 0
limit_request_field_size = 0

