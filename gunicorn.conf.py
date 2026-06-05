# Gunicorn production configuration for the Planilhas service.
# https://docs.gunicorn.org/en/stable/settings.html

import multiprocessing

# ── Binding ───────────────────────────────────────────────────────────────────
bind = "0.0.0.0:5000"

# ── Workers ───────────────────────────────────────────────────────────────────
# Spreadsheet processing is CPU-bound; use (2 × cores + 1) as a starting point
# but cap at 4 to stay within the 1 GB memory budget.
workers = min(multiprocessing.cpu_count() * 2 + 1, 4)
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
# memory that leaked during spreadsheet processing.
max_requests = 200
max_requests_jitter = 40

# ── Request size limits ────────────────────────────────────────────────────────
# By default Gunicorn enforces conservative limits that cause HTTP 413 errors
# for large file uploads before the request even reaches Flask.  Set all three
# to 0 (unlimited) so that Flask's own MAX_CONTENT_LENGTH = 50 MB is the sole
# gatekeeper.
limit_request_line = 0
limit_request_fields = 0
limit_request_field_size = 0
