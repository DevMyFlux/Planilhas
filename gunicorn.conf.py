# Gunicorn production configuration for the Planilhas service.
# https://docs.gunicorn.org/en/stable/settings.html

import multiprocessing

# ── Binding ───────────────────────────────────────────────────────────────────
bind = "0.0.0.0:5000"

# ── Workers ───────────────────────────────────────────────────────────────────
# A single worker processing one of the larger SOULMV Razao exports (tens of
# thousands of raw rows) peaks around 225-240 MB RSS while building the row
# list and the styled output workbook (measured locally on the ~115k-row HCN
# file; matches the ~240 MB reported in production before crashes started).
# It settles back to a low resting baseline once the request finishes - this
# is not a per-request leak that accumulates over a worker's life - but that
# per-request PEAK is real and transient. At 4 workers, two or more big-file
# uploads landing on different workers at the same time is enough to push
# combined RSS past the 1 GB service limit and get OOM-killed. Capping at 2
# keeps the worst realistic case (2 workers both mid-peak) around ~500 MB,
# with real headroom.
workers = min(multiprocessing.cpu_count() * 2 + 1, 2)
worker_class = "sync"

# Load the app (and its heavy imports - openpyxl, pdfplumber, etc, ~35-60 MB)
# once in the master process before forking, so workers share those pages via
# copy-on-write instead of each paying that baseline separately.
preload_app = True

# ── Timeouts ──────────────────────────────────────────────────────────────────
# A 735-page Razao PDF takes ~90 s to parse on a fast dev machine; give each
# request 300 s so a slower container isn't killed by the master mid-request.
timeout = 300
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
