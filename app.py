import json
import os
import re
import tempfile
import threading
import time
import traceback
from pathlib import Path
from uuid import uuid4

from flask import Flask, abort, jsonify, render_template, request, send_file, url_for
from openpyxl.utils.exceptions import InvalidFileException
from werkzeug.utils import secure_filename

from beautifier import beautify_workbook, set_progress_callback


ALLOWED_EXTENSIONS = {".pdf", ".xls", ".xlsx", ".xlsm"}

app = Flask(__name__)
app.config["SECRET_KEY"] = "excel-bonito-secret"
# Stream uploads to disk instead of buffering in memory.
# Werkzeug will spool anything above this threshold to a temp file.
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB hard limit

# Conversions of large PDFs take minutes, far longer than a proxy will hold a
# request open. /upload therefore only stores the file and starts a background
# thread; the page polls /status and downloads from /download. Job state lives
# on disk (not in memory) because the gunicorn workers are separate processes
# and a status poll may land on a different one than the worker running the job.
JOBS_DIR = Path(tempfile.gettempdir()) / "planilhas_jobs"
JOBS_DIR.mkdir(parents=True, exist_ok=True)
JOB_TTL_SECONDS = 2 * 60 * 60
XLSX_MIMETYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_JOB_ID_RE = re.compile(r"^[0-9a-f]{32}$")

# One conversion at a time per worker process bounds peak memory (each worker
# runs at most one job; with 2 workers that is at most 2 concurrent jobs).
_conversion_slot = threading.BoundedSemaphore(1)


def is_allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def _status_path(job_id: str) -> Path:
    return JOBS_DIR / f"{job_id}.json"


def _result_path(job_id: str) -> Path:
    return JOBS_DIR / f"{job_id}.xlsx"


def _write_status(job_id: str, data: dict) -> None:
    path = _status_path(job_id)
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(data), encoding="utf-8")
    os.replace(tmp, path)


def _read_status(job_id: str) -> dict | None:
    try:
        return json.loads(_status_path(job_id).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def _cleanup_old_jobs() -> None:
    cutoff = time.time() - JOB_TTL_SECONDS
    for path in JOBS_DIR.iterdir():
        try:
            if path.stat().st_mtime < cutoff:
                path.unlink()
        except OSError:
            pass


def _worker_alive(pid: int | None) -> bool:
    # A job whose worker process is gone (OOM kill, redeploy) would otherwise
    # look "processing" forever. os.kill(pid, 0) is only a harmless existence
    # probe on POSIX; on Windows signal 0 is CTRL_C_EVENT, so skip the check.
    if pid is None or os.name != "posix":
        return True
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def _run_job(job_id: str, source_path: Path, extension: str, output_name: str) -> None:
    status = {
        "status": "queued",
        "pid": os.getpid(),
        "filename": output_name,
        "progress": None,
        "detail": "Na fila...",
    }
    _write_status(job_id, status)

    last_write = [0.0]

    def on_progress(done: int, total: int) -> None:
        now = time.monotonic()
        if done < total and now - last_write[0] < 1.0:
            return
        last_write[0] = now
        status["progress"] = int(done * 100 / total) if total else None
        status["detail"] = f"Pagina {done} de {total}"
        _write_status(job_id, status)

    try:
        with _conversion_slot:
            status["status"] = "processing"
            status["detail"] = "Processando..."
            _write_status(job_id, status)
            set_progress_callback(on_progress)
            try:
                output_stream = beautify_workbook(str(source_path), input_extension=extension)
                result_tmp = _result_path(job_id).with_suffix(".xlsx.tmp")
                result_tmp.write_bytes(output_stream.getvalue())
                del output_stream
                os.replace(result_tmp, _result_path(job_id))
                status.update(status="done", progress=100, detail="Pronto")
            finally:
                set_progress_callback(None)
    except InvalidFileException:
        status.update(
            status="error",
            message="Nao foi possivel abrir esse arquivo. Confira se ele e um Excel valido.",
        )
    except ValueError as exc:
        status.update(status="error", message=str(exc))
    except Exception:
        traceback.print_exc()
        status.update(
            status="error",
            message="O arquivo foi lido, mas houve um erro inesperado ao organizar a planilha.",
        )
    finally:
        try:
            source_path.unlink()
        except OSError:
            pass
        _write_status(job_id, status)


@app.get("/")
def index():
    return render_template("index.html")


@app.errorhandler(413)
def file_too_large(_error):
    return jsonify(error="O arquivo passa do limite de 50 MB."), 413


@app.post("/upload")
def upload_file():
    uploaded_file = request.files.get("file")

    if uploaded_file is None or uploaded_file.filename == "":
        return jsonify(error="Selecione um arquivo PDF ou Excel para continuar."), 400

    if not is_allowed_file(uploaded_file.filename):
        return jsonify(error="Envie um arquivo .pdf, .xls, .xlsx ou .xlsm."), 400

    original_name = secure_filename(uploaded_file.filename)
    original_extension = Path(original_name).suffix.lower()
    output_name = f"{Path(original_name).stem}_organizado_{uuid4().hex[:8]}.xlsx"

    _cleanup_old_jobs()

    job_id = uuid4().hex
    source_path = JOBS_DIR / f"{job_id}.src{original_extension}"
    uploaded_file.save(source_path)

    threading.Thread(
        target=_run_job,
        args=(job_id, source_path, original_extension, output_name),
        name=f"job-{job_id[:8]}",
    ).start()

    return jsonify(job_id=job_id, status_url=url_for("job_status", job_id=job_id)), 202


@app.get("/status/<job_id>")
def job_status(job_id: str):
    if not _JOB_ID_RE.match(job_id):
        abort(404)

    status = _read_status(job_id)
    if status is None:
        return jsonify(
            status="error",
            message="Trabalho nao encontrado ou expirado. Envie o arquivo novamente.",
        ), 404

    if status["status"] in ("queued", "processing") and not _worker_alive(status.get("pid")):
        return jsonify(
            status="error",
            message="O processamento foi interrompido (o servidor reiniciou). Envie o arquivo novamente.",
        )

    payload = {
        "status": status["status"],
        "progress": status.get("progress"),
        "detail": status.get("detail"),
        "message": status.get("message"),
    }
    if status["status"] == "done":
        payload["download_url"] = url_for("download_result", job_id=job_id)
    return jsonify(payload)


@app.get("/download/<job_id>")
def download_result(job_id: str):
    if not _JOB_ID_RE.match(job_id):
        abort(404)

    status = _read_status(job_id)
    result = _result_path(job_id)
    if status is None or status.get("status") != "done" or not result.exists():
        abort(404)

    return send_file(
        result,
        as_attachment=True,
        download_name=status["filename"],
        mimetype=XLSX_MIMETYPE,
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False, use_reloader=False)
