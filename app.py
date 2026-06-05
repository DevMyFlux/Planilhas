from io import BytesIO
import os
import tempfile
import traceback
from pathlib import Path
from uuid import uuid4

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from openpyxl.utils.exceptions import InvalidFileException
from werkzeug.utils import secure_filename

from beautifier import beautify_workbook


ALLOWED_EXTENSIONS = {".pdf", ".xls", ".xlsx", ".xlsm"}

app = Flask(__name__)
app.config["SECRET_KEY"] = "excel-bonito-secret"
# Stream uploads to disk instead of buffering in memory.
# Werkzeug will spool anything above this threshold to a temp file.
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB hard limit
app.config["MAX_FORM_MEMORY_SIZE"] = 0  # always spool to disk


def is_allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/upload")
def upload_file():
    uploaded_file = request.files.get("file")

    if uploaded_file is None or uploaded_file.filename == "":
        flash("Selecione um arquivo PDF ou Excel para continuar.")
        return redirect(url_for("index"))

    if not is_allowed_file(uploaded_file.filename):
        flash("Envie um arquivo .pdf, .xls, .xlsx ou .xlsm.")
        return redirect(url_for("index"))

    original_name = secure_filename(uploaded_file.filename)
    original_extension = Path(original_name).suffix.lower()
    output_name = f"{Path(original_name).stem}_organizado_{uuid4().hex[:8]}.xlsx"

    # Write the upload to a named temp file so the file is never fully held
    # in process memory — the OS page cache handles buffering instead.
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=original_extension)
    try:
        with os.fdopen(tmp_fd, "wb") as tmp_file:
            uploaded_file.save(tmp_file)

        try:
            output_stream = beautify_workbook(tmp_path, input_extension=original_extension)
        except InvalidFileException:
            flash("Nao foi possivel abrir esse arquivo. Confira se ele e um Excel valido.")
            return redirect(url_for("index"))
        except ValueError as exc:
            flash(str(exc))
            return redirect(url_for("index"))
        except Exception:
            traceback.print_exc()
            flash("O arquivo foi lido, mas houve um erro inesperado ao organizar a planilha.")
            return redirect(url_for("index"))
    finally:
        # Always remove the temp file, even if processing raised an exception.
        try:
            os.unlink(tmp_path)
        except OSError:
            pass

    output_stream.seek(0)

    return send_file(
        output_stream,
        as_attachment=True,
        download_name=output_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
