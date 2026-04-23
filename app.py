from io import BytesIO
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

    workbook_bytes = BytesIO(uploaded_file.read())
    try:
        output_stream = beautify_workbook(workbook_bytes, input_extension=original_extension)
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

    output_stream.seek(0)

    return send_file(
        output_stream,
        as_attachment=True,
        download_name=output_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
