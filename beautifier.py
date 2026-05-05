from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
import re
import unicodedata

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import pdfplumber
import xlrd


# ── HCN output headers ──────────────────────────────────────────────────────
BALANCETE_HCN_HEADERS = [
    "REG", "ContaContabil", "NomeConta", "SaldoAnterior",
    "Debito", "Credito", "SaldoAtual",
    "Estoque", "AtivoImob", "DepreciacaoAcumulativa",
    "ContaFinanceira", "ReservaDeContingência", "Centro de Custos",
]
DIARIO_HCN_HEADERS = [
    "REG", "DATA", "CLASSIFICAÇÃO", "DESCRIÇÃO",
    "HISTÓRICO", "DÉBITO", "CRÉDITO", "Centro de Custos",
]
RAZAO_HCN_HEADERS = [
    "REG", "NOME CONTA", "CONTA CONTÁBIL", "DATA",
    "SALDO ANTERIOR", "HISTÓRICO", "DÉBITO", "CRÉDITO", "Saldo", "Contra Partida",
]

DEFAULT_OUTPUT_HEADERS = ["Data", "Descricao", "Debito", "Credito", "Saldo"]

# ── Styles ───────────────────────────────────────────────────────────────────
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
ALT_ROW_FILL = PatternFill("solid", fgColor="F4F8FB")
TOTAL_BORDER = Border(
    bottom=Side(style="medium", color="1F1F1F"),
    top=Side(style="thin", color="D9E2F3"),
)
THIN_BORDER = Border(
    left=Side(style="thin", color="D9E2F3"),
    right=Side(style="thin", color="D9E2F3"),
    top=Side(style="thin", color="D9E2F3"),
    bottom=Side(style="thin", color="D9E2F3"),
)
ACCOUNTING_FORMAT = '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-'

# ── Regexes ──────────────────────────────────────────────────────────────────
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
MONEY_RE = re.compile(
    r"^-?\s*(?:R\$\s*)?\d{1,3}(?:\.\d{3})*,\d{2}$"
    r"|^-?\s*(?:R\$\s*)?\d+,\d{2}$"
)
ACCOUNT_RE = re.compile(r"^([\d]+)\s+([\d.]+)\s*-\s*(.+)$")
CONTRA_RE = re.compile(r"^(\d+)\s*-\s*([\d.].*)$")

# ── Portuguese month names (normalised, no accents) ──────────────────────────
PT_MONTHS: dict[str, str] = {
    "janeiro": "01", "fevereiro": "02", "marco": "03",
    "abril": "04", "maio": "05", "junho": "06", "julho": "07",
    "agosto": "08", "setembro": "09", "outubro": "10",
    "novembro": "11", "dezembro": "12",
}

HEADER_ALIASES = {
    "date": {"data", "dt"},
    "description": {"historico", "descricao", "descricao historico", "complemento", "detalhe"},
    "debit": {"debito", "valor debito"},
    "credit": {"credito", "valor credito"},
    "saldo": {"saldo", "saldo final"},
}


# ═══════════════════════════════════════════════════════════════════════════════
# Public entry point
# ═══════════════════════════════════════════════════════════════════════════════

def beautify_workbook(file_stream: BytesIO, input_extension: str = ".xlsx") -> BytesIO:
    if input_extension == ".pdf":
        return beautify_pdf(file_stream)

    output_workbook = Workbook()
    output_workbook.remove(output_workbook.active)

    # Accumulate rows per document type so sheets of the same type get merged.
    # Key: tuple of header names → (first_title, parsed_sheet_template, [rows])
    merged: dict[tuple, tuple[str, dict, list]] = {}
    order: list[tuple] = []  # insertion order to preserve sheet ordering

    for original_title, rows in read_input_sheets(file_stream, input_extension):
        parsed_sheet = extract_records(rows)
        if not parsed_sheet:
            continue

        key = tuple(parsed_sheet["headers"])
        if key not in merged:
            merged[key] = (original_title, parsed_sheet, list(parsed_sheet["rows"]))
            order.append(key)
        else:
            merged[key][2].extend(parsed_sheet["rows"])

    created_sheets = 0
    for key in order:
        title, template, all_rows = merged[key]
        if not all_rows:
            continue
        final_sheet = dict(template)
        final_sheet["rows"] = all_rows

        output_sheet = output_workbook.create_sheet(
            title=build_sheet_title(final_sheet["headers"], created_sheets)
        )
        write_records(output_sheet, final_sheet)
        style_output_sheet(output_sheet, final_sheet, created_sheets)
        created_sheets += 1

    if created_sheets == 0:
        raise ValueError(
            "Nao encontrei lancamentos no formato esperado. "
            "Se quiser, me envie um exemplo desse Excel para eu ajustar o parser."
        )

    output = BytesIO()
    output_workbook.save(output)
    output.seek(0)
    return output


def beautify_pdf(file_stream: BytesIO) -> BytesIO:
    parsed_documents = parse_pdf_documents(file_stream)
    if not parsed_documents:
        raise ValueError(
            "Nao consegui interpretar esse PDF ainda. "
            "Se ele for exportado do sistema contabil, me envie o arquivo que eu ajusto o layout."
        )

    output_workbook = Workbook()
    output_workbook.remove(output_workbook.active)

    for sheet_index, (title, parsed_sheet) in enumerate(parsed_documents):
        output_sheet = output_workbook.create_sheet(
            title=build_sheet_title(parsed_sheet["headers"], sheet_index)
        )
        write_records(output_sheet, parsed_sheet)
        style_output_sheet(output_sheet, parsed_sheet, sheet_index)

    output = BytesIO()
    output_workbook.save(output)
    output.seek(0)
    return output


# ═══════════════════════════════════════════════════════════════════════════════
# PDF parsing
# ═══════════════════════════════════════════════════════════════════════════════

def parse_pdf_documents(file_stream: BytesIO) -> list[tuple[str, dict]]:
    file_stream.seek(0)
    with pdfplumber.open(file_stream) as pdf:
        first_page_text = pdf.pages[0].extract_text() or ""
        norm = normalize_text(first_page_text)

        if "balancete" in norm:
            rows = parse_balancete_pdf(pdf)
            if rows:
                return [("BalancetePDF", _balancete_sheet(rows))]

        if "diario geral" in norm or "diario" in norm:
            rows = parse_diario_pdf(pdf)
            if rows:
                return [("DiarioPDF", _diario_sheet(rows))]

        if "razao contabil" in norm or "razao" in norm:
            rows = parse_razao_pdf(pdf)
            if rows:
                return [("RazaoPDF", _razao_sheet(rows))]

    return []


def _balancete_sheet(rows: list[dict]) -> dict:
    return {
        "headers": BALANCETE_HCN_HEADERS,
        "rows": rows,
        "date_columns": set(),
        "money_columns": {4, 5, 6, 7},
        "description_column": 3,
    }


def _diario_sheet(rows: list[dict]) -> dict:
    return {
        "headers": DIARIO_HCN_HEADERS,
        "rows": rows,
        "date_columns": set(),
        "money_columns": {6, 7},
        "description_column": 5,
    }


def _razao_sheet(rows: list[dict]) -> dict:
    return {
        "headers": RAZAO_HCN_HEADERS,
        "rows": rows,
        "date_columns": set(),
        "money_columns": {5, 7, 8, 9},
        "description_column": 6,
    }


# ── PDF page helpers ──────────────────────────────────────────────────────────

def parse_balancete_pdf(pdf) -> list[dict]:
    """Parse TOTVS Balancete PDF lines.

    Line format (no Red column):
      CONTA  DESCRICAO  SALDO_ANTERIOR[D|C]  DEBITO  CREDITO  SALDO_ATUAL[D|C]
    """
    rows: list[dict] = []
    _money = r"-?\d{1,3}(?:\.\d{3})*,\d{2}"
    pattern = re.compile(
        r"^(?P<conta>\d[\d.]*)\s+"
        r"(?P<descricao>.+?)\s+"
        r"(?P<saldo_anterior>" + _money + r")[DC]?\s+"
        r"(?P<debito>" + _money + r")\s+"
        r"(?P<credito>" + _money + r")\s+"
        r"(?P<saldo_atual>" + _money + r")[DC]?$"
    )
    for page in pdf.pages:
        text = page.extract_text() or ""
        for line in text.splitlines():
            m = pattern.match(normalize_spaces(line))
            if not m:
                continue
            rows.append({
                "REG": "0300",
                "ContaContabil": m.group("conta"),
                "NomeConta": m.group("descricao"),
                "SaldoAnterior": parse_money_value(m.group("saldo_anterior")),
                "Debito": parse_money_value(m.group("debito")),
                "Credito": parse_money_value(m.group("credito")),
                "SaldoAtual": parse_money_value(m.group("saldo_atual")),
                "Estoque": "N",
                "AtivoImob": "N",
                "DepreciacaoAcumulativa": "N",
                "ContaFinanceira": "N",
                "ReservaDeContingência": "N",
                "Centro de Custos": None,
            })
    return rows


def parse_diario_pdf(pdf) -> list[dict]:
    rows: list[dict] = []
    current: dict | None = None

    for page in pdf.pages:
        line_map = extract_pdf_lines(page)
        for _, words in line_map:
            if not words:
                continue
            ordered = sorted(words, key=lambda item: item[0])
            texts = [t for _, t in ordered]
            first_text = texts[0]
            line_text = " ".join(texts)
            norm_line = normalize_text(line_text)

            if (
                first_text == "Lote"
                or "diario" in norm_line
                or "pagina:" in norm_line
                or "periodo:" in norm_line
                or "uruacu" in norm_line
                or "hcn - hosp" in norm_line
                or "total do dia:" in norm_line
                or "total da empresa:" in norm_line
                or "total geral:" in norm_line
                or re.fullmatch(r"(?:-?\d{1,3}(?:\.\d{3})*,\d{2}\s*){1,3}", line_text.strip())
                or "19.324.171/0008-70" in line_text
                or "centro-norte goiano" in norm_line
            ):
                continue

            if re.fullmatch(r"\d+", first_text) and len(ordered) >= 2 and re.match(r"^\d{8}", texts[1]):
                if current is not None:
                    rows.extend(_finalise_diario_pdf(current))

                lote = first_text
                nr_mvto, conta_debito_codigo = split_mvto_and_account(texts[1])
                debit_tokens = [conta_debito_codigo] if conta_debito_codigo else []
                debit_tokens.extend(t for x, t in ordered if 100 <= x < 200)
                credit_tokens = [t for x, t in ordered if 200 <= x < 309]
                historico_words = [t for x, t in ordered if 309 <= x < 470]
                amount_values = [
                    (x, parse_money_value(t))
                    for x, t in ordered
                    if x >= 470 and parse_money_value(t) is not None
                ]

                debito = credito = None
                if len(amount_values) >= 2:
                    debito = amount_values[0][1]
                    credito = amount_values[-1][1]
                elif len(amount_values) == 1:
                    amount = amount_values[0][1]
                    if bool(debit_tokens) and not bool(credit_tokens):
                        debito = amount
                    elif bool(credit_tokens) and not bool(debit_tokens):
                        credito = amount
                    elif amount_values[0][0] >= 535:
                        credito = amount
                    else:
                        debito = amount

                current = {
                    "_lote": lote,
                    "_nr_mvto": nr_mvto,
                    "_conta_debito": " ".join(debit_tokens).strip(),
                    "_conta_credito": " ".join(credit_tokens).strip(),
                    "_historico": " ".join(historico_words).strip(),
                    "_debito": debito,
                    "_credito": credito,
                    "_debito_desc": [],
                    "_credito_desc": [],
                    "_historico_extra": [],
                }
                continue

            if current is None:
                continue
            if "emitido por:" in norm_line:
                continue

            for x, text in ordered:
                if 100 <= x < 200:
                    current["_debito_desc"].append(text)
                elif 200 <= x < 309:
                    current["_credito_desc"].append(text)
                elif 309 <= x < 470:
                    current["_historico_extra"].append(text)

    if current is not None:
        rows.extend(_finalise_diario_pdf(current))
    return rows


def _finalise_diario_pdf(c: dict) -> list[dict]:
    conta_debito = clean_diario_account(
        join_description(c["_conta_debito"], " ".join(c.pop("_debito_desc", [])))
    )
    conta_credito = clean_diario_account(
        join_description(c["_conta_credito"], " ".join(c.pop("_credito_desc", [])))
    )
    historico = join_description(
        c["_historico"], " ".join(c.pop("_historico_extra", []))
    ) or "Sem historico"

    rows: list[dict] = []
    if conta_debito:
        code, name = split_account_field(conta_debito)
        rows.append({
            "REG": 1600, "DATA": "", "CLASSIFICAÇÃO": code, "DESCRIÇÃO": name,
            "HISTÓRICO": historico, "DÉBITO": decimal_to_float(c["_debito"]),
            "CRÉDITO": 0.0 if c["_debito"] is not None else None, "Centro de Custos": None,
        })
    if conta_credito:
        code, name = split_account_field(conta_credito)
        rows.append({
            "REG": 1600, "DATA": "", "CLASSIFICAÇÃO": code, "DESCRIÇÃO": name,
            "HISTÓRICO": historico, "DÉBITO": None,
            "CRÉDITO": decimal_to_float(c["_credito"]), "Centro de Custos": None,
        })
    return rows


def parse_razao_pdf(pdf) -> list[dict]:
    rows: list[dict] = []
    current_account_name = ""
    current_account_code = ""
    current_saldo_anterior: Decimal | None = None
    current_date = ""
    pending_record: dict | None = None
    money_pattern = r"-?\d{1,3}(?:\.\d{3})*,\d{2}"

    for page in pdf.pages:
        text = page.extract_text() or ""
        for raw_line in text.splitlines():
            line = normalize_spaces(raw_line)
            if not line:
                continue
            if line.startswith("Conta Anal") and "tica:" in line:
                m = re.search(r"Conta Anal[íi]tica:\s*(.+?)\s+Saldo Anterior:", line)
                account_full = m.group(1).strip() if m else line.split(":", 1)[-1].strip()
                current_account_code, current_account_name = parse_razao_account_header(account_full)
                saldo_m = re.search(r"Saldo Anterior:\s*(" + money_pattern + r")", line)
                current_saldo_anterior = parse_money_value(saldo_m.group(1)) if saldo_m else None
                pending_record = None
                continue
            if line.startswith("Data ") or "Razão" in line or "Pagina:" in normalize_text(line):
                continue
            parsed_date = parse_date_value(line.split(" ")[0])
            if parsed_date:
                current_date = parsed_date
                values = re.findall(money_pattern, line)
                pending_record = {
                    "REG": 1700,
                    "NOME CONTA": current_account_name or "Sem conta",
                    "CONTA CONTÁBIL": current_account_code or "",
                    "DATA": current_date,
                    "SALDO ANTERIOR": decimal_to_float(current_saldo_anterior),
                    "HISTÓRICO": "Sem historico",
                    "DÉBITO": decimal_to_float(parse_money_value(values[0])) if len(values) > 0 else None,
                    "CRÉDITO": decimal_to_float(parse_money_value(values[1])) if len(values) > 1 else None,
                    "Saldo": decimal_to_float(parse_money_value(values[2])) if len(values) > 2 else None,
                    "Contra Partida": "",
                }
                rows.append(pending_record)
                continue
            if line.startswith("Contrapartida:"):
                if pending_record is not None:
                    contra = line.split(":", 1)[1].strip()
                    mm = re.search(rf"\s+{money_pattern}\s+{money_pattern}\s+", contra)
                    if mm:
                        contra = contra[: mm.start()].strip()
                    pending_record["Contra Partida"] = parse_razao_contra_code(contra)
                continue
            if pending_record is not None:
                extra = line.strip()
                if extra and normalize_text(pending_record["HISTÓRICO"]) in {"", "sem historico"}:
                    pending_record["HISTÓRICO"] = extra
    return rows


# ═══════════════════════════════════════════════════════════════════════════════
# XLS / XLSX extraction
# ═══════════════════════════════════════════════════════════════════════════════

def extract_records(rows: list[tuple]) -> dict | None:
    non_empty_rows = [list(row) for row in rows if not row_is_empty(row)]
    if not non_empty_rows:
        return None

    balancete = detect_balancete_layout(non_empty_rows)
    if balancete is not None:
        balancete_rows = parse_balancete_rows(non_empty_rows, balancete)
        if balancete_rows:
            return {
                "headers": BALANCETE_HCN_HEADERS,
                "rows": balancete_rows,
                "date_columns": set(),
                "money_columns": {4, 5, 6, 7},
                "description_column": 3,
            }

    diario = detect_diario_layout(non_empty_rows)
    if diario is not None:
        diario_rows = parse_diario_rows(non_empty_rows, diario)
        if diario_rows:
            return {
                "headers": DIARIO_HCN_HEADERS,
                "rows": diario_rows,
                "date_columns": set(),
                "money_columns": {6, 7},
                "description_column": 5,
            }

    razao = detect_razao_layout(non_empty_rows)
    if razao is not None:
        razao_rows = parse_razao_rows(non_empty_rows, razao)
        if razao_rows:
            return {
                "headers": RAZAO_HCN_HEADERS,
                "rows": razao_rows,
                "date_columns": set(),
                "money_columns": {5, 7, 8, 9},
                "description_column": 6,
            }

    structured = detect_structured_layout(non_empty_rows)
    if structured is not None:
        structured_records = parse_structured_rows(
            non_empty_rows, structured["header_index"], structured["columns"]
        )
        if structured_records:
            return {
                "headers": DEFAULT_OUTPUT_HEADERS,
                "rows": structured_records,
                "date_columns": {1},
                "money_columns": {3, 4, 5},
                "description_column": 2,
            }

    generic_records = parse_generic_rows(non_empty_rows)
    if generic_records:
        return {
            "headers": DEFAULT_OUTPUT_HEADERS,
            "rows": generic_records,
            "date_columns": {1},
            "money_columns": {3, 4, 5},
            "description_column": 2,
        }
    return None


def read_input_sheets(file_stream: BytesIO, input_extension: str) -> list[tuple[str, list[tuple]]]:
    file_stream.seek(0)

    if input_extension == ".pdf":
        return read_pdf_sheets(file_stream)

    if input_extension == ".xls":
        workbook = xlrd.open_workbook(file_contents=file_stream.getvalue())
        sheets: list[tuple[str, list[tuple]]] = []
        for sheet in workbook.sheets():
            rows = []
            for row_index in range(sheet.nrows):
                parsed_row = []
                for col_index in range(sheet.ncols):
                    parsed_row.append(
                        convert_xls_cell(
                            workbook,
                            sheet.cell_value(row_index, col_index),
                            sheet.cell_type(row_index, col_index),
                        )
                    )
                rows.append(tuple(parsed_row))
            sheets.append((sheet.name, rows))
        return sheets

    workbook = load_workbook(file_stream, data_only=True, keep_vba=input_extension == ".xlsm")
    return [
        (sheet.title, list(sheet.iter_rows(values_only=True)))
        for sheet in workbook.worksheets
    ]


def read_pdf_sheets(file_stream: BytesIO) -> list[tuple[str, list[tuple]]]:
    file_stream.seek(0)
    rows: list[tuple] = []
    with pdfplumber.open(file_stream) as pdf:
        for page in pdf.pages:
            page_rows = extract_rows_from_pdf_page(page)
            if page_rows:
                rows.extend(page_rows)
            else:
                rows.extend(extract_rows_from_pdf_text(page))
    return [("PDF_Convertido", rows)]


def extract_rows_from_pdf_page(page) -> list[tuple]:
    extracted: list[tuple] = []
    for table in page.extract_tables():
        for row in table:
            if not row:
                continue
            cleaned = tuple(clean_pdf_cell(cell) for cell in row)
            if any(cleaned):
                extracted.append(cleaned)
    return extracted


def extract_rows_from_pdf_text(page) -> list[tuple]:
    text = page.extract_text() or ""
    rows: list[tuple] = []
    for line in text.splitlines():
        parts = [s.strip() for s in re.split(r"\s{2,}", line) if s.strip()]
        if parts:
            rows.append(tuple(parts))
    return rows


def clean_pdf_cell(value: object) -> str:
    return normalize_spaces(value)


# ═══════════════════════════════════════════════════════════════════════════════
# Layout detection
# ═══════════════════════════════════════════════════════════════════════════════

def detect_balancete_layout(rows: list[list]) -> dict | None:
    for index in range(min(len(rows), 10)):
        norm_row = [normalize_text(v) for v in rows[index]]
        columns = {
            "conta": find_header_index(norm_row, {"conta"}),
            "descricao": find_header_index(norm_row, {"descricao", "nome"}),
            "saldo_anterior": find_header_index(norm_row, {"saldo anterior"}),
            "debito": find_header_index(norm_row, {"valor debito", "debito"}),
            "credito": find_header_index(norm_row, {"valor credito", "credito"}),
            "saldo_atual": find_header_index(norm_row, {"saldo atual"}),
        }
        required = {"conta", "descricao", "saldo_anterior", "debito", "credito", "saldo_atual"}
        if all(columns[k] is not None for k in required):
            columns["header_index"] = index
            columns["red"] = None  # optional – not required
            return columns
    return None


def detect_diario_layout(rows: list[list]) -> dict | None:
    for index in range(min(len(rows), 15)):
        norm_row = [normalize_text(v) for v in rows[index]]
        columns = {
            "lote": find_header_index(norm_row, {"lote"}),
            "nr_mvto": find_header_index(norm_row, {"nr. mvto", "nr mvto"}),
            "conta_debito": find_header_index(norm_row, {"cont. debito", "cont debito"}),
            "conta_credito": find_header_index(norm_row, {"cont. credito", "cont credito"}),
            "historico": find_header_index(norm_row, {"historico", "historico"}),
            "debito": find_header_index(norm_row, {"valor debito", "debito"}),
            "credito": find_header_index(norm_row, {"valor credito", "credito"}),
        }
        if all(v is not None for v in columns.values()):
            columns["header_index"] = index
            return columns
    return None


def detect_razao_layout(rows: list[list]) -> dict | None:
    for index in range(min(len(rows), 20)):
        norm_row = [normalize_text(v) for v in rows[index]]
        if "conta analitica:" in norm_row or "conta analitica" in norm_row:
            return {"account_header_index": index}
    return None


def detect_structured_layout(rows: list[list]) -> dict | None:
    for index in range(min(len(rows), 20)):
        row = rows[index]
        columns: dict[str, int] = {}
        for cell_index, value in enumerate(row):
            norm = normalize_text(value)
            if not norm:
                continue
            for key, aliases in HEADER_ALIASES.items():
                if norm in aliases and key not in columns:
                    columns[key] = cell_index
        if {"date", "description"}.issubset(columns) and columns.keys() & {"debit", "credit", "saldo"}:
            return {"header_index": index, "columns": columns}
    return None


# ═══════════════════════════════════════════════════════════════════════════════
# Balancete parsing
# ═══════════════════════════════════════════════════════════════════════════════

def parse_balancete_rows(rows: list[list], columns: dict) -> list[dict]:
    parsed: list[dict] = []

    for row in rows[columns["header_index"] + 1:]:
        conta = normalize_spaces(get_value(row, columns["conta"]))
        # Auxiliary (per-vendor) detail rows ("Conta Aux.:") must not appear in output.
        # Also skip any TOTVS noise labels (Total:, Subtotal:, etc.).
        conta_norm = normalize_text(conta)
        if conta_norm.startswith("conta aux") or conta_norm.startswith("total") or conta_norm.startswith("subtotal"):
            continue

        descricao = normalize_spaces(get_value(row, columns["descricao"]))
        sa, deb, cred, sat = _extract_balancete_money(row)

        if normalize_text(conta) == "conta" or normalize_text(descricao) == "descricao":
            continue
        if not any([conta, descricao, sa, deb, cred, sat]):
            continue

        parsed.append({
            "REG": "0300",
            "ContaContabil": conta,
            "NomeConta": descricao or "Sem descricao",
            "SaldoAnterior": sa,
            "Debito": deb,
            "Credito": cred,
            "SaldoAtual": sat,
            "Estoque": "N", "AtivoImob": "N",
            "DepreciacaoAcumulativa": "N", "ContaFinanceira": "N",
            "ReservaDeContingência": "N", "Centro de Custos": None,
        })

    return parsed


def _extract_balancete_money(
    row: list,
) -> tuple[Decimal | None, Decimal | None, Decimal | None, Decimal | None]:
    money_values = [parse_money_value(v) for v in row if parse_money_value(v) is not None]
    if len(money_values) >= 4:
        return money_values[-4], money_values[-3], money_values[-2], money_values[-1]
    if len(money_values) == 3:
        return None, money_values[0], money_values[1], money_values[2]
    if len(money_values) == 2:
        return None, money_values[0], money_values[1], None
    if len(money_values) == 1:
        return None, None, None, money_values[0]
    return None, None, None, None


# ═══════════════════════════════════════════════════════════════════════════════
# Diário parsing
# ═══════════════════════════════════════════════════════════════════════════════

def parse_diario_rows(rows: list[list], columns: dict) -> list[dict]:
    """
    Process ALL non-empty rows (including date headers and repeated column headers).
    Each XLS sub-row (debit side OR credit side) becomes one HCN output row.
    Combined rows (both sides in one row) are split into two output rows.
    """
    parsed: list[dict] = []
    current_date = ""

    col_lote = columns["lote"]
    col_debito_conta = columns["conta_debito"]
    col_credito_conta = columns["conta_credito"]
    col_historico = columns["historico"]
    col_debito_val = columns["debito"]
    col_credito_val = columns["credito"]

    for row in rows:
        texts = [normalize_spaces(v) for v in row]

        # ── Date header row: "01 de Março de 2026" in one of the cells ──
        for cell_val in texts:
            pt_date = parse_pt_date_header(cell_val)
            if pt_date:
                current_date = pt_date
                break
        else:
            # Not a date header – process as data row
            lote_text = _safe_col(texts, col_lote)

            # Skip column-header repetition rows
            if normalize_text(lote_text) == "lote":
                continue

            conta_debito = _safe_col(texts, col_debito_conta)
            conta_credito = _safe_col(texts, col_credito_conta)
            historico = _safe_col(texts, col_historico) or "Sem historico"
            debito_val = parse_money_value(_safe_col(texts, col_debito_val))
            credito_val = parse_money_value(_safe_col(texts, col_credito_val))

            if not conta_debito and not conta_credito:
                continue

            if conta_debito:
                code, name = split_account_field(conta_debito)
                parsed.append({
                    "REG": 1600,
                    "DATA": current_date,
                    "CLASSIFICAÇÃO": code,
                    "DESCRIÇÃO": name,
                    "HISTÓRICO": historico,
                    "DÉBITO": decimal_to_float(debito_val),
                    "CRÉDITO": 0.0,  # reference: CRÉDITO is always 0 (never null) on debit rows
                    "Centro de Custos": None,
                })

            if conta_credito:
                code, name = split_account_field(conta_credito)
                parsed.append({
                    "REG": 1600,
                    "DATA": current_date,
                    "CLASSIFICAÇÃO": code,
                    "DESCRIÇÃO": name,
                    "HISTÓRICO": historico,
                    "DÉBITO": None,  # reference: DÉBITO is always null on credit rows
                    "CRÉDITO": decimal_to_float(credito_val),
                    "Centro de Custos": None,
                })

    return parsed


# ═══════════════════════════════════════════════════════════════════════════════
# Razão parsing
# ═══════════════════════════════════════════════════════════════════════════════

def parse_razao_rows(rows: list[list], layout: dict) -> list[dict]:
    parsed: list[dict] = []
    current_account_name = ""
    current_account_code = ""
    current_saldo_anterior: Decimal | None = None
    current_date = ""
    pending_record: dict | None = None
    expecting_contra = False

    for row in rows[layout["account_header_index"]:]:
        texts = [normalize_spaces(v) for v in row]
        non_empty_vals = [v for v in texts if v]

        if not non_empty_vals:
            continue

        col1 = _safe_col(texts, 1)
        col3 = _safe_col(texts, 3)
        col5 = _safe_col(texts, 5)
        col8 = _safe_col(texts, 8)
        col11 = _safe_col(texts, 11)
        col16 = _safe_col(texts, 16)
        col20 = _safe_col(texts, 20)
        col24 = _safe_col(texts, 24)

        # ── (1) Account header row (always checked first) ──────────────────
        if col1 in {"Conta Analitica:", "Conta Analítica:"}:
            expecting_contra = False
            current_account_code, current_account_name = parse_razao_account_header(col5)
            current_saldo_anterior = parse_money_value(col24)
            pending_record = None
            continue

        # ── (2) Account code row after Contrapartida: ──────────────────────
        if expecting_contra:
            # If this row has a parseable date or is a Contrapartida label itself,
            # it is NOT the account-code row – reset the flag and fall through.
            if parse_date_value(col1) is not None or col3 in {"Contrapartida:", "Contrapartida"}:
                expecting_contra = False
                # fall through to date/transaction handling below
            else:
                if col5 and pending_record is not None:
                    pending_record["Contra Partida"] = parse_razao_contra_code(col5)
                expecting_contra = False
                continue

        # ── (3) Skip "Data" column header ─────────────────────────────────
        if col1 == "Data" or (normalize_text(col1) == "data" and parse_date_value(col1) is None):
            continue

        # ── (4) Contrapartida label row ────────────────────────────────────
        if col3 in {"Contrapartida:", "Contrapartida"}:
            expecting_contra = True
            continue

        # ── (5) Transaction row (has date or carries current date) ─────────
        row_date = parse_date_value(col1)
        if row_date:
            current_date = row_date

        debit = parse_money_value(col8)
        credit = parse_money_value(col11)
        saldo = parse_money_value(col16)

        if debit is not None or credit is not None or saldo is not None:
            pending_record = {
                "REG": 1700,
                "NOME CONTA": current_account_name or "Sem conta",
                "CONTA CONTÁBIL": current_account_code or "",
                "DATA": current_date,
                "SALDO ANTERIOR": decimal_to_float(current_saldo_anterior),
                "HISTÓRICO": col20 or "Sem historico",
                "DÉBITO": decimal_to_float(debit),
                "CRÉDITO": decimal_to_float(credit),
                "Saldo": decimal_to_float(saldo),
                "Contra Partida": "",
            }
            parsed.append(pending_record)

    return parsed


# ═══════════════════════════════════════════════════════════════════════════════
# Generic / structured fallback parsers (unchanged)
# ═══════════════════════════════════════════════════════════════════════════════

def parse_structured_rows(
    rows: list[list], header_index: int, columns: dict
) -> list[dict]:
    records: list[dict] = []
    current: dict | None = None
    ignored = set(columns.values())

    for row in rows[header_index + 1:]:
        parsed_date = parse_date_value(get_value(row, columns.get("date")))
        description = normalize_spaces(get_value(row, columns.get("description")))
        debit = parse_money_value(get_value(row, columns.get("debit")))
        credit = parse_money_value(get_value(row, columns.get("credit")))
        saldo = parse_money_value(get_value(row, columns.get("saldo")))

        if parsed_date:
            current = {
                "Data": parsed_date, "Descricao": description,
                "Debito": debit, "Credito": credit, "Saldo": saldo,
            }
            append_extra_description(current, row, ignored)
            records.append(current)
            continue

        if current is None:
            continue

        continued = description or collect_text_fragments(row, ignored)
        if continued:
            current["Descricao"] = join_description(current["Descricao"], continued)

        if current["Debito"] is None and debit is not None:
            current["Debito"] = debit
        if current["Credito"] is None and credit is not None:
            current["Credito"] = credit
        if current["Saldo"] is None and saldo is not None:
            current["Saldo"] = saldo

    return finalize_records(records)


def parse_generic_rows(rows: list[list]) -> list[dict]:
    records: list[dict] = []
    current: dict | None = None

    for row in rows:
        parsed_date = first_date_in_row(row)
        money_cells = extract_money_cells(row)
        text_parts = extract_text_parts(row)

        if parsed_date:
            debit, credit, saldo = distribute_amounts(money_cells)
            current = {
                "Data": parsed_date, "Descricao": " ".join(text_parts),
                "Debito": debit, "Credito": credit, "Saldo": saldo,
            }
            records.append(current)
            continue

        if current is None:
            continue

        continuation = " ".join(text_parts)
        if continuation:
            current["Descricao"] = join_description(current["Descricao"], continuation)

        debit, credit, saldo = distribute_amounts(money_cells)
        if current["Debito"] is None and debit is not None:
            current["Debito"] = debit
        if current["Credito"] is None and credit is not None:
            current["Credito"] = credit
        if current["Saldo"] is None and saldo is not None:
            current["Saldo"] = saldo

    return finalize_records(records)


def finalize_records(records: list[dict]) -> list[dict]:
    cleaned = []
    for r in records:
        if not r["Data"]:
            continue
        r["Descricao"] = normalize_spaces(r["Descricao"]) or "Sem descricao"
        cleaned.append(r)
    return cleaned


# ═══════════════════════════════════════════════════════════════════════════════
# Output writing & styling
# ═══════════════════════════════════════════════════════════════════════════════

def write_records(sheet, parsed_sheet: dict) -> None:
    headers = parsed_sheet["headers"]
    sheet.append(headers)
    for record in parsed_sheet["rows"]:
        row = []
        for header in headers:
            value = record.get(header)
            row.append(decimal_to_float(value) if isinstance(value, Decimal) else value)
        sheet.append(row)


def style_output_sheet(sheet, parsed_sheet: dict, sheet_index: int) -> None:
    headers = parsed_sheet["headers"]
    last_row = sheet.max_row
    last_col = sheet.max_column
    money_columns = parsed_sheet["money_columns"]
    desc_col = parsed_sheet["description_column"]

    is_hcn = headers in (BALANCETE_HCN_HEADERS, DIARIO_HCN_HEADERS, RAZAO_HCN_HEADERS)

    if is_hcn:
        # HCN: céluals planas, sem tabela, sem freeze — replicar estrutura do modelo de referência
        if headers != DIARIO_HCN_HEADERS:
            sheet.auto_filter.ref = sheet.dimensions

        for cell in sheet[1]:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row in range(2, last_row + 1):
            for col in range(1, last_col + 1):
                cell = sheet.cell(row=row, column=col)
                if col in money_columns:
                    cell.number_format = ACCOUNTING_FORMAT
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col == desc_col:
                    cell.alignment = Alignment(wrap_text=True, vertical="center")

    else:
        # Genérico: estilo decorado com tabela, freeze e preenchimento alternado
        sheet.sheet_view.showGridLines = False
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

        date_columns = parsed_sheet["date_columns"]

        for cell in sheet[1]:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER

        for row in range(2, last_row + 1):
            for col in range(1, last_col + 1):
                cell = sheet.cell(row=row, column=col)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(vertical="center", wrap_text=(col == desc_col))

                if row % 2 == 0:
                    cell.fill = ALT_ROW_FILL

                if col in date_columns:
                    cell.number_format = "dd/mm/yyyy"
                elif col in money_columns:
                    cell.number_format = ACCOUNTING_FORMAT
                    cell.alignment = Alignment(horizontal="right", vertical="center")

        highlight_total_rows(sheet, last_row, last_col, desc_col)
        create_table(sheet, last_row, last_col, sheet_index)

    adjust_column_widths(sheet, parsed_sheet)


def highlight_total_rows(sheet, last_row: int, last_col: int, desc_col: int) -> None:
    for row in range(2, last_row + 1):
        desc = sheet.cell(row=row, column=desc_col).value
        if isinstance(desc, str) and any(
            kw in desc.lower() for kw in ("total", "subtotal", "saldo anterior", "resumo")
        ):
            for col in range(1, last_col + 1):
                cell = sheet.cell(row=row, column=col)
                cell.font = Font(bold=True, color="1F1F1F")
                cell.fill = PatternFill("solid", fgColor="D9EAD3")
                cell.border = TOTAL_BORDER


def adjust_column_widths(sheet, parsed_sheet: dict) -> None:
    headers = parsed_sheet["headers"]
    if headers == BALANCETE_HCN_HEADERS:
        widths = {1: 8, 2: 22, 3: 54, 4: 18, 5: 18, 6: 18, 7: 18, 8: 10, 9: 12, 10: 22, 11: 16, 12: 22, 13: 16}
    elif headers == DIARIO_HCN_HEADERS:
        widths = {1: 9, 2: 11, 3: 22, 4: 57, 5: 96, 6: 25, 7: 18, 8: 22}
    elif headers == RAZAO_HCN_HEADERS:
        widths = {1: 8, 2: 44, 3: 22, 4: 16, 5: 18, 6: 56, 7: 18, 8: 18, 9: 18, 10: 22}
    else:
        widths = {1: 14, 2: 60, 3: 16, 4: 16, 5: 16}
    for col_idx, width in widths.items():
        sheet.column_dimensions[get_column_letter(col_idx)].width = width


def create_table(sheet, last_row: int, last_col: int, sheet_index: int) -> None:
    if last_row < 2:
        return
    table_range = f"A1:{get_column_letter(last_col)}{last_row}"
    safe_title = re.sub(r"[^A-Za-z0-9_]", "", sheet.title)[:18] or "Planilha"
    table = Table(
        displayName=f"Tabela_{sheet_index + 1}_{safe_title}",
        ref=table_range,
    )
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    sheet.add_table(table)


def build_sheet_title(headers: list, index: int) -> str:
    if headers == BALANCETE_HCN_HEADERS:
        return "0300"
    if headers == DIARIO_HCN_HEADERS:
        return "1600"
    if headers == RAZAO_HCN_HEADERS:
        return "1700"
    return f"Planilha{index + 1}"


# ═══════════════════════════════════════════════════════════════════════════════
# Helper / utility functions
# ═══════════════════════════════════════════════════════════════════════════════

def parse_pt_date_header(text: str) -> str | None:
    """Parse 'DD de Mês de AAAA' (from TOTVS date headers) → 'DD/MM'."""
    if not text:
        return None
    norm = normalize_text(text)  # removes accents, lowercases
    m = re.match(r"^(\d{1,2})\s+de\s+(\w+)\s+de\s+\d{4}$", norm)
    if m:
        day = m.group(1).zfill(2)
        month = PT_MONTHS.get(m.group(2))
        if month:
            return f"{day}/{month}"
    return None


def split_account_field(text: str) -> tuple[str, str]:
    """Split 'CODE - NAME' into (code, name). Handles missing dash gracefully."""
    stripped = text.strip()
    m = re.match(r"^([\d.]+)\s*-\s*(.+)$", stripped)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return stripped, stripped


def parse_razao_account_header(text: str) -> tuple[str, str]:
    """Parse 'SEQNO CODE - NAME' → (code, name). E.g. '92576 1.1.1.02.02.010 - BCO CEF...'"""
    if not text:
        return "", ""
    m = ACCOUNT_RE.match(text.strip())
    if m:
        return m.group(2).strip(), m.group(3).strip()
    # Fallback: try CODE - NAME without sequence
    m2 = re.match(r"^([\d.]+)\s*-\s*(.+)$", text.strip())
    if m2:
        return m2.group(1).strip(), m2.group(2).strip()
    return "", text.strip()


def parse_razao_contra_code(text: str) -> str:
    """Parse 'SEQNO - CODE' → 'CODE'. E.g. '10028 - 2.1.1.02.01.001' → '2.1.1.02.01.001'"""
    if not text:
        return ""
    m = CONTRA_RE.match(text.strip())
    if m:
        return m.group(2).strip()
    # Maybe already just a code
    return text.strip()


def _safe_col(texts: list[str], index: int) -> str:
    return texts[index] if index < len(texts) else ""


def row_is_empty(row: tuple | list) -> bool:
    return all(normalize_spaces(v) == "" for v in row)


def normalize_text(value: object) -> str:
    text = normalize_spaces(value)
    if not text:
        return ""
    text = (
        unicodedata.normalize("NFKD", text)
        .encode("ascii", "ignore")
        .decode("ascii")
        .lower()
    )
    return re.sub(r"\s+", " ", text).strip()


def normalize_spaces(value: object) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", text).strip()


def parse_date_value(value: object) -> str | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")

    text = normalize_spaces(value)
    if not text:
        return None
    text = text.split(" ")[0]

    if DATE_RE.match(text):
        try:
            return datetime.strptime(text, "%d/%m/%Y").strftime("%d/%m/%Y")
        except ValueError:
            return None

    for fmt in ("%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue

    return None


_US_NUMBER_RE = re.compile(r"^-?\d+(?:\.\d+)?$")


def parse_money_value(value: object) -> Decimal | None:
    if value is None or value == "":
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))

    text = normalize_spaces(value)
    if not text:
        return None

    # XLS/XLSX numeric cells converted to string via normalize_spaces() arrive in
    # US dot-decimal format ("1234.56", "1200.0").  Handle these before the BR regex.
    if _US_NUMBER_RE.match(text):
        try:
            return Decimal(text)
        except InvalidOperation:
            pass

    normalized = text.replace("R$", "").replace(" ", "")
    if not MONEY_RE.match(normalized):
        return None

    try:
        return Decimal(normalized.replace(".", "").replace(",", "."))
    except InvalidOperation:
        return None


def decimal_to_float(value: object) -> float | None:
    if value is None:
        return None
    return float(value)


def convert_xls_cell(workbook, value: object, cell_type: int) -> object:
    if cell_type == xlrd.XL_CELL_DATE:
        return xlrd.xldate.xldate_as_datetime(value, workbook.datemode)
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def find_header_index(norm_row: list[str], candidates: set[str]) -> int | None:
    for i, v in enumerate(norm_row):
        if v in candidates:
            return i
    return None


def get_value(row: list, index: int | None) -> object:
    if index is None or index >= len(row):
        return None
    return row[index]


def first_date_in_row(row: list) -> str | None:
    for v in row:
        parsed = parse_date_value(v)
        if parsed:
            return parsed
    return None


def extract_money_cells(row: list) -> list[Decimal]:
    non_empty_idx = [i for i, v in enumerate(row) if normalize_spaces(v) != ""]
    candidate_idx = set(non_empty_idx[-3:])
    values: list[Decimal] = []
    for i, v in enumerate(row):
        if i not in candidate_idx and not isinstance(v, str):
            continue
        parsed = parse_money_value(v)
        if parsed is not None:
            values.append(parsed)
    return values


def extract_text_parts(row: list) -> list[str]:
    parts: list[str] = []
    for v in row:
        if parse_date_value(v) or parse_money_value(v) is not None:
            continue
        text = normalize_spaces(v)
        if text:
            parts.append(text)
    return parts


def distribute_amounts(
    values: list[Decimal],
) -> tuple[Decimal | None, Decimal | None, Decimal | None]:
    if not values:
        return None, None, None
    if len(values) == 1:
        return None, None, values[0]
    if len(values) == 2:
        return values[0], None, values[1]
    return values[-3], values[-2], values[-1]


def collect_text_fragments(row: list, ignored: set[int]) -> str:
    parts: list[str] = []
    for i, v in enumerate(row):
        if i in ignored:
            continue
        if parse_money_value(v) is not None or parse_date_value(v):
            continue
        text = normalize_spaces(v)
        if text:
            parts.append(text)
    return " ".join(parts)


def append_extra_description(current: dict, row: list, ignored: set[int]) -> None:
    extra = collect_text_fragments(row, ignored)
    if extra:
        current["Descricao"] = join_description(current["Descricao"], extra)


def join_description(current: object, extra: str) -> str:
    base = normalize_spaces(current)
    if not base:
        return extra
    return f"{base} {extra}".strip()


# ── PDF-specific helpers ──────────────────────────────────────────────────────

def extract_pdf_lines(page) -> list[tuple[float, list[tuple[float, str]]]]:
    words = page.extract_words(use_text_flow=True)
    grouped: dict[float, list[tuple[float, str]]] = {}
    for word in words:
        key = round(word["top"], 1)
        grouped.setdefault(key, []).append((round(word["x0"], 1), word["text"]))
    return sorted(grouped.items(), key=lambda item: item[0])


def split_mvto_and_account(value: str) -> tuple[str, str]:
    m = re.match(r"^(\d{8})(.+)$", value)
    if m:
        return m.group(1), m.group(2)
    return value[:8], value[8:]


def clean_diario_account(value: str) -> str:
    cleaned = normalize_spaces(value)
    cleaned = re.sub(r"\s*-\s*", " - ", cleaned)
    return normalize_spaces(cleaned)


