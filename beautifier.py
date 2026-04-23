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
import xlrd


DEFAULT_OUTPUT_HEADERS = ["Data", "Descricao", "Debito", "Credito", "Saldo"]
BALANCETE_HEADERS = [
    "Red",
    "Conta",
    "Descricao",
    "Saldo Anterior",
    "Debito",
    "Credito",
    "Saldo Atual",
]
DIARIO_HEADERS = [
    "Lote",
    "Nr Mvto",
    "Conta Debito",
    "Conta Credito",
    "Historico",
    "Debito",
    "Credito",
]
RAZAO_HEADERS = [
    "Conta Analitica",
    "Data",
    "Historico",
    "Contrapartida",
    "Debito",
    "Credito",
    "Saldo",
]
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
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
MONEY_RE = re.compile(r"^-?\s*(?:R\$\s*)?\d{1,3}(?:\.\d{3})*,\d{2}$|^-?\s*(?:R\$\s*)?\d+,\d{2}$")
HEADER_ALIASES = {
    "date": {"data", "dt"},
    "description": {"historico", "descricao", "descricao historico", "complemento", "detalhe"},
    "debit": {"debito", "valor debito"},
    "credit": {"credito", "valor credito"},
    "saldo": {"saldo", "saldo final"},
}


def beautify_workbook(file_stream: BytesIO, input_extension: str = ".xlsx") -> BytesIO:
    output_workbook = Workbook()
    output_workbook.remove(output_workbook.active)

    created_sheets = 0
    for original_title, rows in read_input_sheets(file_stream, input_extension):
        parsed_sheet = extract_records(rows)
        if not parsed_sheet:
            continue

        output_sheet = output_workbook.create_sheet(title=build_sheet_title(original_title, created_sheets))
        write_records(output_sheet, parsed_sheet)
        style_output_sheet(output_sheet, parsed_sheet, created_sheets)
        created_sheets += 1

    if created_sheets == 0:
        raise ValueError(
            "Nao encontrei lancamentos no formato esperado. Se quiser, me envie um exemplo desse Excel para eu ajustar o parser."
        )

    output = BytesIO()
    output_workbook.save(output)
    output.seek(0)
    return output


def extract_records(rows: list[tuple]) -> dict[str, object] | None:
    non_empty_rows = [list(row) for row in rows if not row_is_empty(row)]
    if not non_empty_rows:
        return None

    balancete = detect_balancete_layout(non_empty_rows)
    if balancete is not None:
        balancete_rows = parse_balancete_rows(non_empty_rows, balancete)
        if balancete_rows:
            return {
                "headers": BALANCETE_HEADERS,
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
                "headers": DIARIO_HEADERS,
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
                "headers": RAZAO_HEADERS,
                "rows": razao_rows,
                "date_columns": {2},
                "money_columns": {5, 6, 7},
                "description_column": 3,
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

    if input_extension == ".xls":
        workbook = xlrd.open_workbook(file_contents=file_stream.getvalue())
        sheets: list[tuple[str, list[tuple]]] = []
        for sheet in workbook.sheets():
            rows = []
            for row_index in range(sheet.nrows):
                parsed_row = []
                for column_index in range(sheet.ncols):
                    parsed_row.append(
                        convert_xls_cell(
                            workbook,
                            sheet.cell_value(row_index, column_index),
                            sheet.cell_type(row_index, column_index),
                        )
                    )
                rows.append(tuple(parsed_row))
            sheets.append((sheet.name, rows))
        return sheets

    workbook = load_workbook(file_stream, data_only=True, keep_vba=input_extension == ".xlsm")
    return [(sheet.title, list(sheet.iter_rows(values_only=True))) for sheet in workbook.worksheets]


def convert_xls_cell(workbook, value: object, cell_type: int) -> object:
    if cell_type == xlrd.XL_CELL_DATE:
        return xlrd.xldate.xldate_as_datetime(value, workbook.datemode)
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def detect_structured_layout(rows: list[list[object]]) -> dict[str, object] | None:
    for index in range(min(len(rows), 20)):
        row = rows[index]
        columns: dict[str, int] = {}

        for cell_index, value in enumerate(row):
            normalized = normalize_text(value)
            if not normalized:
                continue

            for key, aliases in HEADER_ALIASES.items():
                if normalized in aliases and key not in columns:
                    columns[key] = cell_index

        if {"date", "description"}.issubset(columns) and columns.keys() & {"debit", "credit", "saldo"}:
            return {"header_index": index, "columns": columns}

    return None


def parse_structured_rows(
    rows: list[list[object]], header_index: int, columns: dict[str, int]
) -> list[dict[str, object]]:
    records: list[dict[str, object]] = []
    current: dict[str, object] | None = None
    ignored_columns = set(columns.values())

    for row in rows[header_index + 1 :]:
        parsed_date = parse_date_value(get_value(row, columns.get("date")))
        description = normalize_spaces(get_value(row, columns.get("description")))
        debit = parse_money_value(get_value(row, columns.get("debit")))
        credit = parse_money_value(get_value(row, columns.get("credit")))
        saldo = parse_money_value(get_value(row, columns.get("saldo")))

        if parsed_date:
            current = {
                "Data": parsed_date,
                "Descricao": description,
                "Debito": debit,
                "Credito": credit,
                "Saldo": saldo,
            }
            append_extra_description(current, row, ignored_columns)
            records.append(current)
            continue

        if current is None:
            continue

        continued_description = description or collect_text_fragments(row, ignored_columns)
        if continued_description:
            current["Descricao"] = join_description(current["Descricao"], continued_description)

        if current["Debito"] is None and debit is not None:
            current["Debito"] = debit
        if current["Credito"] is None and credit is not None:
            current["Credito"] = credit
        if current["Saldo"] is None and saldo is not None:
            current["Saldo"] = saldo

    return finalize_records(records)


def parse_generic_rows(rows: list[list[object]]) -> list[dict[str, object]]:
    records: list[dict[str, object]] = []
    current: dict[str, object] | None = None

    for row in rows:
        parsed_date = first_date_in_row(row)
        money_cells = extract_money_cells(row)
        text_parts = extract_text_parts(row)

        if parsed_date:
            debit, credit, saldo = distribute_amounts(money_cells)
            current = {
                "Data": parsed_date,
                "Descricao": " ".join(text_parts),
                "Debito": debit,
                "Credito": credit,
                "Saldo": saldo,
            }
            records.append(current)
            continue

        if current is None:
            continue

        continuation_text = " ".join(text_parts)
        if continuation_text:
            current["Descricao"] = join_description(current["Descricao"], continuation_text)

        debit, credit, saldo = distribute_amounts(money_cells)
        if current["Debito"] is None and debit is not None:
            current["Debito"] = debit
        if current["Credito"] is None and credit is not None:
            current["Credito"] = credit
        if current["Saldo"] is None and saldo is not None:
            current["Saldo"] = saldo

    return finalize_records(records)


def finalize_records(records: list[dict[str, object]]) -> list[dict[str, object]]:
    cleaned: list[dict[str, object]] = []
    for record in records:
        if not record["Data"]:
            continue

        record["Descricao"] = normalize_spaces(record["Descricao"]) or "Sem descricao"
        cleaned.append(record)

    return cleaned


def write_records(sheet, parsed_sheet: dict[str, object]) -> None:
    headers = parsed_sheet["headers"]
    sheet.append(headers)
    for record in parsed_sheet["rows"]:
        row = []
        for header in headers:
            value = record.get(header)
            row.append(decimal_to_float(value) if isinstance(value, Decimal) else value)
        sheet.append(row)


def style_output_sheet(sheet, parsed_sheet: dict[str, object], sheet_index: int) -> None:
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    last_row = sheet.max_row
    last_column = sheet.max_column
    date_columns = parsed_sheet["date_columns"]
    money_columns = parsed_sheet["money_columns"]
    description_column = parsed_sheet["description_column"]

    for cell in sheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER

    for row in range(2, last_row + 1):
        for column in range(1, last_column + 1):
            cell = sheet.cell(row=row, column=column)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=column == description_column)

            if row % 2 == 0:
                cell.fill = ALT_ROW_FILL

            if column in date_columns:
                cell.number_format = "dd/mm/yyyy"
            elif column in money_columns:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")

    highlight_total_rows(sheet, last_row, last_column, description_column)
    adjust_column_widths(sheet, parsed_sheet)
    create_table(sheet, last_row, last_column, sheet_index)


def highlight_total_rows(sheet, last_row: int, last_column: int, description_column: int) -> None:
    for row in range(2, last_row + 1):
        description = sheet.cell(row=row, column=description_column).value
        if isinstance(description, str) and any(
            keyword in description.lower() for keyword in ("total", "subtotal", "saldo anterior", "resumo")
        ):
            for column in range(1, last_column + 1):
                cell = sheet.cell(row=row, column=column)
                cell.font = Font(bold=True, color="1F1F1F")
                cell.fill = PatternFill("solid", fgColor="D9EAD3")
                cell.border = TOTAL_BORDER


def adjust_column_widths(sheet, parsed_sheet: dict[str, object]) -> None:
    if parsed_sheet["headers"] == BALANCETE_HEADERS:
        widths = {1: 12, 2: 22, 3: 54, 4: 18, 5: 18, 6: 18, 7: 18}
    elif parsed_sheet["headers"] == DIARIO_HEADERS:
        widths = {1: 12, 2: 14, 3: 36, 4: 36, 5: 56, 6: 18, 7: 18}
    elif parsed_sheet["headers"] == RAZAO_HEADERS:
        widths = {1: 42, 2: 14, 3: 56, 4: 28, 5: 18, 6: 18, 7: 18}
    else:
        widths = {1: 14, 2: 60, 3: 16, 4: 16, 5: 16}
    for column_idx, width in widths.items():
        sheet.column_dimensions[get_column_letter(column_idx)].width = width


def create_table(sheet, last_row: int, last_column: int, sheet_index: int) -> None:
    if last_row < 2:
        return

    table_range = f"A1:{get_column_letter(last_column)}{last_row}"
    safe_title = re.sub(r"[^A-Za-z0-9_]", "", sheet.title)[:18] or "Planilha"
    table = Table(displayName=f"Tabela_{sheet_index + 1}_{safe_title}", ref=table_range)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(table)


def build_sheet_title(original_title: str, index: int) -> str:
    base = re.sub(r"[^A-Za-z0-9]", "", original_title)[:20] or f"Planilha{index + 1}"
    return f"Organizado{base}"[:31]


def row_is_empty(row: tuple | list) -> bool:
    return all(normalize_spaces(value) == "" for value in row)


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

    normalized_text = text.replace("R$", "").replace(" ", "")
    if not MONEY_RE.match(normalized_text):
        return None

    try:
        return Decimal(normalized_text.replace(".", "").replace(",", "."))
    except InvalidOperation:
        return None


def first_date_in_row(row: list[object]) -> str | None:
    for value in row:
        parsed = parse_date_value(value)
        if parsed:
            return parsed
    return None


def extract_money_cells(row: list[object]) -> list[Decimal]:
    non_empty_indexes = [index for index, value in enumerate(row) if normalize_spaces(value) != ""]
    candidate_indexes = set(non_empty_indexes[-3:])
    values: list[Decimal] = []
    for index, value in enumerate(row):
        if index not in candidate_indexes and not isinstance(value, str):
            continue

        parsed = parse_money_value(value)
        if parsed is not None:
            values.append(parsed)
    return values


def extract_text_parts(row: list[object]) -> list[str]:
    parts: list[str] = []
    for value in row:
        if parse_date_value(value) or parse_money_value(value) is not None:
            continue

        text = normalize_spaces(value)
        if text:
            parts.append(text)
    return parts


def distribute_amounts(values: list[Decimal]) -> tuple[Decimal | None, Decimal | None, Decimal | None]:
    if not values:
        return None, None, None
    if len(values) == 1:
        return None, None, values[0]
    if len(values) == 2:
        return values[0], None, values[1]
    return values[-3], values[-2], values[-1]


def get_value(row: list[object], index: int | None) -> object:
    if index is None or index >= len(row):
        return None
    return row[index]


def collect_text_fragments(row: list[object], ignored_columns: set[int]) -> str:
    parts: list[str] = []
    for index, value in enumerate(row):
        if index in ignored_columns:
            continue
        if parse_money_value(value) is not None or parse_date_value(value):
            continue

        text = normalize_spaces(value)
        if text:
            parts.append(text)

    return " ".join(parts)


def append_extra_description(current: dict[str, object], row: list[object], ignored_columns: set[int]) -> None:
    extra = collect_text_fragments(row, ignored_columns)
    if extra:
        current["Descricao"] = join_description(current["Descricao"], extra)


def join_description(current: object, extra: str) -> str:
    base = normalize_spaces(current)
    if not base:
        return extra
    return f"{base} {extra}".strip()


def decimal_to_float(value: object) -> float | None:
    if value is None:
        return None
    return float(value)
def detect_balancete_layout(rows: list[list[object]]) -> dict[str, int] | None:
    for index in range(min(len(rows), 10)):
        normalized_row = [normalize_text(value) for value in rows[index]]
        columns = {
            "red": find_header_index(normalized_row, {"red.", "red"}),
            "conta": find_header_index(normalized_row, {"conta"}),
            "descricao": find_header_index(normalized_row, {"descricao", "descrição"}),
            "saldo_anterior": find_header_index(normalized_row, {"saldo anterior"}),
            "debito": find_header_index(normalized_row, {"valor debito", "debito", "valor débito"}),
            "credito": find_header_index(normalized_row, {"valor credito", "credito", "valor crédito"}),
            "saldo_atual": find_header_index(normalized_row, {"saldo atual"}),
        }
        if all(value is not None for value in columns.values()):
            columns["header_index"] = index
            return columns
    return None


def detect_diario_layout(rows: list[list[object]]) -> dict[str, int] | None:
    for index in range(min(len(rows), 15)):
        normalized_row = [normalize_text(value) for value in rows[index]]
        columns = {
            "lote": find_header_index(normalized_row, {"lote"}),
            "nr_mvto": find_header_index(normalized_row, {"nr. mvto", "nr mvto"}),
            "conta_debito": find_header_index(normalized_row, {"cont. debito", "cont debito"}),
            "conta_credito": find_header_index(normalized_row, {"cont. credito", "cont credito"}),
            "historico": find_header_index(normalized_row, {"historico", "histórico"}),
            "debito": find_header_index(normalized_row, {"valor debito", "valor débito", "debito"}),
            "credito": find_header_index(normalized_row, {"valor credito", "valor crédito", "credito"}),
        }
        if all(value is not None for value in columns.values()):
            columns["header_index"] = index
            return columns
    return None


def parse_diario_rows(rows: list[list[object]], columns: dict[str, int]) -> list[dict[str, object]]:
    parsed_rows: list[dict[str, object]] = []
    for row in rows[columns["header_index"] + 1 :]:
        lote = normalize_spaces(get_value(row, columns["lote"]))
        nr_mvto = normalize_spaces(get_value(row, columns["nr_mvto"]))
        conta_debito = normalize_spaces(get_value(row, columns["conta_debito"]))
        conta_credito = normalize_spaces(get_value(row, columns["conta_credito"]))
        historico = normalize_spaces(get_value(row, columns["historico"]))
        debito = parse_money_value(get_value(row, columns["debito"]))
        credito = parse_money_value(get_value(row, columns["credito"]))

        if not any([lote, nr_mvto, conta_debito, conta_credito, historico, debito, credito]):
            continue

        parsed_rows.append(
            {
                "Lote": lote,
                "Nr Mvto": nr_mvto,
                "Conta Debito": conta_debito,
                "Conta Credito": conta_credito,
                "Historico": historico or "Sem historico",
                "Debito": debito,
                "Credito": credito,
            }
        )
    return parsed_rows


def detect_razao_layout(rows: list[list[object]]) -> dict[str, int] | None:
    for index in range(min(len(rows), 20)):
        normalized_row = [normalize_text(value) for value in rows[index]]
        if "conta analitica:" in normalized_row or "conta analitica" in normalized_row:
            return {"account_header_index": index}
    return None


def parse_razao_rows(rows: list[list[object]], layout: dict[str, int]) -> list[dict[str, object]]:
    parsed_rows: list[dict[str, object]] = []
    current_account = ""
    current_date = ""
    pending_history = ""
    pending_record: dict[str, object] | None = None

    for row in rows[layout["account_header_index"] :]:
        row_texts = [normalize_spaces(value) for value in row]
        non_empty = [(index, value) for index, value in enumerate(row_texts) if value]
        if not non_empty:
            continue

        if any(value == "Conta Analitica:" or value == "Conta Analítica:" for _, value in non_empty):
            current_account = row_texts[5] if len(row_texts) > 5 else ""
            current_date = ""
            pending_history = ""
            pending_record = None
            continue

        if any(value == "Data" for _, value in non_empty):
            continue

        row_date = first_date_in_row(row)
        if row_date:
            current_date = row_date
            debit = first_money_in_indexes(row, [7, 8, 9, 10, 11, 12])
            credit = first_money_in_indexes(row, [13, 14, 15, 16])
            saldo = first_money_in_indexes(row, [17, 18, 19])
            pending_record = {
                "Conta Analitica": current_account or "Sem conta",
                "Data": current_date,
                "Historico": pending_history or "Sem historico",
                "Contrapartida": "",
                "Debito": debit,
                "Credito": credit,
                "Saldo": saldo,
            }
            parsed_rows.append(pending_record)
            pending_history = ""
            continue

        history_candidate = first_text_in_indexes(row, [20, 21, 22, 23, 24])
        if history_candidate and not contains_marker(history_candidate, ["saldo anterior:", "contrapartida:"]):
            pending_history = join_description(pending_history, history_candidate)
            if pending_record is not None and normalize_spaces(pending_record["Historico"]) in {"", "Sem historico"}:
                pending_record["Historico"] = pending_history
            continue

        if any(value == "Contrapartida:" for _, value in non_empty):
            if pending_record is not None:
                alt_history = first_text_in_indexes(row, [20, 21, 22, 23, 24])
                if alt_history:
                    pending_record["Historico"] = join_description(pending_record["Historico"], alt_history)
            continue

        contra = first_text_in_indexes(row, [5, 6, 7, 8])
        if contra and pending_record is not None:
            pending_record["Contrapartida"] = contra

    return parsed_rows


def find_header_index(normalized_row: list[str], candidates: set[str]) -> int | None:
    for index, value in enumerate(normalized_row):
        if value in candidates:
            return index
    return None


def parse_balancete_rows(rows: list[list[object]], columns: dict[str, int]) -> list[dict[str, object]]:
    parsed_rows: list[dict[str, object]] = []

    for row in rows[columns["header_index"] + 1 :]:
        conta = normalize_spaces(get_value(row, columns["conta"]))
        descricao = normalize_spaces(get_value(row, columns["descricao"]))
        red = normalize_spaces(get_value(row, columns["red"]))
        saldo_anterior = parse_money_value(get_value(row, columns["saldo_anterior"]))
        debito = parse_money_value(get_value(row, columns["debito"]))
        credito = parse_money_value(get_value(row, columns["credito"]))
        saldo_atual = parse_money_value(get_value(row, columns["saldo_atual"]))

        if not any([conta, descricao, red, saldo_anterior, debito, credito, saldo_atual]):
            continue

        parsed_rows.append(
            {
                "Red": red,
                "Conta": conta,
                "Descricao": descricao or "Sem descricao",
                "Saldo Anterior": saldo_anterior,
                "Debito": debito,
                "Credito": credito,
                "Saldo Atual": saldo_atual,
            }
        )

    return parsed_rows


def first_money_in_indexes(row: list[object], indexes: list[int]) -> Decimal | None:
    for index in indexes:
        value = get_value(row, index)
        parsed = parse_money_value(value)
        if parsed is not None:
            return parsed
    return None


def first_text_in_indexes(row: list[object], indexes: list[int]) -> str:
    parts: list[str] = []
    for index in indexes:
        value = normalize_spaces(get_value(row, index))
        if value:
            parts.append(value)
    return " ".join(parts).strip()


def contains_marker(text: str, markers: list[str]) -> bool:
    normalized = normalize_text(text)
    return any(marker in normalized for marker in markers)
