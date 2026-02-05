from __future__ import annotations

from datetime import date, datetime, timezone
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import csv
import re
from typing import Callable

from openpyxl import load_workbook

from .bcb_ptax import PtaxQuote
from .investing import Quote
from .valor_globo import BidAskQuote


_DECIMAL_4 = Decimal("0.0001")
_DATE_PATTERN = re.compile(r"^\d{2}/\d{2}/\d{4}$")
_LOCAL_TZ = datetime.now().astimezone().tzinfo or timezone.utc


def _quantize_4(value: Decimal) -> Decimal:
    return value.quantize(_DECIMAL_4, rounding=ROUND_HALF_UP)


def _is_blank(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    return False


def _set_cell(
    sheet,
    address: str,
    value: object,
    *,
    number_format: str | None = None,
    overwrite: bool = True,
) -> bool:
    cell = sheet[address]
    if not overwrite and not _is_blank(cell.value):
        return False
    cell.value = value
    if number_format:
        cell.number_format = number_format
    return True


def _load_and_save_workbook(
    path: Path,
    apply_updates: Callable[[object], object],
) -> object:
    workbook = load_workbook(path)
    try:
        sheet = workbook.active
        result = apply_updates(sheet)
        workbook.save(path)
        return result
    finally:
        close = getattr(workbook, "close", None)
        if callable(close):
            close()


def _coerce_date(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
    return None


def _as_local_date(value: datetime) -> date:
    if value.tzinfo is None:
        localized = value.replace(tzinfo=_LOCAL_TZ)
    else:
        localized = value.astimezone(_LOCAL_TZ)
    return localized.date()


def _as_local_datetime(value: datetime) -> datetime:
    if value.tzinfo is None:
        return value.replace(tzinfo=_LOCAL_TZ)
    return value.astimezone(_LOCAL_TZ)


def _find_row_by_date(sheet, target_date: date) -> int | None:
    for row in range(3, sheet.max_row + 1):
        cell_value = sheet.cell(row=row, column=1).value
        cell_date = _coerce_date(cell_value)
        if cell_date == target_date:
            return row
    return None


def _find_last_date_row(sheet) -> int | None:
    last_date_row = None
    for row in range(3, sheet.max_row + 1):
        if _coerce_date(sheet.cell(row=row, column=1).value):
            last_date_row = row
    return last_date_row


def _find_or_create_row_by_date(sheet, target_date: date) -> int:
    row = _find_row_by_date(sheet, target_date)
    if row is not None:
        return row
    last_date_row = _find_last_date_row(sheet)
    row = (last_date_row or 2) + 1
    sheet[f"A{row}"] = target_date
    sheet[f"A{row}"].number_format = "dd/mm/yyyy"
    return row


def _find_last_updated_row(sheet) -> int:
    last_logged = None
    last_date = None
    for row in range(3, sheet.max_row + 1):
        if _coerce_date(sheet.cell(row=row, column=1).value):
            last_date = row
        if sheet.cell(row=row, column=12).value:
            last_logged = row
    if last_logged is not None:
        return last_logged
    if last_date is not None:
        return last_date
    raise ValueError("nenhuma linha com data encontrada na planilha")


def _format_date_cell(value: object) -> str:
    cell_date = _coerce_date(value)
    if cell_date:
        return cell_date.strftime("%d/%m/%Y")
    return ""


def _to_decimal(value: object) -> Decimal | None:
    if value is None:
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        cleaned = re.sub(r"[^\d,.-]", "", text)
        if "," in cleaned:
            cleaned = cleaned.replace(".", "").replace(",", ".")
        return Decimal(cleaned)
    raise TypeError(f"valor numerico invalido: {value!r}")


def _format_number_cell(value: object) -> str:
    number = _to_decimal(value)
    if number is None:
        return ""
    formatted = _quantize_4(number)
    return f"{formatted:.4f}".replace(".", ",")


def _format_log_cell(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return f"OK {value.strftime('%d/%m/%Y %H:%M:%S')}"
    return str(value).strip()


def update_xlsx_usd_brl(
    path: str | Path,
    quote: Quote,
    spread: Decimal = Decimal("0.0020"),
    *,
    target_date: date | None = None,
    overwrite: bool = True,
) -> None:
    target = Path(path)
    if not target.exists():
        raise FileNotFoundError(target)

    def _apply(sheet) -> None:
        collected_date = target_date or _as_local_date(quote.collected_at)
        compra = _quantize_4(quote.value)

        row = _find_or_create_row_by_date(sheet, collected_date)
        _set_cell(
            sheet,
            f"A{row}",
            collected_date,
            number_format="dd/mm/yyyy",
            overwrite=True,
        )
        buy_address = f"B{row}"
        existing_buy = sheet[buy_address].value
        wrote_buy = _set_cell(
            sheet,
            buy_address,
            compra,
            number_format="0.0000",
            overwrite=overwrite,
        )
        buy_for_sale = compra
        if not wrote_buy and not _is_blank(existing_buy):
            try:
                parsed = _to_decimal(existing_buy)
                if parsed is not None:
                    buy_for_sale = _quantize_4(parsed)
            except Exception:
                buy_for_sale = compra

        venda = _quantize_4(buy_for_sale + spread)
        _set_cell(
            sheet,
            f"C{row}",
            venda,
            number_format="0.0000",
            overwrite=overwrite,
        )

    _load_and_save_workbook(target, _apply)


def update_xlsx_dolar_turismo(
    path: str | Path,
    quote: BidAskQuote,
    target_date: date | None = None,
    *,
    overwrite: bool = True,
) -> None:
    target = Path(path)
    if not target.exists():
        raise FileNotFoundError(target)

    def _apply(sheet) -> None:
        use_date = target_date or _as_local_date(quote.collected_at)
        row = _find_or_create_row_by_date(sheet, use_date)

        compra = _quantize_4(quote.buy)
        venda = _quantize_4(quote.sell)

        _set_cell(
            sheet,
            f"F{row}",
            compra,
            number_format="0.0000",
            overwrite=overwrite,
        )
        _set_cell(
            sheet,
            f"G{row}",
            venda,
            number_format="0.0000",
            overwrite=overwrite,
        )

    _load_and_save_workbook(target, _apply)


def update_xlsx_dolar_ptax(
    path: str | Path,
    quote: PtaxQuote,
    target_date: date | None = None,
    *,
    overwrite: bool = True,
) -> None:
    target = Path(path)
    if not target.exists():
        raise FileNotFoundError(target)

    def _apply(sheet) -> None:
        use_date = target_date or _as_local_date(quote.collected_at)
        row = _find_or_create_row_by_date(sheet, use_date)

        compra = _quantize_4(quote.buy)
        venda = _quantize_4(quote.sell)

        _set_cell(
            sheet,
            f"D{row}",
            compra,
            number_format="0.0000",
            overwrite=overwrite,
        )
        _set_cell(
            sheet,
            f"E{row}",
            venda,
            number_format="0.0000",
            overwrite=overwrite,
        )

    _load_and_save_workbook(target, _apply)


def update_xlsx_euro_ptax(
    path: str | Path,
    quote: PtaxQuote,
    target_date: date | None = None,
    *,
    overwrite: bool = True,
) -> None:
    target = Path(path)
    if not target.exists():
        raise FileNotFoundError(target)

    def _apply(sheet) -> None:
        use_date = target_date or _as_local_date(quote.collected_at)
        row = _find_or_create_row_by_date(sheet, use_date)

        compra = _quantize_4(quote.buy)
        venda = _quantize_4(quote.sell)

        _set_cell(
            sheet,
            f"H{row}",
            compra,
            number_format="0.0000",
            overwrite=overwrite,
        )
        _set_cell(
            sheet,
            f"I{row}",
            venda,
            number_format="0.0000",
            overwrite=overwrite,
        )

    _load_and_save_workbook(target, _apply)


def update_xlsx_chf_ptax(
    path: str | Path,
    quote: PtaxQuote,
    target_date: date | None = None,
    *,
    overwrite: bool = True,
) -> None:
    target = Path(path)
    if not target.exists():
        raise FileNotFoundError(target)

    def _apply(sheet) -> None:
        use_date = target_date or _as_local_date(quote.collected_at)
        row = _find_or_create_row_by_date(sheet, use_date)

        compra = _quantize_4(quote.buy)
        venda = _quantize_4(quote.sell)

        _set_cell(
            sheet,
            f"J{row}",
            compra,
            number_format="0.0000",
            overwrite=overwrite,
        )
        _set_cell(
            sheet,
            f"K{row}",
            venda,
            number_format="0.0000",
            overwrite=overwrite,
        )

    _load_and_save_workbook(target, _apply)


def update_xlsx_log(
    path: str | Path,
    target_date: date | None = None,
    logged_at: datetime | None = None,
    status: str = "OK",
    detail: str | None = None,
) -> None:
    target = Path(path)
    if not target.exists():
        raise FileNotFoundError(target)

    def _apply(sheet) -> None:
        use_date = target_date or datetime.now(_LOCAL_TZ).date()
        row = _find_or_create_row_by_date(sheet, use_date)

        when = _as_local_datetime(logged_at) if logged_at else datetime.now(_LOCAL_TZ)
        timestamp = when.strftime("%d/%m/%Y %H:%M:%S")
        status_text = (status or "OK").strip()
        if detail:
            detail_text = " ".join(str(detail).split())
            cell_value = f"{status_text} {timestamp} - {detail_text}"
        else:
            cell_value = f"{status_text} {timestamp}"

        _set_cell(sheet, f"L{row}", cell_value, overwrite=True)

    _load_and_save_workbook(target, _apply)


def update_xlsx_quotes_and_log(
    path: str | Path,
    *,
    target_date: date,
    usd_brl: Quote | None = None,
    ptax_usd: PtaxQuote | None = None,
    turismo: BidAskQuote | None = None,
    ptax_eur: PtaxQuote | None = None,
    ptax_chf: PtaxQuote | None = None,
    spread: Decimal = Decimal("0.0020"),
    overwrite_quotes: bool = False,
    logged_at: datetime | None = None,
    status: str = "OK",
    detail: str | None = None,
) -> dict[str, tuple[str, ...]]:
    """Atualiza varias cotacoes no XLSX com apenas um save.

    Retorna, por fonte, quais campos foram efetivamente gravados
    (por padrao nao sobrescreve celulas ja preenchidas).
    """
    target = Path(path)
    if not target.exists():
        raise FileNotFoundError(target)

    def _apply(sheet) -> dict[str, tuple[str, ...]]:
        row = _find_or_create_row_by_date(sheet, target_date)
        _set_cell(
            sheet,
            f"A{row}",
            target_date,
            number_format="dd/mm/yyyy",
            overwrite=True,
        )

        written: dict[str, tuple[str, ...]] = {}

        if usd_brl:
            compra = _quantize_4(usd_brl.value)
            updated: list[str] = []

            buy_address = f"B{row}"
            existing_buy = sheet[buy_address].value
            wrote_buy = _set_cell(
                sheet,
                buy_address,
                compra,
                number_format="0.0000",
                overwrite=overwrite_quotes,
            )
            if wrote_buy:
                updated.append("compra")

            buy_for_sale = compra
            if not wrote_buy and not _is_blank(existing_buy):
                try:
                    parsed = _to_decimal(existing_buy)
                    if parsed is not None:
                        buy_for_sale = _quantize_4(parsed)
                except Exception:
                    buy_for_sale = compra

            venda = _quantize_4(buy_for_sale + spread)
            if _set_cell(
                sheet,
                f"C{row}",
                venda,
                number_format="0.0000",
                overwrite=overwrite_quotes,
            ):
                updated.append("venda")
            written["usd_brl"] = tuple(updated)

        if ptax_usd:
            compra = _quantize_4(ptax_usd.buy)
            venda = _quantize_4(ptax_usd.sell)
            updated = []
            if _set_cell(
                sheet,
                f"D{row}",
                compra,
                number_format="0.0000",
                overwrite=overwrite_quotes,
            ):
                updated.append("compra")
            if _set_cell(
                sheet,
                f"E{row}",
                venda,
                number_format="0.0000",
                overwrite=overwrite_quotes,
            ):
                updated.append("venda")
            written["ptax_usd"] = tuple(updated)

        if turismo:
            compra = _quantize_4(turismo.buy)
            venda = _quantize_4(turismo.sell)
            updated = []
            if _set_cell(
                sheet,
                f"F{row}",
                compra,
                number_format="0.0000",
                overwrite=overwrite_quotes,
            ):
                updated.append("compra")
            if _set_cell(
                sheet,
                f"G{row}",
                venda,
                number_format="0.0000",
                overwrite=overwrite_quotes,
            ):
                updated.append("venda")
            written["turismo"] = tuple(updated)

        if ptax_eur:
            compra = _quantize_4(ptax_eur.buy)
            venda = _quantize_4(ptax_eur.sell)
            updated = []
            if _set_cell(
                sheet,
                f"H{row}",
                compra,
                number_format="0.0000",
                overwrite=overwrite_quotes,
            ):
                updated.append("compra")
            if _set_cell(
                sheet,
                f"I{row}",
                venda,
                number_format="0.0000",
                overwrite=overwrite_quotes,
            ):
                updated.append("venda")
            written["ptax_eur"] = tuple(updated)

        if ptax_chf:
            compra = _quantize_4(ptax_chf.buy)
            venda = _quantize_4(ptax_chf.sell)
            updated = []
            if _set_cell(
                sheet,
                f"J{row}",
                compra,
                number_format="0.0000",
                overwrite=overwrite_quotes,
            ):
                updated.append("compra")
            if _set_cell(
                sheet,
                f"K{row}",
                venda,
                number_format="0.0000",
                overwrite=overwrite_quotes,
            ):
                updated.append("venda")
            written["ptax_chf"] = tuple(updated)

        when = _as_local_datetime(logged_at) if logged_at else datetime.now(_LOCAL_TZ)
        timestamp = when.strftime("%d/%m/%Y %H:%M:%S")
        status_text = (status or "OK").strip()
        if detail:
            detail_text = " ".join(str(detail).split())
            cell_value = f"{status_text} {timestamp} - {detail_text}"
        else:
            cell_value = f"{status_text} {timestamp}"

        _set_cell(sheet, f"L{row}", cell_value, overwrite=True)
        return written

    return _load_and_save_workbook(target, _apply)


def update_csv_from_xlsx(
    xlsx_path: str | Path,
    csv_path: str | Path,
) -> None:
    source = Path(xlsx_path)
    if not source.exists():
        raise FileNotFoundError(source)

    workbook = load_workbook(source, data_only=True)
    try:
        sheet = workbook.active
        row = _find_last_updated_row(sheet)
    finally:
        close = getattr(workbook, "close", None)
        if callable(close):
            close()

    date_value = _format_date_cell(sheet.cell(row=row, column=1).value)
    if not date_value:
        raise ValueError("data da linha nao encontrada para atualizar o CSV")

    data_row = [date_value]
    for col in range(2, 12):
        data_row.append(_format_number_cell(sheet.cell(row=row, column=col).value))
    data_row.append(_format_log_cell(sheet.cell(row=row, column=12).value))

    default_header = [
        "Data",
        "Dolar Oficial Compra",
        "Dolar Oficial Venda",
        "Dolar PTAX Compra",
        "Dolar PTAX Venda",
        "Dolar Turismo Compra",
        "Dolar Turismo Venda",
        "Euro PTAX Compra",
        "Euro PTAX Venda",
        "CHF PTAX Compra",
        "CHF PTAX Venda",
        "Log",
    ]

    target = Path(csv_path)
    rows: list[list[str]] = []
    if target.exists():
        for encoding in ("utf-8", "latin-1"):
            try:
                with target.open("r", encoding=encoding, newline="") as handle:
                    reader = csv.reader(handle, delimiter=";")
                    rows = list(reader)
                break
            except UnicodeDecodeError:
                rows = []
                continue

    if not rows or all(len(row) <= 1 for row in rows):
        rows = [default_header, data_row]
    else:
        header_rows: list[list[str]] = []
        data_rows: list[list[str]] = []
        for row_values in rows:
            if row_values and _DATE_PATTERN.match(row_values[0]):
                data_rows.append(row_values)
            else:
                header_rows.append(row_values)

        if not header_rows:
            header_rows = [default_header]

        replaced = False
        new_data_rows: list[list[str]] = []
        for row_values in data_rows:
            if row_values and row_values[0] == date_value:
                new_data_rows.append(data_row)
                replaced = True
            else:
                new_data_rows.append(row_values)
        if not replaced:
            new_data_rows.append(data_row)

        rows = header_rows + new_data_rows

    with target.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle, delimiter=";", quoting=csv.QUOTE_MINIMAL)
        writer.writerows(rows)
