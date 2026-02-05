from datetime import date, datetime, timezone
from decimal import Decimal
from pathlib import Path
import csv

from openpyxl import Workbook, load_workbook

from cotacoes_moedas.bcb_ptax import PtaxQuote
from cotacoes_moedas.investing import Quote
from cotacoes_moedas.valor_globo import BidAskQuote
from cotacoes_moedas.storage import (
    update_csv_from_xlsx,
    update_xlsx_log,
    update_xlsx_quotes_and_log,
    update_xlsx_usd_brl,
)


def _close_workbook(workbook) -> None:
    close = getattr(workbook, "close", None)
    if callable(close):
        close()


def _make_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "Data"
    sheet["B1"] = "Dolar Oficial Compra"
    sheet["C1"] = "Dolar Oficial Venda"
    sheet["L1"] = "Log"
    workbook.save(path)
    _close_workbook(workbook)


def test_update_xlsx_log_error_format(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "cotacoes.xlsx"
    _make_workbook(xlsx_path)

    target_date = date(2026, 1, 23)
    logged_at = datetime(2026, 1, 23, 9, 15, 0)
    update_xlsx_log(
        xlsx_path,
        target_date=target_date,
        logged_at=logged_at,
        status="ERRO",
        detail="ptax_usd: Timeout",
    )

    workbook = load_workbook(xlsx_path)
    sheet = workbook.active
    assert sheet["L3"].value == "ERRO 23/01/2026 09:15:00 - ptax_usd: Timeout"
    _close_workbook(workbook)


def test_update_csv_from_xlsx_replaces_same_date(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "cotacoes.xlsx"
    csv_path = tmp_path / "cotacoes.csv"
    _make_workbook(xlsx_path)

    local_tz = datetime.now().astimezone().tzinfo or timezone.utc
    collected_at = datetime(2026, 1, 23, 12, 0, 0, tzinfo=local_tz)
    quote = Quote(
        symbol="USD/BRL",
        value=Decimal("5.2849"),
        value_raw="5,2849",
        collected_at=collected_at,
    )
    update_xlsx_usd_brl(xlsx_path, quote, spread=Decimal("0.0020"))
    update_xlsx_log(xlsx_path, target_date=collected_at.date())
    update_csv_from_xlsx(xlsx_path, csv_path)

    updated_quote = Quote(
        symbol="USD/BRL",
        value=Decimal("5.3001"),
        value_raw="5,3001",
        collected_at=collected_at,
    )
    update_xlsx_usd_brl(xlsx_path, updated_quote, spread=Decimal("0.0020"))
    update_xlsx_log(xlsx_path, target_date=collected_at.date())
    update_csv_from_xlsx(xlsx_path, csv_path)

    with csv_path.open("r", encoding="utf-8", newline="") as handle:
        rows = list(csv.reader(handle, delimiter=";"))

    data_rows = [row for row in rows if row and row[0] == "23/01/2026"]
    assert len(data_rows) == 1
    assert data_rows[0][1] == "5,3001"
    assert data_rows[0][2] == "5,3021"


def test_update_xlsx_quotes_and_log_fills_only_blanks(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "cotacoes.xlsx"
    _make_workbook(xlsx_path)

    target_date = date(2026, 1, 23)

    workbook = load_workbook(xlsx_path)
    sheet = workbook.active
    sheet["A3"] = target_date
    sheet["A3"].number_format = "dd/mm/yyyy"
    sheet["B3"] = Decimal("5.0000")
    sheet["B3"].number_format = "0.0000"
    workbook.save(xlsx_path)
    _close_workbook(workbook)

    collected_at = datetime(2026, 1, 23, 12, 0, 0, tzinfo=timezone.utc)
    usd = Quote(
        symbol="USD/BRL",
        value=Decimal("5.2849"),
        value_raw="5,2849",
        collected_at=collected_at,
    )
    ptax_usd = PtaxQuote(
        symbol="USD/BRL PTAX",
        buy=Decimal("5.1000"),
        sell=Decimal("5.2000"),
        buy_raw="5,1000",
        sell_raw="5,2000",
        collected_at=collected_at,
    )
    turismo = BidAskQuote(
        symbol="USD/BRL Turismo",
        buy=Decimal("5.5000"),
        sell=Decimal("5.6000"),
        buy_raw="5,5000",
        sell_raw="5,6000",
        collected_at=collected_at,
    )

    written = update_xlsx_quotes_and_log(
        xlsx_path,
        target_date=target_date,
        usd_brl=usd,
        ptax_usd=ptax_usd,
        turismo=turismo,
        spread=Decimal("0.0020"),
        overwrite_quotes=False,
        logged_at=datetime(2026, 1, 23, 9, 15, 0),
    )

    assert written["usd_brl"] == ("venda",)
    assert written["ptax_usd"] == ("compra", "venda")
    assert written["turismo"] == ("compra", "venda")

    workbook = load_workbook(xlsx_path)
    sheet = workbook.active
    assert Decimal(str(sheet["B3"].value)).quantize(Decimal("0.0001")) == Decimal("5.0000")
    assert Decimal(str(sheet["C3"].value)).quantize(Decimal("0.0001")) == Decimal("5.0020")
    assert Decimal(str(sheet["D3"].value)).quantize(Decimal("0.0001")) == Decimal("5.1000")
    assert Decimal(str(sheet["E3"].value)).quantize(Decimal("0.0001")) == Decimal("5.2000")
    assert Decimal(str(sheet["F3"].value)).quantize(Decimal("0.0001")) == Decimal("5.5000")
    assert Decimal(str(sheet["G3"].value)).quantize(Decimal("0.0001")) == Decimal("5.6000")
    assert sheet["L3"].value == "OK 23/01/2026 09:15:00"
    _close_workbook(workbook)


def test_update_xlsx_quotes_and_log_overwrites_when_enabled(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "cotacoes.xlsx"
    _make_workbook(xlsx_path)

    target_date = date(2026, 1, 23)

    workbook = load_workbook(xlsx_path)
    sheet = workbook.active
    sheet["A3"] = target_date
    sheet["A3"].number_format = "dd/mm/yyyy"
    sheet["B3"] = Decimal("4.0000")
    sheet["B3"].number_format = "0.0000"
    sheet["C3"] = Decimal("4.0020")
    sheet["C3"].number_format = "0.0000"
    workbook.save(xlsx_path)
    _close_workbook(workbook)

    collected_at = datetime(2026, 1, 23, 12, 0, 0, tzinfo=timezone.utc)
    usd = Quote(
        symbol="USD/BRL",
        value=Decimal("5.2849"),
        value_raw="5,2849",
        collected_at=collected_at,
    )

    written = update_xlsx_quotes_and_log(
        xlsx_path,
        target_date=target_date,
        usd_brl=usd,
        spread=Decimal("0.0020"),
        overwrite_quotes=True,
        logged_at=datetime(2026, 1, 23, 9, 15, 0),
    )

    assert written["usd_brl"] == ("compra", "venda")

    workbook = load_workbook(xlsx_path)
    sheet = workbook.active
    assert Decimal(str(sheet["B3"].value)).quantize(Decimal("0.0001")) == Decimal("5.2849")
    assert Decimal(str(sheet["C3"].value)).quantize(Decimal("0.0001")) == Decimal("5.2869")
    _close_workbook(workbook)
