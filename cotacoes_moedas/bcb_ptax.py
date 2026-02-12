from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal
import re
import time

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from .parsing import ParseDecimalError, parse_pt_br_decimal
from .page_consistency import (
    PageCheck,
    PageConsistencyError,
    describe_page,
    ensure_page_consistency,
)
from .playwright_utils import chromium_page, proxy_from_env


BCB_HISTORICO_URL = "https://www.bcb.gov.br/estabilidadefinanceira/historicocotacoes"
PTAX_USD_LABEL = "DOLAR DOS EUA"
PTAX_EUR_LABEL = "EURO"
PTAX_CHF_LABEL = "FRANCO SUICO"
_DATE_PATTERN = re.compile(r"^\d{2}/\d{2}/\d{4}$")


class PriceParseError(RuntimeError):
    pass


@dataclass(frozen=True)
class PtaxQuote:
    symbol: str
    buy: Decimal
    sell: Decimal
    buy_raw: str
    sell_raw: str
    collected_at: datetime


def _format_date(value: date) -> str:
    return value.strftime("%d/%m/%Y")


def _parse_ptax_date(value: str) -> date:
    return datetime.strptime(value, "%d/%m/%Y").date()


def _extract_ptax_rows(frame) -> list[tuple[date, str, str, str]]:
    rows: list[tuple[date, str, str, str]] = []
    row_locator = frame.locator("tr")
    for index in range(row_locator.count()):
        cells = row_locator.nth(index).locator("td")
        if cells.count() < 4:
            continue
        date_text = cells.nth(0).inner_text().strip()
        if not _DATE_PATTERN.match(date_text):
            continue
        buy_raw = cells.nth(2).inner_text().strip()
        sell_raw = cells.nth(3).inner_text().strip()
        rows.append((_parse_ptax_date(date_text), date_text, buy_raw, sell_raw))
    return rows


def _load_ptax_rows(frame, timeout_ms: int) -> list[tuple[date, str, str, str]]:
    deadline = time.monotonic() + (timeout_ms / 1000)
    while time.monotonic() < deadline:
        rows = _extract_ptax_rows(frame)
        if rows:
            return rows
        time.sleep(0.5)
    raise PriceParseError("timeout ao carregar tabela PTAX")


def _fetch_ptax(
    currency_label: str,
    symbol: str,
    headless: bool = True,
    timeout_ms: int = 45000,
    lookback_days: int = 7,
) -> PtaxQuote:
    today = date.today()
    start = today - timedelta(days=max(1, lookback_days))
    target_date = _format_date(today)
    buy_raw = ""
    sell_raw = ""

    proxy = proxy_from_env()
    try:
        with chromium_page(
            headless=headless,
            proxy=proxy,
            launch_args=["--no-sandbox"],
        ) as page:
            page.goto(BCB_HISTORICO_URL, wait_until="commit", timeout=timeout_ms)
            iframe_selector = (
                'iframe[src*="ptax.bcb.gov.br/ptax_internet/consultaBoletim.do"]'
            )
            ensure_page_consistency(
                page,
                source=f"BCB PTAX {currency_label}",
                checks=[
                    PageCheck(
                        "url esperada",
                        lambda p: (
                            "bcb.gov.br/estabilidadefinanceira/historicocotacoes"
                            in (p.url or "").lower(),
                            f"url atual: {p.url}",
                        ),
                    ),
                    PageCheck(
                        "iframe PTAX",
                        lambda p: (
                            p.locator(iframe_selector).count() > 0,
                            f"iframe ausente: {iframe_selector}",
                        ),
                    ),
                ],
            )
            frame = page.frame_locator(
                iframe_selector
            )
            required_fields = [
                ('input[name="RadOpcao"][value="1"]', "opcao de periodo"),
                ('input[name="DATAINI"]', "campo DATAINI"),
                ('input[name="DATAFIM"]', "campo DATAFIM"),
                ('select[name="ChkMoeda"]', "combo de moeda"),
                ('input[type="submit"]', "botao de consulta"),
            ]
            for selector, label in required_fields:
                if frame.locator(selector).count() <= 0:
                    raise PriceParseError(
                        "estrutura da pagina possivelmente alterada em "
                        f"BCB PTAX ({currency_label}); "
                        f"campo ausente ({label}): {selector}; "
                        f"{describe_page(page)}"
                    )
            frame.locator('input[name="RadOpcao"][value="1"]').check()
            frame.locator('input[name="DATAINI"]').fill(_format_date(start))
            frame.locator('input[name="DATAFIM"]').fill(target_date)
            frame.locator('select[name="ChkMoeda"]').select_option(
                label=currency_label
            )
            frame.locator('input[type="submit"]').click()

            rows = _load_ptax_rows(frame, timeout_ms)
            target_row = next((row for row in rows if row[0] == today), None)
            if target_row is None:
                last_row = max(rows, key=lambda row: row[0])
                raise PriceParseError(
                    "cotacao PTAX nao disponivel para "
                    f"{target_date}; ultima data disponivel: {last_row[1]}"
                )
            buy_raw = target_row[2]
            sell_raw = target_row[3]
    except PlaywrightTimeoutError as exc:
        raise PriceParseError(
            f"timeout ao buscar PTAX para {currency_label}"
        ) from exc
    except PageConsistencyError as exc:
        raise PriceParseError(str(exc)) from exc

    if not buy_raw or not sell_raw:
        raise PriceParseError(
            f"nao encontrou cotacao PTAX para a data atual ({currency_label})"
        )

    try:
        buy = parse_pt_br_decimal(buy_raw)
        sell = parse_pt_br_decimal(sell_raw)
    except ParseDecimalError as exc:
        raise PriceParseError(str(exc)) from exc
    return PtaxQuote(
        symbol=symbol,
        buy=buy,
        sell=sell,
        buy_raw=buy_raw,
        sell_raw=sell_raw,
        collected_at=datetime.now(timezone.utc),
    )


def fetch_dolar_ptax(
    headless: bool = True,
    timeout_ms: int = 45000,
    lookback_days: int = 7,
) -> PtaxQuote:
    return _fetch_ptax(
        currency_label=PTAX_USD_LABEL,
        symbol="USD/BRL PTAX",
        headless=headless,
        timeout_ms=timeout_ms,
        lookback_days=lookback_days,
    )


def fetch_euro_ptax(
    headless: bool = True,
    timeout_ms: int = 45000,
    lookback_days: int = 7,
) -> PtaxQuote:
    return _fetch_ptax(
        currency_label=PTAX_EUR_LABEL,
        symbol="EUR/BRL PTAX",
        headless=headless,
        timeout_ms=timeout_ms,
        lookback_days=lookback_days,
    )


def fetch_chf_ptax(
    headless: bool = True,
    timeout_ms: int = 45000,
    lookback_days: int = 7,
) -> PtaxQuote:
    return _fetch_ptax(
        currency_label=PTAX_CHF_LABEL,
        symbol="CHF/BRL PTAX",
        headless=headless,
        timeout_ms=timeout_ms,
        lookback_days=lookback_days,
    )
