from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from decimal import Decimal

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from .parsing import ParseDecimalError, parse_pt_br_decimal
from .playwright_utils import chromium_page, proxy_from_env


USD_BRL_URL = "https://br.investing.com/currencies/usd-brl"
USD_BRL_SELECTOR = '[data-test="instrument-price-last"]'


class PriceParseError(RuntimeError):
    pass


@dataclass(frozen=True)
class Quote:
    symbol: str
    value: Decimal
    value_raw: str
    collected_at: datetime


def fetch_usd_brl(
    headless: bool = True,
    timeout_ms: int = 45000,
) -> Quote:
    proxy = proxy_from_env()
    raw_value = ""
    try:
        with chromium_page(headless=headless, proxy=proxy) as page:
            page.goto(USD_BRL_URL, wait_until="commit", timeout=timeout_ms)
            locator = page.locator(USD_BRL_SELECTOR).first
            locator.wait_for(state="visible", timeout=timeout_ms)
            raw_value = locator.inner_text().strip()
    except PlaywrightTimeoutError as exc:
        raise PriceParseError(
            "timeout ao aguardar o valor em instrument-price-last"
        ) from exc

    if not raw_value:
        raise PriceParseError("nao encontrou o valor em instrument-price-last")
    try:
        value = parse_pt_br_decimal(raw_value)
    except ParseDecimalError as exc:
        raise PriceParseError(str(exc)) from exc
    return Quote(
        symbol="USD/BRL",
        value=value,
        value_raw=raw_value,
        collected_at=datetime.now(timezone.utc),
    )
