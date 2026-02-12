from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from decimal import Decimal
import re

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from .parsing import ParseDecimalError, parse_pt_br_decimal
from .page_consistency import (
    PageCheck,
    PageConsistencyError,
    ensure_page_consistency,
)
from .playwright_utils import chromium_page, proxy_from_env


VALOR_GLOBO_URL = "https://valor.globo.com/"
ROW_LABEL_RE = re.compile(r"D.lar Turismo", re.IGNORECASE)
_HAS_DIGIT = re.compile(r"\d")


class PriceParseError(RuntimeError):
    pass


@dataclass(frozen=True)
class BidAskQuote:
    symbol: str
    buy: Decimal
    sell: Decimal
    buy_raw: str
    sell_raw: str
    collected_at: datetime


def fetch_dolar_turismo(
    headless: bool = True,
    timeout_ms: int = 45000,
) -> BidAskQuote:
    proxy = proxy_from_env()
    buy_raw = ""
    sell_raw = ""
    try:
        with chromium_page(headless=headless, proxy=proxy) as page:
            page.goto(VALOR_GLOBO_URL, wait_until="commit", timeout=timeout_ms)

            row = page.locator("tr", has_text=ROW_LABEL_RE).first
            row.wait_for(state="visible", timeout=timeout_ms)
            ensure_page_consistency(
                page,
                source="Valor Dolar Turismo",
                checks=[
                    PageCheck(
                        "url esperada",
                        lambda p: (
                            "valor.globo.com" in (p.url or "").lower(),
                            f"url atual: {p.url}",
                        ),
                    ),
                    PageCheck(
                        "linha Dolar Turismo",
                        lambda p: (
                            p.locator("tr", has_text=ROW_LABEL_RE).count() > 0,
                            "linha de Dolar Turismo nao encontrada",
                        ),
                    ),
                ],
            )

            cells = row.locator("td")
            if cells.count() < 3:
                raise PriceParseError("linha de Dolar Turismo incompleta")

            buy_raw = cells.nth(1).inner_text().strip()
            sell_raw = cells.nth(2).inner_text().strip()
    except PlaywrightTimeoutError as exc:
        raise PriceParseError("timeout ao buscar Dolar Turismo") from exc
    except PageConsistencyError as exc:
        raise PriceParseError(str(exc)) from exc

    if not buy_raw or not sell_raw or not _HAS_DIGIT.search(buy_raw) or not _HAS_DIGIT.search(sell_raw):
        raise PriceParseError(
            "cotacao de Dolar Turismo nao atualizada no Valor"
        )

    if not buy_raw or not sell_raw:
        raise PriceParseError("nao encontrou compra/venda de Dolar Turismo")

    try:
        buy = parse_pt_br_decimal(buy_raw)
        sell = parse_pt_br_decimal(sell_raw)
    except ParseDecimalError as exc:
        raise PriceParseError(str(exc)) from exc

    return BidAskQuote(
        symbol="USD/BRL Turismo",
        buy=buy,
        sell=sell,
        buy_raw=buy_raw,
        sell_raw=sell_raw,
        collected_at=datetime.now(timezone.utc),
    )
