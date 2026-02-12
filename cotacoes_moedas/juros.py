from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timezone
from decimal import Decimal, ROUND_HALF_UP, localcontext
import re
import time

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from .parsing import ParseDecimalError, parse_pt_br_decimal
from .page_consistency import (
    PageCheck,
    PageConsistencyError,
    ensure_page_consistency,
)
from .playwright_utils import chromium_page, proxy_from_env


TJLP_URL = (
    "https://www.bndes.gov.br/wps/portal/site/home/financiamento/guia/"
    "custos-financeiros/taxa-juros-longo-prazo-tjlp"
)
SELIC_URL = "https://www.bcb.gov.br/controleinflacao/historicotaxasjuros"
_DATE_PATTERN = re.compile(r"^\d{2}/\d{2}/\d{4}$")
_PERCENT_RE = re.compile(r"(-?\d[\d\.,]*)\s*%")
_HAS_DIGIT = re.compile(r"\d")
_CDI_SPREAD = Decimal("0.10")
_CDI_QUANTIZER = Decimal("0.0000000001")
_CDI_BUSINESS_DAYS = Decimal("252")


class PriceParseError(RuntimeError):
    pass


@dataclass(frozen=True)
class InterestRateQuote:
    name: str
    value: Decimal
    value_raw: str
    reference_date: date | None
    collected_at: datetime


def _parse_percent_value(raw_text: str) -> Decimal:
    text = " ".join((raw_text or "").split())
    match = _PERCENT_RE.search(text)
    candidate = match.group(1) if match else text
    try:
        return parse_pt_br_decimal(candidate)
    except ParseDecimalError as exc:
        raise PriceParseError(str(exc)) from exc


def _extract_latest_selic_row(page) -> tuple[date, str, str] | None:
    latest: tuple[date, str, str] | None = None
    rows = page.locator("table tr")
    for index in range(rows.count()):
        cells = rows.nth(index).locator("td")
        if cells.count() < 5:
            continue
        date_raw = " ".join(cells.nth(1).inner_text().split())
        if not _DATE_PATTERN.match(date_raw):
            continue
        rate_raw = " ".join(cells.nth(4).inner_text().split())
        if not _HAS_DIGIT.search(rate_raw):
            continue
        try:
            row_date = datetime.strptime(date_raw, "%d/%m/%Y").date()
        except ValueError:
            continue
        if latest is None or row_date > latest[0]:
            latest = (row_date, date_raw, rate_raw)
    return latest


def _wait_latest_selic_row(page, timeout_ms: int) -> tuple[date, str, str]:
    deadline = time.monotonic() + (timeout_ms / 1000)
    while time.monotonic() < deadline:
        latest = _extract_latest_selic_row(page)
        if latest is not None:
            return latest
        time.sleep(0.5)
    raise PriceParseError("timeout ao carregar tabela da SELIC")


def fetch_tjlp(
    headless: bool = True,
    timeout_ms: int = 45000,
) -> InterestRateQuote:
    proxy = proxy_from_env()
    raw_value = ""
    try:
        with chromium_page(headless=headless, proxy=proxy) as page:
            page.goto(TJLP_URL, wait_until="domcontentloaded", timeout=timeout_ms)
            locator = page.locator("div.valor", has_text="%").first
            locator.wait_for(state="visible", timeout=timeout_ms)
            ensure_page_consistency(
                page,
                source="BNDES TJLP",
                checks=[
                    PageCheck(
                        "url esperada",
                        lambda p: (
                            "bndes.gov.br" in (p.url or "").lower(),
                            f"url atual: {p.url}",
                        ),
                    ),
                    PageCheck(
                        "seletor de valor",
                        lambda p: (
                            p.locator("div.valor", has_text="%").count() > 0,
                            "nao encontrou bloco com percentual da TJLP",
                        ),
                    ),
                ],
            )
            raw_value = locator.inner_text().strip()
    except PlaywrightTimeoutError as exc:
        raise PriceParseError("timeout ao buscar TJLP no BNDES") from exc
    except PageConsistencyError as exc:
        raise PriceParseError(str(exc)) from exc

    if not raw_value:
        raise PriceParseError("nao encontrou valor da TJLP")

    value = _parse_percent_value(raw_value)
    return InterestRateQuote(
        name="TJLP",
        value=value,
        value_raw=raw_value,
        reference_date=None,
        collected_at=datetime.now(timezone.utc),
    )


def fetch_selic(
    headless: bool = True,
    timeout_ms: int = 45000,
) -> InterestRateQuote:
    proxy = proxy_from_env()
    reference_date: date | None = None
    raw_value = ""
    try:
        with chromium_page(
            headless=headless,
            proxy=proxy,
            launch_args=["--no-sandbox"],
        ) as page:
            page.goto(SELIC_URL, wait_until="domcontentloaded", timeout=timeout_ms)
            table = page.locator("table").first
            table.wait_for(state="visible", timeout=timeout_ms)
            ensure_page_consistency(
                page,
                source="BCB SELIC",
                checks=[
                    PageCheck(
                        "url esperada",
                        lambda p: (
                            "bcb.gov.br/controleinflacao/historicotaxasjuros"
                            in (p.url or "").lower(),
                            f"url atual: {p.url}",
                        ),
                    ),
                    PageCheck(
                        "linhas da tabela",
                        lambda p: (
                            p.locator("table tr").count() > 1,
                            "tabela sem linhas suficientes para historico",
                        ),
                    ),
                ],
            )
            reference_date, _, raw_value = _wait_latest_selic_row(page, timeout_ms)
    except PlaywrightTimeoutError as exc:
        raise PriceParseError("timeout ao buscar SELIC no BCB") from exc
    except PageConsistencyError as exc:
        raise PriceParseError(str(exc)) from exc

    if not raw_value or reference_date is None:
        raise PriceParseError("nao encontrou valor atual da SELIC")

    value = _parse_percent_value(raw_value)
    return InterestRateQuote(
        name="SELIC",
        value=value,
        value_raw=raw_value,
        reference_date=reference_date,
        collected_at=datetime.now(timezone.utc),
    )


def calculate_cdi_daily_percent(
    selic_annual_percent: Decimal,
    *,
    annual_spread: Decimal = _CDI_SPREAD,
) -> Decimal:
    """Converte SELIC anual (%) em CDI diario (%) com base de 252 dias uteis.

    Regra alinhada ao calculo operacional em HP12C:
    FV = 100 + SELIC - 0,10; PV = 100; N = 252; calcula I.
    """
    hundred = Decimal("100")
    future_value = hundred + selic_annual_percent - annual_spread
    if future_value <= 0:
        raise ValueError("valor final invalido para calcular CDI")

    with localcontext() as context:
        context.prec = 34
        daily = (
            ((future_value / hundred) ** (Decimal("1") / _CDI_BUSINESS_DAYS))
            - Decimal("1")
        ) * hundred
    return daily.quantize(_CDI_QUANTIZER, rounding=ROUND_HALF_UP)
