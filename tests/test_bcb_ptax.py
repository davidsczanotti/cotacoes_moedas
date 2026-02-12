from __future__ import annotations

import pytest

from cotacoes_moedas.bcb_ptax import (
    PriceParseError,
    _find_ptax_frame,
    _load_ptax_frame,
)


class _FakeLocator:
    def __init__(self, count: int) -> None:
        self._count = count

    def count(self) -> int:
        return self._count


class _FakeFrame:
    def __init__(self, url: str, selectors: dict[str, int] | None = None) -> None:
        self.url = url
        self._selectors = selectors or {}

    def locator(self, selector: str) -> _FakeLocator:
        return _FakeLocator(self._selectors.get(selector, 0))


class _FakePage:
    def __init__(self, frames: list[_FakeFrame], iframe_sources: list[str]) -> None:
        self.frames = frames
        self._iframe_sources = iframe_sources
        self.url = "https://www.bcb.gov.br/estabilidadefinanceira/historicocotacoes"

    def title(self) -> str:
        return "Banco Central do Brasil"

    def eval_on_selector_all(self, _selector: str, _script: str) -> list[str]:
        return self._iframe_sources


def test_find_ptax_frame_prefers_frame_url_signature() -> None:
    frame = _FakeFrame("https://ptax.bcb.gov.br/ptax_internet/consultaBoletim.do")
    page = _FakePage(frames=[_FakeFrame("https://example.com"), frame], iframe_sources=[])

    selected = _find_ptax_frame(page)

    assert selected is frame


def test_find_ptax_frame_falls_back_to_expected_fields() -> None:
    frame = _FakeFrame(
        "https://example.com/frame",
        selectors={
            'input[name="DATAINI"]': 1,
            'select[name="ChkMoeda"]': 1,
        },
    )
    page = _FakePage(frames=[_FakeFrame("https://example.com/other"), frame], iframe_sources=[])

    selected = _find_ptax_frame(page)

    assert selected is frame


def test_load_ptax_frame_reports_iframe_sources_when_missing() -> None:
    page = _FakePage(
        frames=[_FakeFrame("https://example.com")],
        iframe_sources=["https://example.com/iframe-x"],
    )

    with pytest.raises(PriceParseError) as exc_info:
        _load_ptax_frame(page, timeout_ms=10)

    message = str(exc_info.value)
    assert "iframe PTAX nao encontrado/carregado" in message
    assert "iframe-x" in message
