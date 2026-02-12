from __future__ import annotations

import pytest

from cotacoes_moedas.page_consistency import (
    PageCheck,
    PageConsistencyError,
    ensure_page_consistency,
)


class _FakePage:
    def __init__(self, url: str, title: str) -> None:
        self.url = url
        self._title = title

    def title(self) -> str:
        return self._title


def test_ensure_page_consistency_no_failure() -> None:
    page = _FakePage("https://example.com", "Example")
    ensure_page_consistency(
        page,
        source="source",
        checks=[PageCheck("ok", lambda _: (True, ""))],
    )


def test_ensure_page_consistency_reports_failures() -> None:
    page = _FakePage("https://example.com/test", "Example Test")

    with pytest.raises(PageConsistencyError) as exc_info:
        ensure_page_consistency(
            page,
            source="fonte-x",
            checks=[
                PageCheck("check-1", lambda _: (False, "falhou seletor")),
                PageCheck("check-2", lambda _: (False, "falhou tabela")),
            ],
        )

    message = str(exc_info.value)
    assert "fonte-x" in message
    assert "check-1" in message
    assert "falhou seletor" in message
    assert "check-2" in message
    assert "url='https://example.com/test'" in message
