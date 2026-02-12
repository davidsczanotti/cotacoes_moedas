from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

from playwright.sync_api import Page


class PageConsistencyError(RuntimeError):
    pass


@dataclass(frozen=True)
class PageCheck:
    name: str
    validate: Callable[[Page], tuple[bool, str]]


def describe_page(page: Page) -> str:
    url = " ".join((page.url or "").split())
    title = ""
    try:
        title = " ".join((page.title() or "").split())
    except Exception:
        title = ""
    if title:
        return f"url={url!r}; title={title!r}"
    return f"url={url!r}"


def ensure_page_consistency(
    page: Page,
    *,
    source: str,
    checks: list[PageCheck],
) -> None:
    failures: list[str] = []
    for check in checks:
        try:
            ok, detail = check.validate(page)
        except Exception as exc:
            failures.append(
                f"{check.name}: excecao {exc.__class__.__name__} {exc}"
            )
            continue
        if not ok:
            failures.append(f"{check.name}: {detail}")

    if failures:
        detail = " | ".join(failures)
        raise PageConsistencyError(
            f"estrutura da pagina possivelmente alterada em {source}; "
            f"falhas: {detail}; {describe_page(page)}"
        )
