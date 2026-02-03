from __future__ import annotations

from contextlib import contextmanager
import os
from typing import Iterator
from urllib.parse import unquote, urlparse

from playwright.sync_api import Page, sync_playwright


DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/122.0 Safari/537.36"
)
DEFAULT_LOCALE = "pt-BR"
DEFAULT_VIEWPORT = {"width": 1366, "height": 768}
DEFAULT_CHROMIUM_ARGS = ["--disable-blink-features=AutomationControlled"]


def proxy_from_env() -> dict[str, str] | None:
    proxy_url = (
        os.environ.get("HTTPS_PROXY")
        or os.environ.get("https_proxy")
        or os.environ.get("HTTP_PROXY")
        or os.environ.get("http_proxy")
    )
    if not proxy_url:
        return None
    parsed = urlparse(proxy_url)
    if not parsed.scheme or not parsed.hostname or not parsed.port:
        return None
    proxy = {"server": f"{parsed.scheme}://{parsed.hostname}:{parsed.port}"}
    if parsed.username:
        proxy["username"] = unquote(parsed.username)
    if parsed.password:
        proxy["password"] = unquote(parsed.password)
    return proxy


@contextmanager
def chromium_page(
    *,
    headless: bool = True,
    proxy: dict[str, str] | None = None,
    launch_args: list[str] | None = None,
    user_agent: str = DEFAULT_USER_AGENT,
    locale: str = DEFAULT_LOCALE,
    viewport: dict[str, int] | None = None,
) -> Iterator[Page]:
    args = list(DEFAULT_CHROMIUM_ARGS)
    if launch_args:
        args.extend(launch_args)
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            headless=headless,
            args=args,
            proxy=proxy,
        )
        try:
            context = browser.new_context(
                user_agent=user_agent,
                locale=locale,
                viewport=viewport or DEFAULT_VIEWPORT,
            )
            page = context.new_page()
            yield page
        finally:
            browser.close()

