from __future__ import annotations

import re


_URL_CREDENTIALS_RE = re.compile(r"([a-zA-Z][a-zA-Z0-9+.-]*://)([^@\s/]+)@")
_PASSWORD_PAIR_RE = re.compile(r"(?i)\b(password|passwd|pwd)\b\s*[:=]\s*\S+")


def redact_secrets(text: str) -> str:
    if not text:
        return text

    def _mask(match: re.Match[str]) -> str:
        scheme = match.group(1)
        credentials = match.group(2)
        if ":" in credentials:
            username = credentials.split(":", 1)[0]
            return f"{scheme}{username}:***@"
        return f"{scheme}***@"

    redacted = _URL_CREDENTIALS_RE.sub(_mask, text)
    return _PASSWORD_PAIR_RE.sub(lambda match: f"{match.group(1)}=***", redacted)

