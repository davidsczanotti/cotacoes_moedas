from __future__ import annotations

from decimal import Decimal, InvalidOperation
import re


class ParseDecimalError(ValueError):
    pass


_NON_NUMERIC = re.compile(r"[^\d,.-]")


def parse_pt_br_decimal(text: str) -> Decimal:
    """Parseia numeros no formato brasileiro, aceitando separadores comuns.

    Exemplos aceitos: "5,2849", "5.284,90", "R$ 5,2849".
    """
    cleaned = _NON_NUMERIC.sub("", (text or "")).strip()
    if "," in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    try:
        return Decimal(cleaned)
    except (InvalidOperation, ValueError) as exc:
        raise ParseDecimalError(f"valor invalido: {text!r}") from exc

