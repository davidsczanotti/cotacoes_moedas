from __future__ import annotations

from datetime import datetime
from pathlib import Path

import main


def _all_unfilled() -> dict[str, bool]:
    return {key: False for key in main._SOURCE_REQUIRED_COLUMNS}


def test_select_fetches_morning_window_runs_usd_and_turismo(monkeypatch) -> None:
    now = datetime(2026, 2, 4, 7, 0, 0, tzinfo=main._LOCAL_TZ)
    planilha_path = Path("C:/fake/cotacoes.xlsx")

    called: dict[str, object] = {}

    def fake_read_filled_sources(path: Path, target_date) -> dict[str, bool]:
        called["path"] = path
        called["date"] = target_date
        return _all_unfilled()

    monkeypatch.setattr(main, "_read_filled_sources", fake_read_filled_sources)

    selected_specs, outcomes = main._select_fetches(now, planilha_path)

    assert called["path"] == planilha_path
    assert called["date"] == now.date()
    assert [spec.key for spec in selected_specs] == ["usd_brl", "turismo"]
    assert outcomes["ptax_usd"].skipped
    assert outcomes["ptax_usd"].skip_reason == "fora do horario (antes de 13:10)"
    assert outcomes["ptax_eur"].skip_reason == "fora do horario (antes de 13:10)"
    assert outcomes["ptax_chf"].skip_reason == "fora do horario (antes de 13:10)"


def test_select_fetches_afternoon_window_runs_ptax(monkeypatch) -> None:
    now = datetime(2026, 2, 4, 14, 31, 0, tzinfo=main._LOCAL_TZ)
    planilha_path = Path("C:/fake/cotacoes.xlsx")

    monkeypatch.setattr(main, "_read_filled_sources", lambda *_: _all_unfilled())

    selected_specs, outcomes = main._select_fetches(now, planilha_path)

    assert [spec.key for spec in selected_specs] == [
        "ptax_usd",
        "ptax_eur",
        "ptax_chf",
    ]
    assert outcomes["usd_brl"].skip_reason == "fora do horario (apos 08:30)"
    assert outcomes["turismo"].skip_reason == "fora do horario (apos 08:30)"


def test_select_fetches_between_windows_runs_nothing(monkeypatch) -> None:
    now = datetime(2026, 2, 4, 9, 0, 0, tzinfo=main._LOCAL_TZ)
    planilha_path = Path("C:/fake/cotacoes.xlsx")

    monkeypatch.setattr(main, "_read_filled_sources", lambda *_: _all_unfilled())

    selected_specs, outcomes = main._select_fetches(now, planilha_path)

    assert selected_specs == []
    assert outcomes["usd_brl"].skip_reason == "fora do horario (apos 08:30)"
    assert outcomes["turismo"].skip_reason == "fora do horario (apos 08:30)"
    assert outcomes["ptax_usd"].skip_reason == "fora do horario (antes de 13:10)"
    assert outcomes["ptax_eur"].skip_reason == "fora do horario (antes de 13:10)"
    assert outcomes["ptax_chf"].skip_reason == "fora do horario (antes de 13:10)"


def test_select_fetches_skips_filled_fields(monkeypatch) -> None:
    now = datetime(2026, 2, 4, 14, 31, 0, tzinfo=main._LOCAL_TZ)
    planilha_path = Path("C:/fake/cotacoes.xlsx")

    filled = _all_unfilled()
    filled["ptax_usd"] = True
    filled["ptax_eur"] = True
    filled["ptax_chf"] = False

    monkeypatch.setattr(main, "_read_filled_sources", lambda *_: filled)

    selected_specs, outcomes = main._select_fetches(now, planilha_path)

    assert [spec.key for spec in selected_specs] == ["ptax_chf"]
    assert outcomes["ptax_usd"].skip_reason == "ja preenchido na data de hoje"
    assert outcomes["ptax_eur"].skip_reason == "ja preenchido na data de hoje"
