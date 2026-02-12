from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

import main


def _all_unfilled() -> dict[str, bool]:
    return {key: False for key in main._SOURCE_REQUIRED_COLUMNS}


def _make_planilha_path(tmp_path: Path, name: str = "cotacoes.xlsx") -> Path:
    path = tmp_path / name
    path.write_text("", encoding="utf-8")
    return path


def test_select_fetches_morning_window_runs_usd_and_turismo(
    monkeypatch,
    tmp_path: Path,
) -> None:
    now = datetime(2026, 2, 4, 7, 0, 0, tzinfo=main._LOCAL_TZ)
    planilha_path = _make_planilha_path(tmp_path)

    called: dict[str, object] = {}

    def fake_read_filled_sources(path: Path, target_date) -> dict[str, bool]:
        called["path"] = path
        called["date"] = target_date
        return _all_unfilled()

    monkeypatch.setattr(main, "_read_filled_sources", fake_read_filled_sources)

    selected_specs, outcomes = main._select_fetches(now, planilha_path)

    assert called["path"] == planilha_path
    assert called["date"] == now.date()
    assert [spec.key for spec in selected_specs] == [
        "usd_brl",
        "turismo",
        "tjlp",
        "selic",
    ]
    assert outcomes["ptax_usd"].skipped
    assert outcomes["ptax_usd"].skip_reason == "fora do horario (antes de 13:10)"
    assert outcomes["ptax_eur"].skip_reason == "fora do horario (antes de 13:10)"
    assert outcomes["ptax_chf"].skip_reason == "fora do horario (antes de 13:10)"


def test_select_fetches_afternoon_window_runs_ptax(
    monkeypatch,
    tmp_path: Path,
) -> None:
    now = datetime(2026, 2, 4, 14, 31, 0, tzinfo=main._LOCAL_TZ)
    planilha_path = _make_planilha_path(tmp_path)

    monkeypatch.setattr(main, "_read_filled_sources", lambda *_: _all_unfilled())

    selected_specs, outcomes = main._select_fetches(now, planilha_path)

    assert [spec.key for spec in selected_specs] == [
        "ptax_usd",
        "ptax_eur",
        "ptax_chf",
    ]
    assert outcomes["usd_brl"].skip_reason == "fora do horario (apos 08:30)"
    assert outcomes["turismo"].skip_reason == "fora do horario (apos 08:30)"
    assert outcomes["tjlp"].skip_reason == "fora do horario (apos 08:30)"
    assert outcomes["selic"].skip_reason == "fora do horario (apos 08:30)"


def test_select_fetches_between_windows_runs_nothing(
    monkeypatch,
    tmp_path: Path,
) -> None:
    now = datetime(2026, 2, 4, 9, 0, 0, tzinfo=main._LOCAL_TZ)
    planilha_path = _make_planilha_path(tmp_path)

    monkeypatch.setattr(main, "_read_filled_sources", lambda *_: _all_unfilled())

    selected_specs, outcomes = main._select_fetches(now, planilha_path)

    assert selected_specs == []
    assert outcomes["usd_brl"].skip_reason == "fora do horario (apos 08:30)"
    assert outcomes["turismo"].skip_reason == "fora do horario (apos 08:30)"
    assert outcomes["ptax_usd"].skip_reason == "fora do horario (antes de 13:10)"
    assert outcomes["ptax_eur"].skip_reason == "fora do horario (antes de 13:10)"
    assert outcomes["ptax_chf"].skip_reason == "fora do horario (antes de 13:10)"
    assert outcomes["tjlp"].skip_reason == "fora do horario (apos 08:30)"
    assert outcomes["selic"].skip_reason == "fora do horario (apos 08:30)"


def test_select_fetches_skips_filled_fields(
    monkeypatch,
    tmp_path: Path,
) -> None:
    now = datetime(2026, 2, 4, 14, 31, 0, tzinfo=main._LOCAL_TZ)
    planilha_path = _make_planilha_path(tmp_path)

    filled = _all_unfilled()
    filled["ptax_usd"] = True
    filled["ptax_eur"] = True
    filled["ptax_chf"] = False

    monkeypatch.setattr(main, "_read_filled_sources", lambda *_: filled)

    selected_specs, outcomes = main._select_fetches(now, planilha_path)

    assert [spec.key for spec in selected_specs] == ["ptax_chf"]
    assert outcomes["ptax_usd"].skip_reason == "ja preenchido na data de hoje"
    assert outcomes["ptax_eur"].skip_reason == "ja preenchido na data de hoje"
    assert outcomes["tjlp"].skip_reason == "fora do horario (apos 08:30)"
    assert outcomes["selic"].skip_reason == "fora do horario (apos 08:30)"


def test_select_reference_planilha_path_prefers_network_when_exists(
    tmp_path: Path,
) -> None:
    local_planilha_path = _make_planilha_path(tmp_path, "local.xlsx")
    network_base = tmp_path / "X_TEMP_Publico"
    network_xlsx = network_base / "cotacoes" / "planilhas" / "cotacoes.xlsx"
    network_xlsx.parent.mkdir(parents=True, exist_ok=True)
    network_xlsx.write_text("", encoding="utf-8")

    selected = main._select_reference_planilha_path(
        local_planilha_path,
        network_dirs=[str(network_base)],
        network_dest_folder="cotacoes",
    )

    assert selected == network_xlsx


def test_select_reference_planilha_path_returns_network_candidate_when_missing(
    tmp_path: Path,
) -> None:
    local_planilha_path = _make_planilha_path(tmp_path, "local.xlsx")
    network_base = tmp_path / "X_TEMP_Publico"
    expected_candidate = network_base / "cotacoes" / "planilhas" / "cotacoes.xlsx"

    selected = main._select_reference_planilha_path(
        local_planilha_path,
        network_dirs=[str(network_base)],
        network_dest_folder="cotacoes",
    )

    assert selected == expected_candidate
    assert not selected.exists()


def test_select_reference_planilha_path_uses_local_when_no_network_dir(
    tmp_path: Path,
) -> None:
    local_planilha_path = _make_planilha_path(tmp_path, "local.xlsx")

    selected = main._select_reference_planilha_path(
        local_planilha_path,
        network_dirs=[],
        network_dest_folder="cotacoes",
    )

    assert selected == local_planilha_path


def test_main_bootstraps_network_reference_planilha_when_missing(
    monkeypatch,
    tmp_path: Path,
) -> None:
    base_dir = tmp_path
    network_base = tmp_path / "network_root"
    now = datetime(2026, 2, 4, 7, 0, 0, tzinfo=main._LOCAL_TZ)

    local_planilhas = base_dir / "planilhas"
    local_planilhas.mkdir(parents=True, exist_ok=True)
    (local_planilhas / "cotacoes.xlsx").write_text("", encoding="utf-8")
    (local_planilhas / "cotacoes.csv").write_text("", encoding="utf-8")

    captured: dict[str, object] = {}

    def fake_select_fetches(*_args, **_kwargs):
        captured["reference"] = _kwargs.get("reference_planilha_path")
        return [], {}

    monkeypatch.setattr(main, "_resolve_base_dir", lambda: base_dir)
    monkeypatch.setattr(main, "_configure_playwright", lambda *_: None)
    monkeypatch.setattr(main, "parse_network_dirs", lambda *_: [str(network_base)])
    monkeypatch.setattr(main, "_now_local", lambda: now)
    monkeypatch.setattr(main, "_select_fetches", fake_select_fetches)
    monkeypatch.setattr(main, "normalize_xlsx_layout", lambda *_: None)
    monkeypatch.setattr(main, "_log", lambda *_: None)

    exit_code = main.main()

    expected_reference = network_base / "cotacoes" / "planilhas" / "cotacoes.xlsx"
    assert exit_code == 0
    assert expected_reference.exists()
    assert captured["reference"] == expected_reference


def test_main_bootstraps_network_reference_when_missing_between_windows(
    monkeypatch,
    tmp_path: Path,
) -> None:
    base_dir = tmp_path
    network_base = tmp_path / "network_root"
    now = datetime(2026, 2, 4, 9, 37, 0, tzinfo=main._LOCAL_TZ)

    local_planilhas = base_dir / "planilhas"
    local_planilhas.mkdir(parents=True, exist_ok=True)
    (local_planilhas / "cotacoes.xlsx").write_text("", encoding="utf-8")
    (local_planilhas / "cotacoes.csv").write_text("", encoding="utf-8")

    captured: dict[str, object] = {}

    def fake_select_fetches(*_args, **_kwargs):
        captured["reference"] = _kwargs.get("reference_planilha_path")
        return [], {}

    monkeypatch.setattr(main, "_resolve_base_dir", lambda: base_dir)
    monkeypatch.setattr(main, "_configure_playwright", lambda *_: None)
    monkeypatch.setattr(main, "parse_network_dirs", lambda *_: [str(network_base)])
    monkeypatch.setattr(main, "_now_local", lambda: now)
    monkeypatch.setattr(main, "_select_fetches", fake_select_fetches)
    monkeypatch.setattr(main, "normalize_xlsx_layout", lambda *_: None)
    monkeypatch.setattr(main, "_log", lambda *_: None)

    exit_code = main.main()

    expected_reference = network_base / "cotacoes" / "planilhas" / "cotacoes.xlsx"
    assert exit_code == 0
    assert expected_reference.exists()
    assert captured["reference"] == expected_reference


def test_main_aborts_when_network_reference_planilha_missing_and_bootstrap_fails(
    monkeypatch,
    tmp_path: Path,
) -> None:
    base_dir = tmp_path
    network_base = tmp_path / "network_root"
    now = datetime(2026, 2, 4, 7, 0, 0, tzinfo=main._LOCAL_TZ)

    monkeypatch.setattr(main, "_resolve_base_dir", lambda: base_dir)
    monkeypatch.setattr(main, "_configure_playwright", lambda *_: None)
    monkeypatch.setattr(main, "parse_network_dirs", lambda *_: [str(network_base)])
    monkeypatch.setattr(main, "_now_local", lambda: now)
    monkeypatch.setattr(main, "normalize_xlsx_layout", lambda *_: None)
    monkeypatch.setattr(main, "_log", lambda *_: None)

    exit_code = main.main()

    assert exit_code == 1


def test_validate_planilha_row_consistency_detects_missing_expected_fields(
    tmp_path: Path,
) -> None:
    planilha_path = tmp_path / "cotacoes.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "Data"
    sheet["A3"] = datetime(2026, 2, 12)
    sheet["M3"] = 0.15
    sheet["N3"] = 0.0551310642
    workbook.save(planilha_path)
    workbook.close()

    outcomes = {
        "usd_brl": main.FetchOutcome(
            label="USD/BRL (Investing)",
            value=None,
            error=None,
            elapsed_s=0.0,
            skipped=True,
            skip_reason="ja preenchido na data de hoje",
        ),
        "ptax_usd": main.FetchOutcome(
            label="PTAX USD",
            value=None,
            error=None,
            elapsed_s=0.0,
            skipped=True,
            skip_reason="fora do horario (antes de 13:10)",
        ),
        "turismo": main.FetchOutcome(
            label="Dolar Turismo (Valor)",
            value=None,
            error=None,
            elapsed_s=0.0,
            skipped=True,
            skip_reason="fora do horario (apos 08:30)",
        ),
        "ptax_eur": main.FetchOutcome(
            label="PTAX EUR",
            value=None,
            error=None,
            elapsed_s=0.0,
            skipped=True,
            skip_reason="fora do horario (antes de 13:10)",
        ),
        "ptax_chf": main.FetchOutcome(
            label="PTAX CHF",
            value=None,
            error=None,
            elapsed_s=0.0,
            skipped=True,
            skip_reason="fora do horario (antes de 13:10)",
        ),
        "tjlp": main.FetchOutcome(
            label="TJLP (BNDES)",
            value=None,
            error=None,
            elapsed_s=0.0,
            skipped=True,
            skip_reason="fora do horario (apos 08:30)",
        ),
        "selic": main.FetchOutcome(
            label="SELIC (BCB)",
            value=None,
            error=None,
            elapsed_s=0.0,
            skipped=True,
            skip_reason="ja preenchido na data de hoje",
        ),
    }

    issues = main._validate_planilha_row_consistency(
        planilha_path,
        target_date=datetime(2026, 2, 12).date(),
        outcomes=outcomes,
    )

    assert any("USD/BRL (Investing)" in issue for issue in issues)


def test_main_aborts_when_bootstrap_layout_formatting_fails(
    monkeypatch,
    tmp_path: Path,
) -> None:
    base_dir = tmp_path
    network_base = tmp_path / "network_root"
    now = datetime(2026, 2, 4, 9, 37, 0, tzinfo=main._LOCAL_TZ)

    local_planilhas = base_dir / "planilhas"
    local_planilhas.mkdir(parents=True, exist_ok=True)
    (local_planilhas / "cotacoes.xlsx").write_text("", encoding="utf-8")
    (local_planilhas / "cotacoes.csv").write_text("", encoding="utf-8")

    monkeypatch.setattr(main, "_resolve_base_dir", lambda: base_dir)
    monkeypatch.setattr(main, "_configure_playwright", lambda *_: None)
    monkeypatch.setattr(main, "parse_network_dirs", lambda *_: [str(network_base)])
    monkeypatch.setattr(main, "_now_local", lambda: now)
    monkeypatch.setattr(
        main,
        "normalize_xlsx_layout",
        lambda *_: (_ for _ in ()).throw(RuntimeError("falha fake")),
    )
    monkeypatch.setattr(main, "_log", lambda *_: None)

    exit_code = main.main()

    assert exit_code == 1


def test_main_syncs_network_even_when_no_sources_selected(
    monkeypatch,
    tmp_path: Path,
) -> None:
    base_dir = tmp_path
    now = datetime(2026, 2, 4, 9, 37, 0, tzinfo=main._LOCAL_TZ)
    network_base = tmp_path / "network_root"

    local_planilhas = base_dir / "planilhas"
    local_planilhas.mkdir(parents=True, exist_ok=True)
    local_planilha_path = local_planilhas / "cotacoes.xlsx"

    workbook = Workbook()
    workbook.save(local_planilha_path)
    workbook.close()
    (local_planilhas / "cotacoes.csv").write_text("", encoding="utf-8")

    copied: dict[str, object] = {}

    def fake_copy(*_args, **_kwargs):
        copied["called"] = True
        destination = network_base / "cotacoes" / "planilhas"
        destination.mkdir(parents=True, exist_ok=True)
        return destination

    outcomes = {
        key: main.FetchOutcome(
            label=main._SOURCE_LABELS[key],
            value=None,
            error=None,
            elapsed_s=0.0,
            skipped=True,
            skip_reason="fora do horario (teste)",
        )
        for key in main._SOURCE_REQUIRED_COLUMNS
    }

    monkeypatch.setattr(main, "_resolve_base_dir", lambda: base_dir)
    monkeypatch.setattr(main, "_configure_playwright", lambda *_: None)
    monkeypatch.setattr(main, "parse_network_dirs", lambda *_: [str(network_base)])
    monkeypatch.setattr(main, "_now_local", lambda: now)
    monkeypatch.setattr(main, "normalize_xlsx_layout", lambda *_: None)
    monkeypatch.setattr(
        main,
        "_select_reference_planilha_path",
        lambda *_args, **_kwargs: local_planilha_path,
    )
    monkeypatch.setattr(
        main,
        "_sync_local_planilhas_from_reference",
        lambda *_args, **_kwargs: True,
    )
    monkeypatch.setattr(main, "_select_fetches", lambda *_args, **_kwargs: ([], outcomes))
    monkeypatch.setattr(main, "_copy_planilhas_to_network", fake_copy)
    monkeypatch.setattr(main, "_log", lambda *_: None)

    exit_code = main.main()

    assert exit_code == 0
    assert copied["called"] is True
