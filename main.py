from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import date, datetime, timezone
from decimal import Decimal
import os
from pathlib import Path
import sys
import time
from typing import Callable

from openpyxl import load_workbook

from cotacoes_moedas import (
    fetch_chf_ptax,
    fetch_dolar_ptax,
    fetch_dolar_turismo,
    fetch_euro_ptax,
    fetch_usd_brl,
    update_csv_from_xlsx,
    update_xlsx_quotes_and_log,
)
from cotacoes_moedas.network_sync import (
    copiar_pasta_para_rede,
    parse_network_dirs,
)
from cotacoes_moedas.redaction import redact_secrets

_USD_SPREAD = Decimal("0.0200")
_LOCAL_TZ = datetime.now().astimezone().tzinfo or timezone.utc
_DEFAULT_NETWORK_COPY_DIR = r"X:\TEMP\_Publico;\\192.168.21.25\users\TEMP\_Publico"
_DEFAULT_NETWORK_DEST_FOLDER = "cotacoes"
_MORNING_QUOTES_CUTOFF_HM = (8, 30)
_PTAX_AVAILABLE_FROM_HM = (13, 10)
_SOURCE_REQUIRED_COLUMNS: dict[str, tuple[str, ...]] = {
    "usd_brl": ("B", "C"),
    "ptax_usd": ("D", "E"),
    "turismo": ("F", "G"),
    "ptax_eur": ("H", "I"),
    "ptax_chf": ("J", "K"),
}
_SOURCE_LABELS: dict[str, str] = {
    "usd_brl": "USD/BRL (Investing)",
    "ptax_usd": "PTAX USD",
    "ptax_eur": "PTAX EUR",
    "ptax_chf": "PTAX CHF",
    "turismo": "Dolar Turismo (Valor)",
}


@dataclass(frozen=True)
class FetchSpec:
    key: str
    label: str
    fetch_fn: Callable[[], object]


@dataclass
class FetchOutcome:
    label: str
    value: object | None
    error: str | None
    elapsed_s: float
    skipped: bool = False
    skip_reason: str | None = None


def _now_local() -> datetime:
    return datetime.now(_LOCAL_TZ)


def _log(message: str) -> None:
    timestamp = _now_local().strftime("%H:%M:%S")
    print(f"[{timestamp}] {message}", flush=True)


def _log_stage(step: int, total: int, message: str) -> None:
    _log(f"Etapa {step}/{total} - {message}")


def _format_duration(seconds: float) -> str:
    total_seconds = int(round(seconds))
    minutes, remainder = divmod(total_seconds, 60)
    return f"{minutes:02d}:{remainder:02d}"


def _resolve_base_dir() -> Path:
    return Path(sys.argv[0]).resolve().parent


def _configure_playwright(base_dir: Path) -> None:
    if os.environ.get("PLAYWRIGHT_BROWSERS_PATH"):
        return
    browsers_path = base_dir / "ms-playwright"
    if browsers_path.exists():
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(browsers_path)
        _log(f"Playwright browsers: {browsers_path}")


def _error_detail(label: str, exc: Exception) -> str:
    message = " ".join(str(exc).split())
    message = redact_secrets(message)
    if message:
        return f"{label}: {exc.__class__.__name__} {message}"
    return f"{label}: {exc.__class__.__name__}"


def _run_fetch(
    label: str,
    fetch_fn: Callable[[], object],
) -> tuple[object | None, str | None, float]:
    _log(f"Coletando {label}...")
    start = time.monotonic()
    try:
        result = fetch_fn()
    except Exception as exc:
        detail = _error_detail(label, exc)
        _log(f"Falha em {label}. Valores nao atualizados. Detalhe: {detail}")
        return None, detail, time.monotonic() - start
    elapsed = time.monotonic() - start
    _log(f"{label} OK em {elapsed:.1f}s")
    return result, None, elapsed


def _run_fetches(fetch_specs: list[FetchSpec]) -> dict[str, FetchOutcome]:
    outcomes: dict[str, FetchOutcome] = {}
    if not fetch_specs:
        return outcomes

    max_workers = len(fetch_specs)
    env_max_workers = os.environ.get("COTACOES_MAX_WORKERS")
    if env_max_workers:
        try:
            max_workers = max(1, min(max_workers, int(env_max_workers)))
        except ValueError:
            _log(
                "Aviso: COTACOES_MAX_WORKERS invalido "
                f"({env_max_workers!r}). Usando {max_workers}."
            )

    if max_workers <= 1 or len(fetch_specs) == 1:
        _log(f"Coleta: {len(fetch_specs)} fonte(s) em modo sequencial.")
        for spec in fetch_specs:
            value, error, elapsed = _run_fetch(spec.label, spec.fetch_fn)
            outcomes[spec.key] = FetchOutcome(
                label=spec.label,
                value=value,
                error=error,
                elapsed_s=elapsed,
            )
        return outcomes

    limit_note = (
        f" (limitado por COTACOES_MAX_WORKERS={env_max_workers})"
        if env_max_workers
        else ""
    )
    _log(
        "Coleta: "
        f"{len(fetch_specs)} fonte(s) em paralelo ({max_workers} workers){limit_note}."
    )
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {
            executor.submit(_run_fetch, spec.label, spec.fetch_fn): spec
            for spec in fetch_specs
        }
        for future in as_completed(futures):
            spec = futures[future]
            try:
                value, error, elapsed = future.result()
            except Exception as exc:
                detail = _error_detail(spec.label, exc)
                _log(
                    f"Falha em {spec.label}. Valores nao atualizados. "
                    f"Detalhe: {detail}"
                )
                value, error, elapsed = None, detail, 0.0
            outcomes[spec.key] = FetchOutcome(
                label=spec.label,
                value=value,
                error=error,
                elapsed_s=elapsed,
            )
    return outcomes


def _collect_errors(outcomes: dict[str, FetchOutcome]) -> list[str]:
    errors: list[str] = []
    for outcome in outcomes.values():
        if outcome.error:
            errors.append(outcome.error)
    return errors


def _log_fetch_summary(outcomes: dict[str, FetchOutcome]) -> None:
    total = len(outcomes)
    skipped = sum(1 for outcome in outcomes.values() if outcome.skipped)
    attempted = total - skipped
    ok_count = sum(
        1
        for outcome in outcomes.values()
        if not outcome.skipped and not outcome.error
    )
    fail_count = sum(
        1
        for outcome in outcomes.values()
        if not outcome.skipped and outcome.error
    )
    _log(
        "Coleta concluida: "
        f"{ok_count}/{attempted} fontes OK, "
        f"{fail_count} com erro, "
        f"{skipped} puladas."
    )


def _log_fetch_plan(
    selected_specs: list[FetchSpec],
    outcomes: dict[str, FetchOutcome],
) -> None:
    if selected_specs:
        labels = "; ".join(spec.label for spec in selected_specs)
        _log(f"Fontes selecionadas para coleta ({len(selected_specs)}): {labels}")
    else:
        _log("Nenhuma fonte selecionada para coleta agora.")

    for key in ("usd_brl", "turismo", "ptax_usd", "ptax_eur", "ptax_chf"):
        outcome = outcomes.get(key)
        if outcome and outcome.skipped:
            _log(f"{outcome.label}: pulado ({outcome.skip_reason})")


def _hm(value: datetime) -> tuple[int, int]:
    return value.hour, value.minute


def _coerce_date(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
    return None


def _find_row_by_date(sheet, target_date: date) -> int | None:
    for row in range(3, sheet.max_row + 1):
        cell_date = _coerce_date(sheet.cell(row=row, column=1).value)
        if cell_date == target_date:
            return row
    return None


def _is_source_filled(sheet, row: int, columns: tuple[str, ...]) -> bool:
    for col in columns:
        value = sheet[f"{col}{row}"].value
        if value is None:
            return False
        if isinstance(value, str) and not value.strip():
            return False
    return True


def _read_filled_sources(planilha_path: Path, target_date: date) -> dict[str, bool]:
    workbook = load_workbook(planilha_path, data_only=True)
    try:
        sheet = workbook.active
        row = _find_row_by_date(sheet, target_date)
        filled: dict[str, bool] = {}
        for key, columns in _SOURCE_REQUIRED_COLUMNS.items():
            filled[key] = (
                row is not None and _is_source_filled(sheet, row, columns)
            )
        return filled
    finally:
        close = getattr(workbook, "close", None)
        if callable(close):
            close()


def _skip_outcome(key: str, reason: str) -> FetchOutcome:
    return FetchOutcome(
        label=_SOURCE_LABELS[key],
        value=None,
        error=None,
        elapsed_s=0.0,
        skipped=True,
        skip_reason=reason,
    )


def _select_fetches(
    now: datetime,
    planilha_path: Path,
) -> tuple[list[FetchSpec], dict[str, FetchOutcome]]:
    today = now.date()
    filled = _read_filled_sources(planilha_path, today)
    allow_morning_quotes = _hm(now) <= _MORNING_QUOTES_CUTOFF_HM
    allow_ptax = _hm(now) >= _PTAX_AVAILABLE_FROM_HM

    all_specs = {
        "usd_brl": FetchSpec(
            key="usd_brl",
            label=_SOURCE_LABELS["usd_brl"],
            fetch_fn=fetch_usd_brl,
        ),
        "ptax_usd": FetchSpec(
            key="ptax_usd",
            label=_SOURCE_LABELS["ptax_usd"],
            fetch_fn=fetch_dolar_ptax,
        ),
        "ptax_eur": FetchSpec(
            key="ptax_eur",
            label=_SOURCE_LABELS["ptax_eur"],
            fetch_fn=fetch_euro_ptax,
        ),
        "ptax_chf": FetchSpec(
            key="ptax_chf",
            label=_SOURCE_LABELS["ptax_chf"],
            fetch_fn=fetch_chf_ptax,
        ),
        "turismo": FetchSpec(
            key="turismo",
            label=_SOURCE_LABELS["turismo"],
            fetch_fn=fetch_dolar_turismo,
        ),
    }

    selected: list[FetchSpec] = []
    outcomes: dict[str, FetchOutcome] = {}

    for key in ("usd_brl", "turismo"):
        if not allow_morning_quotes:
            outcomes[key] = _skip_outcome(key, "fora do horario (apos 08:30)")
        elif filled.get(key, False):
            outcomes[key] = _skip_outcome(key, "ja preenchido na data de hoje")
        else:
            selected.append(all_specs[key])

    for key in ("ptax_usd", "ptax_eur", "ptax_chf"):
        if not allow_ptax:
            outcomes[key] = _skip_outcome(key, "fora do horario (antes de 13:10)")
        elif filled.get(key, False):
            outcomes[key] = _skip_outcome(key, "ja preenchido na data de hoje")
        else:
            selected.append(all_specs[key])

    return selected, outcomes


def _update_planilha(
    planilha_path: Path,
    target_date: date,
    outcomes: dict[str, FetchOutcome],
    errors: list[str],
) -> dict[str, tuple[str, ...]]:
    _log(f"Atualizando planilha: {planilha_path} (gravacao unica)")

    status = "ERRO" if errors else "OK"
    detail = " | ".join(errors) if errors else None

    written = update_xlsx_quotes_and_log(
        planilha_path,
        target_date=target_date,
        usd_brl=outcomes["usd_brl"].value,
        ptax_usd=outcomes["ptax_usd"].value,
        ptax_eur=outcomes["ptax_eur"].value,
        ptax_chf=outcomes["ptax_chf"].value,
        turismo=outcomes["turismo"].value,
        spread=_USD_SPREAD,
        overwrite_quotes=False,
        logged_at=_now_local(),
        status=status,
        detail=detail,
    )

    def _describe_fields(fields: tuple[str, ...]) -> str:
        if not fields:
            return "nao gravou (ja preenchido na planilha)"
        if len(fields) == 2:
            return "gravou compra e venda"
        return "gravou " + " e ".join(fields)

    for key in ("usd_brl", "turismo", "ptax_usd", "ptax_eur", "ptax_chf"):
        outcome = outcomes[key]
        if outcome.skipped:
            _log(f"{outcome.label} pulado: {outcome.skip_reason}")
            continue
        if outcome.value is None:
            _log(f"{outcome.label}: sem dados; planilha nao atualizada para esta fonte.")
            continue
        _log(f"{outcome.label}: {_describe_fields(written.get(key, ()))}.")

    has_quotes = any(outcomes[key].value is not None for key in written)
    wrote_any = any(fields for fields in written.values())
    if has_quotes and not wrote_any:
        _log(
            "Nenhuma cotacao foi gravada (valores ja estavam preenchidos). "
            "Apenas o log foi atualizado."
        )

    return written


def _log_quote_summary(outcomes: dict[str, FetchOutcome]) -> None:
    usd = outcomes["usd_brl"]
    quote = usd.value
    if quote:
        _log(
            "USD/BRL: "
            f"{quote.value} ({quote.value_raw}) em {quote.collected_at.astimezone(_LOCAL_TZ)}"
        )
    elif usd.skipped:
        _log(f"USD/BRL: pulado ({usd.skip_reason})")
    else:
        _log("USD/BRL: sem dados")

    usd_ptax_outcome = outcomes["ptax_usd"]
    ptax = usd_ptax_outcome.value
    if ptax:
        _log(
            "PTAX USD: "
            f"{ptax.buy} / {ptax.sell} em {ptax.collected_at.astimezone(_LOCAL_TZ)}"
        )
    elif usd_ptax_outcome.skipped:
        _log(f"PTAX USD: pulado ({usd_ptax_outcome.skip_reason})")
    else:
        _log("PTAX USD: sem dados")

    eur_ptax_outcome = outcomes["ptax_eur"]
    euro = eur_ptax_outcome.value
    if euro:
        _log(
            "PTAX EUR: "
            f"{euro.buy} / {euro.sell} em {euro.collected_at.astimezone(_LOCAL_TZ)}"
        )
    elif eur_ptax_outcome.skipped:
        _log(f"PTAX EUR: pulado ({eur_ptax_outcome.skip_reason})")
    else:
        _log("PTAX EUR: sem dados")

    chf_ptax_outcome = outcomes["ptax_chf"]
    chf = chf_ptax_outcome.value
    if chf:
        _log(
            "PTAX CHF: "
            f"{chf.buy} / {chf.sell} em {chf.collected_at.astimezone(_LOCAL_TZ)}"
        )
    elif chf_ptax_outcome.skipped:
        _log(f"PTAX CHF: pulado ({chf_ptax_outcome.skip_reason})")
    else:
        _log("PTAX CHF: sem dados")

    turismo_outcome = outcomes["turismo"]
    turismo = turismo_outcome.value
    if turismo:
        _log(
            "Dolar Turismo: "
            f"{turismo.buy} / {turismo.sell} em {turismo.collected_at.astimezone(_LOCAL_TZ)}"
        )
    elif turismo_outcome.skipped:
        _log(f"Dolar Turismo: pulado ({turismo_outcome.skip_reason})")
    else:
        _log("Dolar Turismo: sem dados")


def _copy_planilhas_to_network(
    planilhas_dir: Path,
    network_dirs: list[str],
    *,
    network_dest_folder: str,
) -> None:
    if not planilhas_dir.exists() or not planilhas_dir.is_dir():
        _log(f"Pasta de planilhas nao encontrada: {planilhas_dir}")
        return

    _log("Copiando pasta de planilhas para a rede...")
    destination_dir, unc_error, copy_error = copiar_pasta_para_rede(
        planilhas_dir,
        network_dirs,
        nome_pasta_destino=network_dest_folder,
    )
    if not destination_dir:
        if copy_error:
            detail = redact_secrets(str(copy_error))
            _log(
                "Copia em rede falhou: "
                f"{copy_error.__class__.__name__} {detail}"
            )
        else:
            _log(
                "Copia em rede: nenhum destino valido informado: "
                + "; ".join(network_dirs)
            )
        return

    if unc_error:
        detail = redact_secrets(str(unc_error))
        _log(
            "Aviso: Nao foi possivel converter para UNC, usando caminho original: "
            f"{detail}"
        )

    _log(f"Pasta '{planilhas_dir.name}' copiada para: {destination_dir}")


def main() -> int:
    process_start = time.monotonic()
    total_steps = 6

    try:
        _log("Inicio da coleta de cotacoes.")
        base_dir = _resolve_base_dir()

        _log_stage(1, total_steps, "Preparando ambiente e configuracoes.")
        _log(f"Diretorio base: {base_dir}")
        _configure_playwright(base_dir)
        network_copy_dirs = parse_network_dirs(
            os.environ.get("COTACOES_NETWORK_DIR") or _DEFAULT_NETWORK_COPY_DIR
        )
        planilha_path = base_dir / "planilhas" / "cotacoes.xlsx"
        csv_path = base_dir / "planilhas" / "cotacoes.csv"
        _log(f"Planilha principal: {planilha_path}")
        _log(f"Arquivo CSV: {csv_path}")
        if network_copy_dirs:
            _log("Destinos de copia em rede: " + "; ".join(network_copy_dirs))
        else:
            _log("Destinos de copia em rede: nenhum")
        network_dest_folder = (
            os.environ.get("COTACOES_NETWORK_DEST_FOLDER")
            or _DEFAULT_NETWORK_DEST_FOLDER
        ).strip()
        if not network_dest_folder:
            network_dest_folder = _DEFAULT_NETWORK_DEST_FOLDER
        _log(f"Pasta de destino na rede: {network_dest_folder}")

        _log_stage(2, total_steps, "Coletando cotacoes das fontes.")
        now = _now_local()
        now_hm = _hm(now)
        _log(
            "Horario local atual: "
            f"{now.strftime('%H:%M')} "
            f"(Investing/Valor ate {(_MORNING_QUOTES_CUTOFF_HM[0]):02d}:{(_MORNING_QUOTES_CUTOFF_HM[1]):02d}; "
            f"PTAX apos {(_PTAX_AVAILABLE_FROM_HM[0]):02d}:{(_PTAX_AVAILABLE_FROM_HM[1]):02d})."
        )
        if now_hm > _MORNING_QUOTES_CUTOFF_HM and now_hm < _PTAX_AVAILABLE_FROM_HM:
            _log(
                "Fora da janela de coleta no momento "
                "(apos 08:30 e antes de 13:10). Encerrando sem alteracoes."
            )
            duration = _format_duration(time.monotonic() - process_start)
            _log(f"Processo finalizado em {duration} (minutos:segundos).")
            return 0
        _log(
            "Validando planilha para decidir quais fontes coletar "
            "(nao sobrescreve valores ja preenchidos no dia)."
        )
        selected_specs, outcomes = _select_fetches(now, planilha_path)
        _log_fetch_plan(selected_specs, outcomes)
        if not selected_specs:
            _log(
                "Nenhuma fonte elegivel para atualizar agora "
                "(fora da janela de horario ou ja preenchida). Encerrando sem alteracoes."
            )
            duration = _format_duration(time.monotonic() - process_start)
            _log(f"Processo finalizado em {duration} (minutos:segundos).")
            return 0

        outcomes.update(_run_fetches(selected_specs))
        errors = _collect_errors(outcomes)
        _log_fetch_summary(outcomes)

        _log_stage(3, total_steps, "Atualizando planilha Excel.")
        _update_planilha(planilha_path, now.date(), outcomes, errors)

        _log_stage(4, total_steps, "Atualizando CSV.")
        update_csv_from_xlsx(planilha_path, csv_path)

        _log_stage(5, total_steps, "Resumo das cotacoes coletadas.")
        _log_quote_summary(outcomes)
        if errors:
            _log(f"Falhas na coleta: {len(errors)}. Consulte o log da planilha.")
        else:
            _log("Coleta sem falhas.")

        _log_stage(
            6,
            total_steps,
            f"Validando pasta '{network_dest_folder}' na rede e copiando planilhas.",
        )
        if not network_copy_dirs:
            _log("Copia em rede desabilitada.")
        else:
            _copy_planilhas_to_network(
                base_dir / "planilhas",
                network_copy_dirs,
                network_dest_folder=network_dest_folder,
            )
        duration = _format_duration(time.monotonic() - process_start)
        _log(f"Processo finalizado em {duration} (minutos:segundos).")
        return 0
    except KeyboardInterrupt:
        _log("Execucao interrompida pelo usuario (Ctrl+C).")
        duration = _format_duration(time.monotonic() - process_start)
        _log(f"Processo interrompido em {duration} (minutos:segundos).")
        return 130
    except PermissionError as exc:
        detail = redact_secrets(str(exc))
        _log(
            "ERRO ao gravar arquivos. "
            "Possivel causa: planilha/CSV abertos no Excel ou permissao insuficiente."
        )
        if detail:
            _log(f"Detalhe: {exc.__class__.__name__} {detail}")
        else:
            _log(f"Detalhe: {exc.__class__.__name__}")
        duration = _format_duration(time.monotonic() - process_start)
        _log(f"Processo abortado em {duration} (minutos:segundos).")
        return 1
    except Exception as exc:
        detail = _error_detail("Erro inesperado", exc)
        _log(f"ERRO: {detail}")
        duration = _format_duration(time.monotonic() - process_start)
        _log(f"Processo abortado em {duration} (minutos:segundos).")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
