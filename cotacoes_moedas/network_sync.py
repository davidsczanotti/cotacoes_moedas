from __future__ import annotations

from pathlib import Path
import shutil

from .network_copy import try_to_unc


def parse_network_dirs(value: str | None) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in value.split(";") if item.strip()]


def copiar_pasta_para_rede(
    origem: str | Path,
    destinos_base: list[str],
    nome_pasta_destino: str = "cotacoes",
) -> tuple[Path | None, OSError | None, Exception | None]:
    """Copia uma pasta para um destino de rede.

    Repete a logica do script `teste_copia.py`:
    - tenta converter drive mapeado -> UNC (Windows)
    - cria a pasta `nome_pasta_destino`
    - copia a pasta de origem para dentro dela

    Estrutura final:
    `<destino_base>/<nome_pasta_destino>/<nome_da_pasta_origem>`
    """
    source = Path(origem)
    if not source.exists() or not source.is_dir():
        return None, None, FileNotFoundError(f"Pasta de origem '{source}' nao encontrada")

    last_error: Exception | None = None
    last_unc_error: OSError | None = None
    has_candidate = False

    for destino_base in destinos_base:
        destino_base = (destino_base or "").strip()
        if not destino_base:
            continue
        has_candidate = True
        destino_base_unc, unc_error = try_to_unc(destino_base)
        destino_completo = Path(destino_base_unc) / nome_pasta_destino / source.name
        try:
            destino_completo.parent.mkdir(parents=True, exist_ok=True)
            shutil.copytree(source, destino_completo, dirs_exist_ok=True)
        except Exception as exc:
            last_error = exc
            last_unc_error = unc_error
            continue
        return destino_completo, unc_error, None

    if not has_candidate:
        return None, None, ValueError("Nenhum destino valido informado para copia em rede.")
    return None, last_unc_error, last_error
