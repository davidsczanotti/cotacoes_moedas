from pathlib import Path

import main
from cotacoes_moedas import network_sync


def test_copy_planilhas_to_network(monkeypatch) -> None:
    planilhas_dir = Path("C:/fake/planilhas")
    destino_base = "C:/fake/network"

    def fake_exists(self: Path) -> bool:
        return self == planilhas_dir

    def fake_is_dir(self: Path) -> bool:
        return self == planilhas_dir

    def fake_mkdir(self: Path, *_, **__) -> None:
        return None

    def fake_try_to_unc(path: str) -> tuple[str, OSError | None]:
        return path, None

    copied: list[tuple[Path, Path, bool]] = []

    def fake_copytree(src: Path, dst: Path, *, dirs_exist_ok: bool = False, **_) -> None:
        copied.append((src, dst, dirs_exist_ok))

    monkeypatch.setattr(Path, "exists", fake_exists)
    monkeypatch.setattr(Path, "is_dir", fake_is_dir)
    monkeypatch.setattr(Path, "mkdir", fake_mkdir)
    monkeypatch.setattr(network_sync, "try_to_unc", fake_try_to_unc)
    monkeypatch.setattr(network_sync.shutil, "copytree", fake_copytree)

    main._copy_planilhas_to_network(planilhas_dir, [destino_base])

    target_dir = Path(destino_base) / "cotacoes"
    assert copied == [
        (planilhas_dir, target_dir / planilhas_dir.name, True),
    ]
