"""Script manual para copiar planilhas para a rede (nao usado pelo app)."""

from pathlib import Path
import shutil


from cotacoes_moedas.network_copy import try_to_unc

def copiar_pasta_para_rede(origem: str, destino_base: str, nome_pasta_destino: str = "cotacoes"):
    """
    Copia uma pasta para um destino de rede, criando a pasta de destino se não existir.

    Args:
        origem: Caminho da pasta de origem (ex: "planilhas")
        destino_base: Caminho base de destino (ex: "X:\\TEMP\\_Publico" ou "\\\\servidor\\compartilhamento")
        nome_pasta_destino: Nome da pasta a ser criada no destino (padrão: "cotacoes")
    """
    destino_base_unc, unc_error = try_to_unc(destino_base)
    if unc_error:
        print(
            "Aviso: Não foi possível converter para UNC, usando caminho original: "
            f"{unc_error}"
        )
        destino_base_unc = destino_base

    # Criar caminho completo de destino
    destino_completo = Path(destino_base_unc) / nome_pasta_destino / Path(origem).name

    # Criar pasta de destino se não existir
    try:
        destino_completo.parent.mkdir(parents=True, exist_ok=True)
        print(f"Pasta de destino criada/verificada: {destino_completo.parent}")

        # Copiar pasta de origem para destino
        if Path(origem).exists():
            shutil.copytree(origem, destino_completo, dirs_exist_ok=True)
            print(f"Pasta '{origem}' copiada para: {destino_completo}")
        else:
            print(f"Erro: Pasta de origem '{origem}' não encontrada")
    except Exception as e:
        print(f"Erro ao copiar pasta: {e}")
        print(f"Tentando criar apenas a estrutura de pastas...")

        # Tentar criar apenas a estrutura de pastas
        try:
            destino_completo.mkdir(parents=True, exist_ok=True)
            print(f"Estrutura de pastas criada: {destino_completo}")
        except Exception as e2:
            print(f"Erro ao criar estrutura de pastas: {e2}")

if __name__ == "__main__":
    caminho_destino = r"X:\TEMP\_Publico"
    copiar_pasta_para_rede("planilhas", caminho_destino)

    # Exemplo adicional com caminho local para teste:
    # copiar_pasta_para_rede("planilhas", r"C:\temp", "cotacoes_teste")
