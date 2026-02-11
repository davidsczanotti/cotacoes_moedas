# cotacoes-moedas

Aplicacao para coletar cotacoes de moedas em sites especificos e manter
historico continuo em XLSX e CSV.

## Visao geral

- Coleta USD/BRL (spot) no Investing.
- Coleta PTAX USD/EUR/CHF no Banco Central (BCB).
- Coleta Dolar Turismo no Valor Globo.
- Coleta TJLP no BNDES.
- Coleta SELIC no BCB (linha mais recente) e calcula CDI diario.
- Atualiza `planilhas/cotacoes.xlsx` como fonte de verdade.
- Atualiza `planilhas/cotacoes.csv` com separador `;` e decimal `,`.

## Estrutura do codigo

- `main.py`: orquestracao (coleta, atualiza planilhas e copia para rede).
- `cotacoes_moedas/investing.py`: USD/BRL (Investing).
- `cotacoes_moedas/valor_globo.py`: Dolar Turismo (Valor).
- `cotacoes_moedas/bcb_ptax.py`: PTAX (BCB).
- `cotacoes_moedas/juros.py`: TJLP, SELIC e calculo de CDI.
- `cotacoes_moedas/storage.py`: escrita no XLSX/CSV.
- `cotacoes_moedas/network_copy.py`: conversao de drive mapeado -> UNC (Windows).
- `cotacoes_moedas/network_sync.py`: selecao do destino e copia de `planilhas/` na rede.
- `cotacoes_moedas/playwright_utils.py`: utilitarios Playwright (proxy e pagina Chromium).
- `cotacoes_moedas/parsing.py`: parse de numeros PT-BR.
- `cotacoes_moedas/redaction.py`: mascara credenciais/senhas em mensagens.

## Como funciona

1. Valida a planilha para decidir quais fontes coletar (janela de horario + campos vazios no dia).
2. Busca os dados via Playwright (headless) apenas para as fontes elegiveis (por padrao em paralelo).
3. Para cada fonte, grava os valores na linha da data local do dia.
4. Se a data ja existe, preenche apenas colunas vazias (nao sobrescreve cotacoes ja preenchidas no dia); se nao, cria nova linha.
5. Para TJLP/SELIC/CDI, se nao houver valor novo no dia, repete o ultimo valor disponivel.
6. Atualiza a coluna de log com status e timestamp.
7. Atualiza o CSV com a linha do ultimo log, substituindo a data se ja existir.

Observacoes importantes (para uso no cliente):

- A atualizacao do `cotacoes.xlsx` e feita com **uma unica gravacao** no arquivo (menos chance de erro e mais rapido).
- Se alguma celula do dia ja estiver preenchida, ela **nao e sobrescrita**; o console vai mostrar `nao gravou (ja preenchido na planilha)`.
- Se aparecer `ERRO ao gravar arquivos`, normalmente e porque o `cotacoes.xlsx`/`cotacoes.csv` esta aberto no Excel ou a permissao da pasta nao permite escrita.

## Regras de horario (janelas)

Para economizar processamento e evitar sobrescrever cotacoes do dia, o robo aplica estas janelas no horario local da maquina:

- **USD/BRL (Investing)** e **Dolar Turismo (Valor)**: coleta apenas ate **08:30** (inclusive).
- **PTAX (BCB)**: coleta apenas a partir de **13:10** (inclusive).
- **TJLP/SELIC**: coleta apenas ate **08:30** (inclusive).
- As fontes elegiveis dentro da janela rodam em paralelo por padrao (limite com `COTACOES_MAX_WORKERS`).
- Se nao houver nenhuma fonte elegivel (fora da janela ou colunas do dia ja preenchidas), o robo encerra **sem alterar** `planilhas/cotacoes.xlsx` e `planilhas/cotacoes.csv`.

Cenarios:

- **07:00**: busca apenas USD/BRL (Investing) e Dolar Turismo (Valor); nao aciona PTAX.
- **08:31**: nao busca nenhuma fonte e nao altera nenhum valor.
- **13:10**: busca apenas PTAX (USD/EUR/CHF); nao aciona Investing/Valor.

## Regras do historico

- Uma linha por data.
- Execucoes repetidas no mesmo dia nao duplicam linhas.
- O CSV e atualizado apenas para a data do ultimo log gravado.
- As colunas TJLP/SELIC/CDI repetem o ultimo valor quando nao houver atualizacao no dia.

## Tratamento de erros

- Se uma fonte falhar, os campos dessa fonte nao sao atualizados (preserva o valor do dia, se existir).
- O fluxo continua mesmo com falhas parciais.
- A coluna de log registra `ERRO <timestamp> - <motivo>`.
- Fontes puladas por janela de horario ou por ja estarem preenchidas no dia nao contam como erro.

## Estrutura das colunas

Ordem do XLSX/CSV:

1. Data
2. Dolar Oficial Compra (USD/BRL do Investing)
3. Dolar Oficial Venda (Compra + spread de 0.0020)
4. Dolar PTAX Compra
5. Dolar PTAX Venda
6. Dolar Turismo Compra
7. Dolar Turismo Venda
8. Euro PTAX Compra
9. Euro PTAX Venda
10. CHF PTAX Compra
11. CHF PTAX Venda
12. TJLP
13. SELIC
14. CDI
15. Situacao (log)

## Pre-requisitos

- Python >= 3.12
- Poetry v2
- Playwright + Chromium

## Instalacao

```bash
poetry install
poetry run playwright install chromium
```

No WSL, se faltar dependencia do navegador, use:

```bash
poetry run playwright install --with-deps
```

## Uso

```bash
poetry run python main.py
```

## Logs de progresso

Durante a execucao, o script imprime mensagens de progresso indicando cada fonte
e quando a planilha/CSV sao atualizados.

Principais mensagens:

- Validacao e selecao das fontes (o que vai rodar e o que foi pulado).
- Modo de coleta (sequencial ou paralelo / quantidade de workers).
- Erros de copia em rede (com detalhes das tentativas quando todos os destinos falham).

## Testes

```bash
poetry run pytest
```

Arquivos atualizados:

- `planilhas/cotacoes.xlsx`
- `planilhas/cotacoes.csv`

## Empacotamento (Windows)

Recomendado: Nuitka (standalone) + Inno Setup.

1. Gere o build no Windows:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
scripts\\build_windows.ps1
```

Isso cria `dist\\main.dist` com o executavel e os dados. Se o
Inno Setup estiver no PATH, o instalador tambem e gerado automaticamente.

Se o Inno Setup nao estiver no PATH, defina `INNO_SETUP_PATH`:

```powershell
$env:INNO_SETUP_PATH = "C:\\Program Files (x86)\\Inno Setup 6\\ISCC.exe"
scripts\\build_windows.ps1
```

2. Gere o instalador com Inno Setup:

```powershell
iscc installer\\cotacoes-moedas.iss
```

O instalador sai em `dist\\cotacoes-moedas-setup.exe`.
O script imprime o caminho do instalador ao final do build.

Observacoes:

- O script inclui `planilhas/` e `ms-playwright/` no pacote.
- O script faz um `robocopy` de `planilhas/` e `ms-playwright/` para `dist\\main.dist` e gera `dist\\cotacoes-moedas-portable.zip`.
- O icone do app/instalador vem de `imagem_ico/finance.ico`.
- A versao vem de `pyproject.toml`.
- Para atualizar, substitua o `cotacoes-moedas-setup.exe` na rede e execute novamente.
- Para evitar alertas de antivirus, assine o instalador e o executavel.
- Use o Agendador do Windows apontando para `{app}\\cotacoes-moedas.exe`.

Assinatura (opcional):

- Defina `SIGNING_PFX_PATH` com o caminho do `.pfx`.
- `SIGNING_PFX_PASSWORD` e opcional; se nao informar, o script pede a senha.
- `SIGNING_TIMESTAMP_URL` (opcional) define o servidor de timestamp.
- Certificados self-signed exigem instalar o `.cer` em "Trusted Root" e
  "Trusted Publishers" nas maquinas cliente para evitar "Unknown Publisher".

## Configuracao

Proxy (opcional):

- `HTTP_PROXY` / `HTTPS_PROXY` (com usuario/senha se necessario).

Copia para pasta de rede:

- `COTACOES_NETWORK_DIR` com um ou mais caminhos separados por `;` (ex.: `X:\TEMP\_Publico;Y:\TEMP\_Publico;\\servidor\TEMP\_Publico`).

Pasta na rede (opcional):

- `COTACOES_NETWORK_DEST_FOLDER` define a subpasta criada no destino (padrao: `cotacoes`).

Concorrencia (opcional):

- `COTACOES_MAX_WORKERS` limita quantas fontes rodam em paralelo (ex.: `1` desativa paralelismo).

## Observacoes

- A planilha `planilhas/cotacoes.xlsx` precisa existir (modelo).
- O cliente nao trabalha em fins de semana; recomenda-se executar apenas em dias uteis.
- Se rodar em fim de semana/feriado, e possivel nao ter cotacoes e registrar erro no log.
