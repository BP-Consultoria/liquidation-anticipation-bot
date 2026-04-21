# Liquidation Anticipation Bot

Bot RPA com integrações de API para automação do fluxo de **antecipação e liquidação** de recebíveis em operações de securitização.

## O que o projeto faz

O bot automatiza o ciclo operacional de antecipação de recebíveis:

1. **Consulta o banco de dados** (SQL Server) para buscar borderôs/antecipações pendentes.
2. **Conecta à API do Banco Arbi** para obter valores líquidos (TED REMESSA) por conta cedente.
3. **Atualiza o banco** com o valor líquido final e calcula débito/crédito agregado.
4. **Calcula deságio** (desconto) aplicando regras de cedente/sacado, dias úteis e fatores lidos do banco de securitização.
5. **Automatiza o sistema WBA Securitização** (desktop Windows) via RPA — lançamentos, recompra, ajustes em grid.
6. **Notifica via Microsoft Teams** (Graph API) com resumo da liquidação na conta corrente.

## Estrutura do projeto

```
├── auth/                        # Credenciais Arbi (YAML)
│   ├── paramenters.yml
│   └── parameters_homologacao.yml
├── src/
│   ├── main.py                  # Entry point → rpa.runner.run()
│   ├── config.py                # Settings (pydantic-settings + .env)
│   ├── api/
│   │   └── client.py            # Cliente HTTP genérico (httpx)
│   ├── rpa/
│   │   ├── runner.py            # Orquestração principal do fluxo
│   │   └── Wba.py               # Classe RPA para o WBA (pywinauto/pyautogui)
│   ├── services/
│   │   ├── database.py          # Queries e updates no SQL Server (pyodbc)
│   │   ├── desagio.py           # Cálculo de deságio (regras, dias úteis, fatores)
│   │   └── teams.py             # Auth Microsoft Graph (azure-identity)
│   └── utils/
│       ├── extrair_extratos.py  # API Arbi: token, extrato, valor líquido TED
│       ├── send_message_teams.py# Formatação e envio de mensagens no Teams
│       └── wba_helpers.py       # Helpers de UI/valores para o WBA
├── tests/                       # Testes (pytest)
├── pyproject.toml               # Metadados, dependências e scripts
├── uv.lock                      # Lock de dependências (uv)
└── .env.example                 # Template de variáveis de ambiente
```

## Pré-requisitos

- **Python 3.11+**
- **Windows** (o RPA do WBA usa `pywinauto`/`pyautogui` e depende do executável `Securitizacao.exe`)
- **SQL Server** com ODBC Driver 17 instalado
- **WBA Securitização** instalado na máquina
- Credenciais do **Banco Arbi**, **Microsoft Graph (Teams)** e dos bancos de dados

## Instalação

### Com [uv](https://docs.astral.sh/uv/) (recomendado)

```bash
uv sync
```

### Com pip

```bash
pip install -e .
```

Para dependências de desenvolvimento (pytest, ruff):

```bash
pip install -e ".[dev]"
```

## Configuração

1. Copie o template de variáveis de ambiente:

```bash
cp .env.example .env
```

2. Preencha o `.env` com os valores reais:

```env
# Geral
APP_NAME=liquidation-anticipation-bot
DEBUG=false

# API Arbi
API_BASE_URL=https://...
API_KEY=sua-chave
API_TIMEOUT=30

# Banco de dados (antecipação)
DB_DRIVER=mssql+pyodbc
DB_HOST=seu-servidor
DB_PORT=1433
DB_NAME=anticipation_db
DB_USER=usuario
DB_PASSWORD=senha
DB_ODBC_DRIVER=ODBC Driver 17 for SQL Server

# Banco de dados (securitização / deságio)
SEC_HOST=seu-servidor
SEC_PORT=1433
SEC_NAME=securitizacao_db
SEC_USER=usuario
SEC_PASSWORD=senha
SEC_ODBC_DRIVER=ODBC Driver 17 for SQL Server

# WBA
WBA_USERNAME=usuario-wba
WBA_PASSWORD=senha-wba

# Microsoft Teams (Graph API)
MS_CLIENT_ID=seu-client-id
MS_USERNAME=usuario@dominio.com
MS_PASSWORD=senha
MS_GRAPH_BASE_URL=https://graph.microsoft.com/v1.0
CHAT_ID=id-do-chat
TEAMS_NOME_PORTAL=NomePortal
```

3. Configure as credenciais da API Arbi em `auth/paramenters.yml`.

## Como rodar

### Via script registrado

```bash
bot
```

### Via Python direto

```bash
cd src
python main.py
```

### Fluxo de execução

```
bot / python main.py
  └─► main() → run()
        ├─ Conecta ao SQL Server
        ├─ Busca antecipações pendentes
        ├─ Renova token Arbi e consulta extratos (TED REMESSA)
        ├─ Atualiza Valor_Liquido_Final no banco
        ├─ Calcula deságio por cedente/sacado
        ├─ Agrega débito/crédito
        ├─ Abre WBA e executa lançamentos via RPA
        ├─ Envia notificação no Teams
        └─ Desconecta
```

## Desenvolvimento

### Lint

```bash
ruff check src/
```

### Testes

```bash
pytest
```

## Tecnologias

| Tecnologia | Uso |
|---|---|
| Python 3.11+ | Linguagem principal |
| pywinauto / pyautogui | Automação RPA do WBA (Windows) |
| pyodbc | Conexão direta com SQL Server |
| pandas / numpy | Manipulação de dados tabulares |
| httpx / requests | Chamadas HTTP (API Arbi / Graph) |
| pydantic-settings | Configuração tipada via `.env` |
| azure-identity | Autenticação Microsoft Graph |
| pyyaml | Leitura de credenciais Arbi |
| ruff | Linter e formatador |
| pytest | Framework de testes |
