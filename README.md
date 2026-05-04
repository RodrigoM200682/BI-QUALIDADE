# BI Qualidade Corporativo - Versão estruturada para GitHub/Streamlit Cloud

## Como rodar

1. Suba todos os arquivos e pastas deste pacote para o repositório GitHub.
2. No Streamlit Cloud, configure o arquivo principal como `app.py`.
3. O aplicativo cria automaticamente as pastas de persistência caso o GitHub não envie pastas vazias.

## Primeiro acesso

- Usuário: `admin`
- Senha: `QualidadeRS2026`

## Estrutura obrigatória

```
app.py
corporate_core.py
requirements.txt
modulos/
  __init__.py
  indicadores_qualidade.py
  sqdcp.py
  cpk.py
data/
  corporativo/.gitkeep
  qualidade/.gitkeep
  sqdcp/.gitkeep
  cpk/.gitkeep
  backups/.gitkeep
manual/
  Manual_BI_Qualidade_Corporativo_RM_2026.pdf
```

## Correção aplicada

Esta versão possui mecanismo de autocorreção: se a pasta `modulos` ou algum módulo não estiver no servidor, o `app.py` recria os arquivos essenciais automaticamente a partir de cópias internas.

## Persistência

- Banco corporativo: `data/corporativo/bi_qualidade_corporativo.db`
- Indicadores: `data/qualidade/`
- SQDCP: `data/sqdcp/`
- CPK: `data/cpk/`
- Backups: `data/backups/`


## Persistência reforçada V6

Esta versão garante persistência específica para os módulos SQDCP e CPK:

- SQDCP: salva lançamentos, ações e metas em `data/sqdcp/sqdcp_base.xlsx`, também gravando uma cópia redundante no banco corporativo SQLite.
- CPK: salva os modelos de cartas em `data/cpk/modelos_cartas_cpk.json`, também gravando uma cópia redundante no banco corporativo SQLite.
- Se as pastas `data/sqdcp` ou `data/cpk` não existirem no GitHub/Streamlit Cloud, o `app.py` recria a estrutura automaticamente.
- Para persistência definitiva após redeploy no Streamlit Cloud, configure os secrets do GitHub:

```toml
GITHUB_TOKEN = "seu_token"
GITHUB_REPO = "usuario/repositorio"
GITHUB_BRANCH = "main"
GITHUB_SQDCP_FILE_PATH = "data/sqdcp/sqdcp_base.xlsx"
GITHUB_CPK_FILE_PATH = "data/cpk/modelos_cartas_cpk.json"
```

Sem estes secrets, a persistência funciona durante uso normal do app, mas pode ser perdida em redeploy/rebuild do Streamlit Cloud.
