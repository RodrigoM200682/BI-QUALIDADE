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

## Ajuste V7 - Consulta de modelos CPK

O módulo CPK agora possui a aba **0. Consulta de modelos**, permitindo pesquisar modelos salvos por característica, linha, embalagem ou nome do modelo.

Funcionalidades incluídas:
- visualizar todas as características salvas em modelos de carta;
- carregar o modelo completo;
- incluir somente uma característica salva na inspeção atual;
- manter LIE, LSE, número de amostras e número de medições por amostra do modelo original;
- impedir duplicidade de característica dentro da inspeção aberta.

Para usar em nova inspeção:
1. Abra o módulo CPK.
2. Salve a Carta de dados da nova inspeção.
3. Vá em **Consulta de modelos** ou em **Criar inspeção > Puxar característica de modelo salvo**.
4. Pesquise a característica e clique em incluir.


## Atualização V8 - Backup e restauração do CPK

O módulo CPK agora possui a aba **5. Backup / restauração**.

Funcionalidades incluídas:
- Baixar um Excel completo com modelos salvos, carta atual, características, medições e resultados.
- Restaurar automaticamente a base do CPK por upload do Excel gerado pelo próprio aplicativo.
- Persistir a última inspeção ativa em `data/cpk/cpk_estado_atual.json`.
- Persistir os modelos de carta em `data/cpk/modelos_cartas_cpk.json`.

Fluxo recomendado:
1. Criar/salvar carta e características no módulo CPK.
2. Salvar medições.
3. Acessar `5. Backup / restauração`.
4. Baixar o backup Excel completo.
5. Em caso de perda de histórico, fazer upload deste arquivo e clicar em **Atualizar base CPK com este backup**.

## V9 - Ajuste do cálculo CPK
- Aba de análise estatística do CPK ajustada conforme guia prático anexado.
- O cálculo agora apresenta CPS, CPI e CPK com passo a passo por característica.
- Exportações Excel/PDF do CPK passam a apresentar CPS e CPI.


## V10 - Restauração automática CPK

No módulo CPK, a aba **5. Backup / restauração** permite baixar um Excel completo com modelos de cartas, carta ativa, características, medições e resultados. Ao fazer upload deste Excel, a base CPK é atualizada automaticamente, sem necessidade de botão adicional.

Arquivos de persistência principais:

- `data/cpk/modelos_cartas_cpk.json`
- `data/cpk/cpk_estado_atual.json`
- `data/sqdcp/sqdcp_base.xlsx`
