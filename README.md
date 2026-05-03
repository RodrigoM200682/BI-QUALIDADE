# BI Qualidade Integrado

Aplicativo Streamlit único com três módulos preservados:

1. Indicadores de Qualidade
2. SQDCP / FMDS
3. Projeto CPK

## Arquivos principais

- `app.py`: tela inicial integrada e roteamento dos módulos.
- `modulos/indicadores_qualidade.py`: módulo de indicadores.
- `modulos/sqdcp.py`: módulo SQDCP/FMDS.
- `modulos/cpk.py`: módulo CPK.
- `Manual_BI_Qualidade_Integrado_RM_2026.pdf`: manual disponível para download na tela inicial.
- `requirements.txt`: dependências para Streamlit Cloud.

## Ajustes desta versão

- Visual padronizado em todas as telas.
- Botão verde de retorno para a tela inicial ao lado do título do módulo.
- Persistência por módulo em `.unified_state`.
- Mantida a persistência própria já existente dos módulos:
  - Indicadores: última base Excel carregada.
  - SQDCP: base Excel local e opção GitHub via secrets.
  - CPK: estado das cartas, características e medições.
- Manual integrado na tela inicial com download em PDF.

## Como publicar no Streamlit Cloud

1. Envie todos os arquivos e pastas para o GitHub.
2. Configure o arquivo principal como `app.py`.
3. Mantenha a pasta `modulos` no repositório.
4. Mantenha o `requirements.txt` na raiz.
5. Para persistência definitiva do SQDCP, configure os secrets:
   - `GITHUB_TOKEN`
   - `GITHUB_REPO`
   - `GITHUB_BRANCH`
   - `GITHUB_FILE_PATH`

