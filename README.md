# BI Qualidade Integrado

Aplicativo único em Streamlit com tela inicial para acesso aos três módulos já desenvolvidos:

1. Indicadores de Qualidade
2. SQDCP / FMDS
3. Projeto CPK

## Como rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Estrutura

```text
app.py                         # Tela inicial e roteamento dos módulos
modulos/indicadores_qualidade.py
modulos/sqdcp.py
modulos/cpk.py
requirements.txt
```

## Persistência

- Indicadores de Qualidade: mantém a última base carregada em `.last_input` e fallback no arquivo `Consultas_RNC_APP.xlsx`, conforme código original.
- SQDCP: mantém persistência local em `data/sqdcp_base.xlsx` e pode sincronizar com GitHub quando os secrets estiverem configurados.
- CPK: além dos modelos já salvos em `modelos_cartas_cpk.json`, o integrador salva o estado principal em `.unified_state/cpk_session.pkl`.

## Secrets opcionais para SQDCP no Streamlit Cloud

```toml
GITHUB_TOKEN = "seu_token"
GITHUB_REPO = "usuario/repositorio"
GITHUB_BRANCH = "main"
GITHUB_FILE_PATH = "data/sqdcp_base.xlsx"
```

## Observação técnica

Os programas originais foram mantidos como módulos independentes para preservar as funcionalidades já validadas. O arquivo `app.py` apenas controla a tela inicial, o botão de retorno e a execução do módulo selecionado.
