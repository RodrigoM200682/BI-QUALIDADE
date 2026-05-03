# APP Integrado Qualidade RS

Aplicativo Streamlit integrado com módulos:
- Indicadores de Qualidade
- SQDCP
- Projeto CPK

## Como rodar localmente
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud
No repositório GitHub, mantenha estes arquivos na raiz:
- app.py
- requirements.txt

Em Main file path, informe: `app.py`.

## Persistência
O aplicativo cria automaticamente o banco `qualidade_integrado.db` na pasta do app. Os dados ficam salvos mesmo após reinício do aplicativo, falha temporária ou inatividade, desde que o ambiente preserve arquivos locais.

Para ambientes cloud com reset de container, recomenda-se versionar backups exportados ou conectar futuramente um banco externo.
