# BI Qualidade Corporativo

Aplicativo Streamlit integrado com três módulos: Indicadores de Qualidade, SQDCP/FMDS e Projeto CPK.

## Primeiro acesso

- Usuário: `admin`
- Senha: `QualidadeRS2026`

Após o primeiro acesso, crie usuários nominais em **Administração > Usuários**.

## Como rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Como publicar no Streamlit Cloud

1. Suba todos os arquivos deste pacote no GitHub.
2. Configure o app apontando para `app.py`.
3. Opcional: configure os Secrets do GitHub para o módulo SQDCP usando `.streamlit/secrets.toml.example` como referência.

## Persistência

- Banco corporativo: `data/corporativo/bi_qualidade_corporativo.db`
- Indicadores de Qualidade: `data/qualidade/`
- SQDCP/FMDS: `data/sqdcp/`
- Projeto CPK: `data/cpk/`
- Backups: `data/backups/`

## Perfis

- `admin`: acessa tudo, usuários, auditoria e backup.
- `qualidade`: acessa Indicadores de Qualidade e Projeto CPK.
- `producao`: acessa SQDCP/FMDS.
- `consulta`: acessa os módulos sem administração.
