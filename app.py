from __future__ import annotations

import runpy
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

import corporate_core as core

APP_TITLE = "BI Qualidade Corporativo"
BASE_DIR = Path(__file__).resolve().parent
MODULES_DIR = BASE_DIR / "modulos"
MANUAL_FILE = BASE_DIR / "manual" / "Manual_BI_Qualidade_Corporativo_RM_2026.pdf"

MODULES = {
    "indicadores": {
        "title": "Indicadores de Qualidade",
        "subtitle": "Ocorrências, motivos, comparação entre períodos e exportação gerencial.",
        "icon": "📈",
        "file": MODULES_DIR / "indicadores_qualidade.py",
        "profiles": ["admin", "qualidade", "consulta"],
    },
    "sqdcp": {
        "title": "SQDCP / FMDS",
        "subtitle": "Segurança, Qualidade, Delivery, Custo e Processo com metas, ações e histórico.",
        "icon": "🏭",
        "file": MODULES_DIR / "sqdcp.py",
        "profiles": ["admin", "producao", "consulta"],
    },
    "cpk": {
        "title": "Projeto CPK",
        "subtitle": "Cartas de inspeção, medições, Cp/Cpk, gráficos, parecer e exportações.",
        "icon": "📊",
        "file": MODULES_DIR / "cpk.py",
        "profiles": ["admin", "qualidade", "consulta"],
    },
}

st.set_page_config(page_title=APP_TITLE, page_icon="🧭", layout="wide")

CSS = """
<style>
:root {
  --verde: #0F5132;
  --verde2: #16A34A;
  --verde3: #DCFCE7;
  --texto: #10231D;
  --muted: #64748B;
  --card: rgba(255,255,255,.94);
  --borda: rgba(15,81,50,.18);
}
.stApp {background: linear-gradient(135deg, #ECFDF5 0%, #F8FAFC 45%, #EEF7F0 100%) !important; color: var(--texto) !important;}
.block-container {padding-top: 1rem !important; max-width: 1500px !important;}
section[data-testid="stSidebar"] > div {background: linear-gradient(180deg, #F0FDF4 0%, #FFFFFF 100%) !important; border-right: 1px solid var(--borda) !important;}
h1,h2,h3,h4,label,p,.stMarkdown {color: var(--texto) !important;}
.hero {border-radius: 28px; padding: 28px 30px; background: linear-gradient(135deg, #0F5132, #166534); color: white !important; box-shadow: 0 18px 50px rgba(15,81,50,.24); margin-bottom: 20px;}
.hero h1,.hero p {color:white !important;} .hero h1 {margin:0;font-size:2.2rem;} .hero p {margin-top:8px;color:rgba(255,255,255,.88)!important;}
.pill {display:inline-block; padding:6px 11px; border-radius:999px; background:#DCFCE7; border:1px solid #86EFAC; color:#166534; font-weight:800; font-size:.78rem; margin-right:6px;}
.app-card {border:1px solid var(--borda); border-radius:24px; background:var(--card); padding:22px; min-height:240px; box-shadow:0 14px 36px rgba(15,23,42,.08);}
.app-icon {font-size:2.6rem;margin-bottom:10px;} .app-title {font-size:1.25rem;font-weight:850;color:#12372A;margin-bottom:8px;} .app-subtitle {font-size:.94rem;color:#475569;min-height:70px;}
.module-top {display:flex; align-items:center; justify-content:space-between; gap:14px; border-radius:22px; padding:18px 22px; margin-bottom:16px; background:rgba(255,255,255,.96); border:1px solid var(--borda); box-shadow:0 10px 28px rgba(15,23,42,.07);}
.module-title-box {display:flex; align-items:center; gap:14px;} .module-icon {font-size:2.1rem;} .module-title {font-size:1.55rem;font-weight:850;color:#12372A;line-height:1.1;} .module-subtitle {color:#64748B;font-size:.92rem;margin-top:3px;}
.admin-box {border:1px solid var(--borda); border-radius:18px; background:rgba(255,255,255,.92); padding:16px; box-shadow:0 8px 24px rgba(15,23,42,.06);}
div.stButton > button, div.stDownloadButton > button, div.stFormSubmitButton > button {border-radius:14px !important; font-weight:800 !important; min-height:42px !important; border:1px solid #16A34A !important; background:#16A34A !important; color:white !important;}
div.stButton > button:hover, div.stDownloadButton > button:hover, div.stFormSubmitButton > button:hover {background:#15803D !important; border-color:#15803D !important; color:white !important;}
div[data-testid="stMetric"], div[data-testid="stExpander"], .stDataFrame, div[data-testid="stDataEditor"] {background-color: rgba(255,255,255,.90) !important; border-radius:16px !important; border-color:var(--borda) !important;}
.card, .ok, .wn, .fl, .orange-help {border-radius:14px !important; background:rgba(255,255,255,.92) !important; color:var(--texto) !important; border-color:var(--borda) !important;}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def user() -> dict[str, Any] | None:
    return st.session_state.get("auth_user")


def logout() -> None:
    u = user()
    if u:
        core.audit(u["username"], "login", "logout", "Saída manual")
    st.session_state.pop("auth_user", None)
    st.session_state.pop("module", None)
    st.rerun()


def render_login() -> None:
    core.init_db()
    st.markdown(
        """
        <div class="hero">
            <h1>🧭 BI Qualidade Corporativo</h1>
            <p>Ambiente integrado com controle de acesso, persistência central, auditoria e backup.</p>
            <span class="pill">Login</span><span class="pill">SQLite</span><span class="pill">Auditoria</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    col1, col2, col3 = st.columns([1, 1.1, 1])
    with col2:
        st.subheader("Acesso ao sistema")
        with st.form("login_form"):
            username = st.text_input("Usuário", value="admin")
            password = st.text_input("Senha", type="password")
            submitted = st.form_submit_button("Entrar", use_container_width=True)
        if submitted:
            auth = core.authenticate(username, password)
            if auth:
                st.session_state["auth_user"] = auth
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos.")
        st.info("Primeiro acesso: usuário `admin` e senha `QualidadeRS2026`. Troque a senha na Administração.")


def open_module(key: str) -> None:
    u = user()
    if not core.can_access(u, key):
        st.warning("Seu perfil não possui acesso a este módulo.")
        return
    core.audit(u["username"], key, "open_module", MODULES[key]["title"])
    st.session_state["module"] = key
    st.rerun()


def go_home() -> None:
    st.session_state["module"] = None
    st.rerun()


def render_home() -> None:
    u = user()
    st.markdown(
        f"""
        <div class="hero">
            <h1>🧭 BI Qualidade Corporativo</h1>
            <p>Ambiente único para Indicadores de Qualidade, SQDCP/FMDS e Projeto CPK, com banco corporativo, controle de acesso, backup e rastreabilidade.</p>
            <span class="pill">Usuário: {u['display_name']}</span><span class="pill">Perfil: {u['role']}</span><span class="pill">Persistência por módulo</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    cols = st.columns(3)
    for col, key in zip(cols, MODULES.keys()):
        cfg = MODULES[key]
        permitido = core.can_access(u, key)
        with col:
            st.markdown(
                f"""
                <div class="app-card" style="border-top:7px solid #0F5132; opacity:{1 if permitido else .45};">
                    <div class="app-icon">{cfg['icon']}</div>
                    <div class="app-title">{cfg['title']}</div>
                    <div class="app-subtitle">{cfg['subtitle']}</div>
                    <span class="pill">{'Liberado' if permitido else 'Sem acesso'}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.button(f"Abrir {cfg['title']}", key=f"open_{key}", use_container_width=True, disabled=not permitido, on_click=open_module, args=(key,))

    st.markdown("---")
    with st.expander("📘 Manual do aplicativo", expanded=True):
        st.markdown(
            """
            O manual descreve o uso do sistema integrado, a rotina de cada módulo, a persistência dos dados, o login por perfil, backup e boas práticas operacionais.
            """
        )
        if MANUAL_FILE.exists():
            st.download_button("📥 Baixar manual completo em PDF", data=MANUAL_FILE.read_bytes(), file_name="Manual_BI_Qualidade_Corporativo_RM_2026.pdf", mime="application/pdf", use_container_width=True)
        else:
            st.warning("Manual PDF não encontrado no pacote.")


def render_module_header(key: str) -> None:
    cfg = MODULES[key]
    left, right = st.columns([5, 1.25])
    with left:
        st.markdown(
            f"""
            <div class="module-top">
              <div class="module-title-box">
                <div class="module-icon">{cfg['icon']}</div>
                <div><div class="module-title">{cfg['title']}</div><div class="module-subtitle">{cfg['subtitle']}</div></div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with right:
        st.button("↩ Tela inicial", use_container_width=True, on_click=go_home, key=f"home_{key}")


def run_module(key: str) -> None:
    u = user()
    if not core.can_access(u, key):
        st.error("Acesso não autorizado.")
        return
    cfg = MODULES[key]
    render_module_header(key)
    if not cfg["file"].exists():
        st.error(f"Arquivo do módulo não localizado: {cfg['file'].name}")
        return

    original_set_page_config = st.set_page_config
    original_rerun = st.rerun

    def no_set_page_config(*args, **kwargs):
        return None

    def safe_rerun(*args, **kwargs):
        core.audit(u["username"], key, "module_rerun", "Atualização de tela ou salvamento no módulo")
        original_rerun(*args, **kwargs)

    st.set_page_config = no_set_page_config
    st.rerun = safe_rerun
    try:
        runpy.run_path(str(cfg["file"]), run_name=f"__corporativo_{key}__")
    finally:
        st.set_page_config = original_set_page_config
        st.rerun = original_rerun
        st.markdown(CSS, unsafe_allow_html=True)


def render_admin() -> None:
    u = user()
    if not core.can_access(u, "admin"):
        st.error("Apenas administradores acessam esta área.")
        return
    st.subheader("Administração corporativa")
    tabs = st.tabs(["Usuários", "Auditoria", "Backup", "Diretórios"])
    with tabs[0]:
        st.markdown("### Usuários e perfis")
        st.dataframe(pd.DataFrame(core.list_users()), use_container_width=True, hide_index=True)
        with st.form("user_form"):
            c1, c2, c3 = st.columns(3)
            username = c1.text_input("Usuário")
            display = c2.text_input("Nome de exibição")
            role = c3.selectbox("Perfil", ["admin", "qualidade", "producao", "consulta"])
            password = st.text_input("Nova senha", type="password", help="Obrigatória para novo usuário. Para editar sem trocar senha, deixe em branco.")
            active = st.checkbox("Usuário ativo", value=True)
            if st.form_submit_button("Salvar usuário", use_container_width=True):
                try:
                    core.create_or_update_user(username, display, role, password or None, active)
                    core.audit(u["username"], "admin", "save_user", username)
                    st.success("Usuário salvo.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Não foi possível salvar: {exc}")
    with tabs[1]:
        st.markdown("### Trilha de auditoria")
        st.dataframe(pd.DataFrame(core.recent_audit(200)), use_container_width=True, hide_index=True)
    with tabs[2]:
        st.markdown("### Backup completo")
        st.caption("Gera um pacote ZIP com banco corporativo, dados dos módulos, manuais e códigos principais.")
        if st.button("Gerar backup agora", use_container_width=True):
            path = core.make_backup_zip()
            st.session_state["last_backup"] = str(path)
            st.success(f"Backup gerado: {path.name}")
        if st.session_state.get("last_backup"):
            p = Path(st.session_state["last_backup"])
            if p.exists():
                st.download_button("📥 Baixar último backup", data=p.read_bytes(), file_name=p.name, mime="application/zip", use_container_width=True)
    with tabs[3]:
        st.markdown("### Estrutura de persistência")
        rows = []
        for p in [core.DATA_DIR / "qualidade", core.DATA_DIR / "sqdcp", core.DATA_DIR / "cpk", core.CORP_DIR, core.BACKUP_DIR]:
            rows.append({"Pasta": str(p.relative_to(BASE_DIR)), "Existe": p.exists(), "Arquivos": len(list(p.rglob('*'))) if p.exists() else 0})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def sidebar() -> None:
    u = user()
    with st.sidebar:
        st.markdown("### BI Qualidade")
        st.write(f"**Usuário:** {u['display_name']}")
        st.write(f"**Perfil:** {u['role']}")
        st.divider()
        if st.button("🏠 Tela inicial", use_container_width=True):
            go_home()
        if core.can_access(u, "admin"):
            if st.button("⚙️ Administração", use_container_width=True):
                st.session_state["module"] = "admin"
                st.rerun()
        st.divider()
        st.button("Sair", use_container_width=True, on_click=logout)


def main() -> None:
    core.init_db()
    if not user():
        render_login()
        return
    sidebar()
    module = st.session_state.get("module")
    if not module:
        render_home()
    elif module == "admin":
        render_admin()
    elif module in MODULES:
        run_module(module)
    else:
        render_home()


if __name__ == "__main__":
    main()
