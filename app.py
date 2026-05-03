
from __future__ import annotations

import pickle
import runpy
from pathlib import Path
from typing import Any

import streamlit as st

APP_TITLE = "BI Qualidade Integrado"
BASE_DIR = Path(__file__).resolve().parent
MODULES_DIR = BASE_DIR / "modulos"
STATE_DIR = BASE_DIR / ".unified_state"
STATE_DIR.mkdir(exist_ok=True)

def ensure_modules() -> None:
    MODULES_DIR.mkdir(exist_ok=True)

MODULES = {
    "indicadores": {
        "title": "Indicadores de Qualidade",
        "subtitle": "Consulta de ocorrências, comparação de períodos, motivos e exportação de relatório.",
        "icon": "📈",
        "file": MODULES_DIR / "indicadores_qualidade.py",
        "accent": "#0F5132",
    },
    "sqdcp": {
        "title": "SQDCP / FMDS",
        "subtitle": "Painel industrial com relógios, metas, séries históricas, perdas, eficiência e ações.",
        "icon": "🏭",
        "file": MODULES_DIR / "sqdcp.py",
        "accent": "#0F5132",
    },
    "cpk": {
        "title": "Projeto CPK",
        "subtitle": "Carta de inspeção, medições por amostra, Cp/Cpk, gráficos, parecer e exportações.",
        "icon": "📊",
        "file": MODULES_DIR / "cpk.py",
        "accent": "#0F5132",
    },
}

MODULE_STATE_KEYS = {
    "cpk": ["carta_ok", "caracteristicas", "selected_id", "carta_dados", "modelos", "char_form_nonce"],
    "sqdcp": ["dados", "acoes", "metas", "avisos", "loaded"],
    "indicadores": [],
}

MANUAL_FILE = BASE_DIR / "Manual_BI_Qualidade_Integrado_RM_2026.pdf"

st.set_page_config(page_title=APP_TITLE, page_icon="🧭", layout="wide")

UNIFIED_CSS = """
<style>
:root {
    --verde-principal: #0F5132;
    --verde-acao: #16A34A;
    --verde-claro: #DCFCE7;
    --texto: #10231D;
    --muted: #64748B;
    --card: rgba(255,255,255,.92);
    --borda: rgba(15, 81, 50, .16);
}
.stApp {background: linear-gradient(135deg, #ECFDF5 0%, #F8FAFC 42%, #EEF7F0 100%) !important; color: var(--texto) !important;}
.block-container {padding-top: 1.1rem !important; padding-bottom: 2rem !important; max-width: 1480px !important; background: transparent !important;}
section[data-testid="stSidebar"] > div {background: linear-gradient(180deg, #F0FDF4 0%, #FFFFFF 100%) !important; border-right: 1px solid var(--borda) !important;}
h1, h2, h3, h4, label, .stMarkdown, p {color: var(--texto) !important;}
.hero {border-radius: 28px; padding: 28px 30px; background: linear-gradient(135deg, #0F5132, #166534); color: white !important; box-shadow: 0 18px 50px rgba(15, 81, 50, .24); margin-bottom: 22px;}
.hero h1, .hero p {color: white !important;}
.hero h1 {margin:0; font-size:2.15rem;}
.hero p {font-size:1.02rem; color:rgba(255,255,255,.88) !important; margin-top:8px; max-width:1000px;}
.app-card {border:1px solid var(--borda); border-radius:24px; background:var(--card); padding:22px; min-height:245px; box-shadow:0 14px 36px rgba(15,23,42,.08); transition:transform .18s ease, box-shadow .18s ease;}
.app-card:hover {transform:translateY(-2px); box-shadow:0 18px 45px rgba(15,23,42,.12);}
.app-icon {font-size:2.6rem; margin-bottom:10px;}
.app-title {font-size:1.25rem; font-weight:800; margin-bottom:8px; color:#12372A;}
.app-subtitle {font-size:.94rem; color:#475569; min-height:78px;}
.pill {display:inline-block; padding:6px 10px; border-radius:999px; background:#DCFCE7; border:1px solid #86EFAC; color:#166534; font-weight:700; font-size:.78rem;}
.module-top {display:flex; align-items:center; justify-content:space-between; gap:14px; border-radius:22px; padding:18px 22px; margin-bottom:16px; background:rgba(255,255,255,.94); border:1px solid var(--borda); box-shadow:0 10px 28px rgba(15,23,42,.07);}
.module-title-box {display:flex; align-items:center; gap:14px;}
.module-icon {font-size:2.05rem;}
.module-title {font-size:1.55rem; font-weight:850; color:#12372A; line-height:1.1;}
.module-subtitle {color:#64748B; font-size:.92rem; margin-top:3px;}
div.stButton > button, div.stDownloadButton > button, div.stFormSubmitButton > button {border-radius:14px !important; font-weight:800 !important; min-height:42px !important; border:1px solid #16A34A !important; background:#16A34A !important; color:white !important;}
div.stButton > button:hover, div.stDownloadButton > button:hover, div.stFormSubmitButton > button:hover {background:#15803D !important; border-color:#15803D !important; color:white !important;}
div[data-testid="stMetric"], div[data-testid="stExpander"], .stDataFrame, div[data-testid="stDataEditor"] {background-color: rgba(255,255,255,.88) !important; border-radius:16px !important; border-color:var(--borda) !important;}
.card, .ok, .wn, .fl, .orange-help {border-radius:14px !important; background:rgba(255,255,255,.9) !important; color:var(--texto) !important; border-color:var(--borda) !important;}
@media (max-width: 768px) {.module-top{flex-direction:column; align-items:stretch;} .module-title-box{align-items:flex-start;} }
</style>
"""
st.markdown(UNIFIED_CSS, unsafe_allow_html=True)


def _state_path(module_key: str) -> Path:
    return STATE_DIR / f"{module_key}_session.pkl"


def _save_module_state(module_key: str) -> None:
    keys = MODULE_STATE_KEYS.get(module_key, [])
    data: dict[str, Any] = {}
    if keys:
        for key in keys:
            if key in st.session_state:
                data[key] = st.session_state[key]
    else:
        # Para módulos com persistência própria, salva somente controles de navegação/seleções serializáveis.
        for key, value in dict(st.session_state).items():
            if key.startswith("unified_") or key.startswith("_restore_"):
                continue
            if key in ("unified_module",):
                continue
            try:
                pickle.dumps(value)
            except Exception:
                continue
            data[key] = value
    if data:
        try:
            with open(_state_path(module_key), "wb") as f:
                pickle.dump(data, f)
        except Exception:
            pass


def _restore_module_state_once(module_key: str) -> None:
    marker = f"_restore_{module_key}_state_once"
    if st.session_state.get(marker):
        return
    path = _state_path(module_key)
    if path.exists():
        try:
            with open(path, "rb") as f:
                data = pickle.load(f)
            for key, value in data.items():
                st.session_state.setdefault(key, value)
        except Exception:
            pass
    st.session_state[marker] = True


def go_home() -> None:
    module_key = st.session_state.get("unified_module")
    if module_key in MODULES:
        _save_module_state(module_key)
    st.session_state["unified_module"] = None
    st.rerun()


def open_module(module_key: str) -> None:
    st.session_state["unified_module"] = module_key
    st.rerun()


def render_manual_home() -> None:
    with st.expander("📘 Manual do aplicativo integrado", expanded=False):
        st.markdown(
            """
            Este manual descreve a operação do **BI Qualidade Integrado** e de cada módulo:
            **Indicadores de Qualidade**, **SQDCP/FMDS** e **Projeto CPK**.

            Use o PDF quando precisar treinar novos usuários ou padronizar o uso do aplicativo.
            """
        )
        if MANUAL_FILE.exists():
            st.download_button(
                "📥 Baixar manual completo em PDF",
                data=MANUAL_FILE.read_bytes(),
                file_name="Manual_BI_Qualidade_Integrado_RM_2026.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        else:
            st.warning("Manual PDF não localizado no pacote do aplicativo.")


def render_home() -> None:
    st.markdown(
        """
        <div class="hero">
            <h1>🧭 BI Qualidade Integrado</h1>
            <p>Ambiente único para acessar os aplicativos já desenvolvidos: Indicadores de Qualidade, SQDCP/FMDS e Projeto CPK. A navegação foi padronizada e cada módulo mantém sua lógica funcional original.</p>
            <span class="pill">Visual único • Persistência por módulo • Manual integrado</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    cols = st.columns(3)
    for col, key in zip(cols, MODULES.keys()):
        cfg = MODULES[key]
        with col:
            st.markdown(
                f"""
                <div class="app-card" style="border-top: 7px solid {cfg['accent']};">
                    <div class="app-icon">{cfg['icon']}</div>
                    <div class="app-title">{cfg['title']}</div>
                    <div class="app-subtitle">{cfg['subtitle']}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.button(f"Abrir {cfg['title']}", key=f"open_{key}", use_container_width=True, on_click=open_module, args=(key,))
    st.markdown("---")
    render_manual_home()
    st.info("As últimas informações salvas são preservadas por módulo. Para persistência definitiva no Streamlit Cloud, mantenha o repositório completo publicado e configure GitHub Secrets no módulo SQDCP quando necessário.")


def render_module_header(cfg: dict[str, Any]) -> None:
    left, right = st.columns([5, 1.2])
    with left:
        st.markdown(
            f"""
            <div class="module-top">
                <div class="module-title-box">
                    <div class="module-icon">{cfg['icon']}</div>
                    <div>
                        <div class="module-title">{cfg['title']}</div>
                        <div class="module-subtitle">{cfg['subtitle']}</div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with right:
        st.button("↩ Tela inicial", use_container_width=True, on_click=go_home, key="btn_home_module")


def run_selected_module(module_key: str) -> None:
    ensure_modules()
    cfg = MODULES[module_key]
    module_file = Path(cfg["file"])

    render_module_header(cfg)

    if not module_file.exists():
        st.error(f"Arquivo do módulo não encontrado: {module_file.name}. Publique novamente o pacote completo no GitHub.")
        return

    _restore_module_state_once(module_key)

    original_set_page_config = st.set_page_config
    original_rerun = st.rerun

    def _noop_set_page_config(*args, **kwargs):
        return None

    def _safe_rerun(*args, **kwargs):
        _save_module_state(module_key)
        original_rerun(*args, **kwargs)

    st.set_page_config = _noop_set_page_config
    st.rerun = _safe_rerun
    try:
        runpy.run_path(str(module_file), run_name=f"__unified_{module_key}__")
    finally:
        _save_module_state(module_key)
        st.set_page_config = original_set_page_config
        st.rerun = original_rerun
        st.markdown(UNIFIED_CSS, unsafe_allow_html=True)


def main() -> None:
    module_key = st.session_state.get("unified_module")
    if module_key not in MODULES:
        render_home()
    else:
        run_selected_module(module_key)


if __name__ == "__main__":
    main()
