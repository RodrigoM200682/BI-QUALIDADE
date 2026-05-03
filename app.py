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

MODULES = {
    "indicadores": {
        "title": "Indicadores de Qualidade",
        "subtitle": "Consulta de ocorrências, comparação de períodos, motivos e exportação de relatório.",
        "icon": "📈",
        "file": MODULES_DIR / "indicadores_qualidade.py",
        "accent": "#1F5C4B",
    },
    "sqdcp": {
        "title": "SQDCP / FMDS",
        "subtitle": "Painel industrial com relógios, metas, séries históricas, perdas, eficiência e ações.",
        "icon": "🏭",
        "file": MODULES_DIR / "sqdcp.py",
        "accent": "#24A148",
    },
    "cpk": {
        "title": "Projeto CPK",
        "subtitle": "Carta de inspeção, medições por amostra, Cp/Cpk, gráficos, parecer e exportações.",
        "icon": "📊",
        "file": MODULES_DIR / "cpk.py",
        "accent": "#FF8C00",
    },
}

CPK_STATE_KEYS = [
    "carta_ok",
    "caracteristicas",
    "selected_id",
    "carta_dados",
    "modelos",
    "char_form_nonce",
]

st.set_page_config(page_title=APP_TITLE, page_icon="🧭", layout="wide")

st.markdown(
    """
    <style>
    .stApp {
        background: radial-gradient(circle at top left, #DCFCE7 0%, #F8FAFC 36%, #EEF2FF 100%);
        color: #10231D;
    }
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1480px;}
    .hero {
        border-radius: 28px;
        padding: 28px 30px;
        background: linear-gradient(135deg, rgba(15,81,50,0.97), rgba(25,111,88,0.92));
        color: white;
        box-shadow: 0 18px 50px rgba(15, 81, 50, .22);
        margin-bottom: 22px;
    }
    .hero h1 {margin: 0; font-size: 2.15rem; color: white;}
    .hero p {font-size: 1.02rem; color: rgba(255,255,255,.88); margin-top: 8px; max-width: 950px;}
    .app-card {
        border: 1px solid rgba(17, 94, 89, .15);
        border-radius: 24px;
        background: rgba(255,255,255,.87);
        padding: 22px;
        min-height: 245px;
        box-shadow: 0 14px 36px rgba(15, 23, 42, .08);
        transition: transform .18s ease, box-shadow .18s ease;
    }
    .app-card:hover {transform: translateY(-2px); box-shadow: 0 18px 45px rgba(15, 23, 42, .12);}
    .app-icon {font-size: 2.6rem; margin-bottom: 10px;}
    .app-title {font-size: 1.25rem; font-weight: 800; margin-bottom: 8px; color:#12372A;}
    .app-subtitle {font-size: .94rem; color:#475569; min-height: 78px;}
    .pill {
        display:inline-block; padding:6px 10px; border-radius:999px;
        background:#ECFDF5; border:1px solid #BBF7D0; color:#166534; font-weight:700; font-size:.78rem;
    }
    div.stButton > button {
        border-radius: 14px;
        font-weight: 800;
        min-height: 44px;
    }
    .module-header {
        border-radius: 22px; padding: 18px 22px; margin-bottom: 16px;
        background: rgba(255,255,255,.9); border: 1px solid rgba(15, 23, 42, .08);
        box-shadow: 0 10px 28px rgba(15, 23, 42, .07);
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def _save_cpk_state() -> None:
    data: dict[str, Any] = {}
    for key in CPK_STATE_KEYS:
        if key in st.session_state:
            data[key] = st.session_state[key]
    if data:
        with open(STATE_DIR / "cpk_session.pkl", "wb") as f:
            pickle.dump(data, f)


def _restore_cpk_state_once() -> None:
    marker = "_cpk_state_restored_once"
    if st.session_state.get(marker):
        return
    path = STATE_DIR / "cpk_session.pkl"
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
    st.session_state["unified_module"] = None
    st.rerun()


def open_module(module_key: str) -> None:
    st.session_state["unified_module"] = module_key
    st.rerun()


def render_home() -> None:
    st.markdown(
        """
        <div class="hero">
            <h1>🧭 BI Qualidade Integrado</h1>
            <p>Um único ambiente para acessar os aplicativos já desenvolvidos: Indicadores de Qualidade, SQDCP/FMDS e Projeto CPK. A estrutura foi pensada para preservar as funcionalidades dos programas individuais e facilitar o uso pela tela inicial.</p>
            <span class="pill">Persistência reforçada • Navegação por módulos • Visual moderno</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)
    cols = [c1, c2, c3]
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
    st.info(
        "Orientação: mantenha este arquivo app.py como módulo principal no Streamlit Cloud. "
        "Os programas originais ficam preservados na pasta modulos."
    )


def run_selected_module(module_key: str) -> None:
    cfg = MODULES[module_key]

    top_left, top_right = st.columns([5, 1])
    with top_left:
        st.markdown(
            f"""
            <div class="module-header">
                <div style="font-size:1.8rem">{cfg['icon']}</div>
                <h2 style="margin:0;color:#12372A">{cfg['title']}</h2>
                <div style="color:#64748B">{cfg['subtitle']}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with top_right:
        st.button("↩ Tela inicial", use_container_width=True, on_click=go_home)

    if module_key == "cpk":
        _restore_cpk_state_once()

    original_set_page_config = st.set_page_config
    original_rerun = st.rerun

    def _noop_set_page_config(*args, **kwargs):
        return None

    def _safe_rerun(*args, **kwargs):
        if module_key == "cpk":
            _save_cpk_state()
        original_rerun(*args, **kwargs)

    st.set_page_config = _noop_set_page_config
    st.rerun = _safe_rerun
    try:
        runpy.run_path(str(cfg["file"]), run_name=f"__unified_{module_key}__")
    finally:
        if module_key == "cpk":
            _save_cpk_state()
        st.set_page_config = original_set_page_config
        st.rerun = original_rerun


def main() -> None:
    module_key = st.session_state.get("unified_module")
    if module_key not in MODULES:
        render_home()
    else:
        run_selected_module(module_key)


if __name__ == "__main__":
    main()
