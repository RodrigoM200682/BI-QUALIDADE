import io
import json
import math
import os
import re
import base64
from pathlib import Path
from typing import Optional, Tuple
from datetime import datetime, date

import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import requests

try:
    from corporate_core import read_setting, write_setting, audit
except Exception:
    read_setting = None
    write_setting = None
    audit = None

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.graphics.shapes import Drawing, Line, Circle, String, Rect
except Exception:
    A4 = None

st.set_page_config(page_title="Carta de Inspeção CPK", layout="wide", page_icon="📊")

BASE_DIR = Path(__file__).resolve().parents[1]
DATA_CPK_DIR = BASE_DIR / "data" / "cpk"
DATA_CPK_DIR.mkdir(parents=True, exist_ok=True)
MODELOS_FILE = DATA_CPK_DIR / "modelos_cartas_cpk.json"
CPK_STATE_FILE = DATA_CPK_DIR / "cpk_estado_atual.json"


def get_secret(name: str, default: str = "") -> str:
    try:
        return str(st.secrets.get(name, default))
    except Exception:
        return os.getenv(name, default)


GITHUB_TOKEN = get_secret("GITHUB_TOKEN")
GITHUB_REPO = get_secret("GITHUB_REPO")
GITHUB_BRANCH = get_secret("GITHUB_BRANCH", "main")
GITHUB_CPK_FILE_PATH = get_secret("GITHUB_CPK_FILE_PATH", "data/cpk/modelos_cartas_cpk.json")


def github_cpk_enabled() -> bool:
    return bool(GITHUB_TOKEN and GITHUB_REPO and GITHUB_BRANCH and GITHUB_CPK_FILE_PATH)


def github_headers() -> dict:
    return {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def github_cpk_api_url() -> str:
    return f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_CPK_FILE_PATH}"


def read_github_modelos() -> Tuple[Optional[dict], Optional[str], Optional[str]]:
    if not github_cpk_enabled():
        return None, None, None
    try:
        r = requests.get(github_cpk_api_url(), headers=github_headers(), params={"ref": GITHUB_BRANCH}, timeout=20)
        if r.status_code == 404:
            return None, None, None
        r.raise_for_status()
        payload = r.json()
        raw = base64.b64decode(payload["content"])
        data = json.loads(raw.decode("utf-8"))
        return data if isinstance(data, dict) else {}, payload.get("sha"), None
    except Exception as exc:
        return None, None, f"Não foi possível ler os modelos CPK no GitHub: {exc}"


def write_github_modelos(modelos: dict) -> Optional[str]:
    if not github_cpk_enabled():
        return None
    _, sha, _ = read_github_modelos()
    raw = json.dumps(modelos, ensure_ascii=False, indent=2).encode("utf-8")
    payload = {
        "message": "Atualiza modelos de cartas CPK",
        "content": base64.b64encode(raw).decode("utf-8"),
        "branch": GITHUB_BRANCH,
    }
    if sha:
        payload["sha"] = sha
    try:
        r = requests.put(github_cpk_api_url(), headers=github_headers(), json=payload, timeout=25)
        r.raise_for_status()
        return None
    except Exception as exc:
        return f"Não foi possível gravar os modelos CPK no GitHub: {exc}"


def read_corporate_modelos() -> Optional[dict]:
    if read_setting is None:
        return None
    try:
        data = read_setting("cpk_modelos_cartas", None)
        return data if isinstance(data, dict) else None
    except Exception:
        return None


def write_corporate_modelos(modelos: dict) -> None:
    if write_setting is None:
        return
    try:
        write_setting("cpk_modelos_cartas", modelos)
        if audit:
            audit(None, "cpk", "save_cpk_models", f"Modelos CPK salvos: {len(modelos)}")
    except Exception:
        pass

st.markdown(
    """
    <style>
    .stApp {background: linear-gradient(180deg,#0a0a0d 0%,#12121a 100%); color:#e8e8f0;}
    h1, h2, h3 {color:#e8e8f0;}
    .version {color:#8a8a9e; font-size:10px; font-style:italic; text-align:right;}
    .card {background:#1c1c25; border:1px solid #2a2a32; border-radius:16px; padding:16px; margin-bottom:10px;}
    .ok {border-left:5px solid #00e5a0; padding:10px; background:#11181a; border-radius:10px; margin-bottom:8px;}
    .wn {border-left:5px solid #ffb340; padding:10px; background:#1d1810; border-radius:10px; margin-bottom:8px;}
    .fl {border-left:5px solid #ff4d4d; padding:10px; background:#1d1111; border-radius:10px; margin-bottom:8px;}
    .muted {color:#9a9aad; font-size:0.9rem;}
    .orange-help {border-left:5px solid #ff8c00; padding:10px; background:#20160a; border-radius:10px; margin-bottom:8px;}
    div.stButton > button[kind="primary"] {background:#00a86b !important; border-color:#00a86b !important; color:white !important;}
    div.stFormSubmitButton > button {background:#ff8c00 !important; border-color:#ff8c00 !important; color:white !important; font-weight:700;}
    div.stDownloadButton > button {background:#00a86b !important; border-color:#00a86b !important; color:white !important; font-weight:700;}
    </style>
    """,
    unsafe_allow_html=True,
)

VERSION = f"RM_{datetime.now().strftime('%d_%m_%Y_%H%M')}"
st.markdown(f"<div class='version'>{VERSION}</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# Persistência dos modelos
# ─────────────────────────────────────────────────────────────────────────────
def load_modelos():
    """Carrega modelos salvos de cartas CPK com redundância.

    Ordem de recuperação:
    1) GitHub, quando os secrets estiverem configurados;
    2) arquivo local data/cpk/modelos_cartas_cpk.json;
    3) armazenamento corporativo SQLite;
    4) dicionário vazio.
    """
    gh_data, _, gh_warn = read_github_modelos()
    if isinstance(gh_data, dict):
        try:
            MODELOS_FILE.parent.mkdir(parents=True, exist_ok=True)
            MODELOS_FILE.write_text(json.dumps(gh_data, ensure_ascii=False, indent=2), encoding="utf-8")
            write_corporate_modelos(gh_data)
        except Exception:
            pass
        return gh_data

    if MODELOS_FILE.exists():
        try:
            data = json.loads(MODELOS_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                write_corporate_modelos(data)
                return data
        except Exception:
            pass

    stored = read_corporate_modelos()
    if isinstance(stored, dict):
        try:
            MODELOS_FILE.parent.mkdir(parents=True, exist_ok=True)
            MODELOS_FILE.write_text(json.dumps(stored, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass
        return stored

    try:
        MODELOS_FILE.parent.mkdir(parents=True, exist_ok=True)
        MODELOS_FILE.write_text("{}", encoding="utf-8")
    except Exception:
        pass
    return {}


def save_modelos(modelos):
    modelos = modelos if isinstance(modelos, dict) else {}
    MODELOS_FILE.parent.mkdir(parents=True, exist_ok=True)
    MODELOS_FILE.write_text(json.dumps(modelos, ensure_ascii=False, indent=2), encoding="utf-8")
    write_corporate_modelos(modelos)
    warn = write_github_modelos(modelos)
    if warn:
        try:
            st.warning(warn)
        except Exception:
            pass




# ─────────────────────────────────────────────────────────────────────────────
# Backup Excel e persistência da inspeção atual
# ─────────────────────────────────────────────────────────────────────────────
def _json_default(obj):
    try:
        if pd.isna(obj):
            return None
    except Exception:
        pass
    if isinstance(obj, (datetime, date)):
        return obj.strftime("%d/%m/%Y")
    return str(obj)


def current_cpk_state():
    """Retorna tudo que precisa ser recuperado se a sessão cair ou se for necessário restaurar por Excel."""
    return {
        "versao_backup": "CPK_BACKUP_V1",
        "atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "carta_ok": bool(st.session_state.get("carta_ok", False)),
        "carta_dados": st.session_state.get("carta_dados", {}) or {},
        "caracteristicas": st.session_state.get("caracteristicas", []) or [],
        "selected_id": st.session_state.get("selected_id"),
        "modelos": st.session_state.get("modelos", {}) or load_modelos(),
    }


def save_current_cpk_state():
    """Salva a última inspeção ativa fora do session_state."""
    try:
        state = current_cpk_state()
        CPK_STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
        CPK_STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2, default=_json_default), encoding="utf-8")
        if write_setting is not None:
            write_setting("cpk_estado_atual", state)
        if audit:
            audit(None, "cpk", "save_cpk_state", "Estado atual do CPK salvo")
    except Exception:
        pass


def load_current_cpk_state():
    """Carrega a última inspeção ativa salva."""
    if CPK_STATE_FILE.exists():
        try:
            data = json.loads(CPK_STATE_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except Exception:
            pass
    if read_setting is not None:
        try:
            data = read_setting("cpk_estado_atual", None)
            if isinstance(data, dict):
                try:
                    CPK_STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
                    CPK_STATE_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2, default=_json_default), encoding="utf-8")
                except Exception:
                    pass
                return data
        except Exception:
            pass
    return {}


def apply_cpk_state_to_session(state: dict):
    if not isinstance(state, dict):
        return
    st.session_state.carta_ok = bool(state.get("carta_ok", False))
    st.session_state.carta_dados = state.get("carta_dados", {}) or {}
    st.session_state.caracteristicas = state.get("caracteristicas", []) or []
    st.session_state.selected_id = state.get("selected_id")
    modelos = state.get("modelos")
    if isinstance(modelos, dict):
        st.session_state.modelos = modelos


def cpk_backup_excel_bytes():
    """Gera Excel de backup com modelos, carta atual, características, medições e resultados."""
    state = current_cpk_state()
    modelos = state.get("modelos", {}) or {}
    carta = state.get("carta_dados", {}) or {}
    chars = state.get("caracteristicas", []) or []

    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        pd.DataFrame([{
            "orientacao": "Arquivo de backup do módulo CPK. Para restaurar, faça upload deste Excel na aba Backup / restauração do próprio módulo CPK.",
            "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "qtd_modelos": len(modelos),
            "qtd_caracteristicas_abertas": len(chars),
        }]).to_excel(writer, sheet_name="LEIA_ME", index=False)

        pd.DataFrame({"json_modelos": [json.dumps(modelos, ensure_ascii=False, default=_json_default)]}).to_excel(writer, sheet_name="MODELOS_JSON", index=False)
        pd.DataFrame({"json_estado_atual": [json.dumps(state, ensure_ascii=False, default=_json_default)]}).to_excel(writer, sheet_name="ESTADO_JSON", index=False)
        pd.DataFrame([carta] if carta else [{}]).to_excel(writer, sheet_name="CARTA_ATUAL", index=False)

        linhas_chars = []
        linhas_med = []
        for c in chars:
            linhas_chars.append({
                "id": c.get("id"),
                "descricao": c.get("descricao"),
                "lie": c.get("lie"),
                "lse": c.get("lse"),
                "num_amostras": c.get("num_amostras"),
                "num_medicoes": c.get("num_medicoes", 3),
            })
            for row in c.get("medicoes", []) or []:
                base = {"id": c.get("id"), "descricao": c.get("descricao"), "Amostra": row.get("Amostra")}
                for k, v in row.items():
                    if str(k).startswith("Medida"):
                        linhas_med.append({**base, "Medida": k, "Valor": v})
        pd.DataFrame(linhas_chars).to_excel(writer, sheet_name="CARACTERISTICAS", index=False)
        pd.DataFrame(linhas_med).to_excel(writer, sheet_name="MEDICOES", index=False)

        try:
            resultados = [calc_characteristic(c) for c in chars if characteristic_has_measurements(c)]
            linhas_res = [{
                "Característica": r.get("descricao"),
                "N": r.get("n"),
                "Média": r.get("media"),
                "Desvio": r.get("desvio"),
                "LIE": r.get("lie"),
                "LSE": r.get("lse"),
                "Cp": r.get("cp"),
                "CPS": r.get("cps"),
                "CPI": r.get("cpi"),
                "Cpk": r.get("cpk"),
                "Status": r.get("status"),
            } for r in resultados]
            pd.DataFrame(linhas_res).to_excel(writer, sheet_name="RESULTADOS", index=False)
        except Exception:
            pd.DataFrame([]).to_excel(writer, sheet_name="RESULTADOS", index=False)

    bio.seek(0)
    return bio.getvalue()


def restore_cpk_from_excel(file_bytes: bytes):
    """Restaura automaticamente a base CPK a partir do Excel gerado pelo próprio app."""
    sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, engine="openpyxl")
    state = None
    modelos = None

    if "ESTADO_JSON" in sheets and not sheets["ESTADO_JSON"].empty:
        raw = sheets["ESTADO_JSON"].iloc[0, 0]
        if isinstance(raw, str) and raw.strip():
            state = json.loads(raw)

    if "MODELOS_JSON" in sheets and not sheets["MODELOS_JSON"].empty:
        raw = sheets["MODELOS_JSON"].iloc[0, 0]
        if isinstance(raw, str) and raw.strip():
            modelos = json.loads(raw)

    if state is None:
        # Fallback para arquivos legíveis sem JSON: reconstrói a inspeção atual pelas abas estruturadas.
        carta = {}
        if "CARTA_ATUAL" in sheets and not sheets["CARTA_ATUAL"].empty:
            carta = sheets["CARTA_ATUAL"].fillna("").to_dict("records")[0]
        chars = []
        med_df = sheets.get("MEDICOES", pd.DataFrame())
        if "CARACTERISTICAS" in sheets:
            for _, row in sheets["CARACTERISTICAS"].fillna("").iterrows():
                cid = str(row.get("id", "")).strip() or f"C{len(chars)+1:03d}"
                m = int(row.get("num_medicoes", 3) or 3)
                n = int(row.get("num_amostras", 1) or 1)
                medicoes = []
                for i in range(1, n + 1):
                    linha = {"Amostra": i, **{f"Medida {j}": None for j in range(1, m + 1)}}
                    if not med_df.empty:
                        subset = med_df[(med_df["id"].astype(str) == cid) & (pd.to_numeric(med_df["Amostra"], errors="coerce") == i)]
                        for _, mr in subset.iterrows():
                            medida = str(mr.get("Medida", ""))
                            if medida in linha:
                                linha[medida] = parse_float(mr.get("Valor"))
                    medicoes.append(linha)
                chars.append({
                    "id": cid,
                    "descricao": str(row.get("descricao", "")),
                    "lie": parse_float(row.get("lie")),
                    "lse": parse_float(row.get("lse")),
                    "num_amostras": n,
                    "num_medicoes": m,
                    "medicoes": medicoes,
                })
        state = {
            "versao_backup": "CPK_BACKUP_V1_RECONSTRUIDO",
            "atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "carta_ok": bool(carta),
            "carta_dados": carta,
            "caracteristicas": chars,
            "selected_id": chars[0]["id"] if chars else None,
            "modelos": modelos if isinstance(modelos, dict) else load_modelos(),
        }

    if isinstance(modelos, dict):
        state["modelos"] = modelos
        save_modelos(modelos)
    elif isinstance(state.get("modelos"), dict):
        save_modelos(state.get("modelos"))

    apply_cpk_state_to_session(state)
    save_current_cpk_state()
    return len(st.session_state.get("modelos", {}) or {}), len(st.session_state.get("caracteristicas", []) or [])


def modelo_from_state(nome):
    return {
        "nome": nome,
        "criado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "carta_base": {
            "linha": st.session_state.carta_dados.get("linha", ""),
            "embalagem": st.session_state.carta_dados.get("embalagem", ""),
            "domo_mat": st.session_state.carta_dados.get("domo_mat", ""),
            "esp_domo": st.session_state.carta_dados.get("esp_domo", ""),
            "corpo_mat": st.session_state.carta_dados.get("corpo_mat", ""),
            "esp_corpo": st.session_state.carta_dados.get("esp_corpo", ""),
            "fundo_mat": st.session_state.carta_dados.get("fundo_mat", ""),
            "esp_fundo": st.session_state.carta_dados.get("esp_fundo", ""),
        },
        "caracteristicas": [
            {
                "descricao": c["descricao"],
                "lie": c["lie"],
                "lse": c["lse"],
                "num_amostras": c["num_amostras"],
                "num_medicoes": c.get("num_medicoes", 3),
            }
            for c in st.session_state.caracteristicas
        ],
    }


def aplicar_modelo(modelo):
    base = modelo.get("carta_base", {})
    st.session_state.carta_dados = {
        **st.session_state.get("carta_dados", {}),
        "linha": base.get("linha", ""),
        "embalagem": base.get("embalagem", ""),
        "domo_mat": base.get("domo_mat", ""),
        "esp_domo": base.get("esp_domo", ""),
        "corpo_mat": base.get("corpo_mat", ""),
        "esp_corpo": base.get("esp_corpo", ""),
        "fundo_mat": base.get("fundo_mat", ""),
        "esp_fundo": base.get("esp_fundo", ""),
    }
    chars = []
    for idx, c in enumerate(modelo.get("caracteristicas", []), start=1):
        n = int(c.get("num_amostras", 1) or 1)
        m = int(c.get("num_medicoes", 3) or 3)
        chars.append({
            "id": f"C{idx:03d}_{datetime.now().strftime('%H%M%S')}",
            "descricao": c.get("descricao", ""),
            "lie": float(c.get("lie")),
            "lse": float(c.get("lse")),
            "num_amostras": n,
            "num_medicoes": m,
            "medicoes": [{"Amostra": i, **{f"Medida {j}": None for j in range(1, m + 1)}} for i in range(1, n + 1)],
        })
    st.session_state.caracteristicas = chars
    st.session_state.selected_id = chars[0]["id"] if chars else None
    save_current_cpk_state()


def criar_caracteristica_a_partir_modelo(c, ordem=None):
    """Cria uma característica nova, sem medições, usando parâmetros de um modelo salvo."""
    n = int(c.get("num_amostras", 1) or 1)
    m = int(c.get("num_medicoes", 3) or 3)
    ordem = ordem or (len(st.session_state.get("caracteristicas", [])) + 1)
    descricao = str(c.get("descricao", "")).strip()
    return {
        "id": f"C{ordem:03d}_{datetime.now().strftime('%H%M%S%f')}",
        "descricao": descricao,
        "lie": float(c.get("lie")),
        "lse": float(c.get("lse")),
        "num_amostras": n,
        "num_medicoes": m,
        "medicoes": [{"Amostra": i, **{f"Medida {j}": None for j in range(1, m + 1)}} for i in range(1, n + 1)],
    }


def consultar_caracteristicas_modelos(modelos, termo=""):
    """Retorna uma tabela pesquisável com todas as características salvas nos modelos."""
    termo_norm = descricao_normalizada(termo)
    linhas = []
    for nome_modelo, modelo in sorted((modelos or {}).items()):
        base = modelo.get("carta_base", {}) if isinstance(modelo, dict) else {}
        for idx, c in enumerate(modelo.get("caracteristicas", []) if isinstance(modelo, dict) else [], start=1):
            descricao = str(c.get("descricao", "")).strip()
            texto_busca = descricao_normalizada(" ".join([
                nome_modelo,
                descricao,
                str(base.get("linha", "")),
                str(base.get("embalagem", "")),
            ]))
            if termo_norm and termo_norm not in texto_busca:
                continue
            linhas.append({
                "chave": f"{nome_modelo}||{idx-1}",
                "Modelo": nome_modelo,
                "Característica": descricao,
                "LIE": c.get("lie"),
                "LSE": c.get("lse"),
                "Amostras": c.get("num_amostras"),
                "Medições/amostra": c.get("num_medicoes", 3),
                "Linha modelo": base.get("linha", ""),
                "Embalagem modelo": base.get("embalagem", ""),
                "Criado em": modelo.get("criado_em", ""),
            })
    return linhas


def incluir_caracteristica_modelo(modelo_nome, indice_caracteristica):
    modelos = st.session_state.get("modelos", {})
    modelo = modelos.get(modelo_nome)
    if not modelo:
        return False, "Modelo não localizado."
    caracteristicas_modelo = modelo.get("caracteristicas", [])
    if indice_caracteristica < 0 or indice_caracteristica >= len(caracteristicas_modelo):
        return False, "Característica não localizada dentro do modelo."
    c = caracteristicas_modelo[indice_caracteristica]
    desc_norm = descricao_normalizada(c.get("descricao"))
    if any(descricao_normalizada(x.get("descricao")) == desc_norm for x in st.session_state.get("caracteristicas", [])):
        return False, f"A característica '{c.get('descricao')}' já está aberta nesta inspeção."
    nova = criar_caracteristica_a_partir_modelo(c)
    st.session_state.caracteristicas.append(nova)
    st.session_state.selected_id = nova["id"]
    save_current_cpk_state()
    return True, f"Característica '{nova['descricao']}' incluída na inspeção atual com os limites e plano de amostragem do modelo."

# ─────────────────────────────────────────────────────────────────────────────
# Estado
# ─────────────────────────────────────────────────────────────────────────────
def init_state():
    saved_state = load_current_cpk_state()
    modelos_salvos = load_modelos()
    if isinstance(saved_state.get("modelos"), dict) and saved_state.get("modelos"):
        modelos_salvos = saved_state.get("modelos")
    defaults = {
        "carta_ok": bool(saved_state.get("carta_ok", False)),
        "caracteristicas": saved_state.get("caracteristicas", []) or [],
        "selected_id": saved_state.get("selected_id"),
        "carta_dados": saved_state.get("carta_dados", {}) or {},
        "modelos": modelos_salvos,
        "char_form_nonce": 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ─────────────────────────────────────────────────────────────────────────────
# Cálculo estatístico
# ─────────────────────────────────────────────────────────────────────────────
def parse_float(value):
    if value is None or value == "":
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    try:
        return float(str(value).replace(",", "."))
    except Exception:
        return None


def calc_stats(samples, lse=None, lie=None):
    vals = [float(v) for v in samples if v is not None]
    if len(vals) < 2:
        return None
    n = len(vals)
    mean = sum(vals) / n
    variance = sum((v - mean) ** 2 for v in vals) / (n - 1)
    std = math.sqrt(variance)
    if std == 0:
        return None
    cp = cps = cpi = cpk = None
    if lse is not None and lie is not None:
        # Conforme o guia: CPS = capacidade superior e CPI = capacidade inferior.
        # O CPK é o menor valor entre CPS e CPI.
        cp = (lse - lie) / (6 * std)
        cps = (lse - mean) / (3 * std)
        cpi = (mean - lie) / (3 * std)
        cpk = min(cps, cpi)
    elif lse is not None:
        cps = (lse - mean) / (3 * std)
        cpk = cps
    elif lie is not None:
        cpi = (mean - lie) / (3 * std)
        cpk = cpi
    return {"vals": vals, "n": n, "mean": mean, "std": std, "cp": cp, "cps": cps, "cpi": cpi, "cpk": cpk}


def classify_cpk(cpk):
    if cpk is None:
        return "Sem cálculo"
    if cpk >= 1.33:
        return "Capaz"
    if cpk >= 1.00:
        return "Marginal"
    return "Incapaz"


def measurement_keys(char):
    qtd = int(char.get("num_medicoes", 3) or 3)
    return [f"Medida {i}" for i in range(1, qtd + 1)]


def flatten_measurements(char):
    vals = []
    keys = measurement_keys(char)
    for row in char.get("medicoes", []):
        for key in keys:
            v = parse_float(row.get(key))
            if v is not None:
                vals.append(v)
    return vals


def amostras_completas(char):
    completas = 0
    keys = measurement_keys(char)
    for row in char.get("medicoes", []):
        if all(parse_float(row.get(k)) is not None for k in keys):
            completas += 1
    return completas


def calc_characteristic(char):
    vals = flatten_measurements(char)
    s = calc_stats(vals, char.get("lse"), char.get("lie"))
    result = {
        "id": char["id"],
        "descricao": char["descricao"],
        "lie": char.get("lie"),
        "lse": char.get("lse"),
        "amostras_previstas": char.get("num_amostras", 0),
        "amostras_completas": amostras_completas(char),
        "medicoes_por_amostra": int(char.get("num_medicoes", 3) or 3),
        "medidas_previstas": char.get("num_amostras", 0) * int(char.get("num_medicoes", 3) or 3),
        "medidas_realizadas": len(vals),
        "valores": vals,
        "n": s["n"] if s else 0,
        "media": round(s["mean"], 4) if s else None,
        "desvio": round(s["std"], 4) if s else None,
        "cp": round(s["cp"], 4) if s and s["cp"] is not None else None,
        "cps": round(s["cps"], 4) if s and s["cps"] is not None else None,
        "cpi": round(s["cpi"], 4) if s and s["cpi"] is not None else None,
        "cpk": round(s["cpk"], 4) if s and s["cpk"] is not None else None,
    }
    result["status"] = classify_cpk(result["cpk"])
    return result


def cpk_formula_rows(result):
    """Monta o passo a passo do cálculo seguindo o Guia Prático de CPK."""
    lie = result.get("lie")
    lse = result.get("lse")
    media = result.get("media")
    desvio = result.get("desvio")
    cps = result.get("cps")
    cpi = result.get("cpi")
    cpk = result.get("cpk")
    if media is None or desvio is None or desvio == 0:
        return pd.DataFrame([{
            "Passo": "Pré-requisito",
            "Indicador": "Medições válidas",
            "Fórmula / regra": "É necessário ter pelo menos 2 medições válidas e desvio padrão maior que zero.",
            "Resultado": "Sem cálculo",
        }])
    return pd.DataFrame([
        {"Passo": "1", "Indicador": "LIE", "Fórmula / regra": "Limite inferior definido por especificação", "Resultado": lie},
        {"Passo": "1", "Indicador": "LSE", "Fórmula / regra": "Limite superior definido por especificação", "Resultado": lse},
        {"Passo": "2", "Indicador": "n", "Fórmula / regra": "Quantidade de medições válidas coletadas", "Resultado": result.get("n")},
        {"Passo": "3", "Indicador": "Média x̄", "Fórmula / regra": "Soma das medições ÷ quantidade de medições", "Resultado": media},
        {"Passo": "4", "Indicador": "Desvio padrão σ", "Fórmula / regra": "Dispersão dos valores em relação à média", "Resultado": desvio},
        {"Passo": "5", "Indicador": "CPS", "Fórmula / regra": f"(LSE - média) / (3 × σ) = ({lse} - {media}) / (3 × {desvio})", "Resultado": cps},
        {"Passo": "6", "Indicador": "CPI", "Fórmula / regra": f"(média - LIE) / (3 × σ) = ({media} - {lie}) / (3 × {desvio})", "Resultado": cpi},
        {"Passo": "7", "Indicador": "CPK", "Fórmula / regra": "Menor valor entre CPS e CPI", "Resultado": cpk},
    ])


def render_cpk_reference_card():
    st.markdown("#### Modelo de cálculo utilizado")
    st.markdown(
        "<div class='orange-help'>"
        "O cálculo segue o conceito: <b>CPK = menor valor entre CPS e CPI</b>. "
        "CPS mede a distância da média até o limite superior; CPI mede a distância da média até o limite inferior. "
        "O menor deles indica o lado mais crítico da especificação."
        "</div>",
        unsafe_allow_html=True,
    )
    c1, c2, c3 = st.columns(3)
    c1.markdown("**CPS**  \n`(LSE - média) / (3 × desvio padrão)`")
    c2.markdown("**CPI**  \n`(média - LIE) / (3 × desvio padrão)`")
    c3.markdown("**CPK**  \n`min(CPS; CPI)`")


def interpretar_lado_critico(result):
    cps = result.get("cps")
    cpi = result.get("cpi")
    if cps is None or cpi is None:
        return "Lado crítico não calculado."
    if cps < cpi:
        return "O lado crítico é o limite superior, pois CPS é menor que CPI."
    if cpi < cps:
        return "O lado crítico é o limite inferior, pois CPI é menor que CPS."
    return "O processo está aproximadamente centralizado entre os limites, pois CPS e CPI são iguais ou muito próximos."


def characteristic_has_measurements(char):
    return len(flatten_measurements(char)) > 0

def descricao_normalizada(txt):
    return re.sub(r"\s+", " ", str(txt or "").strip()).casefold()

def build_insights(results):
    valid = [r for r in results if r.get("cpk") is not None]
    if not valid:
        return [{"level":"wn", "text":"Ainda não há medições suficientes para análise. Cada característica deve ter medições válidas e limites de especificação informados."}]
    fail = [r for r in valid if r["cpk"] < 1]
    warn = [r for r in valid if 1 <= r["cpk"] < 1.33]
    ok = [r for r in valid if r["cpk"] >= 1.33]
    insights = []
    if fail:
        names = ", ".join(f"{r['descricao']} (Cpk {r['cpk']:.2f})" for r in fail)
        insights.append({"level":"fl", "text":f"Parecer: processo não capaz para {len(fail)} característica(s): {names}. Minha recomendação é não liberar o lote sem avaliação técnica, revisar setup, segregar material suspeito e repetir a coleta após correção."})
    if warn:
        names = ", ".join(f"{r['descricao']} (Cpk {r['cpk']:.2f})" for r in warn)
        insights.append({"level":"wn", "text":f"Parecer: processo marginal para {len(warn)} característica(s): {names}. Eu manteria o processo em acompanhamento reforçado, com ajuste preventivo e aumento temporário da frequência de inspeção."})
    if ok and not fail and not warn:
        insights.append({"level":"ok", "text":f"Parecer: processo capaz. Todas as {len(ok)} característica(s) avaliadas apresentam Cpk ≥ 1,33, indicando boa condição estatística frente aos limites especificados."})
    for r in valid:
        cp, cpk = r.get("cp"), r.get("cpk")
        if cp and cpk and cp > 0 and (cp - cpk) / cp > 0.15:
            perda = round((cp - cpk) / cp * 100)
            insights.append({"level":"wn", "text":f"Descentramento identificado em {r['descricao']}: Cp={cp:.2f} e Cpk={cpk:.2f}, com perda aproximada de {perda}% da capacidade potencial. A variação pode estar aceitável, mas o processo está deslocado em relação ao centro da especificação."})
    incompletas = [r for r in results if r["amostras_completas"] < r["amostras_previstas"]]
    if incompletas:
        names = ", ".join(f"{r['descricao']} ({r['amostras_completas']}/{r['amostras_previstas']} amostras)" for r in incompletas)
        insights.append({"level":"wn", "text":f"Atenção: existem coletas incompletas em {names}. O parecer estatístico é parcial até completar todas as medições previstas em cada amostra."})
    pct_ok = round(len(ok) / len(valid) * 100)
    level = "ok" if pct_ok >= 80 and not fail else "wn" if pct_ok >= 50 else "fl"
    insights.append({"level":level, "text":f"Resumo geral: {pct_ok}% das características calculadas estão capazes. Total analisado: {len(valid)} característica(s)."})
    return insights


def control_chart(result):
    vals = result.get("valores", [])
    if len(vals) < 2:
        return None
    mean = sum(vals) / len(vals)
    std = math.sqrt(sum((v - mean) ** 2 for v in vals) / (len(vals) - 1))
    ucl = mean + 3 * std
    lcl = mean - 3 * std
    x = list(range(1, len(vals) + 1))
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=x,
        y=vals,
        mode="lines+markers",
        name="Medições",
        line=dict(width=2),
        marker=dict(size=7),
    ))
    # A média aparece no título para não poluir o gráfico com mais uma linha.
    fig.add_hline(y=ucl, line_dash="dash", line_color="#ff4d4d", line_width=3, annotation_text="LSC", annotation_font_color="#ff4d4d")
    fig.add_hline(y=lcl, line_dash="dash", line_color="#ff4d4d", line_width=3, annotation_text="LIC", annotation_font_color="#ff4d4d")
    if result.get("lse") is not None:
        fig.add_hline(y=result["lse"], line_dash="solid", line_color="#ff4d4d", line_width=4, annotation_text="LSE", annotation_font_color="#ff4d4d")
    if result.get("lie") is not None:
        fig.add_hline(y=result["lie"], line_dash="solid", line_color="#ff4d4d", line_width=4, annotation_text="LIE", annotation_font_color="#ff4d4d")
    media_txt = "—" if result.get("media") is None else f"{result['media']:.4f}"
    cpk_txt = "—" if result.get("cpk") is None else f"{result['cpk']:.4f}"
    fig.update_layout(
        height=370,
        title=f"{result['descricao']} | Cpk: {cpk_txt} | Média: {media_txt}",
        paper_bgcolor="#1c1c25",
        plot_bgcolor="#16161d",
        font=dict(color="#e8e8f0"),
        margin=dict(l=20, r=20, t=55, b=20),
        xaxis_title="Sequência das medições",
        yaxis_title="Valor medido",
        showlegend=False,
    )
    return fig


def pdf_chart_drawing(result, W):
    vals = result.get("valores", [])
    if len(vals) < 2:
        return None
    mean = sum(vals) / len(vals)
    std = math.sqrt(sum((v - mean) ** 2 for v in vals) / (len(vals) - 1))
    ucl = mean + 3 * std
    lcl = mean - 3 * std
    lse = result.get("lse")
    lie = result.get("lie")
    all_v = vals + [ucl, lcl] + ([lse] if lse is not None else []) + ([lie] if lie is not None else [])
    vmn, vmx = min(all_v), max(all_v)
    vr = vmx - vmn or 0.001
    DW, DH = W, 42 * mm
    PAD_L, PAD_R, PAD_T, PAD_B = 13 * mm, 16 * mm, 5 * mm, 7 * mm
    CW, CH = DW - PAD_L - PAD_R, DH - PAD_T - PAD_B

    def xp(i):
        return PAD_L + (i / (len(vals) - 1 or 1)) * CW

    def yp(v):
        return PAD_B + ((v - vmn) / vr) * CH

    d = Drawing(DW, DH)
    d.add(Rect(0, 0, DW, DH, fillColor=colors.whitesmoke, strokeColor=colors.lightgrey, strokeWidth=0.5))

    def hline(val, col, label, width=1.0, dash=None):
        y = yp(val)
        ln = Line(PAD_L, y, PAD_L + CW, y, strokeColor=col, strokeWidth=width)
        if dash:
            ln.strokeDashArray = dash
        d.add(ln)
        d.add(String(PAD_L + CW + 2, y - 2, label, fontSize=5.5, fillColor=col))

    red = colors.HexColor("#cc0000")
    hline(ucl, red, "LSC", width=1.4, dash=[2, 2])
    hline(lcl, red, "LIC", width=1.4, dash=[2, 2])
    if lse is not None:
        hline(lse, red, "LSE", width=2.2)
    if lie is not None:
        hline(lie, red, "LIE", width=2.2)

    for i in range(len(vals) - 1):
        d.add(Line(xp(i), yp(vals[i]), xp(i + 1), yp(vals[i + 1]), strokeColor=colors.HexColor("#111111"), strokeWidth=0.9))
    for i, v in enumerate(vals):
        out = (lse is not None and v > lse) or (lie is not None and v < lie)
        col = colors.red if out else colors.HexColor("#00a86b")
        d.add(Circle(xp(i), yp(v), 1.7, fillColor=col, strokeColor=col))
    cpk_txt = "—" if result.get("cpk") is None else f"{result['cpk']:.4f}"
    media_txt = "—" if result.get("media") is None else f"{result['media']:.4f}"
    d.add(String(PAD_L, DH - 11, f"{result['descricao']} | Cpk: {cpk_txt} | Média: {media_txt}", fontSize=7.5, fillColor=colors.black))
    return d


def make_pdf(carta, results):
    if A4 is None:
        return None
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=12*mm, bottomMargin=14*mm)
    W = A4[0] - 24*mm
    styles = getSampleStyleSheet()
    def S(name, **kw):
        return ParagraphStyle(name, parent=styles["Normal"], **kw)
    TEXT = colors.HexColor("#111111")
    MUTED = colors.HexColor("#555555")
    story = []
    story.append(Paragraph("<b>Relatório de Carta de Inspeção e CPK</b>", S("title", fontSize=15, textColor=TEXT, fontName="Helvetica-Bold")))
    story.append(Paragraph(f"Gerado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", S("sub", fontSize=8, textColor=MUTED)))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph("<b>Dados dos materiais</b>", S("sec_mat", fontSize=9, textColor=TEXT, fontName="Helvetica-Bold")))
    dados_materiais = [
        ["Usina domo", carta.get("domo_mat", ""), "Espessura domo", carta.get("esp_domo", "")],
        ["Usina corpo", carta.get("corpo_mat", ""), "Espessura corpo", carta.get("esp_corpo", "")],
        ["Usina fundo", carta.get("fundo_mat", ""), "Espessura fundo", carta.get("esp_fundo", "")],
    ]
    t_mat = Table(dados_materiais, colWidths=[W*.18, W*.32, W*.18, W*.32])
    t_mat.setStyle(TableStyle([("GRID", (0,0),(-1,-1), .25, colors.grey), ("FONTSIZE", (0,0),(-1,-1), 7.5), ("BACKGROUND", (0,0),(-1,-1), colors.whitesmoke)]))
    story.append(t_mat)
    story.append(Spacer(1, 3*mm))

    story.append(Paragraph("<b>Dados do lote</b>", S("sec_lote", fontSize=9, textColor=TEXT, fontName="Helvetica-Bold")))
    dados_lote = [
        ["Linha", carta.get("linha", ""), "Embalagem", carta.get("embalagem", "")],
        ["Ordem de produção", carta.get("op", ""), "Quantidade", carta.get("lote_qtd", "")],
        ["Data", carta.get("data", ""), "Observações", carta.get("obs", "")],
    ]
    t_lote = Table(dados_lote, colWidths=[W*.18, W*.32, W*.18, W*.32])
    t_lote.setStyle(TableStyle([("GRID", (0,0),(-1,-1), .25, colors.grey), ("FONTSIZE", (0,0),(-1,-1), 7.5), ("BACKGROUND", (0,0),(-1,-1), colors.whitesmoke)]))
    story.append(t_lote)
    story.append(Spacer(1, 3*mm))

    story.append(Paragraph("<b>Responsável pela inspeção</b>", S("sec_resp", fontSize=9, textColor=TEXT, fontName="Helvetica-Bold")))
    dados_resp = [["Nome", carta.get("responsavel_nome", ""), "Chapa", carta.get("responsavel_chapa", "")]]
    t_resp = Table(dados_resp, colWidths=[W*.18, W*.32, W*.18, W*.32])
    t_resp.setStyle(TableStyle([("GRID", (0,0),(-1,-1), .25, colors.grey), ("FONTSIZE", (0,0),(-1,-1), 7.5), ("BACKGROUND", (0,0),(-1,-1), colors.whitesmoke)]))
    story.append(t_resp)
    story.append(Spacer(1, 4*mm))
    rows = [["Característica", "Amostras", "Med./am.", "N", "Média", "Desvio", "LIE", "LSE", "Cp", "CPS", "CPI", "Cpk", "Status"]]
    for r in results:
        rows.append([r["descricao"], f"{r['amostras_completas']}/{r['amostras_previstas']}", r.get("medicoes_por_amostra", ""), r["n"], r["media"], r["desvio"], r["lie"], r["lse"], r["cp"], r.get("cps"), r.get("cpi"), r["cpk"], r["status"]])
    tb = Table(rows, colWidths=[W*.19, W*.065, W*.055, W*.035, W*.065, W*.065, W*.055, W*.055, W*.045, W*.045, W*.045, W*.05, W*.08], repeatRows=1)
    tb.setStyle(TableStyle([("GRID", (0,0),(-1,-1), .25, colors.grey), ("FONTSIZE", (0,0),(-1,-1), 6.2), ("BACKGROUND", (0,0),(-1,0), colors.lightgrey)]))
    story.append(tb)
    story.append(Spacer(1, 5*mm))
    story.append(Paragraph("<b>Gráficos das coletas</b>", S("sec1", fontSize=10, textColor=TEXT, fontName="Helvetica-Bold")))
    for r in results:
        d = pdf_chart_drawing(r, W)
        if d:
            story.append(d)
            story.append(Spacer(1, 3*mm))
    story.append(Paragraph("<b>Análise e parecer automático</b>", S("sec2", fontSize=10, textColor=TEXT, fontName="Helvetica-Bold")))
    for ins in build_insights(results):
        story.append(Paragraph(ins["text"], S("body", fontSize=8, leading=11, textColor=TEXT)))
        story.append(Spacer(1, 1.2*mm))
    story.append(Spacer(1, 8*mm))
    story.append(Paragraph(f"Responsável pela inspeção: {carta.get('responsavel_nome','')} | Chapa: {carta.get('responsavel_chapa','')}", S("resp", fontSize=8, textColor=TEXT)))
    story.append(Paragraph("Assinatura: _______________________________", S("sig", fontSize=8, textColor=TEXT)))
    doc.build(story)
    buf.seek(0)
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# Interface
# ─────────────────────────────────────────────────────────────────────────────
st.title("📊 Carta de Inspeção CPK")
st.caption("Fluxo: modelo salvo → carta de dados → características sem duplicidade → medições por amostra → análise somente das cartas preenchidas → parecer automático.")

tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs(["0. Consulta de modelos", "1. Carta de dados", "2. Criar inspeção", "3. Registrar medições", "4. Análise estatística", "5. Backup / restauração"])


with tab0:
    st.subheader("Consulta de modelos de carta")
    st.markdown(
        "<div class='orange-help'>Use esta consulta para localizar uma característica já padronizada em modelos anteriores. "
        "Ao iniciar uma nova inspeção, salve primeiro a Carta de dados e depois inclua a característica pelo modelo para manter LIE, LSE, quantidade de amostras e medições por amostra.</div>",
        unsafe_allow_html=True,
    )
    modelos = st.session_state.get("modelos", {})
    if not modelos:
        st.warning("Ainda não há modelos salvos. Crie uma inspeção, cadastre as características e use 'Salvar modelo da carta'.")
    else:
        termo_modelo = st.text_input("Pesquisar por característica, modelo, linha ou embalagem", placeholder="Ex.: pestana, diâmetro, GL, aerossol")
        linhas_modelos = consultar_caracteristicas_modelos(modelos, termo_modelo)
        if not linhas_modelos:
            st.info("Nenhuma característica localizada para o filtro informado.")
        else:
            df_consulta = pd.DataFrame(linhas_modelos)
            st.dataframe(
                df_consulta.drop(columns=["chave"]),
                use_container_width=True,
                hide_index=True,
                height=320,
            )
            opcoes = {
                f"{r['Característica']} | Modelo: {r['Modelo']} | LIE {r['LIE']} | LSE {r['LSE']}": r["chave"]
                for r in linhas_modelos
            }
            escolha = st.selectbox("Selecionar característica/modelo para reutilizar", list(opcoes.keys()))
            modelo_nome, idx_txt = opcoes[escolha].split("||")
            idx_char = int(idx_txt)
            c1, c2 = st.columns(2)
            if c1.button("Carregar modelo completo", use_container_width=True, type="primary"):
                aplicar_modelo(modelos[modelo_nome])
                st.session_state.carta_ok = False
                st.success("Modelo completo carregado. Complete os dados variáveis na aba 'Carta de dados' e salve a carta.")
                st.rerun()
            if c2.button("Incluir somente esta característica na inspeção atual", use_container_width=True, type="primary", disabled=not st.session_state.get("carta_ok", False)):
                ok, msg = incluir_caracteristica_modelo(modelo_nome, idx_char)
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)
            if not st.session_state.get("carta_ok", False):
                st.caption("Para incluir somente uma característica, primeiro salve a Carta de dados da nova inspeção.")

with tab1:
    st.subheader("Carta de dados — materiais e identificação")

    modelos = st.session_state.modelos
    if modelos:
        st.markdown("#### Utilizar modelo salvo")
        nomes_modelos = ["— Novo preenchimento sem modelo —"] + sorted(modelos.keys())
        modelo_sel = st.selectbox("Modelo de carta", nomes_modelos)
        c_load, c_del = st.columns(2)
        if c_load.button("Carregar modelo selecionado", use_container_width=True, type="primary", disabled=(modelo_sel.startswith("—"))):
            aplicar_modelo(modelos[modelo_sel])
            st.session_state.carta_ok = False
            st.success("Modelo carregado. Complete os dados variáveis da carta e salve para liberar a inspeção.")
            st.rerun()
        if c_del.button("Excluir modelo selecionado", use_container_width=True, disabled=(modelo_sel.startswith("—"))):
            modelos.pop(modelo_sel, None)
            save_modelos(modelos)
            st.session_state.modelos = modelos
            st.warning("Modelo excluído.")
            st.rerun()
    else:
        st.markdown("<div class='orange-help'>Nenhum modelo salvo ainda. Após criar as características, salve o modelo para reutilizar nas próximas inspeções.</div>", unsafe_allow_html=True)

    with st.form("form_carta"):
        st.markdown("#### Materiais")
        m1, m2 = st.columns(2)
        with m1:
            domo_mat = st.text_input("Usina domo", value=st.session_state.carta_dados.get("domo_mat", ""))
        with m2:
            esp_domo = st.text_input("Espessura domo *", value=st.session_state.carta_dados.get("esp_domo", ""))

        m3, m4 = st.columns(2)
        with m3:
            corpo_mat = st.text_input("Usina corpo", value=st.session_state.carta_dados.get("corpo_mat", ""))
        with m4:
            esp_corpo = st.text_input("Espessura corpo *", value=st.session_state.carta_dados.get("esp_corpo", ""))

        m5, m6 = st.columns(2)
        with m5:
            fundo_mat = st.text_input("Usina fundo", value=st.session_state.carta_dados.get("fundo_mat", ""))
        with m6:
            esp_fundo = st.text_input("Espessura fundo *", value=st.session_state.carta_dados.get("esp_fundo", ""))

        st.markdown("#### Dados do lote")
        l1, l2 = st.columns(2)
        with l1:
            linha = st.text_input("Linha *", value=st.session_state.carta_dados.get("linha", "GL"))
        with l2:
            embalagem = st.text_input("Embalagem *", value=st.session_state.carta_dados.get("embalagem", ""))

        l3, l4 = st.columns(2)
        with l3:
            op = st.text_input("Ordem de produção *", value=st.session_state.carta_dados.get("op", ""))
        with l4:
            lote_qtd = st.number_input("Quantidade / lote produzido *", min_value=0, step=1, value=int(st.session_state.carta_dados.get("lote_qtd", 0) or 0))

        data_carta = st.date_input("Data", value=date.today(), format="DD/MM/YYYY")

        st.markdown("#### Responsável pela inspeção")
        r1, r2 = st.columns(2)
        with r1:
            responsavel_nome = st.text_input("Nome do responsável pela inspeção *", value=st.session_state.carta_dados.get("responsavel_nome", ""))
        with r2:
            responsavel_chapa = st.text_input("Chapa do responsável pela inspeção *", value=st.session_state.carta_dados.get("responsavel_chapa", ""))

        obs = st.text_area("Observações gerais", value=st.session_state.carta_dados.get("obs", ""))
        submitted = st.form_submit_button("Salvar carta e liberar criação da inspeção", use_container_width=True)
    if submitted:
        obrig = [linha, embalagem, op, esp_corpo, esp_domo, esp_fundo, responsavel_nome, responsavel_chapa]
        if any(str(v).strip() == "" for v in obrig) or lote_qtd <= 0:
            st.error("Preencha todos os campos obrigatórios marcados com * e informe lote produzido maior que zero.")
        else:
            st.session_state.carta_ok = True
            st.session_state.carta_dados = {
                "linha": linha, "embalagem": embalagem, "op": op, "esp_corpo": esp_corpo,
                "esp_domo": esp_domo, "esp_fundo": esp_fundo, "lote_qtd": lote_qtd,
                "data": data_carta.strftime("%d/%m/%Y"), "domo_mat": domo_mat,
                "corpo_mat": corpo_mat, "fundo_mat": fundo_mat, "obs": obs,
                "responsavel_nome": responsavel_nome.strip(), "responsavel_chapa": responsavel_chapa.strip(),
            }
            save_current_cpk_state()
            st.success("Carta salva. A aba 'Criar inspeção' está liberada.")

    if st.session_state.carta_ok:
        st.markdown("<div class='card'><b>Carta ativa:</b> " + f"Linha {st.session_state.carta_dados.get('linha')} | Embalagem {st.session_state.carta_dados.get('embalagem')} | OP {st.session_state.carta_dados.get('op')} | Responsável {st.session_state.carta_dados.get('responsavel_nome')} - Chapa {st.session_state.carta_dados.get('responsavel_chapa')}" + "</div>", unsafe_allow_html=True)

with tab2:
    st.subheader("Criação das características de inspeção")
    if not st.session_state.carta_ok:
        st.warning("Primeiro salve a Carta de dados na aba 1.")
    else:
        st.markdown("<div class='orange-help'>O botão de criação fica em laranja para indicar inclusão/edição. Após a característica estar criada e disponível para uso, os botões de abertura e exportação aparecem em verde.</div>", unsafe_allow_html=True)

        modelos_rapidos = st.session_state.get("modelos", {})
        if modelos_rapidos:
            with st.expander("🔎 Puxar característica de modelo salvo", expanded=False):
                termo_rapido = st.text_input("Buscar característica salva", placeholder="Ex.: pestana, diâmetro, altura", key="busca_modelo_rapida")
                linhas_rapidas = consultar_caracteristicas_modelos(modelos_rapidos, termo_rapido)
                if linhas_rapidas:
                    opcoes_rapidas = {
                        f"{r['Característica']} | Modelo: {r['Modelo']} | LIE {r['LIE']} | LSE {r['LSE']}": r["chave"]
                        for r in linhas_rapidas
                    }
                    escolha_rapida = st.selectbox("Modelo/característica", list(opcoes_rapidas.keys()), key="sel_modelo_rapido")
                    modelo_nome_rapido, idx_txt_rapido = opcoes_rapidas[escolha_rapida].split("||")
                    if st.button("Incluir característica selecionada nesta inspeção", use_container_width=True, type="primary", key="btn_incluir_modelo_rapido"):
                        ok, msg = incluir_caracteristica_modelo(modelo_nome_rapido, int(idx_txt_rapido))
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                else:
                    st.info("Nenhuma característica encontrada nos modelos salvos para este filtro.")

        nonce = st.session_state.char_form_nonce
        with st.form(f"form_caracteristica_{nonce}"):
            descricao = st.text_input("Descrição da característica *", placeholder="Ex.: Diâmetro interno, pestana, altura, profundidade de expansão", key=f"desc_char_{nonce}")
            c1, c2 = st.columns(2)
            with c1:
                lie = st.text_input("Limite mínimo / LIE *", placeholder="Ex.: 52,10", key=f"lie_char_{nonce}")
            with c2:
                lse = st.text_input("Limite máximo / LSE *", placeholder="Ex.: 52,30", key=f"lse_char_{nonce}")
            c3, c4 = st.columns(2)
            with c3:
                num_amostras = st.number_input("Número de amostras que serão coletadas *", min_value=1, max_value=200, step=1, value=10, key=f"n_amostras_char_{nonce}")
            with c4:
                num_medicoes = st.number_input("Número de medições por amostra *", min_value=1, max_value=10, step=1, value=3, key=f"n_medicoes_char_{nonce}")
            submitted_char = st.form_submit_button("Salvar característica", use_container_width=True)
        if submitted_char:
            lie_v = parse_float(lie)
            lse_v = parse_float(lse)
            desc_norm = descricao_normalizada(descricao)
            duplicada = any(descricao_normalizada(c.get("descricao")) == desc_norm for c in st.session_state.caracteristicas)
            if not descricao.strip() or lie_v is None or lse_v is None or lie_v >= lse_v:
                st.error("Informe descrição, limite mínimo e limite máximo válidos. O limite mínimo deve ser menor que o limite máximo.")
            elif duplicada:
                st.error(f"A característica '{descricao.strip()}' já possui uma carta de inspeção criada. Não é permitido salvar duas cartas com a mesma descrição.")
            else:
                new_id = f"C{len(st.session_state.caracteristicas)+1:03d}_{datetime.now().strftime('%H%M%S')}"
                medicoes = [{"Amostra": i, **{f"Medida {j}": None for j in range(1, int(num_medicoes) + 1)}} for i in range(1, int(num_amostras)+1)]
                st.session_state.caracteristicas.append({
                    "id": new_id, "descricao": descricao.strip(), "lie": lie_v, "lse": lse_v,
                    "num_amostras": int(num_amostras), "num_medicoes": int(num_medicoes), "medicoes": medicoes,
                })
                st.session_state.selected_id = new_id
                st.session_state.char_form_nonce += 1
                save_current_cpk_state()
                st.success("Característica criada e habilitada para utilização. Os campos foram limpos para nova inclusão.")
                st.rerun()

        if st.session_state.caracteristicas:
            st.markdown("#### Características abertas")
            for char in st.session_state.caracteristicas:
                result = calc_characteristic(char)
                cols = st.columns([4, 1, 1, 1])
                cols[0].markdown(f"**{char['descricao']}**  \nLIE: {char['lie']} | LSE: {char['lse']} | Amostras: {char['num_amostras']} | Medições/amostra: {char.get('num_medicoes', 3)} | Medições: {result['medidas_realizadas']}/{result['medidas_previstas']}")
                cols[1].metric("Cpk", "—" if result["cpk"] is None else result["cpk"])
                cols[2].markdown(f"**Status:** {result['status']}")
                if cols[3].button("Abrir", key=f"open_{char['id']}", type="primary"):
                    st.session_state.selected_id = char["id"]
                    st.success(f"Característica selecionada: {char['descricao']}")

            st.markdown("#### Salvar modelo da carta")
            nome_modelo = st.text_input("Nome do modelo para reutilização", value=f"{st.session_state.carta_dados.get('linha','')}_{st.session_state.carta_dados.get('embalagem','')}".strip("_"))
            if st.button("Salvar modelo com estas características", use_container_width=True, type="primary"):
                if not nome_modelo.strip():
                    st.error("Informe um nome para o modelo.")
                else:
                    st.session_state.modelos[nome_modelo.strip()] = modelo_from_state(nome_modelo.strip())
                    save_modelos(st.session_state.modelos)
                    save_current_cpk_state()
                    st.success("Modelo salvo. Ele será carregado automaticamente como opção quando o aplicativo for reiniciado.")

with tab3:
    st.subheader("Registro das medições")
    if not st.session_state.caracteristicas:
        st.warning("Crie pelo menos uma característica na aba 2 ou carregue um modelo salvo.")
    else:
        options = {f"{c['descricao']} | {c['id']}": c["id"] for c in st.session_state.caracteristicas}
        current_key = next((k for k, v in options.items() if v == st.session_state.selected_id), list(options.keys())[0])
        selected_label = st.selectbox("Selecione a característica", list(options.keys()), index=list(options.keys()).index(current_key))
        st.session_state.selected_id = options[selected_label]
        char = next(c for c in st.session_state.caracteristicas if c["id"] == st.session_state.selected_id)
        keys_medicao = measurement_keys(char)
        st.info(f"Cada amostra deve conter obrigatoriamente {char.get('num_medicoes', 3)} medição(ões). Característica: {char['descricao']} | LIE {char['lie']} | LSE {char['lse']}")
        df_med = pd.DataFrame(char["medicoes"])
        expected_cols = ["Amostra"] + keys_medicao
        for col in expected_cols:
            if col not in df_med.columns:
                df_med[col] = None
        df_med = df_med[expected_cols]
        column_config = {"Amostra": st.column_config.NumberColumn("Amostra", disabled=True)}
        column_config.update({k: st.column_config.NumberColumn(k, format="%.4f") for k in keys_medicao})
        edited_med = st.data_editor(
            df_med,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config=column_config,
            key=f"editor_{char['id']}",
        )
        if st.button("Salvar medições desta característica", use_container_width=True, type="primary"):
            registros = edited_med.to_dict("records")
            incompletas = []
            for row in registros:
                vals = [parse_float(row.get(k)) for k in keys_medicao]
                if any(v is not None for v in vals) and not all(v is not None for v in vals):
                    incompletas.append(int(row.get("Amostra", 0)))
            if incompletas:
                st.error(f"As amostras {incompletas} estão parcialmente preenchidas. Cada amostra registrada deve ter {len(keys_medicao)} medição(ões).")
            else:
                char["medicoes"] = registros
                save_current_cpk_state()
                st.success("Medições salvas. A análise estatística já pode ser consultada na aba 4.")

with tab4:
    st.subheader("Análise estatística e cálculo do CPK")
    render_cpk_reference_card()
    if not st.session_state.caracteristicas:
        st.warning("Não há características criadas para análise.")
    else:
        todas_results = [calc_characteristic(c) for c in st.session_state.caracteristicas]
        chars_preenchidas = [c for c in st.session_state.caracteristicas if characteristic_has_measurements(c)]
        results = [calc_characteristic(c) for c in chars_preenchidas]
        if not results:
            st.warning("Nenhuma carta preenchida foi localizada. A análise estatística considera somente características com medições registradas.")
            st.stop()
        valid = [r for r in results if r.get("cpk") is not None]
        ok_n = len([r for r in valid if r["cpk"] >= 1.33])
        wn_n = len([r for r in valid if 1 <= r["cpk"] < 1.33])
        fl_n = len([r for r in valid if r["cpk"] < 1])
        avg_cpk = round(sum(r["cpk"] for r in valid) / len(valid), 2) if valid else None
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Cpk médio", "—" if avg_cpk is None else avg_cpk)
        c2.metric("Capazes ≥ 1,33", ok_n)
        c3.metric("Marginais", wn_n)
        c4.metric("Incapazes", fl_n)
        c5.metric("Cartas preenchidas", len(results))
        cartas_branco = len(todas_results) - len(results)
        if cartas_branco > 0:
            st.info(f"{cartas_branco} carta(s) em branco foram desconsideradas da análise estatística.")

        result_df = pd.DataFrame([{
            "Característica": r["descricao"], "Amostras completas": f"{r['amostras_completas']}/{r['amostras_previstas']}",
            "Medições/amostra": r.get("medicoes_por_amostra"), "Medições realizadas": r["medidas_realizadas"], "N": r["n"], "Média": r["media"],
            "Desvio": r["desvio"], "LIE": r["lie"], "LSE": r["lse"], "Cp": r["cp"],
            "CPS": r["cps"], "CPI": r["cpi"], "Cpk": r["cpk"], "Status": r["status"],
        } for r in results])
        st.dataframe(result_df, use_container_width=True, hide_index=True)

        st.markdown("#### Passo a passo do cálculo por característica")
        labels_calc = {f"{r['descricao']} | CPK {r['cpk'] if r.get('cpk') is not None else '—'}": i for i, r in enumerate(results)}
        label_calc = st.selectbox("Selecione a característica para visualizar o cálculo detalhado", list(labels_calc.keys()))
        result_calc = results[labels_calc[label_calc]]
        st.dataframe(cpk_formula_rows(result_calc), use_container_width=True, hide_index=True)
        st.info(interpretar_lado_critico(result_calc))

        st.markdown("#### Gráficos das coletas por característica")
        for r in results:
            fig = control_chart(r)
            if fig:
                st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### Análise e parecer automático")
        for ins in build_insights(results):
            st.markdown(f"<div class='{ins['level']}'>{ins['text']}</div>", unsafe_allow_html=True)

        st.markdown("#### Exportação")
        col_a, col_b = st.columns(2)
        with col_a:
            xlsx = io.BytesIO()
            with pd.ExcelWriter(xlsx, engine="openpyxl") as writer:
                pd.DataFrame([st.session_state.carta_dados]).to_excel(writer, sheet_name="CARTA", index=False)
                result_df.to_excel(writer, sheet_name="RESULTADOS", index=False)
                for c in chars_preenchidas:
                    safe = re.sub(r"[^A-Za-z0-9_]+", "_", c["descricao"][:20]) or c["id"]
                    pd.DataFrame(c["medicoes"]).to_excel(writer, sheet_name=safe[:31], index=False)
            st.download_button("Baixar Excel da carta", xlsx.getvalue(), file_name=f"carta_cpk_{datetime.now().strftime('%d_%m_%Y_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_b:
            pdf = make_pdf(st.session_state.carta_dados, results)
            if pdf:
                st.download_button("Baixar relatório PDF", pdf, file_name=f"relatorio_cpk_{datetime.now().strftime('%d_%m_%Y_%H%M')}.pdf", mime="application/pdf", use_container_width=True)
            else:
                st.warning("Inclua reportlab no requirements.txt para exportar PDF.")


with tab5:
    st.subheader("Backup e restauração da base CPK")
    st.markdown(
        "<div class='orange-help'>Use esta tela para baixar um Excel com todos os modelos salvos, carta atual, características, medições e resultados. "
        "Se o histórico local for perdido, faça upload deste mesmo arquivo para restaurar automaticamente a base do módulo CPK.</div>",
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)
    c1.metric("Modelos salvos", len(st.session_state.get("modelos", {}) or {}))
    c2.metric("Características abertas", len(st.session_state.get("caracteristicas", []) or []))
    c3.metric("Carta ativa", "Sim" if st.session_state.get("carta_ok") else "Não")

    st.download_button(
        "📥 Baixar backup Excel completo do CPK",
        data=cpk_backup_excel_bytes(),
        file_name=f"backup_cpk_completo_{datetime.now().strftime('%d_%m_%Y_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.markdown("#### Restaurar base a partir de backup")
    st.info(
        "Ao selecionar o arquivo Excel de backup, o aplicativo atualiza automaticamente a base CPK, "
        "incluindo modelos de carta, carta ativa, características, medições e resultados salvos."
    )
    arquivo_backup = st.file_uploader(
        "Enviar backup Excel gerado pelo módulo CPK",
        type=["xlsx"],
        key="upload_backup_cpk_auto",
    )
    if arquivo_backup is not None:
        backup_bytes = arquivo_backup.getvalue()
        backup_key = f"{arquivo_backup.name}_{len(backup_bytes)}"
        if st.session_state.get("ultimo_backup_cpk_restaurado") != backup_key:
            try:
                qtd_modelos, qtd_chars = restore_cpk_from_excel(backup_bytes)
                st.session_state["ultimo_backup_cpk_restaurado"] = backup_key
                st.success(
                    f"Base CPK atualizada automaticamente. Modelos restaurados: {qtd_modelos}. "
                    f"Características abertas restauradas: {qtd_chars}."
                )
                st.rerun()
            except Exception as exc:
                st.error(f"Não foi possível restaurar o backup CPK: {exc}")
        else:
            st.success("Este backup já foi aplicado nesta sessão. A base CPK está atualizada.")
