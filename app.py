import base64
import io
import json
import sqlite3
from datetime import date, datetime
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

APP_TITLE = "APP Integrado | Qualidade RS"
DB_PATH = Path(__file__).with_name("qualidade_integrado.db")

st.set_page_config(page_title=APP_TITLE, page_icon="✅", layout="wide", initial_sidebar_state="collapsed")

CSS = """
<style>
:root{
  --bg:#07130d; --card:#0d2417; --card2:#12341f; --accent:#80EF80; --text:#f3fff5; --muted:#b9d7bd;
  --danger:#ff5b5b; --warn:#ffd166; --ok:#57e389;
}
.stApp {background: radial-gradient(circle at top left,#1d5d32 0,#07130d 36%,#050905 100%); color:var(--text);} 
[data-testid="stHeader"] {background:rgba(0,0,0,0);} 
.block-container{padding-top:1.2rem; padding-bottom:3rem;}
.hero{padding:28px;border-radius:28px;background:linear-gradient(135deg,rgba(128,239,128,.22),rgba(13,36,23,.92));border:1px solid rgba(128,239,128,.26);box-shadow:0 18px 55px rgba(0,0,0,.34);}
.hero h1{margin:0;font-size:2.2rem;letter-spacing:-.04em;color:#fff;}
.hero p{color:var(--muted);font-size:1rem;margin:.5rem 0 0 0;}
.card{padding:22px;border-radius:24px;background:rgba(13,36,23,.82);border:1px solid rgba(128,239,128,.18);box-shadow:0 10px 30px rgba(0,0,0,.25);height:100%;}
.card h3{margin-top:0;color:#fff}.card p{color:var(--muted)}
.metric-card{padding:18px;border-radius:20px;background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.12)}
.metric-card .label{font-size:.82rem;color:var(--muted)}.metric-card .value{font-size:1.8rem;font-weight:800;color:#fff}
.status-ok{color:#061d0c;background:#80EF80;padding:4px 10px;border-radius:999px;font-weight:700}
.status-warn{color:#332600;background:#ffd166;padding:4px 10px;border-radius:999px;font-weight:700}
.status-bad{color:#3a0505;background:#ff8080;padding:4px 10px;border-radius:999px;font-weight:700}
.small-muted{color:var(--muted);font-size:.84rem}
.stButton>button{border-radius:14px;border:1px solid rgba(128,239,128,.35);background:rgba(128,239,128,.12);color:#fff;font-weight:700;min-height:44px}
.stButton>button:hover{border-color:#80EF80;background:rgba(128,239,128,.25);color:#fff}
[data-testid="stMetric"]{background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.11);padding:14px;border-radius:18px}
hr{border-color:rgba(128,239,128,.22)}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

MES_PT = {1:"jan",2:"fev",3:"mar",4:"abr",5:"mai",6:"jun",7:"jul",8:"ago",9:"set",10:"out",11:"nov",12:"dez"}

# ----------------------------- Persistence -----------------------------
def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

conn = get_conn()

def init_db():
    cur = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS app_state(
        key TEXT PRIMARY KEY, value TEXT, updated_at TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS files(
        module TEXT PRIMARY KEY, filename TEXT, content BLOB, updated_at TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS sqdcp_records(
        id INTEGER PRIMARY KEY AUTOINCREMENT, data TEXT, turno TEXT, dimensao TEXT,
        indicador TEXT, valor REAL, meta REAL, observacao TEXT, updated_at TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS actions(
        id INTEGER PRIMARY KEY AUTOINCREMENT, origem TEXT, indicador TEXT, descricao TEXT,
        responsavel TEXT, prazo TEXT, status TEXT, evidencia TEXT, updated_at TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS cpk_cards(
        id INTEGER PRIMARY KEY AUTOINCREMENT, caracteristica TEXT UNIQUE, material_info TEXT,
        lote_info TEXT, lsl REAL, usl REAL, n_amostras INTEGER, medicoes_amostra INTEGER, created_at TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS cpk_measurements(
        id INTEGER PRIMARY KEY AUTOINCREMENT, card_id INTEGER, amostra INTEGER, medicao INTEGER, valor REAL,
        updated_at TEXT, UNIQUE(card_id, amostra, medicao))""")
    conn.commit()

init_db()

def now(): return datetime.now().isoformat(timespec="seconds")

def set_state(key, value):
    conn.execute("REPLACE INTO app_state(key,value,updated_at) VALUES(?,?,?)", (key, json.dumps(value, ensure_ascii=False), now()))
    conn.commit()

def get_state(key, default=None):
    row = conn.execute("SELECT value FROM app_state WHERE key=?", (key,)).fetchone()
    return json.loads(row["value"]) if row else default

def save_file(module, filename, content):
    conn.execute("REPLACE INTO files(module,filename,content,updated_at) VALUES(?,?,?,?)", (module, filename, content, now()))
    conn.commit()

def load_file(module):
    return conn.execute("SELECT * FROM files WHERE module=?", (module,)).fetchone()

# ----------------------------- UI helpers -----------------------------
def header(title, subtitle=""):
    st.markdown(f"""<div class='hero'><h1>{title}</h1><p>{subtitle}</p></div>""", unsafe_allow_html=True)
    st.write("")

def go_home_button(key="home"):
    if st.button("⬅️ Retornar à tela inicial", key=key):
        st.session_state.page = "home"
        st.rerun()

def set_page(page):
    st.session_state.page = page
    set_state("last_page", page)
    st.rerun()

def safe_date(s):
    return pd.to_datetime(s, errors="coerce", dayfirst=True)

def download_excel(df, sheet="DADOS"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheet, index=False)
    return output.getvalue()

# ----------------------------- Home -----------------------------
def home():
    header("APP Integrado Qualidade RS", "Um único ambiente para Indicadores de Qualidade, SQDCP e Projeto CPK, com persistência automática e navegação por módulos.")
    c1, c2, c3 = st.columns(3)
    cards = [
        (c1, "📊 Indicadores de Qualidade", "Reclamações, atrasos, responsáveis, motivos, filtros e exportação de recortes.", "qualidade"),
        (c2, "🟩 SQDCP", "Gestão diária de Segurança, Qualidade, Delivery, Custo e Processo, com sinaleira e plano de ação.", "sqdcp"),
        (c3, "📐 Projeto CPK", "Cartas de inspeção, medições por amostra, Cp, Cpk, média e análise estatística visual.", "cpk"),
    ]
    for col, title, desc, page in cards:
        with col:
            st.markdown(f"<div class='card'><h3>{title}</h3><p>{desc}</p></div>", unsafe_allow_html=True)
            if st.button(f"Abrir {title.split(' ',1)[1]}", key=f"open_{page}", use_container_width=True):
                set_page(page)
    st.write("")
    st.markdown("### Status de persistência")
    f = load_file("qualidade")
    qtd_sqdcp = conn.execute("SELECT COUNT(*) c FROM sqdcp_records").fetchone()["c"]
    qtd_cpk = conn.execute("SELECT COUNT(*) c FROM cpk_cards").fetchone()["c"]
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Base Qualidade", f["filename"] if f else "Não carregada")
    m2.metric("Registros SQDCP", qtd_sqdcp)
    m3.metric("Cartas CPK", qtd_cpk)
    m4.metric("Banco local", DB_PATH.name)

# ----------------------------- Qualidade -----------------------------
def qualidade():
    header("Indicadores de Qualidade", "Carregamento de consulta do sistema de gestão, análise por período, responsáveis, motivos e recorte de ocorrências.")
    go_home_button("home_qualidade")
    uploaded = st.file_uploader("Carregar nova base Excel de consultas/RNC", type=["xlsx","xls"], key="qual_up")
    if uploaded:
        save_file("qualidade", uploaded.name, uploaded.getvalue())
        st.success("Base salva com persistência. Ela será recarregada automaticamente no próximo acesso.")
    row = load_file("qualidade")
    if not row:
        st.info("Carregue uma base Excel para iniciar. A última base ficará salva automaticamente.")
        return
    st.caption(f"Base atual: {row['filename']} | Atualizada em {row['updated_at']}")
    try:
        df = pd.read_excel(io.BytesIO(row["content"]))
    except Exception as e:
        st.error(f"Não foi possível ler a base: {e}")
        return
    if df.empty:
        st.warning("Base sem registros.")
        return
    cols = df.columns.tolist()
    col_data = next((c for c in cols if "data" in c.lower() and ("emiss" in c.lower() or "abert" in c.lower() or c.lower()=="data")), None)
    col_resp = next((c for c in cols if "responsável da análise" in c.lower()), None) or next((c for c in cols if "respons" in c.lower()), None)
    col_motivo = next((c for c in cols if "motivo" in c.lower()), None) or next((c for c in cols if "título" in c.lower()), None)
    col_status = next((c for c in cols if "situa" in c.lower() or "status" in c.lower()), None)
    if col_data:
        df[col_data] = safe_date(df[col_data])
        df["Ano"] = df[col_data].dt.year
        df["Mes"] = df[col_data].dt.month
        df["Mês/ano"] = df[col_data].dt.strftime("%m/%Y")
        df["Semana"] = df[col_data].dt.isocalendar().week.astype("Int64")
    else:
        st.warning("Não encontrei coluna de data. Os filtros de tempo foram desativados.")
    with st.expander("Filtros", expanded=True):
        fc1,fc2,fc3 = st.columns(3)
        dff = df.copy()
        if col_data:
            anos = sorted([int(a) for a in dff["Ano"].dropna().unique() if int(a) >= 2025])
            anos_sel = fc1.multiselect("Ano", anos, default=anos[-1:] if anos else [])
            if anos_sel: dff = dff[dff["Ano"].isin(anos_sel)]
            meses = sorted(dff["Mes"].dropna().unique().astype(int).tolist())
            mes_sel = fc2.multiselect("Mês", meses, format_func=lambda x: MES_PT.get(x,x))
            if mes_sel: dff = dff[dff["Mes"].isin(mes_sel)]
        if col_resp:
            resps = sorted(dff[col_resp].fillna("Avaliação").astype(str).unique().tolist())
            resp_sel = fc3.multiselect("Responsável", resps)
            if resp_sel: dff = dff[dff[col_resp].fillna("Avaliação").astype(str).isin(resp_sel)]
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Ocorrências", len(dff))
    atrasadas = dff[col_status].astype(str).str.contains("ATRAS|VENC", case=False, na=False).sum() if col_status else 0
    k2.metric("Em atraso", int(atrasadas))
    k3.metric("Responsáveis", dff[col_resp].nunique() if col_resp else "-")
    k4.metric("Motivos", dff[col_motivo].nunique() if col_motivo else "-")
    g1,g2 = st.columns(2)
    with g1:
        if col_data:
            xcol = "Mês/ano" if dff["Ano"].nunique() > 1 else "Mes"
            tmp = dff.groupby(xcol).size().reset_index(name="Ocorrências")
            fig = px.bar(tmp, x=xcol, y="Ocorrências", text="Ocorrências", title="Ocorrências por período")
            fig.update_traces(textposition="outside")
            fig.update_layout(yaxis=dict(dtick=1), xaxis_tickangle=-45, height=420)
            st.plotly_chart(fig, use_container_width=True)
    with g2:
        if col_motivo:
            tmp = dff[col_motivo].fillna("Não informado").astype(str).value_counts().head(12).reset_index()
            tmp.columns=["Motivo","Ocorrências"]
            fig = px.bar(tmp, x="Ocorrências", y="Motivo", orientation="h", text="Ocorrências", title="Top motivos")
            fig.update_layout(height=420, yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig, use_container_width=True)
    g3,g4 = st.columns(2)
    with g3:
        if col_resp:
            tmp = dff[col_resp].fillna("Avaliação").astype(str).value_counts().reset_index()
            tmp.columns=["Responsável","Ocorrências"]
            fig = px.bar(tmp, x="Responsável", y="Ocorrências", text="Ocorrências", title="Ocorrências por responsável")
            fig.update_layout(xaxis_tickangle=-45, height=390)
            st.plotly_chart(fig, use_container_width=True)
    with g4:
        if col_status:
            tmp = dff[col_status].fillna("Não informado").astype(str).value_counts().reset_index()
            tmp.columns=["Status","Ocorrências"]
            fig = px.bar(tmp, x="Status", y="Ocorrências", text="Ocorrências", title="Status / Situação")
            fig.update_layout(height=390)
            st.plotly_chart(fig, use_container_width=True)
    st.markdown("### Recorte de ocorrências")
    st.dataframe(dff, use_container_width=True, height=360)
    st.download_button("Exportar recorte para Excel", download_excel(dff, "RECORTE"), "recorte_qualidade.xlsx")

# ----------------------------- SQDCP -----------------------------
def sqdcp():
    header("SQDCP", "Gestão visual diária com sinaleira, metas, série histórica e plano de ação integrado.")
    go_home_button("home_sqdcp")
    dims = {"S":"Segurança", "Q":"Qualidade", "D":"Delivery", "C":"Custo", "P":"Processo"}
    tab1, tab2, tab3 = st.tabs(["Lançamentos", "Dashboard", "Plano de ação"])
    with tab1:
        with st.form("form_sqdcp", clear_on_submit=True):
            c1,c2,c3,c4 = st.columns(4)
            data = c1.date_input("Data", value=date.today())
            turno = c2.selectbox("Turno", ["1º","2º","3º"])
            dim = c3.selectbox("Dimensão", list(dims.keys()), format_func=lambda x:f"{x} - {dims[x]}")
            indicador = c4.text_input("Indicador", placeholder="Ex.: Reclamações, Acidentes, Eficiência")
            c5,c6 = st.columns(2)
            valor = c5.number_input("Valor realizado", value=0.0, step=1.0)
            meta = c6.number_input("Meta", value=0.0, step=1.0)
            obs = st.text_area("Observação")
            if st.form_submit_button("Salvar lançamento"):
                conn.execute("INSERT INTO sqdcp_records(data,turno,dimensao,indicador,valor,meta,observacao,updated_at) VALUES(?,?,?,?,?,?,?,?)", (str(data),turno,dim,indicador,valor,meta,obs,now()))
                conn.commit(); st.success("Lançamento salvo.")
    with tab2:
        df = pd.read_sql_query("SELECT * FROM sqdcp_records", conn)
        if df.empty:
            st.info("Ainda não há lançamentos SQDCP.")
        else:
            df["data"] = pd.to_datetime(df["data"])
            df["mês/ano"] = df["data"].dt.strftime("%m/%Y")
            latest = df.sort_values("data").groupby(["dimensao","indicador"]).tail(1)
            cols = st.columns(5)
            for i, d in enumerate(dims):
                sub = latest[latest["dimensao"]==d]
                total_bad = int((sub["valor"] > sub["meta"]).sum()) if not sub.empty else 0
                status = "🟢" if total_bad == 0 else "🔴"
                cols[i].metric(f"{status} {d} - {dims[d]}", f"{len(sub)} ind.", f"{total_bad} fora da meta")
            st.markdown("### Série histórica")
            sel = st.selectbox("Indicador para visualizar", sorted(df["indicador"].dropna().unique()))
            his = df[df["indicador"]==sel].groupby("mês/ano", as_index=False)[["valor","meta"]].mean()
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=his["mês/ano"], y=his["valor"], mode="lines+markers", name="Realizado"))
            fig.add_trace(go.Scatter(x=his["mês/ano"], y=his["meta"], mode="lines+markers", name="Meta"))
            fig.update_layout(height=420, xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(df.sort_values("data", ascending=False), use_container_width=True, height=320)
    with tab3:
        action_form("SQDCP")
        show_actions("SQDCP")

# ----------------------------- Actions -----------------------------
def action_form(origem):
    st.markdown("### Nova ação")
    with st.form(f"action_{origem}", clear_on_submit=True):
        c1,c2,c3 = st.columns(3)
        indicador = c1.text_input("Indicador / tema")
        responsavel = c2.text_input("Responsável")
        prazo = c3.date_input("Prazo", value=date.today())
        desc = st.text_area("Descrição da ação")
        c4,c5 = st.columns(2)
        status = c4.selectbox("Status", ["Aberto", "Concluído", "Atrasado", "Sem registro"])
        evidencia = c5.text_input("Evidência")
        if st.form_submit_button("Salvar ação"):
            conn.execute("INSERT INTO actions(origem,indicador,descricao,responsavel,prazo,status,evidencia,updated_at) VALUES(?,?,?,?,?,?,?,?)", (origem, indicador, desc, responsavel, str(prazo), status, evidencia, now()))
            conn.commit(); st.success("Ação salva.")

def show_actions(origem):
    df = pd.read_sql_query("SELECT * FROM actions WHERE origem=?", conn, params=(origem,))
    if df.empty:
        st.info("Nenhuma ação cadastrada.")
        return
    def badge(s):
        if s == "Concluído": return "🟢 Concluído"
        if s == "Atrasado": return "🔴 Atrasado"
        if s == "Aberto": return "🟡 Aberto"
        return "⚪ Sem registro"
    df["Sinaleira"] = df["status"].map(badge)
    st.dataframe(df[["Sinaleira","indicador","descricao","responsavel","prazo","evidencia","updated_at"]], use_container_width=True, height=360)
    st.download_button("Exportar plano de ação", download_excel(df, "AÇÕES"), f"acoes_{origem.lower()}.xlsx")

# ----------------------------- CPK -----------------------------
def calc_cp_cpk(vals, lsl, usl):
    vals = pd.Series(vals).dropna().astype(float)
    if len(vals) < 2: return np.nan, np.nan, np.nan, np.nan
    mean = vals.mean(); sd = vals.std(ddof=1)
    if sd == 0 or np.isnan(sd): return mean, sd, np.nan, np.nan
    cp = (usl-lsl)/(6*sd)
    cpk = min((usl-mean)/(3*sd), (mean-lsl)/(3*sd))
    return mean, sd, cp, cpk

def cpk():
    header("Projeto CPK", "Cadastro de cartas de inspeção, coleta de medições e análise de capacidade do processo.")
    go_home_button("home_cpk")
    tab1, tab2, tab3 = st.tabs(["Criar carta", "Coletar medições", "Análise estatística"])
    with tab1:
        with st.form("card_cpk", clear_on_submit=True):
            st.markdown("#### Material")
            m1,m2,m3,m4 = st.columns(4)
            usina_domo = m1.text_input("Usina domo")
            esp_domo = m2.text_input("Espessura domo")
            usina_corpo = m3.text_input("Usina corpo")
            esp_corpo = m4.text_input("Espessura corpo")
            st.markdown("#### Lote")
            l1,l2,l3,l4 = st.columns(4)
            linha = l1.text_input("Linha")
            embalagem = l2.text_input("Embalagem")
            op = l3.text_input("Ordem de produção")
            qtd = l4.text_input("Quantidade")
            st.markdown("#### Característica")
            c1,c2,c3,c4 = st.columns(4)
            car = c1.text_input("Descrição da característica")
            lsl = c2.number_input("Limite inferior", value=0.0, format="%.4f")
            usl = c3.number_input("Limite superior", value=0.0, format="%.4f")
            n_am = c4.number_input("Nº de amostras", min_value=1, value=5, step=1)
            med = st.number_input("Nº de medições por amostra", min_value=1, value=1, step=1)
            if st.form_submit_button("Salvar carta"):
                if not car.strip():
                    st.error("Informe a descrição da característica.")
                elif conn.execute("SELECT 1 FROM cpk_cards WHERE lower(caracteristica)=lower(?)", (car.strip(),)).fetchone():
                    st.error("Esta característica já possui carta de inspeção cadastrada.")
                elif usl <= lsl:
                    st.error("O limite superior deve ser maior que o limite inferior.")
                else:
                    material = json.dumps({"usina_domo":usina_domo,"esp_domo":esp_domo,"usina_corpo":usina_corpo,"esp_corpo":esp_corpo}, ensure_ascii=False)
                    lote = json.dumps({"linha":linha,"embalagem":embalagem,"op":op,"qtd":qtd}, ensure_ascii=False)
                    conn.execute("INSERT INTO cpk_cards(caracteristica,material_info,lote_info,lsl,usl,n_amostras,medicoes_amostra,created_at) VALUES(?,?,?,?,?,?,?,?)", (car.strip(), material, lote, lsl, usl, int(n_am), int(med), now()))
                    conn.commit(); st.success("Carta salva. Os campos foram limpos automaticamente.")
    with tab2:
        cards = pd.read_sql_query("SELECT * FROM cpk_cards ORDER BY caracteristica", conn)
        if cards.empty:
            st.info("Cadastre uma carta de inspeção.")
        else:
            card_id = st.selectbox("Carta", cards["id"], format_func=lambda x: cards.loc[cards["id"]==x,"caracteristica"].iloc[0])
            card = cards[cards["id"]==card_id].iloc[0]
            existing = pd.read_sql_query("SELECT * FROM cpk_measurements WHERE card_id=?", conn, params=(int(card_id),))
            with st.form("measure_form"):
                values = []
                for a in range(1, int(card.n_amostras)+1):
                    cols = st.columns(int(card.medicoes_amostra))
                    for m in range(1, int(card.medicoes_amostra)+1):
                        old = existing[(existing["amostra"]==a)&(existing["medicao"]==m)]["valor"]
                        default = float(old.iloc[0]) if not old.empty else 0.0
                        values.append((a,m,cols[m-1].number_input(f"Amostra {a} | Medição {m}", value=default, format="%.4f", key=f"v_{card_id}_{a}_{m}")))
                if st.form_submit_button("Salvar medições"):
                    for a,m,v in values:
                        conn.execute("INSERT OR REPLACE INTO cpk_measurements(card_id,amostra,medicao,valor,updated_at) VALUES(?,?,?,?,?)", (int(card_id),a,m,float(v),now()))
                    conn.commit(); st.success("Medições salvas com persistência.")
    with tab3:
        cards = pd.read_sql_query("SELECT * FROM cpk_cards ORDER BY caracteristica", conn)
        if cards.empty:
            st.info("Nenhuma carta cadastrada.")
            return
        rows = []
        for _, card in cards.iterrows():
            med = pd.read_sql_query("SELECT valor FROM cpk_measurements WHERE card_id=?", conn, params=(int(card.id),))
            filled = len(med.dropna())
            expected = int(card.n_amostras) * int(card.medicoes_amostra)
            if filled == 0: continue
            mean, sd, cp, cpkv = calc_cp_cpk(med["valor"], card.lsl, card.usl)
            rows.append({"id":card.id,"Característica":card.caracteristica,"Média":mean,"Desvio":sd,"Cp":cp,"Cpk":cpkv,"Preenchimento":f"{filled}/{expected}","LSL":card.lsl,"USL":card.usl})
        if not rows:
            st.info("A análise mostra apenas cartas preenchidas. Registre medições para visualizar.")
            return
        res = pd.DataFrame(rows)
        st.dataframe(res.drop(columns="id"), use_container_width=True)
        sel = st.selectbox("Visualizar gráfico", res["id"], format_func=lambda x: res.loc[res["id"]==x,"Característica"].iloc[0])
        r = res[res["id"]==sel].iloc[0]
        med = pd.read_sql_query("SELECT amostra, medicao, valor FROM cpk_measurements WHERE card_id=? ORDER BY amostra, medicao", conn, params=(int(sel),))
        med["Coleta"] = med["amostra"].astype(str) + "." + med["medicao"].astype(str)
        k1,k2,k3,k4 = st.columns(4)
        k1.metric("Cpk", f"{r['Cpk']:.3f}" if pd.notna(r['Cpk']) else "-")
        k2.metric("Cp", f"{r['Cp']:.3f}" if pd.notna(r['Cp']) else "-")
        k3.metric("Média", f"{r['Média']:.4f}")
        k4.metric("Desvio padrão", f"{r['Desvio']:.4f}" if pd.notna(r['Desvio']) else "-")
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=med["Coleta"], y=med["valor"], mode="lines+markers", name="Medições"))
        fig.add_hline(y=r["LSL"], line_color="red", line_width=3, annotation_text="Limite inferior")
        fig.add_hline(y=r["USL"], line_color="red", line_width=3, annotation_text="Limite superior")
        fig.update_layout(title=f"Carta de análise — {r['Característica']}", height=470, xaxis_title="Amostra.Medição", yaxis_title="Valor")
        st.plotly_chart(fig, use_container_width=True)
        st.download_button("Exportar análise CPK", download_excel(res.drop(columns="id"), "CPK"), "analise_cpk.xlsx")

# ----------------------------- Router -----------------------------
if "page" not in st.session_state:
    st.session_state.page = get_state("last_page", "home")

page = st.session_state.page
if page == "qualidade": qualidade()
elif page == "sqdcp": sqdcp()
elif page == "cpk": cpk()
else: home()
