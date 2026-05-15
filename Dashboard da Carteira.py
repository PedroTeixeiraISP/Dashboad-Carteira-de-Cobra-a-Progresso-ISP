import io
from pathlib import Path
from datetime import datetime
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ==========================================
# VALIDAÇÃO POR DOMÍNIO DE E-MAIL (SEM SENHA)
# ==========================================
def tela_de_login():
    """Renderiza uma interface de validação por e-mail institucional."""
    st.markdown(
        """
        <style>
            .login-container {
                max-width: 450px;
                margin: 4rem auto 1rem auto;
                padding: 30px;
                background: white;
                border-radius: 16px;
                box-shadow: 0 10px 25px rgba(0,0,0,0.05);
                border: 1px solid rgba(11,107,83,0.1);
            }
            .login-title {
                color: #084C3D;
                font-weight: 800;
                font-size: 1.6rem;
                margin-bottom: 5px;
                text-align: center;
            }
            .login-subtitle {
                color: #667085;
                font-size: 0.9rem;
                margin-bottom: 10px;
                text-align: center;
            }
        </style>
        """,
        unsafe_allow_html=True
    )
    
    # Inicializa variáveis de sessão
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False
        st.session_state.usuario_email = ""

    if not st.session_state.autenticado:
        st.markdown(
            """
            <div class="login-container">
                <div class="login-title">🔐 Acesso Institucional</div>
                <div class="login-subtitle">Insira seu e-mail corporativo para acessar o painel</div>
            </div>
            """, 
            unsafe_allow_html=True
        )
        
        c1, c2, c3 = st.columns([1, 3, 1])
        with c2:
            # Entrada de texto simples fora de formulários complexos para evitar loops de cache
            email_input = st.text_input("E-mail corporativo", placeholder="seuemail@ispschools.com").strip().lower()
            bt_entrar = st.button("Verificar e Acessar", use_container_width=True)
            
            if bt_entrar and email_input:
                # Validação direta do sufixo do e-mail
                if email_input.endswith("@ispschools.com") or email_input.endswith("@colegioprogresso.com.br"):
                    st.session_state.autenticado = True
                    st.session_state.usuario_email = email_input
                    st.success("Acesso autorizado! Carregando...")
                    st.rerun()
                else:
                    st.error("E-mail não autorizado. Utilize um domínio @ispschools.com ou @colegioprogresso.com.br")
                    
        st.stop() # Bloqueia o app caso o e-mail não seja válido

# Executa a validação de domínio
tela_de_login()

# ==========================================
# INÍCIO DO CÓDIGO DO DASHBOARD (SÓ RODA SE LOGADO)
# ==========================================

# 1. Configuração da Página
st.set_page_config(
    page_title="Dashboard da Carteira | Online",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

SHEET_NAME = "Base Teste"
DEFAULT_SHAREPOINT_URL = "https://isp-my.sharepoint.com/:x:/r/personal/pteixeira_ispschools_com/Documents/Base%20de%20cobran%C3%A7a%20-%20Teste.xlsm?d=wec297335739b4b388d4a129b6a95e733&csf=1&web=1&e=VjdoE0&nav=MTVfe0FDOEM2MkQ4LUUyOEQtNDAxMC05MEUyLTE1ODYyQkMwODA5Mn0"

# Cores de Identidade Visual
ISP_GREEN = "#0B6B53"
ISP_GREEN_DARK = "#084C3D"
ISP_GREEN_SOFT = "#DFF3EC"
ISP_GREEN_MID = "#1F8A70"
CARD = "#FFFFFF"
TEXT = "#1F2937"
MUTED = "#667085"
BORDER = "rgba(11,107,83,0.10)"

# Estilização CSS personalizada
st.markdown(
    f"""
    <style>
        .stApp {{background: linear-gradient(180deg, #F7FAF8 0%, #EEF5F2 100%);}}
        .block-container {{max-width: 96%; padding-top: 1.1rem; padding-bottom: 1.5rem;}}
        section[data-testid="stSidebar"] {{background: linear-gradient(180deg, {ISP_GREEN_DARK} 0%, {ISP_GREEN} 100%);}}
        section[data-testid="stSidebar"] * {{color: white !important;}}
        .hero {{
            background: linear-gradient(135deg, {ISP_GREEN_DARK} 0%, {ISP_GREEN_MID} 100%);
            border-radius: 22px;
            padding: 22px 26px;
            color: white;
            box-shadow: 0 12px 34px rgba(8,76,61,0.20);
            margin-bottom: 1rem;
        }}
        .hero h1 {{margin: 0 0 4px 0; font-size: 1.9rem;}}
        .hero p {{margin: 0; opacity: 0.96;}}
        .kpi-card {{
            background: {CARD};
            border: 1px solid {BORDER};
            border-radius: 18px;
            padding: 18px 18px 16px 18px;
            box-shadow: 0 8px 28px rgba(11,107,83,0.08);
            min-height: 118px;
        }}
        .kpi-label {{font-size: 0.86rem; font-weight: 700; color: {MUTED}; margin-bottom: 8px;}}
        .kpi-value {{font-size: 1.9rem; line-height: 1.08; font-weight: 800; color: {ISP_GREEN_DARK};}}
        .kpi-foot {{font-size: 0.82rem; color: {MUTED}; margin-top: 8px;}}
        .section-title {{font-size: 1.08rem; font-weight: 800; color: {ISP_GREEN_DARK}; margin: .25rem 0 .85rem 0;}}
        div[data-testid="stMetric"] {{
            background: white;
            border: 1px solid {BORDER};
            border-radius: 16px;
            padding: 10px 12px;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

# 2. Funções Utilitárias e de Formatação
def make_downloadable_sharepoint_url(url: str) -> str:
    parsed = urlparse(url)
    query = parse_qs(parsed.query)
    query.pop("web", None)
    query["download"] = ["1"]
    new_query = urlencode(query, doseq=True)
    return urlunparse((parsed.scheme, parsed.netloc, parsed.path, parsed.params, new_query, parsed.fragment))


def brl(value: float) -> str:
    value = 0 if pd.isna(value) else value
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def brl_short(value: float) -> str:
    value = 0 if pd.isna(value) else value
    v = abs(value)
    if v >= 1_000_000_000:
        txt = f"R$ {value/1_000_000_000:.2f} bi"
    elif v >= 1_000_000:
        txt = f"R$ {value/1_000_000:.2f} mi"
    elif v >= 1_000:
        txt = f"R$ {value/1_000:.1f} mil"
    else:
        txt = f"R$ {value:,.0f}"
    return txt.replace(",", "X").replace(".", ",").replace("X", ".")


def pct(part, total) -> str:
    return f"{(part / total):.1%}" if total else "0,0%"


def parse_valor(serie: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce")
    limpa = (
        serie.astype(str)
        .str.replace("R$", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(limpa, errors="coerce")


def get_sheet_names(file_source):
    with pd.ExcelFile(file_source, engine="openpyxl") as xls:
        return xls.sheet_names


@st.cache_data(show_spinner=False)
def carregar_dados(file_source, sheet_name):
    with pd.ExcelFile(file_source, engine="openpyxl") as xls:
        abas = xls.sheet_names
        if sheet_name not in abas:
            raise ValueError(
                f"A aba '{sheet_name}' não foi encontrada. Abas disponíveis: {', '.join(abas)}"
            )
        df = pd.read_excel(xls, sheet_name=sheet_name)

    df.columns = [str(c).strip() for c in df.columns]

    colunas_esperadas = [
        "Responsável", "UNIDADE", "Vencimento", "Valor", "Lançamento",
        "Status", "Histórico de Acionamento", "Classe de Risco", "Bloqueado"
    ]
    
    if "Unidade" in df.columns and "UNIDADE" not in df.columns:
        df = df.rename(columns={"Unidade": "UNIDADE"})

    for col in colunas_esperadas:
        if col not in df.columns:
            df[col] = np.nan

    df["Vencimento"] = pd.to_datetime(df["Vencimento"], errors="coerce")
    df["Lançamento"] = pd.to_datetime(df["Lançamento"], errors="coerce")
    df["Valor"] = parse_valor(df["Valor"]).fillna(0)

    for col in ["Responsável", "UNIDADE", "Status", "Classe de Risco", "Histórico de Acionamento"]:
        df[col] = df[col].fillna("").astype(str).str.strip()

    df["Responsável"] = df["Responsável"].replace("", "Não informado")
    df["UNIDADE"] = df["UNIDADE"].replace("", "Não informado")
    df["Classe de Risco"] = df["Classe de Risco"].replace("", "Não informado")

    hoje = pd.Timestamp.today().normalize()
    status_lower = df["Status"].str.lower()

    df["Em Aberto"] = ~status_lower.isin(["pago", "liquidado", "baixado", "quitado", "cancelado"])
    df["Vencido"] = df["Vencimento"].notna() & (df["Vencimento"] < hoje)
    df["Bloqueado"] = df["Bloqueado"].fillna(False).astype(str).str.strip().str.lower().isin(["sim", "true", "1", "s"])
    
    df["Dias em Atraso"] = np.where(df["Vencimento"].notna(), (hoje - df["Vencimento"]).dt.days, np.nan)
    df["Mês Vencimento"] = df["Vencimento"].dt.strftime("%Y-%m")
    df.loc[df["Vencimento"].isna(), "Mês Vencimento"] = "Sem data"

    return df


def montar_tabela(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Unidade", "Responsável", "Classe de Risco", "Valor em Aberto", "Vencimento Mais Antigo", "Último Acionamento", "Observação"])

    base = df.copy()

    def concatenar_status(series):
        valores = [str(x).strip() for x in series if str(x).strip() != ""]
        valores_unicos = list(dict.fromkeys(valores))
        return " | ".join(valores_unicos) if valores_unicos else "Não informado"

    def concatenar_historico(series):
        valores = [str(x).strip() for x in series if str(x).strip() != ""]
        valores_unicos = list(dict.fromkeys(valores))
        return " | ".join(valores_unicos) if valores_unicos else ""

    tabela = (
        base.groupby(["UNIDADE", "Responsável", "Classe de Risco"], dropna=False)
        .agg(
            **{
                "Valor em Aberto": ("Valor", "sum"),
                "Vencimento Mais Antigo": ("Vencimento", "min"),
                "Último Acionamento": ("Status", concatenar_status),
                "Observação": ("Histórico de Acionamento", concatenar_historico)
            }
        )
        .reset_index()
        .sort_values("Valor em Aberto", ascending=False)
    )

    tabela["Vencimento Mais Antigo"] = tabela["Vencimento Mais Antigo"].dt.strftime("%d/%m/%Y").fillna("Sem data")
    tabela = tabela.rename(columns={"UNIDADE": "Unidade"})
    return tabela


def excel_bytes(abas: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for nome, tabela in abas.items():
            tabela.to_excel(writer, index=False, sheet_name=nome[:31])
    return output.getvalue()


# 3. Cabeçalho da Aplicação
st.markdown(
    f"""
    <div class='hero'>
        <h1>Dashboard da Carteira de Cobrança ISP</h1>
        <p>Conectado como: <b>{st.session_state.usuario_email}</b> | Domínio Verificado ✔️</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# 4. Barra Lateral (Sidebar) de Configuração
with st.sidebar:
    st.title("Fonte de dados")
    modo = st.radio("Origem", ["Link SharePoint", "Upload manual", "Arquivo local"], index=0)
    sharepoint_url = st.text_area("Link SharePoint", value=DEFAULT_SHAREPOINT_URL, height=120)
    uploaded = st.file_uploader("Selecione o arquivo .xlsm ou .xlsx", type=["xlsm", "xlsx"])
    local_path = st.text_input("Caminho local do arquivo", value="Base de cobrança - Teste.xlsm")
    
    if st.button("Atualizar dados / Limpar Cache", use_container_width=True):
        st.cache_data.clear()
        
    if st.button("🚪 Sair do Painel", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

fonte = None
fonte_label = None

if modo == "Link SharePoint":
    if sharepoint_url.strip():
        fonte = make_downloadable_sharepoint_url(sharepoint_url.strip())
        fonte_label = "SharePoint"
elif modo == "Upload manual":
    if uploaded is not None:
        fonte = uploaded
        fonte_label = "Upload manual"
elif modo == "Arquivo local":
    path = Path(local_path)
    if path.exists() and path.is_file():
        fonte = str(path)
        fonte_label = "Arquivo local"

if fonte is None:
    st.warning("Defina uma origem válida para carregar a base de cobrança.")
    st.stop()

try:
    abas_disponiveis = get_sheet_names(fonte)
except Exception as e:
    st.error(f"Não foi possível abrir o arquivo Excel: {e}")
    st.stop()

if SHEET_NAME not in abas_disponiveis:
    st.error(f"A aba '{SHEET_NAME}' não foi encontrada no arquivo.")
    st.stop()

try:
    df = carregar_dados(fonte, SHEET_NAME)
except Exception as e:
    st.error(f"Erro ao carregar os dados: {e}")
    st.stop()

# Filtros Dinâmicos na Sidebar
with st.sidebar:
    st.success(f"Fonte: {fonte_label}")
    classes = sorted(df["Classe de Risco"].dropna().astype(str).unique().tolist())
    classe_sel = st.multiselect("Classe de Risco", classes, default=classes)
    apenas_vencidos = st.toggle("Apenas vencidos", value=True)
    apenas_abertos = st.toggle("Apenas em aberto", value=True)

unidades_disponiveis = sorted(df["UNIDADE"].dropna().astype(str).unique().tolist())
st.write("**Filtrar por UNIDADE:**")
unidade_sel = st.pills(label="Unidades Filtro", options=unidades_disponiveis, default=unidades_disponiveis, selection_mode="multi", label_visibility="collapsed")

# Aplicação dos Filtros
base = df.copy()
if unidade_sel: base = base[base["UNIDADE"].isin(unidade_sel)]
if classe_sel: base = base[base["Classe de Risco"].isin(classe_sel)]
if apenas_vencidos: base = base[base["Vencido"]]
if apenas_abertos: base = base[base["Em Aberto"]]

valor_total = float(base["Valor"].sum())
clientes_distintos = int(base["Responsável"].nunique())
titulos_aberto = int(len(base))
clientes_bloqueados = int(base.loc[base["Bloqueado"], "Responsável"].nunique())

# Renderização dos KPIs
k1, k2, k3, k4 = st.columns(4)
for col, titulo, valor, rodape in [(k1, "Valor Total da Carteira", brl_short(valor_total), brl(valor_total)), (k2, "Clientes Distintos", f"{clientes_distintos:,}".replace(",", "."), "Responsáveis únicos"), (k3, "Títulos em Aberto", f"{titulos_aberto:,}".replace(",", "."), "Registros filtrados"), (k4, "Clientes Bloqueados", f"{clientes_bloqueados:,}".replace(",", "."), "Coluna Bloqueado = Sim")]:
    with col: st.markdown(f"<div class='kpi-card'><div class='kpi-label'>{titulo}</div><div class='kpi-value'>{valor}</div><div class='kpi-foot'>{rodape}</div></div>", unsafe_allow_html=True)

st.markdown("<div class='section-title'>Visualizações</div>", unsafe_allow_html=True)

# Gráficos
c1, c2 = st.columns([1.05, 1.25])
with c1:
    status_df = base.groupby("Status", as_index=False)["Valor"].sum().sort_values("Valor", ascending=False)
    status_df = status_df[status_df["Status"] != ""]
    topo = pd.DataFrame({"Status": ["Valor Total"], "Valor": [valor_total]})
    funil_df = pd.concat([topo, status_df], ignore_index=True)
    funil_df["Pct"] = np.where(valor_total > 0, funil_df["Valor"] / valor_total, 0)
    textos = [brl_short(v) if i == 0 else f"{brl_short(v)} | {p:.1%}" for i, (v, p) in enumerate(zip(funil_df["Valor"], funil_df["Pct"]))]
    fig_funil = go.Figure(go.Funnel(y=funil_df["Status"], x=funil_df["Valor"], text=textos, textposition="inside", marker={"color": [ISP_GREEN_DARK] + [ISP_GREEN_MID] * max(len(funil_df) - 1, 0)}, connector={"line": {"color": ISP_GREEN_SOFT, "width": 1.2}}, opacity=0.94))
    fig_funil.update_layout(title="Resumo por Status", height=430, margin=dict(l=130, r=40, t=50, b=40), paper_bgcolor=CARD, plot_bgcolor=CARD, font=dict(color=TEXT, size=11), yaxis=dict(visible=True, automargin=True), xaxis=dict(visible=True, automargin=True))
    st.plotly_chart(fig_funil, use_container_width=True, theme=None)

with c2:
    mes_df = base.groupby("Mês Vencimento", as_index=False)["Valor"].sum().sort_values("Mês Vencimento", ascending=True)
    mes_df["Rótulo"] = mes_df["Valor"].apply(lambda v: f"{brl_short(v)} | {pct(v, valor_total)}")
    fig_mes = px.bar(mes_df, x="Valor", y="Mês Vencimento", orientation="h", text="Rótulo", color_discrete_sequence=[ISP_GREEN])
    fig_mes.update_traces(textposition="inside")
    fig_mes.update_layout(title="Mês de Vencimento vs Valor", height=430, margin=dict(l=100, r=40, t=50, b=40), paper_bgcolor=CARD, plot_bgcolor=CARD, font=dict(color=TEXT, size=11), xaxis=dict(title="Valor Acumulado", visible=True, automargin=True), yaxis=dict(title="Meses de Vencimento", visible=True, automargin=True))
    st.plotly_chart(fig_mes, use_container_width=True, theme=None)

c3, c4 = st.columns([1.15, 0.85])
with c3:
    risco_df = base.groupby("Classe de Risco", as_index=False)["Valor"].sum().sort_values("Classe de Risco", ascending=True)
    risco_df["Rótulo"] = risco_df["Valor"].apply(lambda v: f"{brl_short(v)}<br>{pct(v, valor_total)}")
    fig_risco = px.bar(risco_df, x="Classe de Risco", y="Valor", text="Rótulo", color="Valor", color_continuous_scale=[[0, "#CDEBDD"], [1, ISP_GREEN_DARK]])
    fig_risco.update_layout(title="Classe de Risco vs Valor", height=420, margin=dict(l=60, r=40, t=50, b=50), paper_bgcolor=CARD, plot_bgcolor=CARD, font=dict(color=TEXT, size=11), coloraxis_showscale=False, xaxis=dict(title="Classes de Risco", visible=True, automargin=True), yaxis=dict(title="Volume Financeiro", visible=True, automargin=True))
    st.plotly_chart(fig_risco, use_container_width=True, theme=None)

with c4:
    st.markdown("<div class='section-title'>Indicadores adicionais</div>", unsafe_allow_html=True)
    atraso_medio = base.loc[base["Dias em Atraso"].notna(), "Dias em Atraso"].mean()
    maior_unidade = base.groupby("UNIDADE", as_index=False)["Valor"].sum().sort_values("Valor", ascending=False).head(1)
    mais_antigo = base["Vencimento"].min()
    st.metric("Atraso médio", f"{0 if pd.isna(atraso_medio) else int(round(atraso_medio))} dias")
    st.metric("Vencimento mais antigo", mais_antigo.strftime("%d/%m/%Y") if pd.notna(mais_antigo) else "Sem data")
    if not maior_unidade.empty: st.metric("Unidade com maior exposição", maior_unidade.iloc[0]["UNIDADE"], brl_short(maior_unidade.iloc[0]["Valor"]))
    resumo = base.groupby("Status", as_index=False)["Valor"].sum().sort_values("Valor", ascending=False).head(5)
    resumo = resumo[resumo["Status"] != ""]
    resumo["Valor"] = resumo["Valor"].apply(brl)
    st.dataframe(resumo, use_container_width=True, hide_index=True)

st.markdown("<div class='section-title'>Tabelas detalhadas</div>", unsafe_allow_html=True)
tabela_total = montar_tabela(base)
tabela_bloq = montar_tabela(base[base["Bloqueado"]])

aba1, aba2 = st.tabs(["Carteira Total", "Casos Bloqueados"])
with aba1:
    st.dataframe(tabela_total.style.format({"Valor em Aberto": brl}), use_container_width=True, hide_index=True, height=430)
    st.download_button("Baixar Carteira Total (CSV)", tabela_total.to_csv(index=False).encode("utf-8-sig"), file_name="carteira_total.csv", mime="text/csv", use_container_width=True)
with aba2:
    st.dataframe(tabela_bloq.style.format({"Valor em Aberto": brl}), use_container_width=True, hide_index=True, height=430)
    st.download_button("Baixar Casos Bloqueados (CSV)", tabela_bloq.to_csv(index=False).encode("utf-8-sig"), file_name="casos_bloqueados.csv", mime="text/csv", use_container_width=True)

st.download_button("Baixar ambas as tabelas em Excel", data=excel_bytes({"Carteira Total": tabela_total, "Casos Bloqueados": tabela_bloq}), file_name=f"detalhamento_carteira_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")