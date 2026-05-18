import io
from pathlib import Path
from datetime import datetime
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
            email_input = st.text_input("E-mail corporativo", placeholder="Digite seu e-mail").strip().lower()
            bt_entrar = st.button("Verificar e Acessar", use_container_width=True)
            
            if bt_entrar and email_input:
                if email_input.endswith("@ispschools.com") or email_input.endswith("@colegioprogresso.com.br"):
                    st.session_state.autenticado = True
                    st.session_state.usuario_email = email_input
                    st.success("Acesso autorizado! Carregando...")
                    st.rerun()
                else:
                    st.error("E-mail não autorizado. Utilize um domínio @ispschools.com ou @colegioprogresso.com.br")
                    
        st.stop()

tela_de_login()

# ==========================================
# INÍCIO DO CÓDIGO DO DASHBOARD (SÓ RODA SE LOGADO)
# ==========================================

st.set_page_config(
    page_title="Dashboard da Carteira | Online",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Cores de Identidade Visual
ISP_GREEN = "#0B6B53"
ISP_GREEN_DARK = "#084C3D"
ISP_GREEN_SOFT = "#DFF3EC"
ISP_GREEN_MID = "#1F8A70"
CARD = "#FFFFFF"
TEXT = "#1F2937"
MUTED = "#667085"
BORDER = "rgba(11,107,83,0.10)"

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
        .section-title {{font-size: 1.08rem; font-weight: 800; color: {ISP_GREEN_DARK}; margin: 1.5rem 0 .85rem 0;}}
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

def brl(value: float) -> str:
    value = 0 if pd.isna(value) else value
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def brl_short(value: float) -> str:
    value = 0 if pd.isna(value) else value
    v = abs(value)
    if v >= 1_000_000_000: return f"R$ {value/1_000_000_000:.2f} bi"
    elif v >= 1_000_000: return f"R$ {value/1_000_000:.2f} mi"
    elif v >= 1_000: return f"R$ {value/1_000:.1f} mil"
    else: return f"R$ {value:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")

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

@st.cache_data(show_spinner=False)
def carregar_dados(file_source):
    with pd.ExcelFile(file_source, engine="pyxlsb") as xls:
        abas = xls.sheet_names
        aba_alvo = "Base Teste" if "Base Teste" in abas else abas[0]
        df = pd.read_excel(xls, sheet_name=aba_alvo)

    df = df.dropna(how='all')
    df.columns = [str(c).strip() for c in df.columns]

    # Mapeamento Cirúrgico baseado na lista real enviada
    df = df.rename(columns={
        "UNIDADE": "UNIDADE",
        "Responsável": "Responsável",
        "Classe de Risco": "Classe de Risco",
        "Status": "Status",
        "Histórico do Acionamento": "Historico_Real",  # Alvo prioritário apontado por você
        "Valor TT da Divida": "Valor_Divida"
    })

    # Tratamento de Fallback caso falte alguma coluna mapeada
    if "Valor_Divida" not in df.columns and "Valor" in df.columns:
        df["Valor_Divida"] = df["Valor"]
    elif "Valor_Divida" not in df.columns:
        df["Valor_Divida"] = 0

    if "Historico_Real" not in df.columns and "Histórico de Acionamento" in df.columns:
        df["Historico_Real"] = df["Histórico de Acionamento"]
    elif "Historico_Real" not in df.columns:
        df["Historico_Real"] = ""

    df["Valor_Divida"] = parse_valor(df["Valor_Divida"]).fillna(0)
    df["Historico_Real"] = df["Historico_Real"].fillna("").astype(str).str.strip()

    # Tratamento de Datas nativas da sua lista
    for col_data in ["Data de Vencimento", "Vencimento"]:
        if col_data in df.columns:
            if pd.api.types.is_numeric_dtype(df[col_data]):
                df[col_data] = pd.to_datetime(df[col_data], unit='D', origin='1899-12-30', errors="coerce")
            else:
                df[col_data] = pd.to_datetime(df[col_data], errors="coerce")
            df["Data_Vencimento_Tratada"] = df[col_data]
            break
    if "Data_Vencimento_Tratada" not in df.columns:
        df["Data_Vencimento_Tratada"] = pd.NaT

    for col in ["Responsável", "UNIDADE", "Status", "Classe de Risco"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
        else:
            df[col] = "Não informado"

    df["Responsável"] = df["Responsável"].replace("", "Não informado")
    df["UNIDADE"] = df["UNIDADE"].replace("", "Não informado")
    df["Classe de Risco"] = df["Classe de Risco"].replace("", "Não informado")

    # Identificação de Status e Regras de Negócio
    hoje = pd.Timestamp.today().normalize()
    
    if "Em Aberto" in df.columns:
        df["Em_Aberto_Bool"] = df["Em Aberto"].astype(str).str.lower().isin(["sim", "true", "1", "s", "em aberto"])
    else:
        status_lower = df["Status"].str.lower()
        df["Em_Aberto_Bool"] = ~status_lower.isin(["pago", "liquidado", "baixado", "quitado", "cancelado"])

    if "Vencido" in df.columns:
        df["Vencido_Bool"] = df["Vencido"].astype(str).str.lower().isin(["sim", "true", "1", "s", "vencido"])
    else:
        df["Vencido_Bool"] = df["Data_Vencimento_Tratada"].notna() & (df["Data_Vencimento_Tratada"] < hoje)

    df["Bloqueado_Bool"] = df["Bloqueado"].fillna(False).astype(str).str.strip().str.lower().isin(["sim", "true", "1", "s", "sinalizado"])
    
    if "Dias em Atraso" in df.columns:
        df["Dias_Atraso_Num"] = pd.to_numeric(df["Dias em Atraso"], errors="coerce").fillna(0)
    else:
        df["Dias_Atraso_Num"] = np.where(df["Data_Vencimento_Tratada"].notna(), (hoje - df["Data_Vencimento_Tratada"]).dt.days, 0)

    if "Mês Vencimento" in df.columns:
        df["Mes_Vencimento_Txt"] = df["Mês Vencimento"].fillna("Sem data").astype(str)
    elif "Data_Vencimento_Tratada" in df.columns:
        df["Mes_Vencimento_Txt"] = df["Data_Vencimento_Tratada"].dt.strftime("%Y-%m").fillna("Sem data")
    else:
        df["Mes_Vencimento_Txt"] = "Sem data"

    return df

def montar_tabela(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Unidade", "Responsável", "Classe de Risco", "Valor em Aberto", "Vencimento Mais Antigo", "Último Acionamento", "Observação"])
    
    base = df.copy()
    
    def concatenar_status(series):
        v = [str(x).strip() for x in series if str(x).strip() != "" and str(x).lower() != "nan"]
        u = list(dict.fromkeys(v))
        return " | ".join(u) if u else "Não informado"
        
    def concatenar_historico(series):
        # Captura estritamente os textos legítimos inseridos na planilha
        v = [str(x).strip() for x in series if str(x).strip() != "" and str(x).lower() != "nan" and str(x).lower() != "histórico do acionamento" and str(x).lower() != "histórico de acionamento"]
        u = list(dict.fromkeys(v))
        return " | ".join(u) if u else "Sem observações"

    tabela = (
        base.groupby(["UNIDADE", "Responsável", "Classe de Risco"], dropna=False)
        .agg(**{
            "Valor em Aberto": ("Valor_Divida", "sum"), 
            "Vencimento Mais Antigo": ("Data_Vencimento_Tratada", "min"), 
            "Último Acionamento": ("Status", concatenar_status), 
            "Observação": ("Historico_Real", concatenar_historico)
        })
        .reset_index().sort_values("Valor em Aberto", ascending=False)
    )
    tabela["Vencimento Mais Antigo"] = pd.to_datetime(tabela["Vencimento Mais Antigo"]).dt.strftime("%d/%m/%Y").fillna("Sem data")
    return tabela.rename(columns={"UNIDADE": "Unidade"})

def excel_bytes(abas: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for nome, tabela in abas.items(): tabela.to_excel(writer, index=False, sheet_name=nome[:31])
    return output.getvalue()

# Cabeçalho da Aplicação
st.markdown(f"<div class='hero'><h1>Dashboard da Carteira de Cobrança ISP</h1><p>Conectado como: <b>{st.session_state.usuario_email}</b> | Domínio Verificado ✔️</p></div>", unsafe_allow_html=True)

# Barra Lateral (Sidebar) de Configuração
with st.sidebar:
    st.title("Painel de Controle")
    NOME_ARQUIVO_EXCEL = "Base de cobrança - Teste.xlsb" 
    caminho_base_local = Path(__file__).parent / NOME_ARQUIVO_EXCEL
    
    if caminho_base_local.exists():
        fonte = caminho_base_local
        st.success(f"✔️ Arquivo físico lido")
    else:
        st.warning("⚠️ Arquivo não localizado na raiz. Faça o upload manual abaixo:")
        fonte = st.file_uploader("Selecione o arquivo .xlsb", type=["xlsb"])

    if st.button("🔄 Forçar Recarregamento Total (Limpar Cache)", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
        
    if st.button("🚪 Sair do Painel", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

if fonte is None:
    st.error(f"Por favor, faça o upload do arquivo '{NOME_ARQUIVO_EXCEL}' ou salve-o na mesma pasta.")
    st.stop()

try:
    df = carregar_dados(fonte)
except Exception as e:
    st.error(f"Erro crítico ao processar planilha: {e}")
    st.stop()

if df.empty or df["Valor_Divida"].sum() == 0:
    st.error("🚨 ATENÇÃO: O volume total extraído está zerado!")
    st.dataframe(df.head(5))
    st.stop()

# Filtros Dinâmicos na Sidebar
with st.sidebar:
    classes = sorted(df["Classe de Risco"].dropna().astype(str).unique().tolist())
    classe_sel = st.multiselect("Classe de Risco", classes, default=classes)
    apenas_vencidos = st.toggle("Apenas vencidos", value=False)
    apenas_abertos = st.toggle("Apenas em aberto", value=False)

unidades_disponiveis = sorted(df["UNIDADE"].dropna().astype(str).unique().tolist())
st.write("**Filtrar por UNIDADE:**")
unidade_sel = st.pills(label="Unidades Filtro", options=unidades_disponiveis, default=unidades_disponiveis, selection_mode="multi", label_visibility="collapsed")

# Aplicação dos Filtros
base = df.copy()
if unidade_sel: base = base[base["UNIDADE"].isin(unidade_sel)]
if classe_sel: base = base[base["Classe de Risco"].isin(classe_sel)]
if apenas_vencidos: base = base[base["Vencido_Bool"]]
if apenas_abertos: base = base[base["Em_Aberto_Bool"]]

# KPIs
valor_total = float(base["Valor_Divida"].sum())
clientes_distintos = int(base["Responsável"].nunique())
titulos_aberto = int(len(base))
clientes_bloqueados = int(base.loc[base["Bloqueado_Bool"], "Responsável"].nunique())

k1, k2, k3, k4 = st.columns(4)
for col, titulo, valor, rodape in [(k1, "Valor Total da Carteira", brl_short(valor_total), brl(valor_total)), (k2, "Clientes Distintos", f"{clientes_distintos:,}".replace(",", "."), "Responsáveis únicos"), (k3, "Títulos Filtrados", f"{titulos_aberto:,}".replace(",", "."), "Registros"), (k4, "Clientes Bloqueados", f"{clientes_bloqueados:,}".replace(",", "."), "Bloqueado = Sim")]:
    with col: st.markdown(f"<div class='kpi-card'><div class='kpi-label'>{titulo}</div><div class='kpi-value'>{valor}</div><div class='kpi-foot'>{rodape}</div></div>", unsafe_allow_html=True)

st.markdown("<div class='section-title'>Visualizações</div>", unsafe_allow_html=True)

# Gráficos (Linha 1)
c1, c2 = st.columns([1.05, 1.25])
with c1:
    status_df = base.groupby("Status", as_index=False)["Valor_Divida"].sum().sort_values("Valor_Divida", ascending=False)
    status_df = status_df[status_df["Status"] != ""]
    topo = pd.DataFrame({"Status": ["Valor Total"], "Valor_Divida": [valor_total]})
    funil_df = pd.concat([topo, status_df], ignore_index=True)
    funil_df["Pct"] = np.where(valor_total > 0, funil_df["Valor_Divida"] / valor_total, 0)
    textos = [brl_short(v) if i == 0 else f"{brl_short(v)} | {p:.1%}" for i, (v, p) in enumerate(zip(funil_df["Valor_Divida"], funil_df["Pct"]))]
    fig_funil = go.Figure(go.Funnel(y=funil_df["Status"], x=funil_df["Valor_Divida"], text=textos, textposition="inside", marker={"color": [ISP_GREEN_DARK] + [ISP_GREEN_MID] * max(len(funil_df) - 1, 0)}, connector={"line": {"color": ISP_GREEN_SOFT, "width": 1.2}}, opacity=0.94))
    fig_funil.update_layout(title="Resumo por Status", height=430, margin=dict(l=130, r=40, t=50, b=40), paper_bgcolor=CARD, plot_bgcolor=CARD, font=dict(color=TEXT, size=11))
    st.plotly_chart(fig_funil, use_container_width=True, theme=None)

with c2:
    mes_df = base.groupby("Mes_Vencimento_Txt", as_index=False)["Valor_Divida"].sum().sort_values("Mes_Vencimento_Txt", ascending=True)
    mes_df["Rótulo"] = mes_df["Valor_Divida"].apply(lambda v: f"{brl_short(v)} | {pct(v, valor_total)}")
    fig_mes = px.bar(mes_df, x="Valor_Divida", y="Mes_Vencimento_Txt", orientation="h", text="Rótulo", color_discrete_sequence=[ISP_GREEN])
    fig_mes.update_traces(textposition="inside")
    fig_mes.update_layout(title="Mês de Vencimento vs Valor", height=430, margin=dict(l=100, r=40, t=50, b=40), paper_bgcolor=CARD, plot_bgcolor=CARD, font=dict(color=TEXT, size=11))
    st.plotly_chart(fig_mes, use_container_width=True, theme=None)

# Gráficos (Linha 2)
c3, c4 = st.columns([1.15, 0.85])
with c3:
    risco_df = base.groupby("Classe de Risco", as_index=False)["Valor_Divida"].sum().sort_values("Classe de Risco", ascending=True)
    risco_df["Rótulo"] = risco_df["Valor_Divida"].apply(lambda v: f"{brl_short(v)}<br>{pct(v, valor_total)}")
    fig_risco = px.bar(risco_df, x="Classe de Risco", y="Valor_Divida", text="Rótulo", color="Valor_Divida", color_continuous_scale=[[0, "#CDEBDD"], [1, ISP_GREEN_DARK]])
    fig_risco.update_layout(title="Classe de Risco vs Valor", height=420, margin=dict(l=60, r=40, t=50, b=50), paper_bgcolor=CARD, plot_bgcolor=CARD, font=dict(color=TEXT, size=11), coloraxis_showscale=False)
    st.plotly_chart(fig_risco, use_container_width=True, theme=None)

with c4:
    st.markdown("<div class='section-title' style='margin-top:0;'>Indicadores Adicionais</div>", unsafe_allow_html=True)
    atraso_medio = base.loc[base["Dias_Atraso_Num"] > 0, "Dias_Atraso_Num"].mean()
    maior_unidade = base.groupby("UNIDADE", as_index=False)["Valor_Divida"].sum().sort_values("Valor_Divida", ascending=False).head(1)
    mais_antigo = base["Data_Vencimento_Tratada"].min()
    st.metric("Atraso médio", f"{0 if pd.isna(atraso_medio) else int(round(atraso_medio))} dias")
    st.metric("Vencimento mais antigo", mais_antigo.strftime("%d/%m/%Y") if pd.notna(mais_antigo) else "Sem data")
    if not maior_unidade.empty: st.metric("Unidade com maior exposição", maior_unidade.iloc[0]["UNIDADE"], brl_short(maior_unidade.iloc[0]["Valor_Divida"]))

st.markdown("<div class='section-title'>Tabelas detalhadas</div>", unsafe_allow_html=True)
tabela_total = montar_tabela(base)
tabela_bloq = montar_tabela(base[base["Bloqueado_Bool"]])

aba1, aba2 = st.tabs(["Carteira Total", "Casos Bloqueados"])
with aba1:
    st.dataframe(tabela_total.style.format({"Valor em Aberto": brl}), use_container_width=True, hide_index=True, height=430)
    st.download_button("Baixar Carteira Total (CSV)", tabela_total.to_csv(index=False).encode("utf-8-sig"), file_name="carteira_total.csv", mime="text/csv", use_container_width=True)
with aba2:
    st.dataframe(tabela_bloq.style.format({"Valor em Aberto": brl}), use_container_width=True, hide_index=True, height=430)
    st.download_button("Baixar Casos Bloqueados (CSV)", tabela_bloq.to_csv(index=False).encode("utf-8-sig"), file_name="casos_bloqueados.csv", mime="text/csv", use_container_width=True)

st.download_button("Baixar ambas as tabelas em Excel", data=excel_bytes({"Carteira Total": tabela_total, "Casos Bloqueados": tabela_bloq}), file_name=f"detalhamento_carteira_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
