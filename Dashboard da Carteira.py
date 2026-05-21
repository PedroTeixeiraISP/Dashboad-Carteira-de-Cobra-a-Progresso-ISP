import io
from pathlib import Path
from datetime import datetime
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Dashboard da Carteira | Online",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

ISP_GREEN = "#0B6B53"
ISP_GREEN_DARK = "#084C3D"
ISP_GREEN_SOFT = "#DFF3EC"
ISP_GREEN_MID = "#1F8A70"
CARD = "#FFFFFF"
TEXT = "#1F2937"
MUTED = "#667085"
BORDER = "rgba(11,107,83,0.10)"


def brl(value: float) -> str:
    value = 0 if pd.isna(value) else value
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def brl_short(value: float) -> str:
    value = 0 if pd.isna(value) else value
    v = abs(value)
    if v >= 1_000_000_000:
        return f"R$ {value/1_000_000_000:.2f} bi"
    elif v >= 1_000_000:
        return f"R$ {value/1_000_000:.2f} mi"
    elif v >= 1_000:
        return f"R$ {value/1_000:.1f} mil"
    else:
        return f"R$ {value:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")


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

    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]

    renomear = {
        "Histórico do Acionamento": "Historico_Real",
        "Histórico de Acionamento": "Historico_Real",
        "Valor Original": "Valor_Divida",
        "Valor": "Valor_Divida",
    }
    df = df.rename(columns=renomear)

    if "Valor_Divida" not in df.columns:
        df["Valor_Divida"] = 0

    if "Historico_Real" not in df.columns:
        df["Historico_Real"] = ""

    df["Valor_Divida"] = parse_valor(df["Valor_Divida"]).fillna(0)
    df["Historico_Real"] = df["Historico_Real"].fillna("").astype(str).str.strip()

    for col_data in ["Data de Vencimento", "Vencimento"]:
        if col_data in df.columns:
            if pd.api.types.is_numeric_dtype(df[col_data]):
                df[col_data] = pd.to_datetime(df[col_data], unit="D", origin="1899-12-30", errors="coerce")
            else:
                df[col_data] = pd.to_datetime(df[col_data], errors="coerce")
            df["Data_Vencimento_Tratada"] = df[col_data]
            break
    if "Data_Vencimento_Tratada" not in df.columns:
        df["Data_Vencimento_Tratada"] = pd.NaT

    for col in ["Responsável", "UNIDADE", "Status", "Classe de Risco", "Grupo"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
        else:
            df[col] = "Não informado"

    df["Responsável"] = df["Responsável"].replace("", "Não informado")
    df["UNIDADE"] = df["UNIDADE"].replace("", "Não informado")
    df["Classe de Risco"] = df["Classe de Risco"].replace("", "Não informado")
    df["Grupo"] = df["Grupo"].replace("", "Não informado")

    hoje = pd.Timestamp.today().normalize()
    if "Em Aberto" in df.columns:
        df["Em_Aberto_Bool"] = df["Em Aberto"].astype(str).str.lower().isin(["sim", "true", "1", "s", "em aberto"])
    else:
        status_lower = df["Status"].astype(str).str.lower()
        df["Em_Aberto_Bool"] = ~status_lower.isin(["pago", "liquidado", "baixado", "quitado", "cancelado"])

    if "Vencido" in df.columns:
        df["Vencido_Bool"] = df["Vencido"].astype(str).str.lower().isin(["sim", "true", "1", "s", "vencido"])
    else:
        df["Vencido_Bool"] = df["Data_Vencimento_Tratada"].notna() & (df["Data_Vencimento_Tratada"] < hoje)

    if "Bloqueado" in df.columns:
        df["Bloqueado_Bool"] = df["Bloqueado"].fillna(False).astype(str).str.strip().str.lower().isin(["sim", "true", "1", "s", "sinalizado"])
    else:
        df["Bloqueado_Bool"] = False

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
        return pd.DataFrame(columns=["Unidade", "Responsável", "Classe de Risco", "Grupo", "Valor em Aberto", "Vencimento Mais Antigo", "Último Acionamento", "Observação"])

    base = df.copy()

    def concatenar(series):
        v = [str(x).strip() for x in series if str(x).strip() not in ["", "nan"]]
        u = list(dict.fromkeys(v))
        return " | ".join(u) if u else "Não informado"

    if "Chave" in base.columns:
        sub_agrupado = base.groupby(["UNIDADE", "Responsável", "Classe de Risco", "Grupo", "Chave"], dropna=False).agg({
            "Valor_Divida": "max",
            "Data_Vencimento_Tratada": "min",
            "Status": concatenar,
            "Historico_Real": concatenar,
        }).reset_index()
    else:
        sub_agrupado = base

    tabela = (
        sub_agrupado.groupby(["UNIDADE", "Responsável", "Classe de Risco", "Grupo"], dropna=False)
        .agg(**{
            "Valor em Aberto": ("Valor_Divida", "sum"),
            "Vencimento Mais Antigo": ("Data_Vencimento_Tratada", "min"),
            "Último Acionamento": ("Status", concatenar),
            "Observação": ("Historico_Real", concatenar),
        })
        .reset_index()
        .sort_values("Valor em Aberto", ascending=False)
    )
    tabela["Vencimento Mais Antigo"] = pd.to_datetime(tabela["Vencimento Mais Antigo"]).dt.strftime("%d/%m/%Y").fillna("Sem data")
    return tabela.rename(columns={"UNIDADE": "Unidade"})


# --- NOVA FUNÇÃO: Visão detalhada por lançamento sem consolidar ---
def montar_tabela_lancamentos(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Unidade", "Responsável", "Classe de Risco", "Grupo", "Descrição do Lançamento", "Valor do Lançamento", "Vencimento", "Status"])

    base = df.copy()

    # Renomear os campos principais para ficar claro
    renomear = {
        "UNIDADE": "Unidade",
        "Valor_Divida": "Valor do Lançamento",
        "Data_Vencimento_Tratada": "Vencimento",
        "Historico_Real": "Descrição do Lançamento"
    }
    
    tabela = base.rename(columns=renomear)

    if "Vencimento" in tabela.columns:
        tabela["Vencimento"] = tabela["Vencimento"].dt.strftime("%d/%m/%Y").fillna("Sem data")
        
    # Organizar colunas: colocar as essenciais no começo, e manter o resto para dar contexto completo
    cols_base = ["Unidade", "Responsável", "Classe de Risco", "Grupo", "Descrição do Lançamento", "Valor do Lançamento", "Vencimento", "Status"]
    cols_existentes = [c for c in cols_base if c in tabela.columns]
    
    # Adicionar colunas adicionais originais ocultando colunas de controle do app (que terminam em _Bool, _Txt, _Num)
    cols_restantes = [c for c in tabela.columns if c not in cols_existentes and not c.endswith("_Bool") and not c.endswith("_Txt") and not c.endswith("_Num")]
    
    return tabela[cols_existentes + cols_restantes]
# ------------------------------------------------------------------


def excel_bytes(abas: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for nome, tabela in abas.items():
            tabela.to_excel(writer, index=False, sheet_name=nome[:31])
    return output.getvalue()


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

st.markdown(f"<div class='hero'><h1>Dashboard da Carteira de Cobrança ISP</h1><p>Conectado como: <b>validação por domínio institucional</b></p></div>", unsafe_allow_html=True)

with st.sidebar:
    st.title("Painel de Controle")
    NOME_ARQUIVO_EXCEL = "Base de cobrança - Teste.xlsb"
    caminho_base_local = Path(__file__).parent / NOME_ARQUIVO_EXCEL if "__file__" in globals() else Path(NOME_ARQUIVO_EXCEL)

    if caminho_base_local.exists():
        fonte = caminho_base_local
        st.success("✔️ Arquivo físico lido")
    else:
        st.warning("⚠️ Arquivo não localizado na raiz. Faça o upload manual abaixo:")
        fonte = st.file_uploader("Selecione o arquivo .xlsb", type=["xlsb"])

    if st.button("🔄 Forçar Recarregamento Total (Limpar Cache)", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    if st.button("🚪 Sair do Painel", use_container_width=True):
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

with st.sidebar:
    classes = sorted(df["Classe de Risco"].dropna().astype(str).unique().tolist())
    classe_sel = st.multiselect("Classe de Risco", classes, default=classes)
    apenas_vencidos = st.toggle("Apenas vencidos", value=False)
    apenas_abertos = st.toggle("Apenas em aberto", value=False)

    grupos = sorted(df["Grupo"].dropna().astype(str).unique().tolist())
    st.markdown("### Grupo")

    if "grupo_sel" not in st.session_state:
        st.session_state.grupo_sel = grupos.copy()

    c_limpar, c_todos = st.columns(2)
    with c_limpar:
        if st.button("🧹 Limpar Unidades", use_container_width=True):
            st.session_state.unidade_sel = []
            st.rerun()
    with c_todos:
        if st.button("✅ Todas", use_container_width=True):
            st.session_state.grupo_sel = grupos.copy()
            st.rerun()

    for g in grupos:
        key = f"grp_{g}"
        if key not in st.session_state:
            st.session_state[key] = True

    sel_cols = st.columns(2)
    with sel_cols[0]:
        if st.button("Selecionar todos", use_container_width=True):
            for g in grupos:
                st.session_state[f"grp_{g}"] = True
            st.session_state.grupo_sel = grupos.copy()
            st.rerun()
    with sel_cols[1]:
        if st.button("Limpar todos", use_container_width=True):
            for g in grupos:
                st.session_state[f"grp_{g}"] = False
            st.session_state.grupo_sel = []
            st.rerun()

    grupo_sel = []
    for g in grupos:
        st.session_state[f"grp_{g}"] = st.checkbox(g, value=st.session_state.get(f"grp_{g}", True), key=f"chk_{g}")
        if st.session_state[f"grp_{g}"]:
            grupo_sel.append(g)
    st.session_state.grupo_sel = grupo_sel

unidades_disponiveis = sorted(df["UNIDADE"].dropna().astype(str).unique().tolist())
if "unidade_sel" not in st.session_state:
    st.session_state.unidade_sel = unidades_disponiveis.copy()

lbl_col1, lbl_col2 = st.columns([8, 2])
with lbl_col1:
    st.write("**Filtrar por UNIDADE:**")
with lbl_col2:
    if st.button("🧹 Limpar", type="secondary", use_container_width=True):
        st.session_state.unidade_sel = []
        st.rerun()

unidade_sel = st.pills(
    label="Unidades Filtro",
    options=unidades_disponiveis,
    default=st.session_state.unidade_sel,
    selection_mode="multi",
    label_visibility="collapsed",
)
st.session_state.unidade_sel = unidade_sel if unidade_sel is not None else []

base = df.copy()
if st.session_state.unidade_sel:
    base = base[base["UNIDADE"].isin(st.session_state.unidade_sel)]
else:
    base = base.iloc[0:0]

if classe_sel:
    base = base[base["Classe de Risco"].isin(classe_sel)]
if grupo_sel:
    base = base[base["Grupo"].isin(grupo_sel)]
if apenas_vencidos:
    base = base[base["Vencido_Bool"]]
if apenas_abertos:
    base = base[base["Em_Aberto_Bool"]]

if "Chave" in base.columns:
    base_unicos_kpi = base.drop_duplicates(subset=["Chave"])
else:
    base_unicos_kpi = base

valor_total = float(base_unicos_kpi["Valor_Divida"].sum())
clientes_distintos = int(base["Responsável"].nunique())
titulos_aberto = int(len(base_unicos_kpi))
clientes_bloqueados = int(base.loc[base["Bloqueado_Bool"], "Responsável"].nunique())

k1, k2, k3, k4 = st.columns(4)
for col, titulo, valor, rodape in [
    (k1, "Valor Total da Carteira", brl_short(valor_total), brl(valor_total)),
    (k2, "Clientes Distintos", f"{clientes_distintos:,}".replace(",", "."), "Responsáveis únicos"),
    (k3, "Títulos Filtrados", f"{titulos_aberto:,}".replace(",", "."), "Registros Únicos"),
    (k4, "Clientes Bloqueados", f"{clientes_bloqueados:,}".replace(",", "."), "Bloqueado = Sim"),
]:
    with col:
        st.markdown(f"<div class='kpi-card'><div class='kpi-label'>{titulo}</div><div class='kpi-value'>{valor}</div><div class='kpi-foot'>{rodape}</div></div>", unsafe_allow_html=True)

st.markdown("<div class='section-title'>Visualizações</div>", unsafe_allow_html=True)

c1, c2 = st.columns([1.05, 1.25])
with c1:
    status_df = base_unicos_kpi.groupby("Status", as_index=False)["Valor_Divida"].sum().sort_values("Valor_Divida", ascending=False)
    status_df = status_df[status_df["Status"] != ""]
    topo = pd.DataFrame({"Status": ["Valor Total"], "Valor_Divida": [valor_total]})
    funil_df = pd.concat([topo, status_df], ignore_index=True)
    funil_df["Pct"] = np.where(valor_total > 0, funil_df["Valor_Divida"] / valor_total, 0)
    textos = [brl_short(v) if i == 0 else f"{brl_short(v)} | {p:.1%}" for i, (v, p) in enumerate(zip(funil_df["Valor_Divida"], funil_df["Pct"]))]
    fig_funil = go.Figure(go.Funnel(y=funil_df["Status"], x=funil_df["Valor_Divida"], text=textos, textposition="inside", marker={"color": [ISP_GREEN_DARK] + [ISP_GREEN_MID] * max(len(funil_df) - 1, 0)}, connector={"line": {"color": ISP_GREEN_SOFT, "width": 1.2}}, opacity=0.94))
    fig_funil.update_layout(title="Resumo por Status", height=430, margin=dict(l=130, r=40, t=50, b=40), paper_bgcolor=CARD, plot_bgcolor=CARD, font=dict(color=TEXT, size=11))
    st.plotly_chart(fig_funil, use_container_width=True, theme=None)

with c2:
    mes_df = base_unicos_kpi.groupby("Mes_Vencimento_Txt", as_index=False)["Valor_Divida"].sum().sort_values("Mes_Vencimento_Txt", ascending=True)
    mes_df["Rótulo"] = mes_df["Valor_Divida"].apply(lambda v: f"{brl_short(v)} | {pct(v, valor_total)}")
    fig_mes = px.bar(mes_df, x="Valor_Divida", y="Mes_Vencimento_Txt", orientation="h", text="Rótulo", color_discrete_sequence=[ISP_GREEN])
    fig_mes.update_traces(textposition="inside")
    fig_mes.update_layout(title="Mês de Vencimento vs Valor", height=430, margin=dict(l=100, r=40, t=50, b=40), paper_bgcolor=CARD, plot_bgcolor=CARD, font=dict(color=TEXT, size=11))
    st.plotly_chart(fig_mes, use_container_width=True, theme=None)

c3, c4 = st.columns([1.05, 0.95])
with c3:
    risco_df = base_unicos_kpi.groupby("Classe de Risco", as_index=False)["Valor_Divida"].sum().sort_values("Classe de Risco", ascending=True)
    risco_df["Rótulo"] = risco_df["Valor_Divida"].apply(lambda v: f"{brl_short(v)}<br>{pct(v, valor_total)}")
    fig_risco = px.bar(risco_df, x="Classe de Risco", y="Valor_Divida", text="Rótulo", color="Valor_Divida", color_continuous_scale=[[0, "#CDEBDD"], [1, ISP_GREEN_DARK]])
    fig_risco.update_layout(title="Classe de Risco vs Valor", height=420, margin=dict(l=60, r=40, t=50, b=50), paper_bgcolor=CARD, plot_bgcolor=CARD, font=dict(color=TEXT, size=11), coloraxis_showscale=False)
    st.plotly_chart(fig_risco, use_container_width=True, theme=None)

with c4:
    st.markdown("<div class='section-title' style='margin-top:0;'>Análise de Parcelas em Aberto</div>", unsafe_allow_html=True)
    if not base.empty:
        df_aberto = base[base["Em_Aberto_Bool"]] if "Em_Aberto_Bool" in base.columns else base
        if not df_aberto.empty:
            analise_resp = df_aberto.groupby("Responsável").agg(Qtd_Parcelas=("Valor_Divida", "size"), Valor_Total=("Valor_Divida", "sum")).reset_index()
            def categorizar_faixas(qtd):
                if qtd == 1: return "1 parcela"
                elif qtd == 2: return "2 parcelas"
                elif qtd == 3: return "3 parcelas"
                elif qtd == 4: return "4 parcelas"
                else: return "+5 parcelas"
            analise_resp["Faixa"] = analise_resp["Qtd_Parcelas"].apply(categorizar_faixas)
            resumo_faixas = analise_resp.groupby("Faixa").agg(Qtd_Responsaveis=("Responsável", "nunique"), Valor_Carteira=("Valor_Total", "sum")).reindex(["1 parcela", "2 parcelas", "3 parcelas", "4 parcelas", "+5 parcelas"], fill_value=0).reset_index()
            total_valor_faixas = resumo_faixas["Valor_Carteira"].sum()
            resumo_faixas["% Carteira"] = resumo_faixas["Valor_Carteira"].apply(lambda x: pct(x, total_valor_faixas))
            resumo_exibicao = resumo_faixas.copy()
            resumo_exibicao["Valor na Carteira"] = resumo_exibicao["Valor_Carteira"].apply(brl)
            resumo_exibicao = resumo_exibicao.rename(columns={"Faixa": "Parcelas em Aberto", "Qtd_Responsaveis": "Qtd. Responsáveis"})
            st.dataframe(resumo_exibicao[["Parcelas em Aberto", "Qtd. Responsáveis", "Valor na Carteira", "% Carteira"]], use_container_width=True, hide_index=True, height=245)
            atraso_medio = base_unicos_kpi.loc[base_unicos_kpi["Dias_Atraso_Num"] > 0, "Dias_Atraso_Num"].mean()
            st.caption(f"💡 Atraso médio geral da carteira selecionada: {0 if pd.isna(atraso_medio) else int(round(atraso_medio))} dias.")
        else:
            st.info("Nenhum título em aberto identificado para os filtros selecionados.")
    else:
        st.info("Sem dados para analisar os indicadores adicionais.")

st.markdown("<div class='section-title'>Análise por Grupo</div>", unsafe_allow_html=True)
g1, g2 = st.columns([1.15, 0.85])
with g1:
    grupo_df = base_unicos_kpi.groupby("Grupo", as_index=False)["Valor_Divida"].sum().sort_values("Valor_Divida", ascending=False)
    if not grupo_df.empty:
        grupo_df["Rótulo"] = grupo_df["Valor_Divida"].apply(lambda v: f"{brl_short(v)} | {pct(v, valor_total)}")
        fig_grupo = px.bar(grupo_df, x="Grupo", y="Valor_Divida", text="Rótulo", color="Valor_Divida", color_continuous_scale=[[0, "#DFF3EC"], [1, ISP_GREEN_DARK]])
        fig_grupo.update_layout(title="Valor por Grupo", height=430, margin=dict(l=40, r=40, t=50, b=80), paper_bgcolor=CARD, plot_bgcolor=CARD, font=dict(color=TEXT, size=11), coloraxis_showscale=False)
        fig_grupo.update_traces(textposition="outside")
        st.plotly_chart(fig_grupo, use_container_width=True, theme=None)
    else:
        st.info("Sem dados para o gráfico por Grupo.")
with g2:
    if not grupo_df.empty:
        fig_pie = px.pie(grupo_df, names="Grupo", values="Valor_Divida", hole=0.45, title="Participação por Grupo")
        fig_pie.update_traces(textinfo="percent+label")
        fig_pie.update_layout(height=430, paper_bgcolor=CARD, plot_bgcolor=CARD, font=dict(color=TEXT, size=11))
        st.plotly_chart(fig_pie, use_container_width=True, theme=None)


st.markdown("<div class='section-title'>Tabelas detalhadas</div>", unsafe_allow_html=True)

# Geração das tabelas para exibição
tabela_total = montar_tabela(base)
tabela_bloq = montar_tabela(base[base["Bloqueado_Bool"]])
tabela_nao_aluno = montar_tabela(base[base["Status Aluno"].astype(str).str.contains("Não é aluno", case=False, na=False)])
tabela_lancamentos = montar_tabela_lancamentos(base) # <- Nova Tabela

# Adicionado a nova aba 'Visão por Lançamento'
aba1, aba2, aba3, aba4 = st.tabs(["Carteira Total", "Casos Bloqueados", "Responsável = Não é aluno", "Visão por Lançamento"])

with aba1:
    st.dataframe(tabela_total.style.format({"Valor em Aberto": brl}), use_container_width=True, hide_index=True, height=430)
    st.download_button("Baixar Carteira Total (CSV)", tabela_total.to_csv(index=False).encode("utf-8-sig"), file_name="carteira_total.csv", mime="text/csv", use_container_width=True)
with aba2:
    st.dataframe(tabela_bloq.style.format({"Valor em Aberto": brl}), use_container_width=True, hide_index=True, height=430)
    st.download_button("Baixar Casos Bloqueados (CSV)", tabela_bloq.to_csv(index=False).encode("utf-8-sig"), file_name="casos_bloqueados.csv", mime="text/csv", use_container_width=True)
with aba3:
    st.dataframe(tabela_nao_aluno.style.format({"Valor em Aberto": brl}), use_container_width=True, hide_index=True, height=430)
    st.download_button("Baixar Não é aluno (CSV)", tabela_nao_aluno.to_csv(index=False).encode("utf-8-sig"), file_name="nao_e_aluno.csv", mime="text/csv", use_container_width=True)
with aba4:
    # Nova aba exibindo detalhe por linha sem consolidar
    st.dataframe(tabela_lancamentos.style.format({"Valor do Lançamento": brl}), use_container_width=True, hide_index=True, height=430)
    st.download_button("Baixar Visão por Lançamento (CSV)", tabela_lancamentos.to_csv(index=False).encode("utf-8-sig"), file_name="visao_por_lancamento.csv", mime="text/csv", use_container_width=True)

# Botão global que também faz o download de todas as abas juntas no Excel, agora incluindo os Lançamentos
st.download_button(
    "Baixar todas as tabelas em Excel", 
    data=excel_bytes({
        "Carteira Total": tabela_total, 
        "Casos Bloqueados": tabela_bloq, 
        "Não é aluno": tabela_nao_aluno,
        "Lançamentos Individuais": tabela_lancamentos # Aba adicional no arquivo exportado
    }), 
    file_name=f"detalhamento_carteira_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", 
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
