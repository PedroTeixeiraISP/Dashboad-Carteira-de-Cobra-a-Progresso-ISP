import streamlit as st
import pandas as pd
import numpy as np
import requests
from io import BytesIO

# Configuração da página
st.set_page_config(
    page_title="Dashboard Financeiro & Inadimplência",
    page_icon="📊",
    layout="wide"
)

# -----------------------------------------------------------------------------
# FUNÇÕES DE TRATAMENTO E CARREGAMENTO
# -----------------------------------------------------------------------------

def parse_valor(s):
    """Converte valores numéricos/monetários de string para float."""
    if pd.api.types.is_numeric_dtype(s):
        return s.astype(float)
    
    s = s.astype(str).str.strip()
    s = s.str.replace("R$", "", regex=False).str.replace(" ", "", regex=False)
    
    # Trata formato brasileiro (1.000,00 -> 1000.00)
    s = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce")


@st.cache_data(show_spinner=False, ttl=300)
def carregar_dados_github(url_raw: str):
    """Baixa o arquivo do GitHub e trata as abas e colunas."""
    resp = requests.get(url_raw)
    resp.raise_for_status()
    file_bytes = resp.content

    with pd.ExcelFile(BytesIO(file_bytes), engine="pyxlsb") as xls:
        abas = xls.sheet_names
        
        # 1. Identificar Aba Principal
        aba_alvo = "Base Teste" if "Base Teste" in abas else abas[0]
        df = pd.read_excel(xls, sheet_name=aba_alvo)

        # 2. Identificar e Carregar Aba de Bloqueados (procura variações com 'bloq')
        df_bloqueados_aba = pd.DataFrame()
        aba_bloq_alvo = None
        
        for a in abas:
            if "bloq" in str(a).lower():
                aba_bloq_alvo = a
                break
                
        if aba_bloq_alvo:
            df_bloqueados_aba = pd.read_excel(xls, sheet_name=aba_bloq_alvo)
            df_bloqueados_aba = df_bloqueados_aba.dropna(how="all")
            df_bloqueados_aba.columns = [str(c).strip() for c in df_bloqueados_aba.columns]

    # --- Tratamento da Base Principal ---
    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]

    # Mapeamento de colunas flexível
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

    # Tratamento de Datas
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

    # Colunas de Filtro
    for col in ["Responsável", "UNIDADE", "Status", "Classe de Risco", "Grupo"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
        else:
            df[col] = "Não informado"

    df["Responsável"] = df["Responsável"].replace("", "Não informado")
    df["UNIDADE"] = df["UNIDADE"].replace("", "Não informado")
    df["Classe de Risco"] = df["Classe de Risco"].replace("", "Não informado")
    df["Grupo"] = df["Grupo"].replace("", "Não informado")

    # Regras e Flags
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
        df["Dias_Atraso_Num"] = np.where(
            df["Data_Vencimento_Tratada"].notna(),
            (hoje - df["Data_Vencimento_Tratada"]).dt.days,
            0
        )

    if "Mês Vencimento" in df.columns:
        df["Mes_Vencimento_Txt"] = df["Mês Vencimento"].fillna("Sem data").astype(str)
    elif "Data_Vencimento_Tratada" in df.columns:
        df["Mes_Vencimento_Txt"] = df["Data_Vencimento_Tratada"].dt.strftime("%Y-%m").fillna("Sem data")
    else:
        df["Mes_Vencimento_Txt"] = "Sem data"

    return df, df_bloqueados_aba, abas


# -----------------------------------------------------------------------------
# CARREGAMENTO DOS DADOS (SUBSTITUA SUA URL SE NECESSÁRIO)
# -----------------------------------------------------------------------------
URL_GITHUB = "https://raw.githubusercontent.com/seu-usuario/seu-repositorio/main/sua_planilha.xlsb"

try:
    df_raw, base_bloqueados_aba, list_abas = carregar_dados_github(URL_GITHUB)
except Exception as e:
    st.error(f"Erro ao carregar os dados do GitHub: {e}")
    st.stop()


# -----------------------------------------------------------------------------
# BARRA LATERAL - FILTROS
# -----------------------------------------------------------------------------
st.sidebar.header("🔍 Filtros de Consulta")

unidades = sorted(list(df_raw["UNIDADE"].unique()))
unidade_sel = st.sidebar.multiselect("Unidade", unidades, default=[])

responsaveis = sorted(list(df_raw["Responsável"].unique()))
resp_sel = st.sidebar.multiselect("Responsável", responsaveis, default=[])

classes = sorted(list(df_raw["Classe de Risco"].unique()))
classe_sel = st.sidebar.multiselect("Classe de Risco", classes, default=[])

grupos = sorted(list(df_raw["Grupo"].unique()))
grupo_sel = st.sidebar.multiselect("Grupo", grupos, default=[])

apenas_vencidos = st.sidebar.checkbox("Apenas Vencidos", value=False)
apenas_bloqueados = st.sidebar.checkbox("Apenas Bloqueados (na base principal)", value=False)

# Aplicação dos Filtros
df_filtrado = df_raw.copy()

if unidade_sel:
    df_filtrado = df_filtrado[df_filtrado["UNIDADE"].isin(unidade_sel)]
if resp_sel:
    df_filtrado = df_filtrado[df_filtrado["Responsável"].isin(resp_sel)]
if classe_sel:
    df_filtrado = df_filtrado[df_filtrado["Classe de Risco"].isin(classe_sel)]
if grupo_sel:
    df_filtrado = df_filtrado[df_filtrado["Grupo"].isin(grupo_sel)]
if apenas_vencidos:
    df_filtrado = df_filtrado[df_filtrado["Vencido_Bool"]]
if apenas_bloqueados:
    df_filtrado = df_filtrado[df_filtrado["Bloqueado_Bool"]]


# -----------------------------------------------------------------------------
# CONTEÚDO PRINCIPAL - CARDS E ABAS
# -----------------------------------------------------------------------------
st.title("📊 Painel de Controle Financeiro & Inadimplência")
st.write("---")

# Métricas no Topo
col1, col2, col3, col4 = st.columns(4)
total_registros = len(df_filtrado)
total_valor = df_filtrado["Valor_Divida"].sum()
total_vencidos = df_filtrado[df_filtrado["Vencido_Bool"]]["Valor_Divida"].sum()
total_bloqueados = len(df_filtrado[df_filtrado["Bloqueado_Bool"]])

col1.metric("Total de Registros", f"{total_registros:,}")
col2.metric("Valor Total", f"R$ {total_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
col3.metric("Valor Vencido", f"R$ {total_vencidos:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
col4.metric("Qtd. Bloqueados (Base)", f"{total_bloqueados}")

st.write("---")

# Abas de Visualização
aba1, aba2, aba3 = st.tabs(["📋 Base Filtrada", "📈 Resumo por Unidade", "🚫 Aba Bloqueados (Planilha)"])

with aba1:
    st.subheader("Base de Dados Filtrada")
    st.dataframe(df_filtrado, use_container_width=True, hide_index=True, height=430)
    
    st.download_button(
        label="Baixar Base Filtrada (CSV)",
        data=df_filtrado.to_csv(index=False).encode("utf-8-sig"),
        file_name="base_filtrada.csv",
        mime="text/csv",
        use_container_width=True
    )

with aba2:
    st.subheader("Inadimplência por Unidade")
    if not df_filtrado.empty:
        resumo_unidade = df_filtrado.groupby("UNIDADE").agg(
            Qtd_Titulos=("Valor_Divida", "count"),
            Valor_Total=("Valor_Divida", "sum"),
            Média_Atraso=("Dias_Atraso_Num", "mean")
        ).reset_index()
        
        resumo_unidade["Média_Atraso"] = resumo_unidade["Média_Atraso"].round(1)
        st.dataframe(resumo_unidade, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhum dado encontrado para os filtros selecionados.")

with aba3:
    st.subheader("Registros da Aba de Bloqueados")
    
    if not base_bloqueados_aba.empty:
        st.dataframe(base_bloqueados_aba, use_container_width=True, hide_index=True, height=430)
        
        st.download_button(
            label="Baixar Aba Bloqueados (CSV)",
            data=base_bloqueados_aba.to_csv(index=False).encode("utf-8-sig"),
            file_name="aba_bloqueados.csv",
            mime="text/csv",
            use_container_width=True
        )
    else:
        st.warning("⚠️ Não foi encontrada nenhuma aba que contenha a palavra 'bloq' (ou ela está vazia).")
        st.info(f"📋 **Abas encontradas na sua planilha:** `{list_abas}`")
        st.caption("Verifique na lista acima qual é o nome exato da aba de bloqueados do seu arquivo Excel.")
