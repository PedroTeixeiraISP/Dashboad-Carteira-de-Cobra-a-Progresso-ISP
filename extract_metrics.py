import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime

# Funções do Dashboard
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

# Carregando dados
arquivo = "Base de cobrança - Teste.xlsb"
df = carregar_dados(arquivo)

print("=" * 80)
print("ANÁLISE DA BASE DE COBRANÇA - TESTE")
print("=" * 80)
print()

# Cálculos SEM FILTROS (dados brutos)
print("📊 DADOS BRUTOS (SEM NENHUM FILTRO):")
print("-" * 80)

valor_total_bruto = float(df["Valor_Divida"].sum())
clientes_bloqueados_bruto = int(df.loc[df["Bloqueado_Bool"], "Responsável"].nunique())
valor_vencido_bruto = float(df.loc[df["Vencido_Bool"], "Valor_Divida"].sum())

print(f"✓ Clientes Distintos Bloqueados: {clientes_bloqueados_bruto}")
print(f"✓ Valor Total da Carteira Vencida: R$ {valor_vencido_bruto:,.2f}")
print()

# Detalhamento dos bloqueados
print("📋 DETALHAMENTO DOS CLIENTES BLOQUEADOS:")
print("-" * 80)
bloqueados_df = df[df["Bloqueado_Bool"]].copy()
if not bloqueados_df.empty:
    clientes_unicos_bloqueados = bloqueados_df[["Responsável", "Valor_Divida"]].drop_duplicates(subset=["Responsável"])
    print(f"Total de Clientes Bloqueados: {len(clientes_unicos_bloqueados)}")
    print()
    print("Clientes Bloqueados (Responsáveis):")
    for idx, (resp, val) in enumerate(bloqueados_df.groupby("Responsável")["Valor_Divida"].sum().items(), 1):
        print(f"  {idx}. {resp}: R$ {val:,.2f}")
else:
    print("Nenhum cliente bloqueado encontrado.")

print()
print("=" * 80)
print("RESUMO FINAL")
print("=" * 80)
print(f"🔴 Clientes Distintos como BLOQUEADOS: {clientes_bloqueados_bruto}")
print(f"💰 Valor Total da Carteira VENCIDA: R$ {valor_vencido_bruto:,.2f}")
print("=" * 80)
