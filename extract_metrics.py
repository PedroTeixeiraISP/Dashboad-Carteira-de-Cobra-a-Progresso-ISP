#!/usr/bin/env python3
"""
Script para extrair métricas do Dashboard da Carteira de Cobrança
Executa: python extract_metrics.py
"""

import sys
import pandas as pd
import numpy as np
from datetime import datetime

def parse_valor(serie):
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
    try:
        with pd.ExcelFile(file_source, engine="pyxlsb") as xls:
            abas = xls.sheet_names
            print(f"Abas disponíveis: {abas}")
            aba_alvo = "Base Teste" if "Base Teste" in abas else abas[0]
            print(f"Usando aba: {aba_alvo}")
            df = pd.read_excel(xls, sheet_name=aba_alvo)
    except Exception as e:
        print(f"❌ Erro ao carregar arquivo: {e}")
        return None

    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]

    print(f"Colunas disponíveis: {list(df.columns)}")

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

    # Processar datas de vencimento
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

    # Preencher colunas de texto
    for col in ["Responsável", "UNIDADE", "Status", "Classe de Risco", "Grupo"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
        else:
            df[col] = "Não informado"

    df["Responsável"] = df["Responsável"].replace("", "Não informado")
    df["UNIDADE"] = df["UNIDADE"].replace("", "Não informado")
    df["Classe de Risco"] = df["Classe de Risco"].replace("", "Não informado")
    df["Grupo"] = df["Grupo"].replace("", "Não informado")

    # Processar status
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

    return df

def main():
    print("\n" + "="*100)
    print("EXTRAÇÃO DE MÉTRICAS - DASHBOARD DA CARTEIRA DE COBRANÇA ISP")
    print("="*100 + "\n")
    
    arquivo = "Base de cobrança - Teste.xlsb"
    
    print(f"📂 Carregando arquivo: {arquivo}\n")
    df = carregar_dados(arquivo)
    
    if df is None:
        print("❌ Falha ao carregar dados. Encerrando.")
        return
    
    print(f"✓ Arquivo carregado com sucesso!")
    print(f"✓ Total de registros: {len(df)}\n")
    
    # CÁLCULOS - SEM FILTROS (dados brutos)
    print("="*100)
    print("📊 DADOS BRUTOS (SEM NENHUM FILTRO)")
    print("="*100 + "\n")
    
    # Clientes bloqueados
    clientes_bloqueados_bruto = int(df.loc[df["Bloqueado_Bool"], "Responsável"].nunique())
    
    # Valor vencido
    valor_vencido_bruto = float(df.loc[df["Vencido_Bool"], "Valor_Divida"].sum())
    
    # Valor total
    valor_total_bruto = float(df["Valor_Divida"].sum())
    
    print(f"🔴 CLIENTES DISTINTOS BLOQUEADOS: {clientes_bloqueados_bruto}")
    print(f"💰 VALOR TOTAL DA CARTEIRA VENCIDA: R$ {valor_vencido_bruto:,.2f}")
    print(f"📈 VALOR TOTAL DA CARTEIRA: R$ {valor_total_bruto:,.2f}\n")
    
    # Detalhamento dos bloqueados
    print("="*100)
    print("📋 DETALHAMENTO - CLIENTES BLOQUEADOS")
    print("="*100 + "\n")
    
    bloqueados_df = df[df["Bloqueado_Bool"]].copy()
    
    if not bloqueados_df.empty:
        # Agrupar por responsável
        bloqueados_por_cliente = bloqueados_df.groupby("Responsável").agg({
            "Valor_Divida": "sum",
            "Status": lambda x: x.iloc[0] if len(x) > 0 else "N/A"
        }).reset_index()
        bloqueados_por_cliente = bloqueados_por_cliente.sort_values("Valor_Divida", ascending=False)
        
        print(f"Total de Clientes Distintos Bloqueados: {len(bloqueados_por_cliente)}\n")
        print(f"{'#':<4} {'Responsável':<50} {'Valor Bloqueado':>20}")
        print("-" * 100)
        
        for idx, row in bloqueados_por_cliente.iterrows():
            print(f"{idx+1:<4} {row['Responsável']:<50} R$ {row['Valor_Divida']:>18,.2f}")
        
        print("-" * 100)
        print(f"{'TOTAL':<54} R$ {bloqueados_por_cliente['Valor_Divida'].sum():>18,.2f}\n")
    else:
        print("✓ Nenhum cliente bloqueado encontrado.\n")
    
    # Detalhamento dos vencidos
    print("="*100)
    print("📋 DETALHAMENTO - CARTEIRA VENCIDA")
    print("="*100 + "\n")
    
    vencidos_df = df[df["Vencido_Bool"]].copy()
    
    if not vencidos_df.empty:
        vencidos_por_cliente = vencidos_df.groupby("Responsável").agg({
            "Valor_Divida": "sum"
        }).reset_index()
        vencidos_por_cliente = vencidos_por_cliente.sort_values("Valor_Divida", ascending=False)
        
        print(f"Total de Responsáveis com Carteira Vencida: {len(vencidos_por_cliente)}")
        print(f"Total de Registros Vencidos: {len(vencidos_df)}\n")
        
        print(f"{'#':<4} {'Responsável':<50} {'Valor Vencido':>20}")
        print("-" * 100)
        
        # Mostrar top 10
        for idx, row in vencidos_por_cliente.head(10).iterrows():
            print(f"{idx+1:<4} {row['Responsável']:<50} R$ {row['Valor_Divida']:>18,.2f}")
        
        if len(vencidos_por_cliente) > 10:
            print(f"... ({len(vencidos_por_cliente) - 10} outros clientes)")
        
        print("-" * 100)
        print(f"{'TOTAL':<54} R$ {vencidos_por_cliente['Valor_Divida'].sum():>18,.2f}\n")
    else:
        print("✓ Nenhum registro vencido encontrado.\n")
    
    # Resumo final
    print("="*100)
    print("📊 RESUMO FINAL")
    print("="*100)
    print(f"\n🔴 Clientes Distintos como BLOQUEADOS: {clientes_bloqueados_bruto}")
    print(f"💰 Valor Total da Carteira VENCIDA: R$ {valor_vencido_bruto:,.2f}")
    print(f"📈 Valor Total da Carteira: R$ {valor_total_bruto:,.2f}")
    print(f"📅 Data/Hora da Extração: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
    print("="*100 + "\n")

if __name__ == "__main__":
    main()
