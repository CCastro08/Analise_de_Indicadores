import streamlit as st
import pandas as pd
from datetime import datetime

from utils import (
    carregar_e_limpar_csv,
    montar_endereco,
    buscar_coluna_flexivel,
    gerar_excel_padrao
)

def render():
    st.header("🌸 Auditoria Indicador C7 — Prevenção do Câncer e Saúde da Mulher")
    uploaded_file = st.file_uploader("Anexe o relatório CSV de Saúde da Mulher:", type=["csv"], key="uploader_c7")
    
    if uploaded_file and st.button("🚀 Processar Dashboard C7", key="btn_c7"):
        with st.spinner("Processando C7..."):
            df, data_ref = carregar_e_limpar_csv(uploaded_file.getvalue())
            df_micro, df_nom = processar_c7(df, data_ref)
            cols_status = [c for c in df_nom.columns if "Prática" in c or "status" in c]
            excel_bytes = gerar_excel_padrao(df_micro, df_nom, "INDICADOR C7 — SAÚDE DA MULHER E PREVENÇÃO DO CÂNCER", data_ref, cols_status, "Score_C7_%")
            
            st.success("✅ Dashboard C7 gerado!")
            st.download_button("📥 Baixar Dashboard_Indicador_C7.xlsx", excel_bytes, "Dashboard_Indicador_C7.xlsx")

def processar_c7(df, data_ref):
    df['Endereço'] = df.apply(montar_endereco, axis=1)
    col_nasc = buscar_coluna_flexivel(df, ['nascimento'])
    if col_nasc:
        df['Nascimento'] = pd.to_datetime(df[col_nasc], format='%d/%m/%Y', errors='coerce')
        df['Idade'] = ((data_ref - df['Nascimento']).dt.days / 365.25).fillna(30).astype(int)
    else: df['Idade'] = 30

    def eval_c7(row):
        idade = row['Idade']

        p_a = "Fora da faixa etária"
        if 25 <= idade <= 64:
            col_sol = buscar_coluna_flexivel(df, ['colo', 'colo de utero'])
            p_a = "Não atendida (0/1)"
            if col_sol and row.get(col_sol):
                try:
                    if (data_ref - datetime.strptime(str(row.get(col_sol))[:10], '%d/%m/%Y')).days <= 1095:
                        p_a = "Atendida (1/1)"
                except: pass

        p_b = "Fora da faixa etária"
        if 9 <= idade <= 14:
            col_hpv = buscar_coluna_flexivel(df, ['hpv'])
            hpv = str(row.get(col_hpv, '') or '').strip() if col_hpv else ''
            p_b = "Atendida (1/1)" if hpv and hpv not in ['-', 'Sem registro', 'None'] else "Não atendida (0/1)"

        p_c = "Fora da faixa etária"
        if 14 <= idade <= 69:
            col_ssr = buscar_coluna_flexivel(df, ['saude sexual', 'reprodutiva'])
            p_c = "Não atendida (0/1)"
            if col_ssr and row.get(col_ssr):
                try:
                    if (data_ref - datetime.strptime(str(row.get(col_ssr))[:10], '%d/%m/%Y')).days <= 365:
                        p_c = "Atendida (1/1)"
                except: pass

        p_d = "Fora da faixa etária"
        if 50 <= idade <= 69:
            col_mam = buscar_coluna_flexivel(df, ['mama', 'canser de mama'])
            p_d = "Não atendida (0/1)"
            if col_mam and row.get(col_mam):
                try:
                    if (data_ref - datetime.strptime(str(row.get(col_mam))[:10], '%d/%m/%Y')).days <= 730:
                        p_d = "Atendida (1/1)"
                except: pass

        praticas = [p_a, p_b, p_c, p_d]
        aplicaveis = [p for p in praticas if p != "Fora da faixa etária"]
        atendidas = [p for p in praticas if "Atendida (1/1)" in p]
        score = (len(atendidas) / len(aplicaveis) * 100) if aplicaveis else 100.0

        return pd.Series([p_a, p_b, p_c, p_d, score],
                         index=['Prática A — colo do útero 25-64a (36m) — status', 'Prática B — vacina HPV 9-14a — status',
                                'Prática C — saúde sexual/reprodutiva 14-69a (12m) — status', 'Prática D — câncer de mama 50-69a (24m) — status', 'Score_C7_%'])

    res = df.apply(eval_c7, axis=1)
    for c in res.columns: df[c] = res[c]

    cols_nom = ['Microárea', 'Nome', 'CPF', 'Idade', 'Telefone celular', 'Endereço',
                'Prática A — colo do útero 25-64a (36m) — status', 'Prática B — vacina HPV 9-14a — status',
                'Prática C — saúde sexual/reprodutiva 14-69a (12m) — status', 'Prática D — câncer de mama 50-69a (24m) — status', 'Score_C7_%']
    for c in cols_nom:
        if c not in df.columns: df[c] = "-"

    df_micro = df.groupby('Microárea').agg(
        Mulheres_Relatório=('Nome', 'count'),
        Score_Médio_C7=('Score_C7_%', 'mean'),
        Pct_100_Aplicáveis=('Score_C7_%', lambda x: (x == 100).mean() * 100)
    ).reset_index()

    return df_micro, df[cols_nom]
