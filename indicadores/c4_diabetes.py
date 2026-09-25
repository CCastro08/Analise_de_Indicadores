import streamlit as st
import pandas as pd
from datetime import datetime

from utils import (
    carregar_e_limpar_csv,
    montar_endereco,
    buscar_coluna_flexivel,
    extrair_datas_validas,
    gerar_excel_padrao
)

def render():
    st.header("🩸 Auditoria Indicador C4 — Cuidado da Pessoa com Diabetes")
    uploaded_file = st.file_uploader("Anexe o relatório CSV de Diabetes:", type=["csv"], key="uploader_c4")
    
    if uploaded_file and st.button("🚀 Processar Dashboard C4", key="btn_c4"):
        with st.spinner("Processando C4..."):
            df, data_ref = carregar_e_limpar_csv(uploaded_file.getvalue())
            df_micro, df_nom = processar_c4(df, data_ref)
            cols_status = [c for c in df_nom.columns if "Prática" in c or "status" in c]
            excel_bytes = gerar_excel_padrao(df_micro, df_nom, "INDICADOR C4 — CUIDADO DA PESSOA COM DIABETES", data_ref, cols_status, "Score_C4_%")
            
            st.success("✅ Dashboard C4 gerado!")
            st.download_button("📥 Baixar Dashboard_Indicador_C4.xlsx", excel_bytes, "Dashboard_Indicador_C4.xlsx")

def processar_c4(df, data_ref):
    df['Endereço'] = df.apply(montar_endereco, axis=1)

    def eval_c4(row):
        def check_janela(termos, dias):
            col = buscar_coluna_flexivel(df, termos)
            if col and row.get(col):
                try:
                    dt = datetime.strptime(str(row.get(col))[:10], '%d/%m/%Y')
                    if (data_ref - dt).days <= dias: return "Atendida (1/1)"
                except: pass
            return "Não atendida (0/1)"

        p_a = check_janela(['ultima consulta'], 183)
        p_b = check_janela(['pressao arterial'], 183)
        p_d = check_janela(['peso e altura'], 365)
        p_e = check_janela(['hemoglobina glicada', 'glicada'], 365)
        p_f = check_janela(['avaliacao dos pes', 'pes'], 365)

        col_vis = buscar_coluna_flexivel(df, ['visitas'])
        dts_v = extrair_datas_validas(row.get(col_vis), data_ref, 365) if col_vis else []
        p_c = "Não atendida (0/2)"
        if len(dts_v) >= 2:
            for i in range(len(dts_v)-1):
                if (dts_v[i+1] - dts_v[i]).days >= 30:
                    p_c = "Atendida (2/2)"
                    break

        pts = sum([1 for s in [p_a, p_b, p_c, p_d, p_e, p_f] if "Atendida" in s])
        score = (pts / 6.0) * 100
        return pd.Series([p_a, p_b, p_c, p_d, p_e, p_f, score],
                         index=['Prática A — consulta 6m — status', 'Prática B — pressão 6m — status', 'Prática C — 2 visitas 12m — status',
                                'Prática D — peso/altura 12m — status', 'Prática E — HbA1c 12m — status', 'Prática F — pés 12m — status', 'Score_C4_%'])

    res = df.apply(eval_c4, axis=1)
    for c in res.columns: df[c] = res[c]

    cols_nom = ['Microárea', 'Nome', 'CPF', 'Telefone celular', 'Endereço',
                'Prática A — consulta 6m — status', 'Prática B — pressão 6m — status', 'Prática C — 2 visitas 12m — status',
                'Prática D — peso/altura 12m — status', 'Prática E — HbA1c 12m — status', 'Prática F — pés 12m — status', 'Score_C4_%']
    for c in cols_nom:
        if c not in df.columns: df[c] = "-"

    df_micro = df.groupby('Microárea').agg(
        Diabéticos_Ativos=('Nome', 'count'),
        Score_Médio_C4=('Score_C4_%', 'mean'),
        Pct_100_Práticas=('Score_C4_%', lambda x: (x == 100).mean() * 100)
    ).reset_index()

    return df_micro, df[cols_nom]
