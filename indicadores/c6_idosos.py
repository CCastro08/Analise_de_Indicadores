import streamlit as st
import pandas as pd

from utils import (
    carregar_e_limpar_csv,
    montar_endereco,
    buscar_coluna_flexivel,
    extrair_datas_validas,
    gerar_excel_padrao
)

def render():
    st.header("👴 Auditoria Indicador C6 — Cuidado da Pessoa Idosa")
    uploaded_file = st.file_uploader("Anexe o relatório CSV de Idosos:", type=["csv"], key="uploader_c6")
    
    if uploaded_file and st.button("🚀 Processar Dashboard C6", key="btn_c6"):
        with st.spinner("Processando C6..."):
            df, data_ref = carregar_e_limpar_csv(uploaded_file.getvalue())
            df_micro, df_nom = processar_c6(df, data_ref)
            cols_status = [c for c in df_nom.columns if "Prática" in c or "status" in c]
            excel_bytes = gerar_excel_padrao(df_micro, df_nom, "INDICADOR C6 — CUIDADO DA PESSOA IDOSA", data_ref, cols_status, "Score_C6_%")
            
            st.success("✅ Dashboard C6 gerado!")
            st.download_button("📥 Baixar Dashboard_Indicador_C6.xlsx", excel_bytes, "Dashboard_Indicador_C6.xlsx")

def processar_c6(df, data_ref):
    df['Endereço'] = df.apply(montar_endereco, axis=1)

    def eval_c6(row):
        col_med = buscar_coluna_flexivel(df, ['atendimento medico'])
        col_enf = buscar_coluna_flexivel(df, ['atendimento de enfermagem'])
        d_med = pd.to_numeric(row.get(col_med), errors='coerce') if col_med else None
        d_enf = pd.to_numeric(row.get(col_enf), errors='coerce') if col_enf else None
        p_a = "Atendida (1/1)" if ((pd.notnull(d_med) and 0 <= d_med <= 365) or (pd.notnull(d_enf) and 0 <= d_enf <= 365)) else "Não atendida (0/1)"

        col_ant = buscar_coluna_flexivel(df, ['peso e altura simultaneos', 'simultaneos'])
        qtd_ant = min(int(pd.to_numeric(row.get(col_ant, 0), errors='coerce') or 0), 2) if col_ant else 0
        p_b = f"Atendida ({qtd_ant}/2)" if qtd_ant >= 2 else f"Não atendida ({qtd_ant}/2)"

        col_vis = buscar_coluna_flexivel(df, ['visitas'])
        dts_v = extrair_datas_validas(row.get(col_vis), data_ref, 365) if col_vis else []
        p_c = "Não atendida (0/2)"
        if len(dts_v) >= 2:
            for i in range(len(dts_v)-1):
                if (dts_v[i+1] - dts_v[i]).days >= 30:
                    p_c = "Atendida (2/2)"
                    break

        col_flu = buscar_coluna_flexivel(df, ['influenza'])
        flu = str(row.get(col_flu, '') or '').strip() if col_flu else ''
        p_d = "Atendida (1/1)" if flu and flu not in ['-', 'Sem registro', 'None'] else "Não atendida (0/1)"

        pts = sum([1 for s in [p_a, p_b, p_c, p_d] if "Atendida" in s])
        score = (pts / 4.0) * 100
        return pd.Series([p_a, p_b, p_c, p_d, score],
                         index=['Prática A — consulta 12m — status', 'Prática B — 2 antropometrias 12m — status',
                                'Prática C — 2 visitas 12m — status', 'Prática D — vacina influenza 12m — status', 'Score_C6_%'])

    res = df.apply(eval_c6, axis=1)
    for c in res.columns: df[c] = res[c]

    cols_nom = ['Microárea', 'Nome', 'CPF', 'Telefone celular', 'Endereço',
                'Prática A — consulta 12m — status', 'Prática B — 2 antropometrias 12m — status',
                'Prática C — 2 visitas 12m — status', 'Prática D — vacina influenza 12m — status', 'Score_C6_%']
    for c in cols_nom:
        if c not in df.columns: df[c] = "-"

    df_micro = df.groupby('Microárea').agg(
        Idosos_Ativos=('Nome', 'count'),
        Score_Médio_C6=('Score_C6_%', 'mean'),
        Pct_100_Práticas=('Score_C6_%', lambda x: (x == 100).mean() * 100)
    ).reset_index()

    return df_micro, df[cols_nom]
