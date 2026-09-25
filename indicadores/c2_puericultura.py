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
    st.header("👶 Auditoria Indicador C2 — Puericultura (Crianças < 3 anos)")
    uploaded_file = st.file_uploader("Anexe o relatório CSV de Puericultura:", type=["csv"], key="uploader_c2")
    
    if uploaded_file and st.button("🚀 Processar Dashboard C2", key="btn_c2"):
        with st.spinner("Processando C2..."):
            df, data_ref = carregar_e_limpar_csv(uploaded_file.getvalue())
            df_micro, df_nom = processar_c2(df, data_ref)
            cols_status = [c for c in df_nom.columns if "Prática" in c or "status" in c]
            excel_bytes = gerar_excel_padrao(df_micro, df_nom, "INDICADOR C2 — PUERICULTURA", data_ref, cols_status, "Score_C2_%")
            
            st.success("✅ Dashboard C2 gerado!")
            st.download_button("📥 Baixar Dashboard_Indicador_C2.xlsx", excel_bytes, "Dashboard_Indicador_C2.xlsx")

def processar_c2(df, data_ref):
    df['Endereço'] = df.apply(montar_endereco, axis=1)
    col_nasc = buscar_coluna_flexivel(df, ['nascimento', 'dt_nasc'])
    if col_nasc:
        df['Nascimento'] = pd.to_datetime(df[col_nasc], format='%d/%m/%Y', errors='coerce')
        df['Idade_Dias'] = (data_ref - df['Nascimento']).dt.days
        df = df[df['Idade_Dias'] < 1096].copy()
    else: df['Idade_Dias'] = 300

    def eval_c2(row):
        dias = row.get('Idade_Dias', 300)
        col_cons1 = buscar_coluna_flexivel(df, ['primeira consulta', 'precoce'])
        p_a = "Não atendida (0/1)"
        if col_cons1 and row.get(col_cons1):
            try:
                dt = datetime.strptime(str(row.get(col_cons1))[:10], '%d/%m/%Y')
                if (dt - row.get('Nascimento', dt)).days <= 30: p_a = "Atendida (1/1)"
            except: pass
        if p_a != "Atendida (1/1)" and dias <= 30: p_a = "Aguardando idade"

        col_q_cons = buscar_coluna_flexivel(df, ['consultas de puericultura', 'quantidade de consultas'])
        qtd_cons = min(int(pd.to_numeric(row.get(col_q_cons, 0), errors='coerce') or 0), 9) if col_q_cons else 0
        p_b = f"Atendida ({qtd_cons}/9)" if qtd_cons >= 9 else ("Aguardando idade" if dias < 730 else f"Não atendida ({qtd_cons}/9)")

        col_q_ant = buscar_coluna_flexivel(df, ['simultaneas de peso e altura', 'antropometria'])
        qtd_ant = min(int(pd.to_numeric(row.get(col_q_ant, 0), errors='coerce') or 0), 9) if col_q_ant else 0
        p_c = f"Atendida ({qtd_ant}/9)" if qtd_ant >= 9 else ("Aguardando idade" if dias < 730 else f"Não atendida ({qtd_ant}/9)")

        col_vis = buscar_coluna_flexivel(df, ['visita', 'visitas'])
        dts_v = extrair_datas_validas(row.get(col_vis), data_ref, 180) if col_vis else []
        p_d = "Atendida (2/2)" if len(dts_v) >= 2 else ("Aguardando idade" if dias < 180 else "Não atendida (0/2)")

        col_vac = buscar_coluna_flexivel(df, ['vacina', 'situacao vacinal'])
        p_e = "Atendida (1/1)" if col_vac and str(row.get(col_vac, '')).lower() in ['em dia', 'sim', 'completa'] else ("Aguardando idade" if dias < 365 else "Não atendida (0/1)")

        praticas = [p_a, p_b, p_c, p_d, p_e]
        avaliaveis = [p for p in praticas if "Atendida" in p or "Não atendida" in p]
        atendidas = [p for p in praticas if "Atendida" in p]
        score = (len(atendidas) / len(avaliaveis) * 100) if avaliaveis else 100.0

        return pd.Series([p_a, p_b, p_c, p_d, p_e, score],
                         index=['Prática A — 1ª consulta até 30d — status', 'Prática B — 9 consultas — status', 'Prática C — 9 antropometrias — status',
                                'Prática D — visitas ACS — status', 'Prática E — vacinação em dia — status', 'Score_C2_%'])

    res = df.apply(eval_c2, axis=1)
    for c in res.columns: df[c] = res[c]

    cols_nom = ['Microárea', 'Nome', 'CPF', 'Telefone celular', 'Endereço', 'Prática A — 1ª consulta até 30d — status', 'Prática B — 9 consultas — status', 
                'Prática C — 9 antropometrias — status', 'Prática D — visitas ACS — status', 'Prática E — vacinação em dia — status', 'Score_C2_%']
    for c in cols_nom:
        if c not in df.columns: df[c] = "-"

    df_micro = df.groupby('Microárea').agg(
        Crianças_Ativas=('Nome', 'count'),
        Score_Médio=('Score_C2_%', 'mean'),
        Pct_100_Avaliável=('Score_C2_%', lambda x: (x == 100).mean() * 100)
    ).reset_index()

    return df_micro, df[cols_nom]
