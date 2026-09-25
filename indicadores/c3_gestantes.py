import streamlit as st
import pandas as pd
import re

from utils import (
    carregar_e_limpar_csv,
    montar_endereco,
    buscar_coluna_flexivel,
    gerar_excel_padrao
)

def render():
    st.header("🤰 Auditoria Indicador C3 — Saúde da Gestante e Puérpera")
    st.caption("População alvo: Gestantes ativas e Puérperas")
    
    uploaded_file = st.file_uploader("Anexe o relatório CSV de Gestantes/Puérperas:", type=["csv"], key="uploader_c3")
    
    if uploaded_file and st.button("🚀 Processar Dashboard C3", key="btn_c3"):
        with st.spinner("Processando dados do C3..."):
            df, data_ref = carregar_e_limpar_csv(uploaded_file.getvalue())
            df_micro, df_nom = processar_c3(df, data_ref)
            
            cols_status = [c for c in df_nom.columns if "Prática" in c or "status" in c]
            excel_bytes = gerar_excel_padrao(df_micro, df_nom, "INDICADOR C3 — SAÚDE DA GESTANTE E PUÉRPERA", data_ref, cols_status, "Score_C3_%")
            
            st.success("✅ Dashboard do C3 gerado com sucesso!")
            st.download_button(
                label="📥 Baixar Dashboard_Indicador_C3.xlsx",
                data=excel_bytes,
                file_name="Dashboard_Indicador_C3.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

def processar_c3(df, data_ref):
    df['Endereço'] = df.apply(montar_endereco, axis=1)
    
    col_ig_sem = buscar_coluna_flexivel(df, ['ig (dum) (semanas)', 'ig semanas'])
    col_ig_dias = buscar_coluna_flexivel(df, ['ig (dum) (dias)', 'ig dias'])
    
    df['IG_sem_num'] = pd.to_numeric(df[col_ig_sem], errors='coerce').fillna(0).astype(int) if col_ig_sem else 0
    df['IG_dias_num'] = pd.to_numeric(df[col_ig_dias], errors='coerce').fillna(0).astype(int) if col_ig_dias else 0
    df['IG_total_dias'] = df['IG_sem_num'] * 7 + df['IG_dias_num']

    df_filtered = df[(df['IG_total_dias'] >= 0) & (df['IG_total_dias'] <= 336)].copy()
    if df_filtered.empty:
        df_filtered = df.copy()

    def converter_ig_texto(sem):
        if sem <= 4: mes = "1º Mês"
        elif sem <= 8: mes = "2º Mês"
        elif sem <= 13: mes = "3º Mês"
        elif sem <= 17: mes = "4º Mês"
        elif sem <= 22: mes = "5º Mês"
        elif sem <= 27: mes = "6º Mês"
        elif sem <= 31: mes = "7º Mês"
        elif sem <= 35: mes = "8º Mês"
        elif sem <= 40: mes = "9º Mês"
        elif sem < 42: mes = "Pós-data"
        else: mes = "Pós-termo: ≥42 semanas"
        return f"{sem} semanas ({mes})"

    df_filtered['Idade gestacional atual'] = df_filtered['IG_sem_num'].apply(converter_ig_texto)
    
    col_ult_cons = buscar_coluna_flexivel(df_filtered, ['ultima consulta de pre-natal', 'ultima consulta'])
    df_filtered['Data da última consulta'] = df_filtered[col_ult_cons].fillna('Não informada') if col_ult_cons else 'Não informada'

    col_ult_vis_meses = buscar_coluna_flexivel(df_filtered, ['meses desde a ultima visita'])
    def format_visita(v):
        if pd.notnull(v):
            return f"{v} meses"
        return "Não informada"
    
    df_filtered['Data da última visita'] = df_filtered[col_ult_vis_meses].apply(format_visita) if col_ult_vis_meses else 'Não informada'

    def eval_c3_row(row):
        sem = row['IG_sem_num']
        ig_dias = row['IG_total_dias']
        is_puerpera = sem >= 40 or ig_dias > 294

        def get_num(termos):
            c = buscar_coluna_flexivel(df_filtered, termos)
            if c and pd.notnull(row.get(c)):
                try:
                    nums = re.findall(r'\d+', str(row.get(c)))
                    if nums: return int(nums[0])
                except: pass
            return 0

        if sem <= 28: esperadas = max(1, sem // 4)
        elif sem <= 36: esperadas = 7 + max(0, (sem - 28) // 2)
        else: esperadas = 11 + max(0, sem - 36)
        meta_7 = min(esperadas, 7)
        meta_3 = min(max(1, sem // 10), 3)

        p_a = "Atendida" if get_num(['ate 12 semanas', '12 semanas']) >= 1 else "Não atendida"

        q_b = min(get_num(['quantidade de atendimentos no pre-natal']), 7)
        p_b = f"Atendida ({q_b}/{meta_7})" if q_b >= meta_7 else f"Não atendida ({q_b}/{meta_7})"

        q_c = min(get_num(['medicoes de pressao arterial', 'pressao arterial']), 7)
        p_c = f"Atendida ({q_c}/{meta_7})" if q_c >= meta_7 else f"Não atendida ({q_c}/{meta_7})"

        q_d = min(get_num(['medicoes simultaneas de peso e altura', 'simultaneas de peso e altura']), 7)
        p_d = f"Atendida ({q_d}/{meta_7})" if q_d >= meta_7 else f"Não atendida ({q_d}/{meta_7})"

        q_e = min(get_num(['quantidade de visitas domiciliares no pre-natal']), 3)
        p_e = f"Atendida ({q_e}/{meta_3})" if q_e >= meta_3 else f"Não atendida ({q_e}/{meta_3})"

        col_dtpa = buscar_coluna_flexivel(df_filtered, ['dtpa'])
        dtpa_v = str(row.get(col_dtpa, '') or '').strip() if col_dtpa else ''
        p_f = "Atendida (1/1)" if dtpa_v and dtpa_v not in ['-', 'None', 'nan', ''] else ("Aguardando idade" if sem < 20 else "Não atendida (0/1)")

        col_hiv1 = buscar_coluna_flexivel(df_filtered, ['hiv no primeiro'])
        col_sif1 = buscar_coluna_flexivel(df_filtered, ['sifilis no primeiro'])
        col_hepb1 = buscar_coluna_flexivel(df_filtered, ['hepatite b no primeiro'])
        col_hepc1 = buscar_coluna_flexivel(df_filtered, ['hepatite c no primeiro'])
        
        ex_1t = [row.get(c) for c in [col_hiv1, col_sif1, col_hepb1, col_hepc1] if c]
        all_1t_sim = len(ex_1t) == 4 and all([str(x or '').upper() == 'SIM' for x in ex_1t])
        p_g = "Atendida (1/1)" if all_1t_sim else ("Aguardando idade" if sem <= 13 else "Não atendida (0/1)")

        col_hiv3 = buscar_coluna_flexivel(df_filtered, ['hiv no terceiro'])
        col_sif3 = buscar_coluna_flexivel(df_filtered, ['sifilis no terceiro'])
        
        ex_3t = [row.get(c) for c in [col_hiv3, col_sif3] if c]
        all_3t_sim = len(ex_3t) == 2 and all([str(x or '').upper() == 'SIM' for x in ex_3t])
        p_h = "Atendida (1/1)" if all_3t_sim else ("Aguardando idade" if sem < 28 else "Não atendida (0/1)")

        q_i = get_num(['atendimentos no puerperio'])
        p_i = ("Atendida (1/1)" if q_i > 0 else "Não atendida (0/1)") if is_puerpera else "Não se aplica"

        q_j = get_num(['visitas domiciliares no puerperio'])
        p_j = ("Atendida (1/1)" if q_j > 0 else "Não atendida (0/1)") if is_puerpera else "Não se aplica"

        q_k = get_num(['atendimentos odontologicos no pre-natal'])
        p_k = "Atendida (1/1)" if q_k >= 1 else ("Em acompanhamento (0/1)" if not is_puerpera else "Não atendida (0/1)")

        praticas = [p_a, p_b, p_c, p_d, p_e, p_f, p_g, p_h, p_i, p_j, p_k]
        avaliaveis = [p for p in praticas if ("Atendida" in p or "Não atendida" in p) and "Não se aplica" not in p and "Aguardando" not in p and "acompanhamento" not in p]
        atendidas = [p for p in praticas if "Atendida" in p]
        score = (len(atendidas) / len(avaliaveis) * 100) if avaliaveis else 100.0

        return pd.Series([
            p_a, p_b, p_c, p_d, p_e, p_f, p_g, p_h, p_i, p_j, p_k, score
        ], index=[
            'Prática A — 1ª consulta até 12ª sem — status',
            'Prática B — consultas de pré-natal — status',
            'Prática C — aferição de pressão arterial — status',
            'Prática D — peso e altura — status',
            'Prática E — visitas ACS pré-natal — status',
            'Prática F — vacina dTpa a partir de 20 sem — status',
            'Prática G — exames 1º trimestre (HIV/Sífilis/HepB/HepC) — status',
            'Prática H — exames 3º trimestre (HIV/Sífilis) — status',
            'Prática I — consulta puerpério — status',
            'Prática J — visita ACS puerpério — status',
            'Prática K — atendimento odontológico — status',
            'Score_C3_%'
        ])

    res = df_filtered.apply(eval_c3_row, axis=1)
    for col in res.columns:
        df_filtered[col] = res[col]

    cols_nom = [
        'Microárea', 'Nome', 'CPF', 'Telefone celular', 'Endereço', 'Risco gestacional',
        'Idade gestacional atual', 'Data da última consulta', 'Data da última visita',
        'Prática A — 1ª consulta até 12ª sem — status',
        'Prática B — consultas de pré-natal — status',
        'Prática C — aferição de pressão arterial — status',
        'Prática D — peso e altura — status',
        'Prática E — visitas ACS pré-natal — status',
        'Prática F — vacina dTpa a partir de 20 sem — status',
        'Prática G — exames 1º trimestre (HIV/Sífilis/HepB/HepC) — status',
        'Prática H — exames 3º trimestre (HIV/Sífilis) — status',
        'Prática I — consulta puerpério — status',
        'Prática J — visita ACS puerpério — status',
        'Prática K — atendimento odontológico — status',
        'Score_C3_%'
    ]
    
    for c in cols_nom:
        if c not in df_filtered.columns: df_filtered[c] = "-"

    df_micro = df_filtered.groupby('Microárea').agg(
        Gestantes_Puerperas_Ativas=('Nome', 'count'),
        Score_Médio_C3=('Score_C3_%', 'mean'),
        Pct_100_Avaliável=('Score_C3_%', lambda x: (x == 100).mean() * 100)
    ).reset_index()

    return df_micro, df_filtered[cols_nom]
