import streamlit as st
import pandas as pd
import re
import calendar
from datetime import datetime

from utils import (
    carregar_e_limpar_csv,
    montar_endereco,
    buscar_coluna_flexivel,
    gerar_excel_padrao
)

def render():
    st.header("👶 Auditoria Indicador C2 — Puericultura (Crianças < 3 anos)")
    st.caption("Acompanhamento do Desenvolvimento Infantil conforme Nota Metodológica SAPS/MS")
    
    uploaded_file = st.file_uploader("Anexe o relatório CSV de Puericultura:", type=["csv"], key="uploader_c2")
    
    if uploaded_file and st.button("🚀 Processar Dashboard C2", key="btn_c2"):
        with st.spinner("Processando dados do Indicador C2..."):
            df, data_ref = carregar_e_limpar_csv(uploaded_file.getvalue())
            df_micro, df_nom = processar_c2(df, data_ref)
            
            cols_status = [c for c in df_nom.columns if "Prática" in c or "status" in c]
            excel_bytes = gerar_excel_padrao(df_micro, df_nom, "INDICADOR C2 — PUERICULTURA", data_ref, cols_status, "Score_C2_%")
            
            st.success("✅ Dashboard do C2 gerado com sucesso!")
            st.download_button(
                label="📥 Baixar Dashboard_Indicador_C2.xlsx",
                data=excel_bytes,
                file_name="Dashboard_Indicador_C2.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

def calcular_idade_extenso(dt_nasc, data_ref):
    if not dt_nasc or str(dt_nasc).strip() in ['-', 'None', 'nan', '']:
        return "Não informada"
    try:
        dt = datetime.strptime(str(dt_nasc)[:10], '%d/%m/%Y')
    except:
        return "Não informada"
    
    y1, m1, d1 = dt.year, dt.month, dt.day
    y2, m2, d2 = data_ref.year, data_ref.month, data_ref.day
    
    anos = y2 - y1
    meses = m2 - m1
    dias = d2 - d1
    
    if dias < 0:
        meses -= 1
        prev_month = m2 - 1 if m2 > 1 else 12
        prev_year = y2 if m2 > 1 else y2 - 1
        dias += calendar.monthrange(prev_year, prev_month)[1]
        
    if meses < 0:
        anos -= 1
        meses += 12
        
    partes = []
    if anos > 0:
        partes.append(f"{anos} {'ano' if anos == 1 else 'anos'}")
    if meses > 0:
        partes.append(f"{meses} {'mês' if meses == 1 else 'meses'}")
    if dias > 0 or not partes:
        partes.append(f"{dias} {'dia' if dias == 1 else 'dias'}")
        
    return " e ".join(partes)

def contar_doses_penta(txt):
    if not txt or str(txt).strip() in ['-', 'Sem registro', 'None', 'nan', '']: return 0
    doses = re.findall(r'\b(D1|D2|D3|R1|R2)\b', str(txt))
    return len(set(doses)) if doses else (1 if 'D -' in str(txt) or 'D1' in str(txt) else 0)

def contar_doses_vip(txt):
    if not txt or str(txt).strip() in ['-', 'Sem registro', 'None', 'nan', '']: return 0
    doses = re.findall(r'\b(D1|D2|D3|REF|R1)\b', str(txt))
    return len(set(doses)) if doses else (1 if 'D1' in str(txt) or 'VIP' in str(txt) else 0)

def contar_doses_vpc(txt):
    if not txt or str(txt).strip() in ['-', 'Sem registro', 'None', 'nan', '']: return 0
    doses = re.findall(r'\b(D1|D2|REF|DU)\b', str(txt))
    return len(set(doses)) if doses else (1 if 'D1' in str(txt) or 'VPC' in str(txt) else 0)

def contar_doses_scr(txt):
    if not txt or str(txt).strip() in ['-', 'Sem registro', 'None', 'nan', '']: return 0
    doses = re.findall(r'\b(D1|D2|DU)\b', str(txt))
    return len(set(doses)) if doses else (1 if 'D1' in str(txt) or 'DU' in str(txt) or 'SCR' in str(txt) else 0)

def avaliar_vacinas_detalhado(row, data_ref, df_ref):
    col_penta = buscar_coluna_flexivel(df_ref, ['difteria, tetano, pertusis'])
    col_vip = buscar_coluna_flexivel(df_ref, ['poliomielite'])
    col_scr = buscar_coluna_flexivel(df_ref, ['sarampo, caxumba'])
    col_vpc = buscar_coluna_flexivel(df_ref, ['pneumococica'])

    nasc_str = str(row.get('Data de nascimento', '') or '').strip()
    if not nasc_str or nasc_str in ['-', 'None', 'nan']:
        return "Não informada"
    
    try:
        dt = datetime.strptime(nasc_str[:10], '%d/%m/%Y')
        dias = (data_ref - dt).days
    except:
        return "Não informada"

    d_penta = contar_doses_penta(row.get(col_penta))
    d_vip = contar_doses_vip(row.get(col_vip))
    d_vpc = contar_doses_vpc(row.get(col_vpc))
    d_scr = contar_doses_scr(row.get(col_scr))

    # Doses esperadas para a idade atual
    esp_penta = 0 if dias < 60 else (1 if dias < 120 else (2 if dias < 180 else 3))
    esp_vip = 0 if dias < 60 else (1 if dias < 120 else (2 if dias < 180 else 3))
    esp_vpc = 0 if dias < 60 else (1 if dias < 120 else (2 if dias < 365 else 3))
    esp_scr = 0 if dias < 365 else (1 if dias < 450 else 2)

    # Identificar quais vacinas específicas estão pendentes para a idade atual
    pendentes = []
    if d_penta < esp_penta:
        for dose_num in range(d_penta + 1, esp_penta + 1):
            pendentes.append(f"Penta D{dose_num}")
            
    if d_vip < esp_vip:
        for dose_num in range(d_vip + 1, esp_vip + 1):
            pendentes.append(f"VIP D{dose_num}")

    if d_vpc < esp_vpc:
        for dose_num in range(d_vpc + 1, esp_vpc + 1):
            label = f"Pneumo D{dose_num}" if dose_num <= 2 else "Pneumo Reforço"
            pendentes.append(label)

    if d_scr < esp_scr:
        for dose_num in range(d_scr + 1, esp_scr + 1):
            label = "SCR D1" if dose_num == 1 else "SCR D2/Tetraviral"
            pendentes.append(label)

    # Esquema completo do indicador C2 (Penta 3, VIP 3, Pneumo 3, SCR 2)
    esquema_completo = (d_penta >= 3) and (d_vip >= 3) and (d_vpc >= 3) and (d_scr >= 2)

    if pendentes:
        msg_pend = ", ".join(pendentes)
        return f"Não atendida (Pendente: {msg_pend})"
    elif esquema_completo:
        return "Adequado (Esquema completo)"
    else:
        return "Em acompanhamento (Vacinação em dia)"

def processar_c2(df, data_ref):
    df['Endereço'] = df.apply(montar_endereco, axis=1)

    col_nasc = buscar_coluna_flexivel(df, ['data de nascimento', 'nascimento'])
    
    def parse_idade_dias(row):
        nasc_str = str(row.get(col_nasc, '') or '').strip() if col_nasc else ''
        if nasc_str and nasc_str not in ['-', 'None', 'nan']:
            try:
                dt = datetime.strptime(nasc_str[:10], '%d/%m/%Y')
                return (data_ref - dt).days
            except: pass
        return 0

    df['Idade_Dias'] = df.apply(parse_idade_dias, axis=1)
    df['Idade'] = df.apply(lambda r: calcular_idade_extenso(r.get(col_nasc), data_ref), axis=1)
    
    # Filtro da população ativa do C2 (até 3 anos incompletos / 1095 dias)
    df_filtered = df[df['Idade_Dias'] <= 1095].copy()
    if df_filtered.empty:
        df_filtered = df.copy()

    def eval_c2_row(row):
        dias = row['Idade_Dias']
        
        # 1. Prática A: 1ª consulta até 30d
        col_1c = buscar_coluna_flexivel(df_filtered, ['idade na primeira consulta', 'primeira consulta'])
        val_1c = str(row.get(col_1c, '') or '').strip().lower() if col_1c else ''
        
        p_a = "Não atendida (0/1)"
        if val_1c and val_1c not in ['-', 'none', 'nan', '']:
            if 'ano' not in val_1c:
                if 'mês' not in val_1c and 'mes' not in val_1c:
                    p_a = "Atendida (1/1)"
                elif '0 mês' in val_1c or '0 mes' in val_1c:
                    p_a = "Atendida (1/1)"
                elif '1 mês' in val_1c or '1 mes' in val_1c:
                    if ' e ' not in val_1c or '0 dia' in val_1c:
                        p_a = "Atendida (1/1)"
        
        if p_a != "Atendida (1/1)" and dias <= 30:
            p_a = "Aguardando idade"

        # Cronograma de metas esperadas por Idade Atual
        if dias < 7: meta_esp = 0
        elif dias < 30: meta_esp = 1
        elif dias < 60: meta_esp = 2
        elif dias < 120: meta_esp = 3
        elif dias < 180: meta_esp = 4
        elif dias < 270: meta_esp = 5
        elif dias < 365: meta_esp = 6
        elif dias < 540: meta_esp = 7
        elif dias < 730: meta_esp = 8
        else: meta_esp = 9

        # 2. Prática B: Consultas até 24 meses
        col_qcons = buscar_coluna_flexivel(df_filtered, ['quantidade de consultas ate 24 meses', 'consultas ate 24'])
        q_b = int(pd.to_numeric(row.get(col_qcons, 0), errors='coerce') or 0) if col_qcons else 0
        
        if dias < 730:
            if q_b >= meta_esp:
                p_b = f"Em acompanhamento ({q_b}/{meta_esp})"
            else:
                p_b = f"Não atendida ({q_b}/{meta_esp})"
        else:
            if q_b >= 9:
                p_b = "Atendida (9/9)"
            else:
                p_b = f"Não atendida ({q_b}/9)"

        # 3. Prática C: Antropometrias até 24 meses
        col_qant = buscar_coluna_flexivel(df_filtered, ['medicoes de peso/altura simultaneas', 'simultaneas ate 24'])
        q_c = int(pd.to_numeric(row.get(col_qant, 0), errors='coerce') or 0) if col_qant else 0
        
        if dias < 730:
            if q_c >= meta_esp:
                p_c = f"Em acompanhamento ({q_c}/{meta_esp})"
            else:
                p_c = f"Não atendida ({q_c}/{meta_esp})"
        else:
            if q_c >= 9:
                p_c = "Atendida (9/9)"
            else:
                p_c = f"Não atendida ({q_c}/9)"

        # 4. Prática D: Visitas ACS até 24 meses
        meta_vis = 1 if dias < 30 else 2
        col_qvis = buscar_coluna_flexivel(df_filtered, ['visitas domiciliares ate os 24 meses', 'visitas ate os 24'])
        q_d = int(pd.to_numeric(row.get(col_qvis, 0), errors='coerce') or 0) if col_qvis else 0
        
        if q_d >= meta_vis:
            p_d = f"Atendida ({q_d}/{meta_vis})"
        else:
            p_d = f"Não atendida ({q_d}/{meta_vis})"

        # 5. Prática E: Vacinação em dia detalhada
        p_e = avaliar_vacinas_detalhado(row, data_ref, df_filtered)

        # Cálculo do Score C2 (%)
        pts = 0
        if "Atendida" in p_a: pts += 1
        if "Atendida" in p_b or "Em acompanhamento" in p_b: pts += 1
        if "Atendida" in p_c or "Em acompanhamento" in p_c: pts += 1
        if "Atendida" in p_d: pts += 1
        if "Adequado" in p_e or "Em acompanhamento" in p_e: pts += 1

        score = (pts / 5.0) * 100

        return pd.Series([
            p_a, p_b, p_c, p_d, p_e, score
        ], index=[
            'Prática A — 1ª consulta até 30d — status',
            'Prática B — 9 consultas — status',
            'Prática C — 9 antropometrias — status',
            'Prática D — visitas ACS — status',
            'Prática E — vacinação em dia — status',
            'Score_C2_%'
        ])

    res = df_filtered.apply(eval_c2_row, axis=1)
    for col in res.columns:
        df_filtered[col] = res[col]

    cols_nom = [
        'Microárea', 'Nome', 'CPF', 'Idade', 'Telefone celular', 'Endereço',
        'Prática A — 1ª consulta até 30d — status',
        'Prática B — 9 consultas — status',
        'Prática C — 9 antropometrias — status',
        'Prática D — visitas ACS — status',
        'Prática E — vacinação em dia — status',
        'Score_C2_%'
    ]
    
    for c in cols_nom:
        if c not in df_filtered.columns: df_filtered[c] = "-"

    df_micro = df_filtered.groupby('Microárea').agg(
        Crianças_Ativas=('Nome', 'count'),
        Score_Médio_C2=('Score_C2_%', 'mean'),
        Pct_100_Avaliável=('Score_C2_%', lambda x: (x == 100).mean() * 100)
    ).reset_index()

    return df_micro, df_filtered[cols_nom]
