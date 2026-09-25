import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import io
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="Auditoria e-SUS APS - SAPS/MS", layout="wide")

st.title("🏥 Sistema de Auditoria de Indicadores e-SUS APS (SAPS/MS)")
st.subheader("Processamento Automático e Determinístico (C2, C3, C4, C5, C6, C7)")

# --- UTILITÁRIOS DE LEITURA E HIGIENIZAÇÃO DE CSV ---

def extrair_data_referencia(conteudo_bytes):
    try:
        texto = conteudo_bytes.decode('latin1', errors='ignore')
        match = re.search(r'Gerado em;?\s*(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE)
        if match:
            return datetime.strptime(match.group(1), '%d/%m/%Y')
    except Exception:
        pass
    return datetime.now()

def carregar_e_limpar_csv(file_bytes):
    data_ref = extrair_data_referencia(file_bytes)
    lines = file_bytes.decode('latin1', errors='ignore').splitlines()
    
    header_idx = 0
    for idx, line in enumerate(lines[:30]):
        if line.startswith("Nome;") or ";Nome;" in line or line.startswith('"Nome";'):
            header_idx = idx
            break
            
    df = pd.read_csv(
        io.BytesIO(file_bytes),
        sep=';',
        encoding='latin1',
        skiprows=header_idx,
        dtype=str
    )
    
    df.columns = [str(c).strip().replace('"', '') for c in df.columns]
    df = df.map(lambda x: x.strip().replace('"', '') if isinstance(x, str) else x)
    df = df.replace({'': None, '-': None, 'nan': None, 'NaN': None})
    return df, data_ref

def montar_endereco(row):
    rua = str(row.get('Rua', '') or '').strip()
    num = str(row.get('Número', '') or '').strip()
    comp = str(row.get('Complemento', '') or '').strip()
    partes = [p for p in [rua, num, comp] if p and p not in ['None', 'nan', '']]
    return ", ".join(partes) if partes else "Endereço não informado"

def extrair_datas_validas(visitas_txt, data_ref, janela_dias=365):
    if not visitas_txt: return []
    datas_v = re.findall(r'\d{2}/\d{2}/\d{4}', str(visitas_txt))
    dts = []
    for d in datas_v:
        try:
            dt = datetime.strptime(d, '%d/%m/%Y')
            if 0 <= (data_ref - dt).days <= janela_dias:
                dts.append(dt)
        except: pass
    return sorted(list(set(dts)))

# --- LÓGICA DE GERADOR DE EXCEL PADRONIZADO ---

def gerar_excel_padrao(df_micro, df_nominal, titulo, data_ref, cols_status, col_score):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    blue_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    red_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    gray_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

    # Aba 1: Dashboard_Microarea
    ws1 = wb.active
    ws1.title = "Dashboard_Microarea"
    ws1.append([titulo])
    ws1.append([f"Data de Referência: {data_ref.strftime('%d/%m/%Y')} | Unidade de Saúde / e-SUS APS"])
    ws1.append([])
    
    ws1.cell(row=1, column=1).font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    ws1.cell(row=2, column=1).font = Font(name="Calibri", size=10, italic=True)

    headers_micro = list(df_micro.columns)
    ws1.append(headers_micro)
    
    for col_num in range(1, len(headers_micro) + 1):
        cell = ws1.cell(row=4, column=col_num)
        cell.fill = blue_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r in dataframe_to_rows(df_micro, index=False, header=False):
        ws1.append(r)

    for row in ws1.iter_rows(min_row=5, max_row=ws1.max_row, min_col=3, max_col=len(headers_micro)):
        for cell in row:
            cell.number_format = '0.0"%"'
            cell.border = border_thin

    # Aba 2: Auditoria_Nominal
    ws2 = wb.create_sheet(title="Auditoria_Nominal")
    headers_nom = list(df_nominal.columns)
    ws2.append(headers_nom)
    
    for col_num in range(1, len(headers_nom) + 1):
        cell = ws2.cell(row=1, column=col_num)
        cell.fill = blue_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, r in enumerate(dataframe_to_rows(df_nominal, index=False, header=False), start=2):
        ws2.append(r)
        for c_idx, col_name in enumerate(headers_nom, start=1):
            cell = ws2.cell(row=r_idx, column=c_idx)
            cell.border = border_thin
            val = str(cell.value or '')
            
            if col_name in cols_status or "Prática" in col_name or "status" in col_name:
                if "Atendida" in val:
                    cell.fill = green_fill
                elif "Não atendida" in val:
                    cell.fill = red_fill
                elif "Fora da faixa" in val or "Aguardando" in val or "Não se aplica" in val:
                    cell.fill = gray_fill
            elif col_name == col_score:
                try:
                    cell.value = float(cell.value)
                    cell.number_format = '0.0"%"'
                except: pass

    for ws in [ws1, ws2]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output)
    return output.getvalue()

# --- REGRAS DOS INDICADORES (C2 - C7) ---

def processar_c2(df, data_ref):
    df['Nascimento'] = pd.to_datetime(df.get('Data de nascimento'), format='%d/%m/%Y', errors='coerce')
    df['Idade_Dias'] = (data_ref - df['Nascimento']).dt.days
    df = df[df['Idade_Dias'] < 1096].copy() # Menores de 3 anos
    
    def eval_c2(row):
        dias = row['Idade_Dias']
        # Prática A (1ª Consulta <= 30d)
        p_a = "Não atendida"
        d_cons = row.get('Data da primeira consulta')
        if d_cons:
            try:
                dt = datetime.strptime(str(d_cons)[:10], '%d/%m/%Y')
                if (dt - row['Nascimento']).days <= 30: p_a = "Atendida"
            except: pass
        if p_a != "Atendida" and dias <= 30: p_a = "Aguardando idade"

        # Prática B (Consultas e-SUS 9 marcos)
        qtd_cons = int(pd.to_numeric(row.get('Quantidade de consultas de puericultura', 0), errors='coerce') or 0)
        p_b = "Atendida" if qtd_cons >= 9 else ("Aguardando idade" if dias < 730 else "Não atendida")

        # Prática C (Antropometria simultânea)
        qtd_ant = int(pd.to_numeric(row.get('Quantidade de medições simultâneas de peso e altura', 0), errors='coerce') or 0)
        p_c = "Atendida" if qtd_ant >= 9 else ("Aguardando idade" if dias < 730 else "Não atendida")

        # Prática D (Visitas ACS)
        dts_v = extrair_datas_validas(row.get('Últimas visitas domiciliares'), data_ref, 180)
        p_d = "Atendida" if len(dts_v) >= 2 else ("Aguardando idade" if dias < 180 else "Não atendida")

        # Prática E (Vacinação)
        p_e = "Atendida" if str(row.get('Situação vacinal', '')).lower() == 'em dia' else ("Aguardando idade" if dias < 365 else "Não atendida")

        praticas = [p_a, p_b, p_c, p_d, p_e]
        avaliaveis = [p for p in praticas if p in ["Atendida", "Não atendida"]]
        atendidas = [p for p in praticas if p == "Atendida"]
        score = (len(atendidas) / len(avaliaveis) * 100) if avaliaveis else 100.0

        return pd.Series([p_a, p_b, p_c, p_d, p_e, score])

    df[['Prática A (1ª Consulta <=30d)', 'Prática B (9 Consultas)', 'Prática C (9 Antropometrias)',
        'Prática D (Visitas ACS)', 'Prática E (Vacinas)', 'Score_C2_%']] = df.apply(eval_c2, axis=1)

    cols_nom = ['Microárea', 'Nome', 'CPF', 'Prática A (1ª Consulta <=30d)', 'Prática B (9 Consultas)', 
                'Prática C (9 Antropometrias)', 'Prática D (Visitas ACS)', 'Prática E (Vacinas)', 'Score_C2_%']
    df_nom = df[cols_nom].copy()
    
    df_micro = df.groupby('Microárea').agg(
        Crianças_Ativas=('Nome', 'count'),
        Score_Médio=('Score_C2_%', 'mean'),
        Pct_100_Avaliável=('Score_C2_%', lambda x: (x == 100).mean() * 100)
    ).reset_index()

    return df_micro, df_nom

def processar_c3(df, data_ref):
    df['DUM_dt'] = pd.to_datetime(df.get('DUM'), format='%d/%m/%Y', errors='coerce')
    df = df[df['DUM_dt'].notnull()].copy()
    
    # Se não houver gestantes com DUM válida no arquivo
    if df.empty:
        cols_c3 = ['Prática A', 'Prática B', 'Prática C', 'Prática D', 'Prática E', 'Prática F',
                   'Prática G', 'Prática H', 'Prática I', 'Prática J', 'Prática K', 'Score_C3_%']
        for c in cols_c3:
            df[c] = []
        cols_nom = ['Microárea', 'Nome', 'CPF', 'Telefone celular', 'Risco gestacional'] + cols_c3
        for c in cols_nom:
            if c not in df.columns: df[c] = []
        df_micro = pd.DataFrame(columns=['Microárea', 'Gestantes_Puerperas_Ativas', 'Score_Médio_C3', 'Pct_100_Avaliável'])
        return df_micro, df[cols_nom]

    df['IG_Dias'] = (data_ref - df['DUM_dt']).dt.days
    df = df[(df['IG_Dias'] >= 0) & (df['IG_Dias'] <= 336)].copy() # Gestantes (0-294d) e Puérperas (295-336d)

    if df.empty:
        cols_c3 = ['Prática A', 'Prática B', 'Prática C', 'Prática D', 'Prática E', 'Prática F',
                   'Prática G', 'Prática H', 'Prática I', 'Prática J', 'Prática K', 'Score_C3_%']
        for c in cols_c3:
            df[c] = []
        cols_nom = ['Microárea', 'Nome', 'CPF', 'Telefone celular', 'Risco gestacional'] + cols_c3
        for c in cols_nom:
            if c not in df.columns: df[c] = []
        df_micro = pd.DataFrame(columns=['Microárea', 'Gestantes_Puerperas_Ativas', 'Score_Médio_C3', 'Pct_100_Avaliável'])
        return df_micro, df[cols_nom]

    def eval_c3(row):
        ig = row.get('IG_Dias', 0)
        is_puerpera = ig > 294
        
        p_a = "Atendida" if int(pd.to_numeric(row.get('Quantidade de atendimentos até 12 semanas no pré-natal', 0), errors='coerce') or 0) >= 1 else "Não atendida"
        p_b = "Atendida" if int(pd.to_numeric(row.get('Quantidade de consultas de pré-natal', 0), errors='coerce') or 0) >= 7 else "Não atendida"
        p_c = "Atendida" if int(pd.to_numeric(row.get('Quantidade de aferições de pressão arterial', 0), errors='coerce') or 0) >= 7 else "Não atendida"
        p_d = "Atendida" if int(pd.to_numeric(row.get('Quantidade de medições de peso e altura', 0), errors='coerce') or 0) >= 7 else "Não atendida"
        p_e = "Atendida" if int(pd.to_numeric(row.get('Quantidade de visitas domiciliares', 0), errors='coerce') or 0) >= 3 else "Não atendida"
        
        dtpa_val = str(row.get('dTpa', '') or '').strip()
        p_f = "Atendida" if dtpa_val and dtpa_val not in ['-', 'None', 'nan'] else ("Aguardando idade" if ig < 140 else "Não atendida")
        
        ex_1t = [row.get('HIV 1ºT'), row.get('Sífilis 1ºT'), row.get('Hep B 1ºT'), row.get('Hep C 1ºT')]
        p_g = "Atendida" if all([str(x or '').upper() in ['SIM', 'REALIZADO'] for x in ex_1t]) else ("Aguardando idade" if ig <= 97 else "Não atendida")

        ex_3t = [row.get('HIV 3ºT'), row.get('Sífilis 3ºT')]
        p_h = "Atendida" if all([str(x or '').upper() in ['SIM', 'REALIZADO'] for x in ex_3t]) else ("Aguardando idade" if ig < 196 else "Não atendida")

        p_i = ("Atendida (1/1)" if row.get('Consulta Puerpério') else "Não atendida (0/1)") if is_puerpera else "Não se aplica"
        p_j = ("Atendida (1/1)" if row.get('Visita Puerpério') else "Não atendida (0/1)") if is_puerpera else "Não se aplica"
        p_k = "Atendida (1/1)" if int(pd.to_numeric(row.get('Atendimento Odontológico', 0), errors='coerce') or 0) >= 1 else "Não atendida"

        praticas = [p_a, p_b, p_c, p_d, p_e, p_f, p_g, p_h, p_i, p_j, p_k]
        avaliaveis = [p for p in praticas if "Atendida" in p or "Não atendida" in p]
        atendidas = [p for p in praticas if "Atendida" in p]
        score = (len(atendidas) / len(avaliaveis) * 100) if avaliaveis else 100.0

        return pd.Series([p_a, p_b, p_c, p_d, p_e, p_f, p_g, p_h, p_i, p_j, p_k, score],
                         index=['Prática A', 'Prática B', 'Prática C', 'Prática D', 'Prática E', 'Prática F',
                                'Prática G', 'Prática H', 'Prática I', 'Prática J', 'Prática K', 'Score_C3_%'])

    res_c3 = df.apply(eval_c3, axis=1)
    for col in res_c3.columns:
        df[col] = res_c3[col]

    cols_nom = ['Microárea', 'Nome', 'CPF', 'Telefone celular', 'Risco gestacional', 
                'Prática A', 'Prática B', 'Prática C', 'Prática D', 'Prática E', 'Prática F', 
                'Prática G', 'Prática H', 'Prática I', 'Prática J', 'Prática K', 'Score_C3_%']
    
    for c in cols_nom:
        if c not in df.columns: df[c] = "-"
        
    df_nom = df[cols_nom].copy()

    df_micro = df.groupby('Microárea').agg(
        Gestantes_Puerperas_Ativas=('Nome', 'count'),
        Score_Médio_C3=('Score_C3_%', 'mean'),
        Pct_100_Avaliável=('Score_C3_%', lambda x: (x == 100).mean() * 100)
    ).reset_index()

    return df_micro, df_nom

def processar_c4(df, data_ref):
    df['Endereço'] = df.apply(montar_endereco, axis=1)

    def eval_c4(row):
        def check_janela(col, dias):
            val = row.get(col)
            if val:
                try:
                    dt = datetime.strptime(str(val)[:10], '%d/%m/%Y')
                    if (data_ref - dt).days <= dias: return "Atendida (1/1)"
                except: pass
            return "Não atendida (0/1)"

        p_a = check_janela('Data da última consulta', 183)
        p_b = check_janela('Data da última medição de pressão arterial', 183)
        p_d = check_janela('Data da última medição de peso e altura', 365)
        
        # HbA1c
        d_eval = row.get('Data da última avaliação de hemoglobina glicada')
        d_sol = row.get('Data da última solicitação de hemoglobina glicada')
        p_e = "Não atendida (0/1)"
        for d in [d_eval, d_sol]:
            if d:
                try:
                    if (data_ref - datetime.strptime(str(d)[:10], '%d/%m/%Y')).days <= 365:
                        p_e = "Atendida (1/1)"
                        break
                except: pass

        p_f = check_janela('Data da avaliação dos pés', 365)

        # Visitas ACS
        dts_v = extrair_datas_validas(row.get('Últimas visitas domiciliares'), data_ref, 365)
        p_c = "Não atendida (0/2)"
        if len(dts_v) >= 2:
            for i in range(len(dts_v)-1):
                if (dts_v[i+1] - dts_v[i]).days >= 30:
                    p_c = "Atendida (2/2)"
                    break

        pts = sum([1 for s in [p_a, p_b, p_c, p_d, p_e, p_f] if "Atendida" in s])
        score = (pts / 6.0) * 100
        return pd.Series([p_a, p_b, p_c, p_d, p_e, p_f, score])

    df[['Prática A — consulta 6m', 'Prática B — pressão 6m', 'Prática C — 2 visitas 12m',
        'Prática D — peso/altura 12m', 'Prática E — HbA1c 12m', 'Prática F — pés 12m', 'Score_C4_%']] = df.apply(eval_c4, axis=1)

    cols_nom = ['Microárea', 'Nome', 'CPF', 'Telefone celular', 'Endereço',
                'Prática A — consulta 6m', 'Prática B — pressão 6m', 'Prática C — 2 visitas 12m',
                'Prática D — peso/altura 12m', 'Prática E — HbA1c 12m', 'Prática F — pés 12m', 'Score_C4_%']
    df_nom = df[cols_nom].copy()

    df_micro = df.groupby('Microárea').agg(
        Diabéticos_Ativos=('Nome', 'count'),
        Score_Médio_C4=('Score_C4_%', 'mean'),
        Pct_100_Práticas=('Score_C4_%', lambda x: (x == 100).mean() * 100)
    ).reset_index()

    return df_micro, df_nom

def processar_c5(df, data_ref):
    df['Endereço'] = df.apply(montar_endereco, axis=1)

    def eval_c5(row):
        def check_janela(col, dias):
            val = row.get(col)
            if val:
                try:
                    dt = datetime.strptime(str(val)[:10], '%d/%m/%Y')
                    if (data_ref - dt).days <= dias: return "Atendida (1/1)"
                except: pass
            return "Não atendida (0/1)"

        p_a = check_janela('Data da última consulta', 183)
        p_b = check_janela('Data da última medição de pressão arterial', 183)
        p_d = check_janela('Data da última medição de peso e altura', 365)

        dts_v = extrair_datas_validas(row.get('Últimas visitas domiciliares'), data_ref, 365)
        p_c = "Não atendida (0/2)"
        if len(dts_v) >= 2:
            for i in range(len(dts_v)-1):
                if (dts_v[i+1] - dts_v[i]).days >= 30:
                    p_c = "Atendida (2/2)"
                    break

        pts = sum([1 for s in [p_a, p_b, p_c, p_d] if "Atendida" in s])
        score = (pts / 4.0) * 100
        return pd.Series([p_a, p_b, p_c, p_d, score])

    df[['Prática A — consulta 6m', 'Prática B — pressão 6m', 
        'Prática C — 2 visitas 12m', 'Prática D — peso/altura 12m', 'Score_C5_%']] = df.apply(eval_c5, axis=1)

    cols_nom = ['Microárea', 'Nome', 'CPF', 'Telefone celular', 'Endereço',
                'Prática A — consulta 6m', 'Prática B — pressão 6m', 
                'Prática C — 2 visitas 12m', 'Prática D — peso/altura 12m', 'Score_C5_%']
    df_nom = df[cols_nom].copy()

    df_micro = df.groupby('Microárea').agg(
        Hipertensos_Ativos=('Nome', 'count'),
        Score_Médio_C5=('Score_C5_%', 'mean'),
        Pct_100_Práticas=('Score_C5_%', lambda x: (x == 100).mean() * 100)
    ).reset_index()

    return df_micro, df_nom

def processar_c6(df, data_ref):
    df['Endereço'] = df.apply(montar_endereco, axis=1)

    def eval_c6(row):
        d_med = pd.to_numeric(row.get('Dias desde o último atendimento médico'), errors='coerce')
        d_enf = pd.to_numeric(row.get('Dias desde o último atendimento de enfermagem'), errors='coerce')
        p_a = "Atendida (1/1)" if ((pd.notnull(d_med) and 0 <= d_med <= 365) or (pd.notnull(d_enf) and 0 <= d_enf <= 365)) else "Não atendida (0/1)"

        qtd_ant = int(pd.to_numeric(row.get('Registros de peso e altura simultâneos nos últimos 12 meses', 0), errors='coerce') or 0)
        p_b = "Atendida (2/2)" if qtd_ant >= 2 else "Não atendida (0/2)"

        dts_v = extrair_datas_validas(row.get('Últimas visitas domiciliares'), data_ref, 365)
        p_c = "Não atendida (0/2)"
        if len(dts_v) >= 2:
            for i in range(len(dts_v)-1):
                if (dts_v[i+1] - dts_v[i]).days >= 30:
                    p_c = "Atendida (2/2)"
                    break

        flu = str(row.get('Influenza (últimos 12 meses)', '') or '').strip()
        p_d = "Atendida (1/1)" if flu and flu not in ['-', 'Sem registro', 'None'] else "Não atendida (0/1)"

        pts = sum([1 for s in [p_a, p_b, p_c, p_d] if "Atendida" in s])
        score = (pts / 4.0) * 100
        return pd.Series([p_a, p_b, p_c, p_d, score])

    df[['Prática A — consulta 12m', 'Prática B — 2 antropometrias 12m',
        'Prática C — 2 visitas 12m', 'Prática D — vacina influenza 12m', 'Score_C6_%']] = df.apply(eval_c6, axis=1)

    cols_nom = ['Microárea', 'Nome', 'CPF', 'Telefone celular', 'Endereço',
                'Prática A — consulta 12m', 'Prática B — 2 antropometrias 12m',
                'Prática C — 2 visitas 12m', 'Prática D — vacina influenza 12m', 'Score_C6_%']
    df_nom = df[cols_nom].copy()

    df_micro = df.groupby('Microárea').agg(
        Idosos_Ativos=('Nome', 'count'),
        Score_Médio_C6=('Score_C6_%', 'mean'),
        Pct_100_Práticas=('Score_C6_%', lambda x: (x == 100).mean() * 100)
    ).reset_index()

    return df_micro, df_nom

def processar_c7(df, data_ref):
    df['Endereço'] = df.apply(montar_endereco, axis=1)
    df['Nascimento'] = pd.to_datetime(df.get('Data de nascimento'), format='%d/%m/%Y', errors='coerce')
    df['Idade'] = ((data_ref - df['Nascimento']).dt.days / 365.25).fillna(0).astype(int)

    def eval_c7(row):
        idade = row['Idade']

        # Prática A (Colo do Útero 25-64a, 36m)
        p_a = "Fora da faixa etária"
        if 25 <= idade <= 64:
            d_sol = row.get('Exame de rastreamento de câncer de colo de útero data última solicitação')
            d_eval = row.get('Exame de rastreamento de câncer de colo de útero data última avaliação')
            p_a = "Não atendida (0/1)"
            for d in [d_sol, d_eval]:
                if d:
                    try:
                        if (data_ref - datetime.strptime(str(d)[:10], '%d/%m/%Y')).days <= 1095:
                            p_a = "Atendida (1/1)"
                            break
                    except: pass

        # Prática B (HPV 9-14a)
        p_b = "Fora da faixa etária"
        if 9 <= idade <= 14:
            hpv = str(row.get('HPV', '') or '').strip()
            p_b = "Atendida (1/1)" if hpv and hpv not in ['-', 'Sem registro', 'None'] else "Não atendida (0/1)"

        # Prática C (Saúde Sexual/Reprodutiva 14-69a, 12m)
        p_c = "Fora da faixa etária"
        if 14 <= idade <= 69:
            d_ssr = row.get('Data da última consulta de saúde sexual e reprodutiva')
            p_c = "Não atendida (0/1)"
            if d_ssr:
                try:
                    if (data_ref - datetime.strptime(str(d_ssr)[:10], '%d/%m/%Y')).days <= 365:
                        p_c = "Atendida (1/1)"
                except: pass

        # Prática D (Mama 50-69a, 24m)
        p_d = "Fora da faixa etária"
        if 50 <= idade <= 69:
            d_sol = row.get('Exame de rastreamento de câncer de mama data Última solicitação')
            d_real = row.get('Exame de rastreamento de câncer de mama data Última realização')
            d_eval = row.get('Exame de rastreamento de câncer de mama data Última avaliação')
            p_d = "Não atendida (0/1)"
            for d in [d_sol, d_real, d_eval]:
                if d:
                    try:
                        if (data_ref - datetime.strptime(str(d)[:10], '%d/%m/%Y')).days <= 730:
                            p_d = "Atendida (1/1)"
                            break
                    except: pass

        praticas = [p_a, p_b, p_c, p_d]
        aplicaveis = [p for p in praticas if p != "Fora da faixa etária"]
        atendidas = [p for p in praticas if "Atendida (1/1)" in p]
        score = (len(atendidas) / len(aplicaveis) * 100) if aplicaveis else 100.0

        return pd.Series([p_a, p_b, p_c, p_d, score])

    df[['Prática A — colo do útero 25-64a (36m)', 'Prática B — vacina HPV 9-14a',
        'Prática C — saúde sexual/reprodutiva 14-69a (12m)', 'Prática D — câncer de mama 50-69a (24m)', 'Score_C7_%']] = df.apply(eval_c7, axis=1)

    cols_nom = ['Microárea', 'Nome', 'CPF', 'Idade', 'Telefone celular', 'Endereço',
                'Prática A — colo do útero 25-64a (36m)', 'Prática B — vacina HPV 9-14a',
                'Prática C — saúde sexual/reprodutiva 14-69a (12m)', 'Prática D — câncer de mama 50-69a (24m)', 'Score_C7_%']
    df_nom = df[cols_nom].copy()

    df_micro = df.groupby('Microárea').agg(
        Mulheres_Relatório=('Nome', 'count'),
        Score_Médio_C7=('Score_C7_%', 'mean'),
        Pct_100_Aplicáveis=('Score_C7_%', lambda x: (x == 100).mean() * 100)
    ).reset_index()

    return df_micro, df_nom


# --- INTERFACE GRÁFICA PRINCIPAL ---

mapa_indicadores = {
    "Indicador C2 — Puericultura (Crianças <3 anos)": ("C2", processar_c2, "INDICADOR C2 — PUERICULTURA", "Score_C2_%"),
    "Indicador C3 — Gestantes e Puérperas": ("C3", processar_c3, "INDICADOR C3 — SAÚDE DA GESTANTE E PUÉRPERA", "Score_C3_%"),
    "Indicador C4 — Cuidado da Pessoa com Diabetes": ("C4", processar_c4, "INDICADOR C4 — CUIDADO DA PESSOA COM DIABETES", "Score_C4_%"),
    "Indicador C5 — Cuidado da Pessoa com Hipertensão": ("C5", processar_c5, "INDICADOR C5 — CUIDADO DA PESSOA COM HIPERTENSÃO", "Score_C5_%"),
    "Indicador C6 — Cuidado da Pessoa Idosa": ("C6", processar_c6, "INDICADOR C6 — CUIDADO DA PESSOA IDOSA", "Score_C6_%"),
    "Indicador C7 — Prevenção do Câncer e Saúde da Mulher": ("C7", processar_c7, "INDICADOR C7 — SAÚDE DA MULHER E PREVENÇÃO DO CÂNCER", "Score_C7_%")
}

opção_selecionada = st.selectbox(
    "Selecione o Indicador para Auditoria:",
    list(mapa_indicadores.keys())
)

uploaded_file = st.file_uploader("Anexe o relatório CSV exportado do e-SUS PEC:", type=["csv"])

if uploaded_file and st.button("🚀 Processar Dashboard Agora"):
    with st.spinner("Lendo relatório e aplicando regras oficiais da Nota Metodológica..."):
        bytes_data = uploaded_file.getvalue()
        df_bruto, data_ref = carregar_e_limpar_csv(bytes_data)
        
        sigla, func_proc, titulo_doc, col_score = mapa_indicadores[opção_selecionada]
        
        df_micro, df_nom = func_proc(df_bruto, data_ref)
        
        cols_status = [c for c in df_nom.columns if "Prática" in c or "status" in c]
        excel_bytes = gerar_excel_padrao(df_micro, df_nom, titulo_doc, data_ref, cols_status, col_score)
        
        st.success(f"✅ Dashboard do {sigla} gerado com sucesso!")
        st.download_button(
            label=f"📥 Baixar Dashboard_Indicador_{sigla}.xlsx",
            data=excel_bytes,
            file_name=f"Dashboard_Indicador_{sigla}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
