import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import io
import re
from datetime import datetime

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
    
    def clean_val(x):
        if pd.notnull(x):
            s = str(x).strip().replace('"', '')
            return s if s not in ['', '-', 'nan', 'NaN', 'None'] else None
        return None

    try:
        df = df.map(clean_val)
    except AttributeError:
        df = df.applymap(clean_val)

    return df, data_ref

def montar_endereco(row):
    rua = str(row.get('Rua', '') or '').strip()
    num = str(row.get('Número', '') or '').strip()
    comp = str(row.get('Complemento', '') or '').strip()
    partes = [p for p in [rua, num, comp] if p and p not in ['None', 'nan', '', '-']]
    return ", ".join(partes) if partes else "Endereço não informado"

def buscar_coluna_flexivel(df, termos):
    for col in df.columns:
        col_norm = col.lower().replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u').replace('ã','a').replace('ç','c')
        for t in termos:
            if t in col_norm:
                return col
    return None

# --- GERADOR DE EXCEL PADRONIZADO ---

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

    for row in ws1.iter_rows(min_row=5, max_row=ws1.max_row, min_col=2, max_col=len(headers_micro)):
        for cell in row:
            if cell.column >= 3:
                cell.number_format = '0.0"%"'
            else:
                cell.number_format = '#,##0'
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
                elif "Fora da faixa" in val or "Aguardando" in val or "Não se aplica" in val or "acompanhamento" in val:
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

# --- MÓDULO INDICADOR C3 (GESTANTES E PUÉRPERAS) ---

def processar_c3(df, data_ref):
    df['Endereço'] = df.apply(montar_endereco, axis=1)
    
    col_ig_sem = buscar_coluna_flexivel(df, ['ig (dum) (semanas)', 'ig semanas'])
    col_ig_dias = buscar_coluna_flexivel(df, ['ig (dum) (dias)', 'ig dias'])
    
    df['IG_sem_num'] = pd.to_numeric(df[col_ig_sem], errors='coerce').fillna(0).astype(int) if col_ig_sem else 0
    df['IG_dias_num'] = pd.to_numeric(df[col_ig_dias], errors='coerce').fillna(0).astype(int) if col_ig_dias else 0
    df['IG_total_dias'] = df['IG_sem_num'] * 7 + df['IG_dias_num']

    # Filtro da população ativa (0 a 336 dias / ≤48 semanas)
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

        # Meta esperada por IG
        if sem <= 28: esperadas = max(1, sem // 4)
        elif sem <= 36: esperadas = 7 + max(0, (sem - 28) // 2)
        else: esperadas = 11 + max(0, sem - 36)
        meta_7 = min(esperadas, 7)
        meta_3 = min(max(1, sem // 10), 3)

        # Prática A (1ª Consulta <= 12 sem)
        q_a = get_num(['ate 12 semanas', '12 semanas'])
        p_a = "Atendida" if q_a >= 1 else "Não atendida"

        # Prática B (Consultas Pré-natal)
        q_b = min(get_num(['quantidade de atendimentos no pre-natal']), 7)
        p_b = f"Atendida ({q_b}/{meta_7})" if q_b >= meta_7 else f"Não atendida ({q_b}/{meta_7})"

        # Prática C (Pressão Arterial)
        q_c = min(get_num(['quantidade de medicoes de pressao arterial']), 7)
        p_c = f"Atendida ({q_c}/{meta_7})" if q_c >= meta_7 else f"Não atendida ({q_c}/{meta_7})"

        # Prática D (Peso e Altura)
        q_d = min(get_num(['quantidade de medicoes simultaneas de peso e altura']), 7)
        p_d = f"Atendida ({q_d}/{meta_7})" if q_d >= meta_7 else f"Não atendida ({q_d}/{meta_7})"

        # Prática E (Visitas ACS Pré-natal)
        q_e = min(get_num(['quantidade de visitas domiciliares no pre-natal']), 3)
        p_e = f"Atendida ({q_e}/{meta_3})" if q_e >= meta_3 else f"Não atendida ({q_e}/{meta_3})"

        # Prática F (dTpa >= 20 sem)
        col_dtpa = buscar_coluna_flexivel(df_filtered, ['dtpa'])
        dtpa_v = str(row.get(col_dtpa, '') or '').strip() if col_dtpa else ''
        p_f = "Atendida (1/1)" if dtpa_v and dtpa_v not in ['-', 'None', 'nan', ''] else ("Aguardando idade" if sem < 20 else "Não atendida (0/1)")

        # Prática G (Exames 1ºT)
        col_hiv1 = buscar_coluna_flexivel(df_filtered, ['hiv no primeiro'])
        col_sif1 = buscar_coluna_flexivel(df_filtered, ['sifilis no primeiro'])
        col_hepb1 = buscar_coluna_flexivel(df_filtered, ['hepatite b no primeiro'])
        col_hepc1 = buscar_coluna_flexivel(df_filtered, ['hepatite c no primeiro'])
        
        ex_1t = [row.get(c) for c in [col_hiv1, col_sif1, col_hepb1, col_hepc1] if c]
        all_1t_sim = len(ex_1t) == 4 and all([str(x or '').upper() == 'SIM' for x in ex_1t])
        p_g = "Atendida (1/1)" if all_1t_sim else ("Aguardando idade" if sem <= 13 else "Não atendida (0/1)")

        # Prática H (Exames 3ºT)
        col_hiv3 = buscar_coluna_flexivel(df_filtered, ['hiv no terceiro'])
        col_sif3 = buscar_coluna_flexivel(df_filtered, ['sifilis no terceiro'])
        
        ex_3t = [row.get(c) for c in [col_hiv3, col_sif3] if c]
        all_3t_sim = len(ex_3t) == 2 and all([str(x or '').upper() == 'SIM' for x in ex_3t])
        p_h = "Atendida (1/1)" if all_3t_sim else ("Aguardando idade" if sem < 28 else "Não atendida (0/1)")

        # Prática I (Consulta Puerpério)
        q_i = get_num(['atendimentos no puerperio'])
        p_i = ("Atendida (1/1)" if q_i > 0 else "Não atendida (0/1)") if is_puerpera else "Não se aplica"

        # Prática J (Visita Puerpério)
        q_j = get_num(['visitas domiciliares no puerperio'])
        p_j = ("Atendida (1/1)" if q_j > 0 else "Não atendida (0/1)") if is_puerpera else "Não se aplica"

        # Prática K (Odonto)
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

# --- INTERFACE GRÁFICA PRINCIPAL ---

mapa_indicadores = {
    "Indicador C3 — Gestantes e Puérperas": ("C3", processar_c3, "INDICADOR C3 — SAÚDE DA GESTANTE E PUÉRPERA", "Score_C3_%")
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
