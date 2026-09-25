import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import io
import re
from datetime import datetime

st.set_page_config(page_title="Auditoria Automática e-SUS APS", layout="wide")

st.title("🏥 Sistema de Auditoria de Indicadores e-SUS APS")
st.subheader("Processamento Automático e Determinístico (Sem consumo de tokens)")

# --- MÓDULOS DE PROCESSAMENTO E-SUS APS ---

def extrair_data_referencia(conteudo_bytes):
    """Extrai a data 'Gerado em' do cabeçalho do e-SUS PEC ou assume a data atual."""
    try:
        texto = conteudo_bytes.decode('latin1', errors='ignore')
        match = re.search(r'Gerado em;?\s*(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE)
        if match:
            return datetime.strptime(match.group(1), '%d/%m/%Y')
    except Exception:
        pass
    return datetime.now()

def carregar_e_limpar_csv(file_bytes):
    """Carrega o CSV descartando metadados superiores e localizando a linha do cabeçalho 'Nome'."""
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
    
    # Limpeza básica de colunas e strings (Compatível com Pandas >= 2.1.0 usando .map em vez de .applymap)
    df.columns = [str(c).strip().replace('"', '') for c in df.columns]
    
    # Aplica a limpeza célula por célula usando .map()
    df = df.map(lambda x: x.strip().replace('"', '') if isinstance(x, str) else x)
    df = df.replace({'': None, '-': None, 'nan': None, 'NaN': None})
    
    return df, data_ref

# --- LÓGICA DO INDICADOR C5 (HIPERTENSÃO) ---
def processar_indicador_c5(df, data_ref):
    # Unificação de Endereço
    def montar_endereco(row):
        rua = str(row.get('Rua', '') or '').strip()
        num = str(row.get('Número', '') or '').strip()
        comp = str(row.get('Complemento', '') or '').strip()
        partes = [p for p in [rua, num, comp] if p and p not in ['None', 'nan', '']]
        return ", ".join(partes) if partes else "Endereço não informado"
    
    df['Endereço'] = df.apply(montar_endereco, axis=1)
    
    # Avaliação das Práticas
    def avaliar_c5(row):
        # Prática A (Consulta 6m)
        d_cons = row.get('Data da última consulta')
        status_a = "Não atendida (0/1)"
        if d_cons:
            try:
                dt = datetime.strptime(str(d_cons)[:10], '%d/%m/%Y')
                if (data_ref - dt).days <= 183:
                    status_a = "Atendida (1/1)"
            except: pass

        # Prática B (PA 6m)
        d_pa = row.get('Data da última medição de pressão arterial')
        status_b = "Não atendida (0/1)"
        if d_pa:
            try:
                dt = datetime.strptime(str(d_pa)[:10], '%d/%m/%Y')
                if (data_ref - dt).days <= 183:
                    status_b = "Atendida (1/1)"
            except: pass

        # Prática C (2 Visitas ACS 12m com intervalo >=30 dias)
        visitas_txt = str(row.get('Últimas visitas domiciliares', '') or '')
        datas_v = re.findall(r'\d{2}/\d{2}/\d{4}', visitas_txt)
        status_c = "Não atendida (0/2)"
        
        dts_validas = []
        for d in datas_v:
            try:
                dt = datetime.strptime(d, '%d/%m/%Y')
                if 0 <= (data_ref - dt).days <= 365:
                    dts_validas.append(dt)
            except: pass
            
        dts_validas = sorted(list(set(dts_validas)))
        if len(dts_validas) >= 2:
            for i in range(len(dts_validas)-1):
                if (dts_validas[i+1] - dts_validas[i]).days >= 30:
                    status_c = "Atendida (2/2)"
                    break

        # Prática D (Peso e Altura 12m)
        d_ant = row.get('Data da última medição de peso e altura')
        status_d = "Não atendida (0/1)"
        if d_ant:
            try:
                dt = datetime.strptime(str(d_ant)[:10], '%d/%m/%Y')
                if (data_ref - dt).days <= 365:
                    status_d = "Atendida (1/1)"
            except: pass

        pts = sum([1 for s in [status_a, status_b, status_c, status_d] if "Atendida" in s])
        score = (pts / 4.0) * 100

        return pd.Series([
            status_a, status_b, status_c, status_d, score
        ])

    df[['Prática A — consulta últimos 6 meses — status',
        'Prática B — pressão últimos 6 meses — status',
        'Prática C — 2 visitas em 12 meses — status',
        'Prática D — peso e altura últimos 12 meses — status',
        'Score_C5_%']] = df.apply(avaliar_c5, axis=1)

    # Organização das Colunas Nominais
    cols_nominais = [
        'Microárea', 'Nome', 'CPF', 'Telefone celular', 'Endereço',
        'Prática A — consulta últimos 6 meses — status',
        'Prática B — pressão últimos 6 meses — status',
        'Prática C — 2 visitas em 12 meses — status',
        'Prática D — peso e altura últimos 12 meses — status',
        'Score_C5_%'
    ]
    
    for c in cols_nominais:
        if c not in df.columns: df[c] = "-"
        
    df_nominal = df[cols_nominais].copy()
    
    # Agrupamento por Microárea
    df_micro = df.groupby('Microárea').agg(
        Pessoas_hipertensao_ativas=('Nome', 'count'),
        Score_medio_C5=('Score_C5_%', 'mean'),
        Pct_100_praticas=('Score_C5_%', lambda x: (x == 100).mean() * 100),
        Pct_Pratica_A=('Prática A — consulta últimos 6 meses — status', lambda x: (x.astype(str).str.contains("Atendida")).mean() * 100),
        Pct_Pratica_B=('Prática B — pressão últimos 6 meses — status', lambda x: (x.astype(str).str.contains("Atendida")).mean() * 100),
        Pct_Pratica_C=('Prática C — 2 visitas em 12 meses — status', lambda x: (x.astype(str).str.contains("Atendida")).mean() * 100),
        Pct_Pratica_D=('Prática D — peso e altura últimos 12 meses — status', lambda x: (x.astype(str).str.contains("Atendida")).mean() * 100)
    ).reset_index()

    return df_micro, df_nominal

# --- GERAÇÃO DO ARQUIVO EXCEL COM FORMATAÇÃO ---
def gerar_excel_c5(df_micro, df_nominal, data_ref):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    # Estilos
    blue_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    red_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

    # Aba 1: Dashboard_Microarea
    ws1 = wb.active
    ws1.title = "Dashboard_Microarea"
    ws1.append(["INDICADOR C5 — CUIDADO DA PESSOA COM HIPERTENSÃO"])
    ws1.append([f"Data de Referência: {data_ref.strftime('%d/%m/%Y')} | Relatório Oficial e-SUS APS"])
    ws1.append([])
    
    ws1.cell(row=1, column=1).font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    ws1.cell(row=2, column=1).font = Font(name="Calibri", size=10, italic=True)

    # Tabela Microárea
    headers_micro = ['Microárea', 'Hipertensos Ativos', 'Score Médio C5 (%)', '% com 100% Práticas',
                     '% Atendida Prática A', '% Atendida Prática B', '% Atendida Prática C', '% Atendida Prática D']
    ws1.append(headers_micro)
    
    for col_num in range(1, len(headers_micro) + 1):
        cell = ws1.cell(row=4, column=col_num)
        cell.fill = blue_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r in dataframe_to_rows(df_micro, index=False, header=False):
        ws1.append(r)

    # Formatação de Percentual
    for row in ws1.iter_rows(min_row=5, max_row=ws1.max_row, min_col=3, max_col=8):
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
        for c_idx in range(1, len(headers_nom) + 1):
            cell = ws2.cell(row=r_idx, column=c_idx)
            cell.border = border_thin
            val = str(cell.value)
            
            # Formatação condicional de status
            if "Atendida" in val:
                cell.fill = green_fill
            elif "Não atendida" in val:
                cell.fill = red_fill
            elif c_idx == len(headers_nom): # Coluna Score
                try:
                    cell.value = float(cell.value)
                    cell.number_format = '0.0"%"'
                except: pass

    # Ajuste automático de largura de colunas
    for ws in [ws1, ws2]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output)
    return output.getvalue()


# --- INTERFACE GRÁFICA STREAMLIT ---

indicador = st.selectbox(
    "Selecione o Indicador para Auditoria:",
    ["Indicador C5 - Hipertensão (Disponível Automático)", "Indicador C2, C3, C4, C6, C7 (Em Atualização)"]
)

uploaded_file = st.file_uploader("Anexe o arquivo CSV do e-SUS PEC:", type=["csv"])

if uploaded_file and st.button("🚀 Processar Dashboard Agora"):
    with st.spinner("Lendo relatório e aplicando regras oficiais do e-SUS APS..."):
        bytes_data = uploaded_file.getvalue()
        df_bruto, data_ref = carregar_e_limpar_csv(bytes_data)
        
        if "C5" in indicador or "Hipertensão" in indicador:
            df_micro, df_nominal = processar_indicador_c5(df_bruto, data_ref)
            excel_bytes = gerar_excel_c5(df_micro, df_nominal, data_ref)
            
            st.success("✅ Dashboard gerado com sucesso!")
            st.download_button(
                label="📥 Baixar Dashboard_Indicador_C5_Hipertensao.xlsx",
                data=excel_bytes,
                file_name="Dashboard_Indicador_C5_Hipertensao.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Selecione o Indicador C5 para demonstrar o processamento instantâneo.")
