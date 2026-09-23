import streamlit as st
import google.generativeai as genai
import os
import re
import tempfile
import sys
import traceback

st.set_page_config(page_title="Auditoria e-SUS APS", layout="wide")

# Dicionário de Prompts Técnicos de Auditoria
PROMPTS = {
    "Indicador C2 - Puericultura": """Você é Engenheiro de Dados em Saúde especialista em e-SUS APS e indicadores da Atenção Primária do Ministério da Saúde. A partir do CSV anexado ("acompanhamento-condicao-saude.csv") e da Nota Metodológica C2 oficial, processe os dados em Python e gere a planilha "Dashboard_Indicador_C2_Microareas.xlsx".
1. Leitura e Limpeza: Encoding latin1, separador ;. Identificar cabeçalho (iniciado por "Nome;") e descartar metadados. Extrair data em "Gerado em" para referência.
2. Filtro: Crianças ativas menores de 3 anos (< 1.096 dias de vida).
3. Práticas (A-E): A(1ª Consulta <=30d); B(9 consultas proporcionais aos marcos até 24m); C(9 antropometrias simultâneas proporcionais); D(Visitas ACS: 1ª<=30d, 2ª<=180d); E(Vacinas: Penta D3, Polio D3, Pneumo D2, SCR D2 adequadas à idade).
4. Score C2 (%): 20 pontos por prática. Desconsiderar "Aguardando idade" do denominador.
5. Abas: Dashboard_Microarea (Cards, tabela consolidada por microárea) e Auditoria_Nominal (Filtros, regras aplicadas, dados nominais, remoção de colunas brutas desnecessárias).""",
    
    "Indicador C3 - Gestante e Puérpera": """Você é Engenheiro de Dados em Saúde especialista em e-SUS APS e indicadores do Ministério da Saúde. Processar o arquivo CSV anexado de acordo com a Nota Metodológica Oficial do Indicador C3 e gerar o arquivo Excel "Dashboard_Indicador_C3_Gestantes_Puerperas.xlsx".
1. Leitura e Limpeza: Encoding latin1, separador ;. Localizar cabeçalho "Nome", extrair data "Gerado em".
2. Filtro: DUM válida. Gestantes ativas (IG 0-294 dias). Puérperas ativas (IG 295-336 dias).
3. 11 Práticas (A-K): A(1ª Cons<=12sem); B,C,D(7 Consultas, PA, Peso/Altura proporcionais); E(3 Visitas ACS proporcionais); F(dTpa >=20 sem); G(Exames 1ºT <=13w6d); H(Exames 3ºT >=28w); I,J(Cons e Visita Puerpério); K(Saúde Bucal).
4. IG Atual: X semanas e Y dias (Mês/Classificação).
5. Score C3: Prática A = 10 pts; B a K = 9 pts (Total 100). Excluir não avaliáveis do denominador.
6. Abas: Dashboard_Microarea (Cards, tabela sintética) e Auditoria_Nominal (Dados nominais limpos, score individual).""",
    
    "Indicador C4 - Diabetes": """Atue como Engenheiro de Dados em Saúde e especialista no e-SUS APS. Processe o CSV "Acompanhamento - Diabetes" e gere "Dashboard_Indicador_C4_Diabetes.xlsx".
1. Leitura e Limpeza: Encoding latin1, separador ;. Cabeçalho em "Nome", extrair data "Gerado em".
2. População: Totalidade do relatório (já filtrado pelo PEC).
3. 6 Práticas (A-F): A(Consulta 6m); B(PA 6m); C(2 Visitas ACS 12m com >=30 dias de intervalo); D(Peso/Altura simultâneo 12m); E(HbA1c solicitada/avaliada 12m); F(Pé diabético 12m).
4. Score C4: 1 ponto por prática (máx 6). Score em % (atendidas/6).
5. Abas: Dashboard_Microarea (KPIs, tabela consolidada) e Auditoria_Nominal (Nome, CPF, Celular, Endereço unificado, status de cada prática, Score C4).""",
    
    "Indicador C5 - Hipertensão": """Atue como Engenheiro de Dados em Saúde e especialista no e-SUS APS. Processe o CSV "Acompanhamento - Hipertensão" e gere "Dashboard_Indicador_C5_Hipertensao.xlsx".
1. Leitura e Limpeza: Encoding latin1, separador ;. Cabeçalho em "Nome", extrair data "Gerado em".
2. População: Totalidade do relatório.
3. 4 Práticas (A-D): A(Consulta 6m); B(PA 6m); C(2 Visitas ACS 12m com >=30 dias de intervalo); D(Peso/Altura simultâneo 12m).
4. Score C5: 25 pontos por prática. Score em % (atendidas/4).
5. Abas: Dashboard_Microarea (KPIs, tabela consolidada) e Auditoria_Nominal (Endereço unificado, remoção de lixo de dados, status por prática, Score C5).""",
    
    "Indicador C6 - Pessoa Idosa": """Atue como Engenheiro de Dados em Saúde e especialista no e-SUS APS. Processe o CSV "Acompanhamento - Pessoa Idosa" e gere "Dashboard_Indicador_C6_Pessoa_Idosa.xlsx".
1. Leitura e Limpeza: Encoding latin1, separador ;. Cabeçalho em "Nome", extrair data "Gerado em".
2. População: Idade >= 60 anos, janela global de 12 meses.
3. 4 Práticas (A-D): A(Consulta Med/Enf 12m); B(>=2 Antropometrias simultâneas 12m - usar coluna de contagem); C(2 Visitas ACS 12m com >=30 dias de intervalo); D(Vacina Influenza 12m).
4. Score C6: 25 pontos por prática. Score em %.
5. Abas: Dashboard_Microarea e Auditoria_Nominal (limpa de variáveis excedentes, formatação condicional).""",
    
    "Indicador C7 - Saúde da Mulher": """Atue como Engenheiro de Dados em Saúde e especialista no e-SUS APS. Processe o CSV "Acompanhamento - Saúde da Mulher" e gere "Dashboard_Indicador_C7_Prevencao_Cancer.xlsx".
1. Leitura e Limpeza: Encoding latin1, separador ;. Cabeçalho "Nome", extrair data "Gerado em". Calcular idade até a data de referência.
2. 4 Práticas e Faixas: A(Colo Útero 25-64a, janela 36m); B(HPV 9-14a); C(Saúde Sexual/Reprodutiva 14-69a, janela 12m); D(Mama 50-69a, janela 24m). Atribuir 'Fora da faixa etária' se não aplicável.
3. Score C7: Calculado estritamente sobre as práticas aplicáveis à idade da pessoa.
4. Abas: Dashboard_Microarea (Denominadores ajustados por elegibilidade de cada prática) e Auditoria_Nominal (Idade calculada, exclusão de dados desnecessários, cores por status)."""
}

st.title("Orquestrador Clínico e-SUS APS via Gemini")

with st.sidebar:
    st.header("Configurações")
    api_key = st.text_input("Chave API (Google Gemini)", type="password")
    indicador_selecionado = st.selectbox("Selecione o Indicador para Auditoria", list(PROMPTS.keys()))

uploaded_file = st.file_uploader("Anexe o relatório CSV exportado do e-SUS PEC", type=["csv"])

if st.button("Processar e Gerar Dashboard") and uploaded_file and api_key:
    genai.configure(api_key=api_key)
    modelo = genai.GenerativeModel('gemini-1.5-pro')
    
    with st.spinner("Analisando premissas, gerando código e executando ETL..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = os.path.join(tmpdir, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            try:
                # Upload do arquivo para a API do Gemini
                gemini_file = genai.upload_file(file_path)
                
                # Instrução rigorosa para isolamento de código
                system_instruction = f"""
                Você atuará estritamente como um gerador de script Python. Baseado nas regras a seguir, crie um script que leia o arquivo '{uploaded_file.name}' e gere a planilha Excel de saída correspondente no mesmo diretório. 
                REGRAS OBRIGATÓRIAS:
                1. Utilize as bibliotecas 'pandas' e 'openpyxl'.
                2. O código não deve conter placeholders ou dados fictícios.
                3. Trate erros de encoding tentando 'latin1', 'utf-8' e 'cp1252'.
                4. Retorne EXCLUSIVAMENTE o bloco de código Python delimitado por ```python e ```, sem nenhuma explicação antes ou depois.
                
                REGRAS DE NEGÓCIO:
                {PROMPTS[indicador_selecionado]}
                """
                
                resposta = modelo.generate_content([gemini_file, system_instruction])
                
                # Extração do bloco de código
                match = re.search(r'```python\n(.*?)\n```', resposta.text, re.DOTALL)
                codigo_python = match.group(1) if match else resposta.text.replace('```python', '').replace('```', '')
                
                # Mudança de diretório para o sandbox e execução do código LLM
                cwd_original = os.getcwd()
                os.chdir(tmpdir)
                
                try:
                    exec(codigo_python, globals())
                except Exception as e:
                    st.error("Erro na execução do código gerado pelo modelo.")
                    st.code(traceback.format_exc())
                    st.expander("Ver Código Gerado").code(codigo_python, language='python')
                    os.chdir(cwd_original)
                    st.stop()
                
                os.chdir(cwd_original)
                
                # Identificação do arquivo Excel gerado
                arquivos_gerados = [f for f in os.listdir(tmpdir) if f.endswith('.xlsx')]
                
                if arquivos_gerados:
                    excel_path = os.path.join(tmpdir, arquivos_gerados[0])
                    with open(excel_path, "rb") as f:
                        st.success("Dashboard gerado com sucesso!")
                        st.download_button(
                            label="📥 Baixar Dashboard Excel",
                            data=f,
                            file_name=arquivos_gerados[0],
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.warning("O modelo executou o código, mas não salvou o arquivo .xlsx final.")
                    st.expander("Ver Código Gerado").code(codigo_python, language='python')
                    
            except Exception as e:
                st.error(f"Falha na comunicação com a API ou processamento: {str(e)}")
