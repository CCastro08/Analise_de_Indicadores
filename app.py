import streamlit as st

from indicadores import (
    c2_puericultura,
    c3_gestantes,
    c4_diabetes,
    c5_hipertensao,
    c6_idosos,
    c7_mulher
)

st.set_page_config(page_title="Auditoria e-SUS APS - SAPS/MS", layout="wide")

st.title("🏥 Sistema de Auditoria de Indicadores e-SUS APS (SAPS/MS)")
st.caption("Painel Modular de Auditoria Determinística — Módulos Isolados por Abas")

tab_c3, tab_c2, tab_c4, tab_c5, tab_c6, tab_c7 = st.tabs([
    "🤰 C3 — Gestantes e Puérperas",
    "👶 C2 — Puericultura",
    "🩸 C4 — Diabetes",
    "🫀 C5 — Hipertensão",
    "👴 C6 — Idosos",
    "🌸 C7 — Saúde da Mulher"
])

with tab_c3:
    c3_gestantes.render()

with tab_c2:
    c2_puericultura.render()

with tab_c4:
    c4_diabetes.render()

with tab_c5:
    c5_hipertensao.render()

with tab_c6:
    c6_idosos.render()

with tab_c7:
    c7_mulher.render()
