"""Aplicação NIP multipáginas. Inicie com streamlit run app_nip.py."""
import streamlit as st

st.set_page_config(page_title="NIP | Dashboard Demanda", page_icon="🗺️", layout="wide", initial_sidebar_state="expanded")

navigation = st.navigation([
    st.Page("pages/01_visao_geral.py", title="Visão Geral", icon="🗺️", default=True, url_path="visao-geral"),
    st.Page("pages/02_levantadores.py", title="Levantadores", icon="👷", url_path="levantadores"),
    st.Page("pages/03_prazos_prioridades.py", title="Prazos e Prioridades", icon="⏱️", url_path="prazos-prioridades"),
    st.Page("pages/04_relatorios.py", title="Relatórios", icon="📊", url_path="relatorios"),
    st.Page("pages/05_qualidade_levantamento.py", title="Qualidade Levantamento", icon="✅", url_path="qualidade"),
    st.Page("pages/06_produtividade.py", title="Produtividade", icon="📈", url_path="produtividade"),
    st.Page("pages/07_pre_analise.py", title="Pré-Análise", icon="📋", url_path="pre-analise"),
    st.Page("pages/08_faq_duvidas.py", title="FAQ / Dúvidas", icon="❓", url_path="faq"),
], position="sidebar", expanded=True)
navigation.run()
