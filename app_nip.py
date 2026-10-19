"""Dashboard NIP: layout original, com páginas Python independentes.
O menu visual permanece no HTML, sem sidebar Streamlit.
"""
import streamlit as st
st.set_page_config(page_title='NIP | Dashboard Demanda', page_icon='🗺️', layout='wide', initial_sidebar_state='collapsed')
nav = st.navigation([
    st.Page('pages/01_visao_geral.py', title='Visão Geral', default=True, url_path='visao-geral'),
    st.Page('pages/02_levantadores.py', title='Levantadores', url_path='levantadores'),
    st.Page('pages/03_prazos_prioridades.py', title='Prazos e Prioridades', url_path='prazos-prioridades'),
    st.Page('pages/04_relatorios.py', title='Relatórios', url_path='relatorios'),
    st.Page('pages/05_qualidade_levantamento.py', title='Qualidade Levantamento', url_path='qualidade'),
    st.Page('pages/06_produtividade.py', title='Produtividade', url_path='produtividade'),
    st.Page('pages/07_pre_analise.py', title='Pré-Análise', url_path='pre-analise'),
    st.Page('pages/08_faq_duvidas.py', title='FAQ / Dúvidas', url_path='faq'),
    st.Page('pages/09_croquis.py', title='Croquis', url_path='croquis'),
    st.Page('pages/10_administracao.py', title='Administração NIP', url_path='administracao'),
], position='hidden')
nav.run()
