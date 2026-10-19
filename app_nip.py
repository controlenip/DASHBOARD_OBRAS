"""Dashboard NIP - inicializacao robusta de paginas e layout original."""
from pathlib import Path
import streamlit as st

st.set_page_config(
    page_title="NIP | Dashboard Demanda",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="collapsed",
    menu_items={"Get Help": None, "Report a bug": None, "About": None},
)

# Registre somente paginas fisicamente presentes no repositorio publicado.
# Uma pagina opcional ausente nao pode impedir o dashboard inteiro de abrir.
ROOT = Path(__file__).resolve().parent
PAGES = [
    ("01_visao_geral.py", "Visão Geral", "visao-geral"),
    ("02_levantadores.py", "Levantadores", "levantadores"),
    ("03_prazos_prioridades.py", "Prazos e Prioridades", "prazos-prioridades"),
    ("04_relatorios.py", "Relatórios", "relatorios"),
    ("05_qualidade_levantamento.py", "Qualidade Levantamento", "qualidade"),
    ("06_produtividade.py", "Produtividade", "produtividade"),
    ("07_pre_analise.py", "Pré-Análise", "pre-analise"),
    ("08_faq_duvidas.py", "FAQ / Dúvidas", "faq"),
    ("09_croquis.py", "Croquis", "croquis"),
    ("10_administracao.py", "Administração NIP", "administracao"),
]

available = [(name, title, path) for name, title, path in PAGES if (ROOT / "pages" / name).is_file()]
if not available:
    st.error("Nenhuma página foi encontrada na pasta pages/. Verifique se ela foi enviada ao GitHub com os arquivos .py.")
    st.stop()

# As abas horizontais permanecem no HTML original. Sem menu lateral Streamlit.
nav_pages = [
    st.Page(
        str(ROOT / "pages" / name),
        title=title,
        url_path=path,
        default=(index == 0),
    )
    for index, (name, title, path) in enumerate(available)
]
st.navigation(nav_pages, position="hidden").run()
