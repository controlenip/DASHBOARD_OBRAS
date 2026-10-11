"""NIP — Dashboard HTML original hospedado por Streamlit.
Consulta livre. Publicação da aba NOTAS protegida por ADMIN_PASSWORD.
Planilha de referência: BASE_LEVANTAMENTO_ATUALIZADA.xlsx.
Persistência compartilhada: commits em um repositório GitHub configurado.
"""
from __future__ import annotations
import base64
import hashlib
import hmac
import io
import json
import os
from pathlib import Path
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title='NIP | Dashboard de Obras', page_icon='🗺️', layout='wide', initial_sidebar_state='collapsed')
ROOT = Path(__file__).resolve().parent
EXCEL_NAME = 'BASE_LEVANTAMENTO_ATUALIZADA.xlsx'
TEMPLATE_NAME = 'nip_dashboard_template.html'
REQUIRED = ['ID SISCO','STATUS LIST','STATUS SAP','STATUS SISCO','REGIONAL','MUNICIPIO','LEVANTADOR','DATA DESPACHO']

def setting(name, default=''):
    try:
        return st.secrets.get(name, os.getenv(name, default))
    except Exception:
        return os.getenv(name, default)

GITHUB_REPO = str(setting('GITHUB_REPO', 'controlenip/DASHBOARD_OBRAS')).strip()
GITHUB_BRANCH = str(setting('GITHUB_BRANCH', 'main')).strip()
GITHUB_TOKEN = str(setting('GITHUB_TOKEN')).strip()
ADMIN_PASSWORD = str(setting('ADMIN_PASSWORD')).strip()

st.markdown('''<style>
.stApp {background:#edf2f6;}
[data-testid="stMainBlockContainer"], .block-container {padding:0!important;max-width:100%!important;}
[data-testid="stMain"] {padding:0!important;}
[data-testid="stHeader"] {background:#123653;height:0;}
[data-testid="stToolbar"] {visibility:hidden;}
[data-testid="stSidebar"] {background:white;}
[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {gap:0.8rem;}
[data-testid="stSidebar"] h2 {color:#123653;}
</style>''', unsafe_allow_html=True)


def headers(auth=False):
    h = {'Accept':'application/vnd.github+json', 'X-GitHub-Api-Version':'2022-11-28',
         'User-Agent':'NIP-dashboard'}
    if auth and GITHUB_TOKEN:
        h['Authorization'] = 'Bearer '+GITHUB_TOKEN
    return h

def api(path):
    return 'https://api.github.com/repos/'+GITHUB_REPO+'/'+path


def get_remote_file(ref=None):
    if '/' not in GITHUB_REPO:
        raise ValueError('GITHUB_REPO deve ter o formato usuario/repositorio.')
    h=headers(bool(GITHUB_TOKEN))
    r=requests.get(api('contents/'+EXCEL_NAME),params={'ref':ref or GITHUB_BRANCH},headers=h,timeout=35)
    if r.status_code==404:return None
    r.raise_for_status()
    item=r.json()
    if item.get('content') and item.get('encoding')=='base64':
        return base64.b64decode(item['content'])
    b=requests.get(api('git/blobs/'+item['sha']),headers=h,timeout=90)
    b.raise_for_status()
    return base64.b64decode(b.json()['content'])

@st.cache_data(ttl=90,show_spinner=False)
def get_shared_base():
    try:
        data=get_remote_file()
        if data:return data,'Base compartilhada · GitHub'
    except Exception:
        pass
    local=ROOT/EXCEL_NAME
    if local.exists():return local.read_bytes(),'Cópia inicial do repositório — pode estar desatualizada'
    raise FileNotFoundError('A planilha não foi encontrada. Adicione '+EXCEL_NAME+' ao repositório.')

@st.cache_data(show_spinner=False,max_entries=4)
def parse_notas(blob:bytes):
    df=pd.read_excel(io.BytesIO(blob),sheet_name='NOTAS',engine='openpyxl',dtype=object)
    df.columns=[str(c).strip() for c in df.columns]
    missing=[c for c in REQUIRED if c not in df.columns]
    if missing:raise ValueError('Aba NOTAS sem as colunas: '+', '.join(missing))
    # Match exactly the 22 columns and date formats of the original HTML.
    original_fields = ['ID SISCO','PROTOCOLO','TIPO NOTA','STATUS SAP','STATUS SISCO',
        'STATUS LIST','CONTRATO','REGIONAL','MUNICIPIO','LEVANTADOR','DATA DESPACHO',
        'NOME','LATITUDE','LONGITUDE','DISTANCIA BT','DISTANCIA MT',
        'POSTES PREVISTOS TOTAL','ENDEREÇO','LOCALIDADE','INFORMAÇÕES',
        'ÁREA DE ATUAÇÃO (LEVANTADOR)','INCLUIR_DASH']
    for col in original_fields:
        if col not in df.columns: df[col]=None
    df=df[original_fields].copy()
    df['DATA DESPACHO']=pd.to_datetime(df['DATA DESPACHO'],errors='coerce').dt.strftime('%Y-%m-%d')
    for col in ['LATITUDE','LONGITUDE']:
        df[col]=pd.to_numeric(df[col].astype(str).str.replace(',','.',regex=False),errors='coerce')
    payload=df.to_json(orient='records',force_ascii=False,date_format='iso')
    return payload,len(df)


def publish(blob:bytes, actor:str):
    if not GITHUB_TOKEN:
        raise RuntimeError('Configure GITHUB_TOKEN nos Secrets com permissão Contents: Read and write.')
    if '/' not in GITHUB_REPO:
        raise RuntimeError('GITHUB_REPO inválido.')
    _,count=parse_notas(blob)  # Validate BEFORE publication
    link=api('contents/'+EXCEL_NAME)
    current=requests.get(link,headers=headers(True),params={'ref':GITHUB_BRANCH},timeout=45)
    if current.status_code not in (200,404):current.raise_for_status()
    data={'message':f'Atualização da aba NOTAS — NIP ({actor})',
          'branch':GITHUB_BRANCH,'content':base64.b64encode(blob).decode('ascii')}
    if current.status_code==200:data['sha']=current.json()['sha']
    result=requests.put(link,headers=headers(True),json=data,timeout=180)
    if not result.ok:
        raise RuntimeError(f'Falha na publicação (HTTP {result.status_code}). Confirme permissão de escrita e branch.')
    get_shared_base.clear()
    parse_notas.clear()
    return count


def admin_sidebar():
    with st.sidebar:
        logo=ROOT/'logo_nip.png'
        if logo.exists():st.image(str(logo),width=200)
        st.markdown('### Administração NIP')
        st.caption('A visualização do dashboard é pública. A senha é solicitada somente para publicar uma nova base.')
        if not st.session_state.get('nip_admin'):
            if not ADMIN_PASSWORD:
                st.warning('Para habilitar publicação, configure ADMIN_PASSWORD nos Secrets.')
                return
            with st.form('admin_access',clear_on_submit=True):
                actor=st.text_input('Identificação do responsável')
                password=st.text_input('Senha de administrador',type='password')
                ok=st.form_submit_button('Desbloquear atualização',use_container_width=True)
            if ok:
                if hmac.compare_digest(password,ADMIN_PASSWORD):
                    st.session_state['nip_admin']=True
                    st.session_state['nip_actor']=actor.strip() or 'Administrador'
                    st.rerun()
                else:st.error('Senha incorreta.')
            return
        st.success('Publicação autorizada nesta sessão')
        if st.button('Encerrar acesso administrativo',use_container_width=True):
            st.session_state.pop('nip_admin',None)
            st.rerun()
        file=st.file_uploader('Enviar Excel atualizado (aba NOTAS)',type=['xlsx'],key='nip_upload')
        if file is not None:
            try:
                payload,n=parse_notas(file.getvalue())
                st.caption(f'Arquivo validado · {n:,} registros na aba NOTAS'.replace(',','.'))
            except Exception as e:
                st.error('Arquivo inválido: '+str(e))
                file=None
        if file is not None and st.button('Publicar para todos',type='primary',use_container_width=True):
            try:
                with st.spinner('Publicando versão compartilhada...'):
                    n=publish(file.getvalue(),st.session_state.get('nip_actor','Administrador'))
                st.success(f'{n:,} obras publicadas. Todos verão a nova versão ao atualizar a página.'.replace(',','.'))
                st.rerun()
            except Exception as e:st.error(str(e))
        if not GITHUB_TOKEN:st.info('Para persistir atualizações para todos, configure GITHUB_TOKEN nos Secrets.')


def main():
    admin_sidebar()
    try:
        blob,source=get_shared_base()
        payload,count=parse_notas(blob)
        template=(ROOT/TEMPLATE_NAME).read_text(encoding='utf-8')
        if template.count('__NIP_DATA_PAYLOAD__')!=1:
            raise ValueError('Template visual inválido: marcador da base ausente ou duplicado.')
        # Inject the shared Excel directly into the ORIGINAL HTML/JS implementation.
        # Escape </script> and HTML specials so spreadsheet cell text cannot break out.
        payload=(payload.replace('<','\\u003c').replace('>','\\u003e')
                 .replace('&','\\u0026').replace('\u2028','\\u2028').replace('\u2029','\\u2029'))
        html=template.replace('__NIP_DATA_PAYLOAD__',payload)
        html=html.replace('Base incorporada •',source+' •')
        st.session_state['nip_loaded_count']=count
        components.html(html,height=1100,scrolling=True)
        st.caption(f'Fonte: {source} · {count:,} registros · aba NOTAS · revisão automática em até 90 segundos após recarregar a página.'.replace(',','.'))
    except Exception as e:
        st.error('Não foi possível carregar o Dashboard NIP.')
        st.warning(str(e))
        st.info('Confira se app_nip.py, nip_dashboard_template.html, logo_nip.png e BASE_LEVANTAMENTO_ATUALIZADA.xlsx estão na raiz do repositório.')

if __name__=='__main__':main()
