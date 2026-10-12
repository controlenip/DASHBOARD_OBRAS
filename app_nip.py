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
import zipfile
from collections import Counter
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
LIST_NAME = 'BASE_LIST.xlsx'
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
[data-testid="stHeader"] {height:0!important; min-height:0!important;}
/* Keep Streamlit's outer page fixed: all scrolling belongs to the dashboard iframe. */
html, body, [data-testid="stAppViewContainer"], [data-testid="stMain"],
[data-testid="stMainBlockContainer"], .block-container {
    height:100dvh!important;
    max-height:100dvh!important;
    overflow:hidden!important;
}
[data-testid="stMainBlockContainer"] > div,
[data-testid="stMainBlockContainer"] [data-testid="stVerticalBlock"],
[data-testid="stMainBlockContainer"] [data-testid="stElementContainer"],
[data-testid="stMainBlockContainer"] [data-testid="stCustomComponentV1"] {
    margin:0!important;
    padding:0!important;
    gap:0!important;
}
[data-testid="stCustomComponentV1"] iframe {
    display:block!important;
    width:100%!important;
    height:100dvh!important;
    border:0!important;
}

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


def get_remote_file(ref=None, filename=None):
    if '/' not in GITHUB_REPO:
        raise ValueError('GITHUB_REPO deve ter o formato usuario/repositorio.')
    h=headers(bool(GITHUB_TOKEN))
    r=requests.get(api('contents/'+(filename or EXCEL_NAME)),params={'ref':ref or GITHUB_BRANCH},headers=h,timeout=35)
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
        'STATUS LIST','CONTRATO','PRIORIDADE','SGO','Nº SGO','N° SGO','NÚMERO SGO','REGIONAL','MUNICIPIO','LEVANTADOR','DATA DESPACHO',
        'NOME','LATITUDE','LONGITUDE','DISTANCIA BT','DISTANCIA MT',
        'POSTES PREVISTOS TOTAL','ENDEREÇO','LOCALIDADE','INFORMAÇÕES',
        'ÁREA DE ATUAÇÃO (LEVANTADOR)','INCLUIR_DASH',
        'ORÇAMENTO MODULAR (NÃO ENVIAR A CAMPO)']
    for col in original_fields:
        if col not in df.columns: df[col]=None
    # Os demais campos permitidos da aba NOTAS também abastecem os filtros
    # e as colunas da tabela detalhada, sem alterar os campos existentes.
    hidden_table_fields = {
        'FASE','NOME','CONTRATO','INSTALAÇÃO','ENDEREÇO','LATITUDE','LONGITUDE',
        'INFORMAÇÕES','INFORMAÇÕES EXTRAS','INCLUIR_DASH','POSTES PREVISTOS TOTAL',
        'DIAS ABERTURA-DESPACHO','FILA OPERACIONAL','ETAPA OPERACIONAL',
        'DIAS EM LEVANTAMENTO','FAIXA AGING','SEM LEVANTADOR ATIVO',
        'PRIORIDADE FORA LIST','DIVERGENCIA SISCO-LIST','AÇÃO NECESSÁRIA',
        'SEVERIDADE','INCLUIR_DASH_SEM_DATA','PRONTO_PARA_LIST'}
    extra_fields = [c for c in df.columns if c not in original_fields and c not in hidden_table_fields]
    df=df[original_fields + extra_fields].copy()
    for date_col in ['DATA ABERTURA']:
        if date_col in df.columns:
            df[date_col]=pd.to_datetime(df[date_col],errors='coerce').dt.strftime('%Y-%m-%d')
    df['DATA DESPACHO']=pd.to_datetime(df['DATA DESPACHO'],errors='coerce').dt.strftime('%Y-%m-%d')
    for col in ['LATITUDE','LONGITUDE']:
        df[col]=pd.to_numeric(df[col].astype(str).str.replace(',','.',regex=False),errors='coerce')
    payload=df.to_json(orient='records',force_ascii=False,date_format='iso')
    return payload,len(df)


@st.cache_data(ttl=90,show_spinner=False)
def get_shared_list():
    """Independent, read-only source. Missing BASE_LIST never blocks the dashboard."""
    try:
        remote=get_remote_file(filename=LIST_NAME)
        if remote is not None:
            return remote, 'BASE_LIST.xlsx · GitHub'
    except Exception:
        pass
    local=ROOT/LIST_NAME
    if local.exists():
        return local.read_bytes(), 'BASE_LIST.xlsx · cópia local (verifique a versão no GitHub)'
    return None, 'BASE_LIST.xlsx não encontrada no repositório'

@st.cache_data(show_spinner=False,max_entries=4)
def parse_list_quality(blob:bytes):
    """E=status, O=levantador, T=analista de pre-analise, AF=observacoes."""
    raw=pd.read_excel(io.BytesIO(blob),sheet_name=0,header=0,engine='openpyxl',dtype=object)
    if raw.shape[1] < 32:
        raise ValueError('BASE_LIST precisa ter pelo menos 32 colunas (até AF).')
    def norm(v):
        import unicodedata
        return ''.join(c for c in unicodedata.normalize('NFKD',str(v or '')) if not unicodedata.combining(c)).upper().strip()
    headings=[norm(raw.columns[i]) for i in (4,14,19,31)]
    expected=('STATUS DO PROJETO','LEVANTADOR','ANALISTA(PRE-ANALISE)','OBSERVACOES')
    if any(w not in h for w,h in zip(expected,headings)):
        raise ValueError('Cabeçalhos da BASE_LIST não correspondem às colunas E, O e AF esperadas: '+', '.join(map(str,[raw.columns[i] for i in (4,14,19,31)])))
    result=[]
    for ix,row in raw.iterrows():
        status=row.iloc[4]
        if norm(status) not in ('CORRECAO DE LEVANTAMENTO','EM CORRECAO DE LEVANTAMENTO'):
            continue
        levantador=str(row.iloc[14]).strip() if pd.notna(row.iloc[14]) else ''
        analista=str(row.iloc[19]).strip() if pd.notna(row.iloc[19]) else ''
        if analista.lower() in ('nan','none','null'):analista=''
        obs=str(row.iloc[31]).strip() if pd.notna(row.iloc[31]) else ''
        if obs.lower() in ('nan','none','null'):obs=''
        nota=str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        result.append({'linha':int(ix)+2,'nota':nota,'levantador':levantador or '(Sem levantador)', 'analista':analista or '(Sem analista informado)', 'observacoes':obs})
    return json.dumps(result,ensure_ascii=False),len(raw),len(result)

def validate_workbook(blob:bytes):
    """Validation is read-only and never changes spreadsheet data."""
    issues=[]
    if len(blob)>90*1024*1024:
        return {'fatal':['Arquivo excede 90 MB.'], 'warnings':[], 'rows':0}
    try:
        with zipfile.ZipFile(io.BytesIO(blob)) as z:
            if z.testzip() is not None:raise ValueError('Arquivo Excel danificado.')
        raw=pd.read_excel(io.BytesIO(blob),sheet_name='NOTAS',engine='openpyxl',dtype=object)
    except Exception as exc:
        return {'fatal':['Não foi possível ler a aba NOTAS: '+str(exc)],'warnings':[], 'rows':0}
    raw.columns=[str(c).strip() for c in raw.columns]
    missing=[c for c in REQUIRED if c not in raw.columns]
    fatal=['Colunas obrigatórias ausentes: '+', '.join(missing)] if missing else []
    if not len(raw):fatal.append('A aba NOTAS está vazia.')
    if 'ID SISCO' in raw:
        ids=raw['ID SISCO'].dropna().astype(str).str.strip()
        repeated=int(ids[ids.ne('')].duplicated().sum())
        if repeated:issues.append(f'{repeated} IDs SISCO repetidos (revisar; não removidos).')
    for c in ('LATITUDE','LONGITUDE'):
        if c in raw:
            nums=pd.to_numeric(raw[c].dropna().astype(str).str.replace(',','.',regex=False),errors='coerce')
            invalid=int(nums.isna().sum())
            if invalid:issues.append(f'{invalid} valores não numéricos em {c}.')
    if 'DATA DESPACHO' in raw:
        v=raw['DATA DESPACHO'].dropna()
        invalid=int(pd.to_datetime(v,errors='coerce',dayfirst=False).isna().sum())
        if invalid:issues.append(f'{invalid} datas de despacho inválidas.')
    return {'fatal':fatal,'warnings':issues,'rows':len(raw)}


def github_put(path:str, content:bytes, message:str):
    link=api('contents/'+path)
    current=requests.get(link,headers=headers(True),params={'ref':GITHUB_BRANCH},timeout=45)
    if current.status_code not in (200,404):current.raise_for_status()
    body={'message':message,'branch':GITHUB_BRANCH,'content':base64.b64encode(content).decode('ascii')}
    if current.status_code==200:body['sha']=current.json()['sha']
    result=requests.put(link,headers=headers(True),json=body,timeout=180)
    result.raise_for_status()
    return result.json()

def publish(blob:bytes, actor:str):
    if not GITHUB_TOKEN:
        raise RuntimeError('Configure GITHUB_TOKEN nos Secrets com permissão Contents: Read and write.')
    if '/' not in GITHUB_REPO:
        raise RuntimeError('GITHUB_REPO inválido.')
    check=validate_workbook(blob)
    if check['fatal']:raise ValueError('; '.join(check['fatal']))
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
    audit={'timestamp_utc':datetime.utcnow().isoformat()+'Z','actor':actor,
           'records':count,'sha256':hashlib.sha256(blob).hexdigest(),
           'source':EXCEL_NAME,'commit':result.json().get('commit',{}).get('sha','')}
    try:
        github_put('nip_publicacao_atual.json',json.dumps(audit,ensure_ascii=False,indent=2).encode('utf-8'),
                   'NIP: auditoria de publicação da base')
        st.session_state['nip_audit_status']='Registro de auditoria publicado no GitHub.'
    except Exception:
        st.session_state['nip_audit_status']='A planilha foi publicada, mas o registro de auditoria falhou. Verifique permissões do GitHub.'
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
        if st.session_state.get('nip_audit_status'):st.info(st.session_state['nip_audit_status'])
        if st.button('Encerrar acesso administrativo',use_container_width=True):
            st.session_state.pop('nip_admin',None)
            st.rerun()
        file=st.file_uploader('Enviar Excel atualizado (aba NOTAS)',type=['xlsx'],key='nip_upload')
        if file is not None:
            try:
                check=validate_workbook(file.getvalue())
                if check['fatal']:raise ValueError('; '.join(check['fatal']))
                payload,n=parse_notas(file.getvalue())
                st.caption(f'Arquivo validado · {n:,} registros na aba NOTAS'.replace(',','.'))
                for warning in check['warnings'][:12]:st.warning(warning)
                if check['warnings']:
                    st.caption('Alertas não impedem a publicação. Confira os registros antes de confirmar.')
            except Exception as e:
                st.error('Arquivo inválido: '+str(e))
                file=None
        if file is not None and st.button('Publicar para todos',type='primary',use_container_width=True):
            try:
                with st.spinner('Publicando versão compartilhada...'):
                    n=publish(file.getvalue(),st.session_state.get('nip_actor','Administrador'))
                st.success(f'{n:,} obras publicadas. Todos verão a nova versão ao atualizar a página.'.replace(',','.'))
                st.info(st.session_state.get('nip_audit_status',''))
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
        list_payload='[]'
        list_status='BASE_LIST.xlsx não encontrada; adicione-a à raiz do GitHub.'
        try:
            list_blob,list_source=get_shared_list()
            if list_blob is not None:
                list_payload,all_list_rows,quality_rows=parse_list_quality(list_blob)
                list_status=f'{list_source} · {all_list_rows:,} registros · {quality_rows:,} em correção'.replace(',','.')
            else:
                list_status=list_source
        except Exception as exc:
            list_status='BASE_LIST indisponível: '+str(exc)
        if template.count('__NIP_LIST_PAYLOAD__')!=1 or template.count('__NIP_LIST_STATUS__')!=1:
            raise ValueError('Template de qualidade inválido: marcador BASE_LIST ausente ou duplicado.')
        safe_list=(list_payload.replace('<','\\u003c').replace('>','\\u003e')
                   .replace('&','\\u0026').replace('\u2028','\\u2028').replace('\u2029','\\u2029'))
        html=html.replace('__NIP_LIST_PAYLOAD__',safe_list)
        html=html.replace('__NIP_LIST_STATUS__',json.dumps(list_status,ensure_ascii=False))
        html=html.replace('Base incorporada •',source+' •')
        st.session_state['nip_loaded_count']=count
        # The iframe is viewport-height, not a fixed 1100px. This prevents the
        # outer Streamlit document from scrolling the embedded header off-screen.
        components.html(html,height=900,scrolling=False)
        st.caption(f'Fonte: {source} · {count:,} registros · aba NOTAS · revisão automática em até 90 segundos após recarregar a página.'.replace(',','.'))
    except Exception as e:
        st.error('Não foi possível carregar o Dashboard NIP.')
        st.warning(str(e))
        st.info('Confira se app_nip.py, nip_dashboard_template.html, logo_nip.png e BASE_LEVANTAMENTO_ATUALIZADA.xlsx estão na raiz do repositório.')

if __name__=='__main__':main()
