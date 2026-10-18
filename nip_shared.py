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
from nip_analytics import build_payload
import streamlit.components.v1 as components


ROOT = Path(__file__).resolve().parent
EXCEL_NAME = 'BASE_LEVANTAMENTO_ATUALIZADA.xlsx'
TEMPLATE_NAME = 'nip_dashboard_template.html'
LIST_NAME = 'BASE_LIST.xlsx'
LOCALIDADE_NAME = 'LOCALIDADE LEVANTADORES-SANEAMENTO.xlsx'
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
[data-testid="stSidebar"] {background:#fff;}
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
        if data:return data,'Dados atualizados'
    except Exception:
        pass
    local=ROOT/EXCEL_NAME
    if local.exists():return local.read_bytes(),'Dados disponíveis'
    raise FileNotFoundError('Dados operacionais indisponíveis no momento.')

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
            return remote, 'Dados atualizados'
    except Exception:
        pass
    local=ROOT/LIST_NAME
    if local.exists():
        return local.read_bytes(), 'Dados disponíveis'
    return None, 'Dados de qualidade e produtividade indisponíveis no momento.'

@st.cache_data(ttl=90, show_spinner=False)
def get_shared_localidade():
    """Read-only residence reference; never assigns or changes an obra."""
    try:
        remote = get_remote_file(filename=LOCALIDADE_NAME)
        if remote is not None:
            return remote
    except Exception:
        pass
    local = ROOT / LOCALIDADE_NAME
    return local.read_bytes() if local.exists() else None


@st.cache_data(show_spinner=False, max_entries=4)
def parse_localidade(blob: bytes):
    """Residence points for geographic comparison ONLY."""
    import math
    import unicodedata
    def normalized(value):
        return ''.join(c for c in unicodedata.normalize('NFKD', str(value or ''))
                       if not unicodedata.combining(c)).strip().upper()
    book = pd.read_excel(io.BytesIO(blob), sheet_name=None, engine='openpyxl', dtype=object)
    people = []
    for sheet, team in [('LEVANTADORES', 'Levantamento'), ('SANEAMENTO', 'Saneamento')]:
        name = next((key for key in book if normalized(key) == sheet), None)
        if name is None:
            raise ValueError(f'A planilha de localidade nao possui a aba {sheet}.')
        df = book[name]
        lookup = {normalized(c): c for c in df.columns}
        required = ['NOME', 'LATITUDE', 'LONGITUDE']
        if any(c not in lookup for c in required):
            raise ValueError(f'Aba {sheet}: colunas Nome, Latitude e Longitude obrigatorias.')
        city_col = lookup.get('CIDADES') or lookup.get('CIDADE')
        for _, row in df.iterrows():
            person = str(row[lookup['NOME']]).strip() if pd.notna(row[lookup['NOME']]) else ''
            if not person:
                continue
            try:
                lat = float(str(row[lookup['LATITUDE']]).strip().replace(',', '.'))
                lon = float(str(row[lookup['LONGITUDE']]).strip().replace(',', '.'))
            except (TypeError, ValueError):
                continue
            if not (math.isfinite(lat) and math.isfinite(lon) and -90 <= lat <= 90 and -180 <= lon <= 180):
                continue
            city = row[city_col] if city_col else ''
            people.append({'nome': person, 'cidade': str(city).strip() if pd.notna(city) else '',
                           'lat': lat, 'lon': lon, 'equipe': team})
    return json.dumps(people, ensure_ascii=False, separators=(',', ':'))


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

@st.cache_data(show_spinner=False,max_entries=4)
def parse_list_productivity(blob:bytes):
    """BASE_LIST: H municipio, O levantador, Q data levantamento, R/S postes, U data analise.
    Preserve rows with Q or U; production in field is keyed to Q, analyses to U.
    """
    import unicodedata
    df=pd.read_excel(io.BytesIO(blob),sheet_name=0,engine='openpyxl',dtype=object)
    if df.shape[1]<21:raise ValueError('BASE_LIST precisa conter as colunas H, O, Q, R, S e U.')
    def norm(v):
        return ''.join(c for c in unicodedata.normalize('NFKD',str(v or '')) if not unicodedata.combining(c)).upper().strip()
    expected={7:'MUNICIPIO',14:'LEVANTADOR',16:'DATA DO LEVANTAMENTO',17:'QNT. POSTES LEVANTADOS',18:'P L N',20:'DATA DA ANALISE DE LEVANTAMENTO'}
    invalid=[f'{i+1}: {df.columns[i]}' for i,v in expected.items() if v not in norm(df.columns[i])]
    if invalid:raise ValueError('Cabeçalhos inesperados na BASE_LIST: '+', '.join(invalid))
    def date(v):
        if pd.isna(v):return ''
        try:
            d=pd.to_datetime(v,errors='coerce',dayfirst=True)
            if pd.isna(d) or d.year<2000 or d.year>2100:return ''
            return d.strftime('%Y-%m-%d')
        except Exception:return ''
    def number(v):
        if pd.isna(v) or str(v).strip()=='':return 0
        try:
            n=float(str(v).replace(' ','').replace(',','.'))
            return int(n) if n>=0 and n.is_integer() and n<=100000 else 0
        except Exception:return 0
    result=[]
    for ix,row in df.iterrows():
        u=date(row.iloc[20]);q=date(row.iloc[16]);
        if not (q or u):continue
        staff=str(row.iloc[14]).strip() if pd.notna(row.iloc[14]) else ''
        town=str(row.iloc[7]).strip() if pd.notna(row.iloc[7]) else ''
        note=str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        result.append({'linha':int(ix)+2,'nota':note,'municipio':town or '(Sem município)',
            'levantador':staff or '(Sem levantador)','data':u,'data_levantamento':q,
            'postes_existentes':number(row.iloc[17]),'postes_novos':number(row.iloc[18])})
    return json.dumps(result,ensure_ascii=False),len(df),len(result)


@st.cache_data(show_spinner=False, max_entries=4)
def parse_preanalise(blob: bytes):
    """Leitura independente da BASE_LIST: C, E, Q, T, U. Sem modificar a fonte."""
    import unicodedata
    df = pd.read_excel(io.BytesIO(blob), sheet_name=0, engine='openpyxl', dtype=object)
    if df.shape[1] < 21:
        raise ValueError('BASE_LIST precisa conter as colunas C, E, Q, T e U.')
    def norm(value):
        return ''.join(c for c in unicodedata.normalize('NFKD', str(value or '')) if not unicodedata.combining(c)).upper().strip()
    expected = {4:'STATUS DO PROJETO', 16:'DATA DO LEVANTAMENTO', 19:'ANALISTA', 20:'DATA DA ANALISE'}
    invalid = [f'{i+1}: {df.columns[i]}' for i, word in expected.items() if word not in norm(df.columns[i])]
    if invalid:
        raise ValueError('Cabeçalhos inesperados na BASE_LIST: '+', '.join(invalid))
    def clean(value):
        if pd.isna(value): return ''
        s = str(value).strip()
        return '' if s.lower() in ('nan','none','null') else s
    def date(value):
        if pd.isna(value) or str(value).strip() == '': return ''
        try:
            d = pd.to_datetime(value, errors='coerce', dayfirst=True)
            return d.strftime('%Y-%m-%d') if pd.notna(d) and 2000 <= d.year <= 2100 else ''
        except (ValueError, TypeError): return ''
    records=[]
    for ix, row in enumerate(df.itertuples(index=False, name=None)):
        records.append({'linha':ix+2, 'nota':clean(row[0]), 'nota_sgo':clean(row[2]),
            'status':clean(row[4]), 'analista':clean(row[19]),
            'q':date(row[16]), 'u':date(row[20])})
    return json.dumps(records, ensure_ascii=False, separators=(',', ':'))

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
        st.session_state['nip_audit_status']='Registro de auditoria salvo com sucesso.'
    except Exception:
        st.session_state['nip_audit_status']='A atualização foi concluída, mas o registro de auditoria não pôde ser confirmado.'
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
            except Exception as e:st.error('A atualização não pôde ser concluída. Verifique as permissões com o administrador.')
        if not GITHUB_TOKEN:st.info('A publicação compartilhada está indisponível. Solicite a configuração ao administrador do sistema.')


def render_page(page_id: str):
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
        if template.count('__NIP_LOCALIDADE_PAYLOAD__') != 1:
            raise ValueError('Marcador de localidade ausente ou duplicado no HTML.')
        localidade_payload = '[]'
        localidade_blob = get_shared_localidade()
        if localidade_blob is not None:
            try:
                localidade_payload = parse_localidade(localidade_blob)
            except Exception:
                # Do not replace residence locations with stale hardcoded data.
                localidade_payload = '[]'
        localidade_payload = (localidade_payload.replace('<', '\\u003c')
                             .replace('>', '\\u003e').replace('&', '\\u0026')
                             .replace('\u2028', '\\u2028').replace('\u2029', '\\u2029'))
        html = html.replace('__NIP_LOCALIDADE_PAYLOAD__', localidade_payload)
        list_payload='[]'
        productivity_payload='[]'
        productivity_status='Dados de produtividade indisponíveis no momento.'
        analytics_payload='null'
        preanalise_payload='[]'
        preanalise_status='BASE_LIST indisponível'
        list_status='Dados de qualidade indisponíveis no momento.'
        try:
            list_blob,list_source=get_shared_list()
            if list_blob is not None:
                analytics_payload=build_payload(list_blob)
                try:
                    preanalise_payload=parse_preanalise(list_blob)
                    preanalise_status='Dados atualizados da BASE_LIST'
                except Exception as pre_exc:
                    preanalise_status='Não foi possível validar as colunas C, E, Q, T e U: '+str(pre_exc)
                list_payload,all_list_rows,quality_rows=parse_list_quality(list_blob)
                productivity_payload,prod_total,prod_count=parse_list_productivity(list_blob)
                productivity_status=f'{prod_total:,} registros · {prod_count:,} com data Q ou U'.replace(',','.')
                list_status=f'{all_list_rows:,} registros · {quality_rows:,} em correção'.replace(',','.')
            else:
                list_status='Dados de qualidade indisponíveis no momento.'
        except Exception as exc:
            list_status='Dados de qualidade indisponíveis no momento.'
            productivity_status=list_status
        if template.count('__NIP_LIST_PAYLOAD__')!=1 or template.count('__NIP_LIST_STATUS__')!=1:
            raise ValueError('Template de qualidade inválido: marcador BASE_LIST ausente ou duplicado.')
        safe_list=(list_payload.replace('<','\\u003c').replace('>','\\u003e')
                   .replace('&','\\u0026').replace('\u2028','\\u2028').replace('\u2029','\\u2029'))
        html=html.replace('__NIP_LIST_PAYLOAD__',safe_list)
        html=html.replace('__NIP_LIST_STATUS__',json.dumps(list_status,ensure_ascii=False))
        if template.count('__NIP_PRODUCTIVITY_PAYLOAD__')!=1 or template.count('__NIP_PRODUCTIVITY_STATUS__')!=1:
            raise ValueError('Template de produtividade inválido: marcador ausente ou duplicado.')
        safe_productivity=(productivity_payload.replace('<','\\u003c').replace('>','\\u003e').replace('&','\\u0026'))
        html=html.replace('__NIP_PRODUCTIVITY_PAYLOAD__',safe_productivity)
        html=html.replace('__NIP_PRODUCTIVITY_STATUS__',json.dumps(productivity_status,ensure_ascii=False))
        if template.count('__NIP_ANALYTICS_PAYLOAD__')!=1:
            raise ValueError('Marcador da analise integrada ausente ou duplicado.')
        safe_analytics=analytics_payload.replace('<','\\u003c').replace('>','\\u003e').replace('&','\\u0026')
        html=html.replace('__NIP_ANALYTICS_PAYLOAD__',safe_analytics)
        if template.count('__NIP_PREANALISE_PAYLOAD__')!=1 or template.count('__NIP_PREANALISE_STATUS__')!=1:
            raise ValueError('Marcador de Pré-Análise ausente ou duplicado no HTML.')
        safe_preanalise=(preanalise_payload.replace('<','\\u003c').replace('>','\\u003e').replace('&','\\u0026'))
        html=html.replace('__NIP_PREANALISE_PAYLOAD__',safe_preanalise)
        html=html.replace('__NIP_PREANALISE_STATUS__',json.dumps(preanalise_status,ensure_ascii=False))
        html=html.replace('Base incorporada •','')
        st.session_state['nip_loaded_count']=count
        # The iframe is viewport-height, not a fixed 1100px. This prevents the
        # outer Streamlit document from scrolling the embedded header off-screen.
        html = add_page_controller(html, page_id)
        components.html(html,height=900,scrolling=False)
        st.caption(f'{count:,} registros disponíveis · atualização automática de indicadores.'.replace(',','.'))
    except Exception as e:
        st.error('Não foi possível carregar o Dashboard NIP.')
        st.warning('Os dados não puderam ser carregados neste momento.')
        st.info('Solicite ao administrador que verifique a disponibilidade do sistema.')




# Only activate the requested tab on initial load. Never hide the original
# horizontal navigation or override the original dashboard CSS.
_PAGE_TABS = {
    'geral': 'nip-tab-geral',
    'equipe': 'nip-tab-equipe',
    'prazos': 'nip-tab-prazos',
    'relatorios': 'nip-tab-relatorios',
    'qualidade-list': 'nip-tab-qualidade-list',
    'produtividade': 'nip-tab-produtividade',
    'preanalise': 'nip-tab-preanalise',
    'faq': 'nip-tab-faq',
}

def add_page_controller(html: str, page_id: str) -> str:
    if page_id not in _PAGE_TABS:
        raise ValueError('Página não reconhecida: ' + page_id)
    # The default route uses the untouched original tab selection.
    if page_id == 'geral':
        return html
    tab = json.dumps(_PAGE_TABS[page_id])
    js = """<script>
window.addEventListener('load', function(){
  const target = document.getElementById(__TAB__);
  if (!target) return;
  document.querySelectorAll('.nip-tab').forEach(function(p){
    p.classList.toggle('nip-hidden', p !== target);
  });
  const nav = document.getElementById('nipNav');
  if(nav){
    nav.querySelectorAll('button').forEach(function(b){
      b.classList.toggle('selected', b.getAttribute('data-target') === __TAB__);
    });
  }
}, {once:true});
</script>""".replace('__TAB__', tab)
    return html.replace('</body>', js+'</body>', 1)
