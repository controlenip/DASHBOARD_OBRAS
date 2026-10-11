"""NIP | Gestao Integrada de Obras - Streamlit.
Fonte: arquivo Excel BASE_LEVANTAMENTO_ATUALIZADA.xlsx (aba NOTAS).
Consulta publica. Publicacao compartilhada via GitHub com token administrativo nos Secrets.
"""
import os, io, re, json, time, hmac, hashlib, secrets, unicodedata, base64
from datetime import datetime, date, timedelta
from collections import Counter
from pathlib import Path

import pandas as pd
import numpy as np
import streamlit as st
import requests
import folium
from folium.plugins import MarkerCluster, HeatMap
from streamlit_folium import st_folium
import plotly.express as px
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

st.set_page_config(page_title="NIP | Gestão Integrada", page_icon="🗺️", layout="wide", initial_sidebar_state="expanded")

C = {"blue":"#092d55", "teal":"#00a6b2", "bg":"#f3f6fb", "danger":"#c23c46", "orange":"#e4a11a"}
CSS = """<style>.stApp{background:#f3f6fb;}[data-testid="stSidebar"]{background:#eaf0f7;}
.block-container{padding-top:1.1rem;max-width:1700px}h1,h2,h3{color:#092d55}
.niphero{background:linear-gradient(110deg,#092d55,#175285);padding:17px 22px;border-radius:13px;color:white;margin-bottom:12px}
.niphero h2{color:white;margin:0}.niphero small{color:#dcecfb}
[data-testid="stMetric"]{background:white;padding:13px;border-radius:11px;border-left:4px solid #00a6b2;box-shadow:0 2px 8px #06244812}
[data-testid="stMetricLabel"]{color:#37516c}
</style>"""
st.markdown(CSS, unsafe_allow_html=True)

# Secrets are never displayed and are never copied into session data.
def config(k, default=None):
    try:
        v=st.secrets.get(k)
        if v is not None: return v
    except Exception: pass
    return os.environ.get(k,default)

ADMIN_PASSWORD=config('ADMIN_PASSWORD')
GITHUB_TOKEN=config('GITHUB_TOKEN','')
GITHUB_REPO=config('GITHUB_REPO','controlenip/DASHBOARD_OBRAS')
GITHUB_BRANCH=config('GITHUB_BRANCH','main')
EXCEL_FILE='BASE_LEVANTAMENTO_ATUALIZADA.xlsx'
LOCAL_EXCEL=Path(__file__).with_name(EXCEL_FILE)
LOGO_PATH=Path(__file__).with_name('logo_nip.png')

# Public GitHub repository hosts the shared Excel; visitors do not require credentials.
# A private repository can be read with GITHUB_TOKEN, but then visitors need a
# server-side configured token with repository read access.
def github_api(suffix):
    return f'https://api.github.com/repos/{GITHUB_REPO}/{suffix}'

def github_headers(auth=False):
    h={'Accept':'application/vnd.github+json','X-GitHub-Api-Version':'2022-11-28','User-Agent':'nip-dashboard-streamlit'}
    if GITHUB_TOKEN and auth: h['Authorization']='Bearer '+GITHUB_TOKEN
    return h

def github_file():
    if not GITHUB_REPO or '/' not in GITHUB_REPO:
        raise ValueError('GITHUB_REPO precisa estar no formato usuario/repositorio')
    r=requests.get(github_api('contents/'+EXCEL_FILE),headers=github_headers(bool(GITHUB_TOKEN)),params={'ref':GITHUB_BRANCH},timeout=35)
    if r.status_code==404: return None
    r.raise_for_status()
    item=r.json()
    if item.get('content'):
        return base64.b64decode(item['content'].replace('\n',''))
    # API omits inline content for larger files; use Git blob API by SHA instead.
    blob=requests.get(github_api('git/blobs/'+item['sha']),headers=github_headers(bool(GITHUB_TOKEN)),timeout=60)
    blob.raise_for_status()
    data=blob.json()
    if data.get('encoding')!='base64':raise RuntimeError('GitHub não retornou a planilha em base64')
    return base64.b64decode(data['content'].replace('\n',''))

def source_bytes():
    try:
        blob=github_file()
        if blob: return blob,'GitHub (base compartilhada)'
    except Exception as exc:
        if not LOCAL_EXCEL.exists():
            raise RuntimeError('Não foi possível consultar o Excel compartilhado no GitHub: '+str(exc)) from exc
        st.warning('GitHub indisponível. Exibindo cópia inicial do Excel incluída no aplicativo, que pode estar desatualizada.')
    if LOCAL_EXCEL.exists():return LOCAL_EXCEL.read_bytes(),'Excel inicial do repositório'
    raise FileNotFoundError('Envie BASE_LEVANTAMENTO_ATUALIZADA.xlsx para a raiz do repositório GitHub.')

def latest_id():
    # Kept for compatibility with old main() calling pattern.
    return 1 if (LOCAL_EXCEL.exists() or bool(GITHUB_REPO)) else None

def snapshots():
    # Historical versions are individual commits containing the Excel.
    try:
        r=requests.get(github_api('commits'),params={'path':EXCEL_FILE,'sha':GITHUB_BRANCH,'per_page':15},headers=github_headers(bool(GITHUB_TOKEN)),timeout=25)
        if not r.ok:return pd.DataFrame()
        return pd.DataFrame([{'id':x['sha'],'imported_at':x['commit']['committer']['date'],'actor':x['commit']['author'].get('name','Administrador'),'filename':EXCEL_FILE,'row_count':0} for x in r.json()])
    except Exception:return pd.DataFrame()

def get_snapshot(sid):
    if sid==1:
        contents,source=source_bytes()
        st.session_state['_excel_source']=source
        return read_excel(io.BytesIO(contents))
    # Read older Excel directly from the git tree at commit SHA.
    ref=str(sid)
    r=requests.get(github_api('contents/'+EXCEL_FILE),params={'ref':ref},headers=github_headers(bool(GITHUB_TOKEN)),timeout=30)
    r.raise_for_status()
    item=r.json()
    blob=requests.get(github_api('git/blobs/'+item['sha']),headers=github_headers(bool(GITHUB_TOKEN)),timeout=60)
    blob.raise_for_status()
    return read_excel(io.BytesIO(base64.b64decode(blob.json()['content'].replace('\n',''))))

def insert_snapshot(upload,filename,actor):
    if not GITHUB_TOKEN:
        raise RuntimeError('Para publicar para todos, configure GITHUB_TOKEN nos Secrets do Streamlit com acesso Contents: Read and write.')
    if '/' not in GITHUB_REPO:raise RuntimeError('GITHUB_REPO inválido.')
    contents=upload.getvalue() if hasattr(upload,'getvalue') else bytes(upload)
    # Validation before publishing prevents a broken dashboard.
    checked=read_excel(io.BytesIO(contents))
    if checked.empty:raise ValueError('A aba NOTAS está vazia.')
    url=github_api('contents/'+EXCEL_FILE)
    headers=github_headers(True)
    old=requests.get(url,headers=headers,params={'ref':GITHUB_BRANCH},timeout=40)
    if old.status_code not in (200,404):old.raise_for_status()
    data={'message':f'Atualizar Excel do dashboard NIP - {actor}', 'content':base64.b64encode(contents).decode('ascii'),'branch':GITHUB_BRANCH}
    if old.status_code==200:data['sha']=old.json()['sha']
    res=requests.put(url,json=data,headers=headers,timeout=120)
    if not res.ok:raise RuntimeError(f'GitHub recusou publicação (HTTP {res.status_code}): verifique permissão Contents Write e branch.')
    st.cache_data.clear()
    return len(checked)

def admin_panel():
    """Public dashboard; only spreadsheet publication requires administrator authentication."""
    if 'admin_authenticated' not in st.session_state:
        st.session_state.admin_authenticated = False
    with st.sidebar.expander('🔒 Administração / Atualizar base', expanded=False):
        if not st.session_state.admin_authenticated:
            st.caption('O dashboard é público. A senha é necessária apenas para enviar uma nova planilha.')
            if not ADMIN_PASSWORD:
                st.warning('Importação bloqueada: configure ADMIN_PASSWORD nos Secrets do Streamlit.')
                return
            with st.form('admin_login_form', clear_on_submit=True):
                identification = st.text_input('Identificação do administrador')
                password = st.text_input('Senha de administrador', type='password')
                submit = st.form_submit_button('Desbloquear importação', use_container_width=True)
            if submit:
                if password and hmac.compare_digest(password, str(ADMIN_PASSWORD)):
                    st.session_state.admin_authenticated = True
                    st.session_state.admin_actor = identification.strip() or 'Administrador'
                    st.rerun()
                else:
                    st.error('Senha incorreta.')
            return

        st.success('Publicação de bases autorizada nesta sessão.')
        if st.button('Bloquear administração', key='admin_logout'):
            st.session_state.admin_authenticated = False
            st.session_state.pop('admin_actor', None)
            st.rerun()
        f = st.file_uploader('Excel com aba NOTAS', type=['xlsx'], key='admin_file')
        if f is not None and st.button('Publicar base para todos', type='primary', key='admin_publish'):
            if not st.session_state.admin_authenticated:
                st.error('Acesso administrativo necessário.')
                return
            try:
                with st.spinner('Validando e gravando versão compartilhada...'):
                    insert_snapshot(f, f.name, st.session_state.get('admin_actor', 'Administrador'))
                st.success('Base publicada para todos os visitantes.')
                st.rerun()
            except Exception as ex:
                st.error(f'Importação recusada: {ex}')
        hist = snapshots()
        if not hist.empty:
            st.caption('Versões publicadas')
            st.dataframe(hist[['id','imported_at','row_count','filename']].head(8), hide_index=True)

REQUIRED=['ID SISCO','LEVANTADOR','STATUS LIST','STATUS SAP','STATUS SISCO','REGIONAL','MUNICIPIO','DATA DESPACHO']
FIELDS=['ID SISCO','PROTOCOLO','TIPO NOTA','LEVANTADOR','DATA DESPACHO','DATA ABERTURA','PRIORIDADE','STATUS SAP','STATUS SISCO','STATUS LIST','CONTRATO','REGIONAL','MUNICIPIO','NOME','LATITUDE','LONGITUDE','ENDEREÇO','LOCALIDADE','ÁREA DE ATUAÇÃO (LEVANTADOR)']

def norm(x):
    if pd.isna(x):return ''
    return unicodedata.normalize('NFKD',str(x)).encode('ascii','ignore').decode().strip().casefold()

def status(x):
    v=norm(x)
    return '0' if v in ('0','0.0','') else v

def read_excel(upload):
    # Read only needed worksheet to avoid slow formula-heavy other sheets.
    frame=pd.read_excel(upload,sheet_name='NOTAS',dtype=object,engine='openpyxl')
    frame.columns=[str(c).strip() for c in frame.columns]
    missing=[c for c in REQUIRED if c not in frame.columns]
    if missing:raise ValueError('Colunas obrigatórias ausentes: '+', '.join(missing))
    for c in FIELDS:
        if c not in frame:frame[c]=None
    frame=frame[FIELDS].copy()
    for c in ['LATITUDE','LONGITUDE']:
        frame[c]=pd.to_numeric(frame[c].astype(str).str.replace(',','.',regex=False),errors='coerce')
    for c in ['DATA DESPACHO','DATA ABERTURA']:
        frame[c]=pd.to_datetime(frame[c],errors='coerce').dt.strftime('%Y-%m-%d')
    frame['ID SISCO']=frame['ID SISCO'].apply(lambda v:str(int(v)) if pd.notna(v) and isinstance(v,(int,float)) and float(v).is_integer() else (str(v).strip() if pd.notna(v) else ''))
    return frame.replace({np.nan:None}).reset_index(drop=True)

def prepared(df):
    if df.empty:return df.copy()
    d=df.copy()
    for c in FIELDS:
        if c not in d:d[c]=None
    for c in ['LATITUDE','LONGITUDE']:d[c]=pd.to_numeric(d[c],errors='coerce')
    d['COORD_OK']=d['LATITUDE'].between(-11,-0.5)&d['LONGITUDE'].between(-49.8,-41.0)&~((d['LATITUDE']==0)|(d['LONGITUDE']==0))
    d['SEM_NOME']=d['NOME'].map(lambda x:norm(x) in ('','0','sem informacao','nan','none'))
    d['LIST_N']=d['STATUS LIST'].map(status)
    d['SAP_N']=d['STATUS SAP'].map(status)
    d['SISCO_N']=d['STATUS SISCO'].map(status)
    d['CONTRATO_N']=d['CONTRATO'].map(status)
    d['LEV_N']=d['LEVANTADOR'].fillna('SEM LEVANTADOR').astype(str).str.strip().replace('','SEM LEVANTADOR')
    d['ELEGIVEL']=(d['LIST_N'].isin(['0','em levantamento','correcao de levantamento'])&~d['SAP_N'].isin(['canc','finl'])&d['SISCO_N'].isin(['0','liberado para levantamento','pre analise']))
    d['LIST_GRUPO']=d['LIST_N'].map({'0':'0','em levantamento':'Em levantamento','correcao de levantamento':'Correção de levantamento'}).fillna('Outros')
    dates=pd.to_datetime(d['DATA DESPACHO'],errors='coerce')
    d['DIAS']=((pd.Timestamp.now().normalize()-dates).dt.days).astype('Int64')
    d['ATRASADA']=d['ELEGIVEL'] & d['DIAS'].gt(10).fillna(False)
    d['PRIORITARIA']=d['PRIORIDADE'].map(lambda x:status(x) in ('sim','alta','urgente','critica','critico','1'))
    d['CHAVE']=d['ID SISCO'].fillna('').astype(str).str.strip()
    return d

def filters(d):
    st.sidebar.subheader('Filtros operacionais')
    search=st.sidebar.text_input('Pesquisar nota / SGO / protocolo / nome',key='search')
    opts=['REGIONAL','MUNICIPIO','TIPO NOTA','LEV_N','STATUS SAP','STATUS SISCO','STATUS LIST','CONTRATO','PRIORIDADE']
    labels={'LEV_N':'LEVANTADOR'}
    selections={}
    for col in opts:
        vals=sorted(d[col].dropna().astype(str).unique().tolist()) if col in d else []
        selections[col]=st.sidebar.multiselect(labels.get(col,col),vals,key='filter_'+col)
    a,b=st.sidebar.columns(2)
    dt1=a.date_input('Despacho de',value=None,key='dt1')
    dt2=b.date_input('Até',value=None,key='dt2')
    only_nocoord=st.sidebar.checkbox('Somente sem coordenadas')
    only_noname=st.sidebar.checkbox('Somente sem NOME')
    only_elig=st.sidebar.checkbox('Somente carteira elegível')
    out=d.copy()
    for col,sel in selections.items():
        if sel:out=out[out[col].astype(str).isin(sel)]
    if search:
        q=re.escape(search.strip())
        mask=pd.Series(False,index=out.index)
        for c in ['ID SISCO','PROTOCOLO','NOME','LOCALIDADE']:
            mask|=out[c].fillna('').astype(str).str.contains(q,case=False,regex=True)
        out=out[mask]
    if dt1:out=out[pd.to_datetime(out['DATA DESPACHO'],errors='coerce')>=pd.Timestamp(dt1)]
    if dt2:out=out[pd.to_datetime(out['DATA DESPACHO'],errors='coerce')<=pd.Timestamp(dt2)]
    if only_nocoord:out=out[~out.COORD_OK]
    if only_noname:out=out[out.SEM_NOME]
    if only_elig:out=out[out.ELEGIVEL]
    if st.sidebar.button('Limpar filtros'):
        for k in list(st.session_state):
            if k.startswith('filter_') or k in ('search','dt1','dt2'):del st.session_state[k]
        st.rerun()
    return out

def metric_grid(items):
    cs=st.columns(len(items))
    for col,(name,value) in zip(cs,items):col.metric(name,f'{value:,}'.replace(',','.'))

def map_view(d):
    st.subheader('Mapa geográfico de obras — Maranhão')
    mapdata=d[d.COORD_OK].copy()
    layer=st.radio('Camada do mapa',['Maps','Satélite','Relevo'],horizontal=True)
    heat=st.checkbox('Mapa de calor de obras atrasadas',value=False)
    tile={
      'Maps':'https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}',
      'Satélite':'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
      'Relevo':'https://server.arcgisonline.com/ArcGIS/rest/services/World_Topo_Map/MapServer/tile/{z}/{y}/{x}'}[layer]
    m=folium.Map(location=[-5.2,-45.0],zoom_start=7,tiles=None,min_zoom=6,max_zoom=18,control_scale=True)
    folium.TileLayer(tiles=tile,attr='Tiles © Esri',name=layer,max_zoom=18).add_to(m)
    folium.TileLayer('OpenStreetMap',name='Alternativa OSM',show=False).add_to(m)
    if mapdata.empty:
        m.fit_bounds([[-10.5,-48.9],[-1.0,-41.5]])
    folium.Rectangle(bounds=[[-10.6,-49.0],[-0.8,-41.3]],color='transparent',fill=False).add_to(m)
    cluster=MarkerCluster(name='Obras').add_to(m)
    # Preserve all points on map. Cap popup field text, never discard rows.
    for _,r in mapdata.iterrows():
        desc=''.join(f'<div style="margin:3px 0"><b>{label}:</b> {str(r.get(key) if pd.notna(r.get(key)) else "")[:180]}</div>' for label,key in [('Nota','ID SISCO'),('Nome','NOME'),('Levantador','LEVANTADOR'),('Status LIST','STATUS LIST'),('SAP','STATUS SAP'),('SISCO','STATUS SISCO'),('Município','MUNICIPIO'),('Regional','REGIONAL'),('Dias despacho','DIAS')])
        bg='#c23c46' if r.ATRASADA else '#0d4677'
        marker=folium.DivIcon(html=f'<div style="background:{bg};color:white;border:2px solid white;border-radius:8px;padding:5px;font-weight:bold;box-shadow:0 2px 8px #333;width:30px;text-align:center">●</div>')
        folium.Marker([float(r.LATITUDE),float(r.LONGITUDE)],icon=marker,popup=folium.Popup(f'<div style="min-width:230px;background:#092d55;color:white;padding:12px;border-radius:7px"><b>OBRA {r.get("ID SISCO","")}</b></div><div style="padding:8px">{desc}</div>',max_width=350)).add_to(cluster)
    if heat:
        pts=mapdata[mapdata.ATRASADA][['LATITUDE','LONGITUDE']].dropna().values.tolist()
        if pts:HeatMap(pts,radius=17,blur=14,name='Concentração de atrasos').add_to(m)
    if len(mapdata):
        latmin,latmax=float(mapdata.LATITUDE.min()),float(mapdata.LATITUDE.max())
        lonmin,lonmax=float(mapdata.LONGITUDE.min()),float(mapdata.LONGITUDE.max())
        if len(mapdata)==1:m.location=[latmin,lonmin];m.options['zoom']=13
        else:m.fit_bounds([[latmin,lonmin],[latmax,lonmax]],padding=(18,18))
    folium.LayerControl().add_to(m)
    st_folium(m,height=560,width=None,returned_objects=[],key='main_map')
    if (~d.COORD_OK).any():
        with st.expander(f'Obras sem coordenadas: {(~d.COORD_OK).sum():,}'):
            st.dataframe(d.loc[~d.COORD_OK,['ID SISCO','REGIONAL','MUNICIPIO','LEVANTADOR','STATUS LIST']],hide_index=True,use_container_width=True)

def chart_counts(d,col,title,top=15):
    counts=d[col].fillna('SEM INFORMAÇÃO').astype(str).value_counts().head(top).reset_index()
    counts.columns=[col,'Obras']
    if len(counts):st.plotly_chart(px.bar(counts,x='Obras',y=col,orientation='h',title=title).update_layout(yaxis={'categoryorder':'total ascending'},height=max(300,28*len(counts)+100)),use_container_width=True)

def load_table(d):
    work=d[d.ELEGIVEL & ~d.LEV_N.str.upper().eq('SEM LEVANTADOR')]
    if work.empty:return pd.DataFrame()
    t=pd.crosstab(work.LEV_N,work.LIST_GRUPO)
    for col in ['0','Em levantamento','Correção de levantamento']:
        if col not in t:t[col]=0
    t=t[['0','Em levantamento','Correção de levantamento']].reset_index().rename(columns={'LEV_N':'LEVANTADOR'})
    t['TOTAL']=t[['0','Em levantamento','Correção de levantamento']].sum(axis=1)
    t['SITUAÇÃO']=np.select([t.TOTAL<10,t.TOTAL>50],['⚠️ Recarregar demanda','🔴 Verificar acúmulo'],default='✅ Faixa normal')
    return t.sort_values('TOTAL',ascending=False)

def redistributions(d,table):
    st.markdown('#### Sugestões de redistribuição')
    if table.empty:return st.info('Sem equipes para comparar.')
    lows=table[table.TOTAL<10]; highs=table[table.TOTAL>50]
    if lows.empty or highs.empty:return st.info('Não há simultaneamente levantadores abaixo de 10 e acima de 50 obras.')
    # Match within municipality or regional; no coordinate-based distance estimated.
    loc=d[d.ELEGIVEL & d.COORD_OK]
    suggestions=[]
    for _,h in highs.iterrows():
        hm=loc[loc.LEV_N==h.LEVANTADOR]
        for _,l in lows.iterrows():
            lm=loc[loc.LEV_N==l.LEVANTADOR]
            hmuns=set(hm.MUNICIPIO.dropna().astype(str));lmuns=set(lm.MUNICIPIO.dropna().astype(str))
            hregs=set(hm.REGIONAL.dropna().astype(str));lregs=set(lm.REGIONAL.dropna().astype(str))
            common=hmuns&lmuns
            if common or hregs&lregs:
                suggestions.append({'ORIGEM':h.LEVANTADOR,'OBRAS ORIGEM':int(h.TOTAL),'DESTINO':l.LEVANTADOR,'OBRAS DESTINO':int(l.TOTAL),'CRITÉRIO':'Mesmo município: '+', '.join(sorted(common)[:2]) if common else 'Mesma regional','TRANSFERÊNCIA SUGERIDA':min(10-int(l.TOTAL),int(h.TOTAL)-50)})
    if suggestions:st.dataframe(pd.DataFrame(suggestions).sort_values('TRANSFERÊNCIA SUGERIDA',ascending=False),hide_index=True,use_container_width=True)
    else:st.info('Sem correspondências de município ou regional entre os grupos. Avaliar manualmente.')
    st.caption('Propostas indicativas, sem movimentação automática; não estimam distância viária nem consideram capacidade de campo.')

def aging(d):
    work=d[d.ELEGIVEL].copy()
    today=pd.Timestamp.now().normalize()
    valid=work[work.DIAS.notna()].copy()
    labels=['Até 5 dias','6–10 dias','11–15 dias','16–30 dias','Mais de 30 dias']
    valid['FAIXA']=pd.cut(valid.DIAS.astype(float),[-1e10,5,10,15,30,1e10],labels=labels)
    counts=valid.FAIXA.value_counts().reindex(labels,fill_value=0)
    metric_grid([(label,int(v)) for label,v in counts.items()])
    st.plotly_chart(px.bar(x=counts.index,y=counts.values,labels={'x':'Faixa desde despacho','y':'Obras'},title='Envelhecimento da carteira elegível'),use_container_width=True)
    due=work[work.DIAS.between(8,10,inclusive='both')].copy()
    st.write(f'**Próximas a ultrapassar 10 dias (8–10 dias completos): {len(due)} obras**')
    if len(due):st.dataframe(due[['ID SISCO','MUNICIPIO','REGIONAL','LEVANTADOR','DATA DESPACHO','DIAS']],hide_index=True,use_container_width=True)
    st.caption('Obras com 10 dias completos passarão à condição de atraso quando atingirem 11 dias. Obras sem data não entram nas faixas.')

def critical(d):
    w=d[d.ELEGIVEL].copy()
    w['RISCO']=w.DIAS.fillna(0).clip(lower=0).astype(float)+w.PRIORITARIA.astype(int)*15+w.SEM_NOME.astype(int)*3+(~w.COORD_OK).astype(int)*3
    columns=['ID SISCO','REGIONAL','MUNICIPIO','LEVANTADOR','STATUS LIST','PRIORIDADE','DATA DESPACHO','DIAS','RISCO']
    st.dataframe(w.sort_values('RISCO',ascending=False)[columns].head(30),hide_index=True,use_container_width=True)
    st.caption('Ranking indicativo: dias desde despacho + 15 para prioridade + 3 para ausência de nome + 3 para coordenada ausente. Não é um SLA contratual.')

def quality(d):
    ids=d.CHAVE
    bad_date=pd.to_datetime(d['DATA DESPACHO'],errors='coerce').gt(pd.Timestamp.now())
    count={'Sem coordenadas':int((~d.COORD_OK).sum()),'Sem NOME':int(d.SEM_NOME.sum()),'IDs ausentes':int(ids.eq('').sum()),'IDs repetidos':int((ids.duplicated(keep=False)&ids.ne('')).sum()),'Despacho futuro':int(bad_date.sum()),'Sem data de despacho':int(pd.to_datetime(d['DATA DESPACHO'],errors='coerce').isna().sum())}
    metric_grid(list(count.items())[:3]);metric_grid(list(count.items())[3:])
    selection=st.selectbox('Investigar inconsistência',list(count))
    tests={'Sem coordenadas':~d.COORD_OK,'Sem NOME':d.SEM_NOME,'IDs ausentes':ids.eq(''),'IDs repetidos':ids.duplicated(keep=False)&ids.ne(''),'Despacho futuro':bad_date,'Sem data de despacho':pd.to_datetime(d['DATA DESPACHO'],errors='coerce').isna()}
    st.dataframe(d.loc[tests[selection],['ID SISCO','REGIONAL','MUNICIPIO','LEVANTADOR','STATUS LIST','DATA DESPACHO','LATITUDE','LONGITUDE']],hide_index=True,use_container_width=True)

def comparison(current,old):
    a=current[current.CHAVE.ne('')].drop_duplicates('CHAVE',keep='last').set_index('CHAVE')
    b=old[old.CHAVE.ne('')].drop_duplicates('CHAVE',keep='last').set_index('CHAVE')
    new=a.index.difference(b.index);gone=b.index.difference(a.index);both=a.index.intersection(b.index)
    changes=[]
    for col in ['STATUS LIST','STATUS SAP','STATUS SISCO','LEVANTADOR','REGIONAL','MUNICIPIO']:
        before=b.loc[both,col].fillna('').astype(str);after=a.loc[both,col].fillna('').astype(str)
        for k in both[before.ne(after)]:changes.append({'ID SISCO':k,'CAMPO':col,'ANTES':before.loc[k],'DEPOIS':after.loc[k]})
    return len(new),len(gone),pd.DataFrame(changes)

def xlsx_bytes(tables):
    f=io.BytesIO()
    with pd.ExcelWriter(f,engine='openpyxl') as w:
        for name,df in tables.items():
            # Excel sheet name max 31 chars, no invalid characters
            safe=re.sub(r'[:\\/?*\[\]]','_',name)[:31]
            df.drop(columns=[c for c in df if c in ['LIST_N','SAP_N','SISCO_N','CONTRATO_N','LEV_N','CHAVE']],errors='ignore').to_excel(w,sheet_name=safe,index=False)
    return f.getvalue()

def pdf_bytes(tables):
    f=io.BytesIO();doc=SimpleDocTemplate(f,pagesize=A4,rightMargin=28,leftMargin=28,topMargin=28,bottomMargin=28)
    styles=getSampleStyleSheet();story=[Paragraph('NIP | Relatório Gerencial de Obras',styles['Title']),Paragraph('Emissão: '+datetime.now().strftime('%d/%m/%Y %H:%M'),styles['Normal']),Spacer(1,16)]
    for title,df in tables.items():
        story.append(Paragraph(title,styles['Heading2']))
        short=df.head(20).copy()
        if short.shape[1]>6:short=short.iloc[:,:6]
        values=[list(map(str,short.columns))]+[[str(v)[:23] for v in r] for r in short.fillna('').itertuples(index=False,name=None)]
        width=535/max(1,len(values[0]));tab=Table(values,colWidths=[width]*len(values[0]),repeatRows=1,hAlign='LEFT')
        tab.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#092d55')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTSIZE',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,-1),7),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.whitesmoke,colors.white])]))
        story.extend([tab,Spacer(1,16)])
    doc.build(story);return f.getvalue()

def main():
    if LOGO_PATH.exists():
        st.sidebar.image(str(LOGO_PATH), use_container_width=True)
    else:
        st.sidebar.markdown('## NIP')
    st.sidebar.caption('🔓 Consulta pública — sem necessidade de senha')
    admin_panel()
    if not GITHUB_TOKEN:
        st.sidebar.info('Consulta pública habilitada. Para publicação pela interface, configure GITHUB_TOKEN nos Secrets.')
    try:
        raw=get_snapshot(1)
        d=prepared(raw)
    except Exception as exc:
        st.error('Não foi possível ler a aba NOTAS do Excel compartilhado.')
        st.warning(str(exc))
        st.info('Adicione o arquivo BASE_LEVANTAMENTO_ATUALIZADA.xlsx na raiz do repositório GitHub, ao lado de app_nip.py.')
        st.stop()
    if st.sidebar.button('Verificar atualização da base'):
        st.rerun()
    st.markdown('<div class="niphero"><h2>GRUPO ÍGNEO / NIP | Gestão Integrada de Obras</h2><small>Dashboard geográfico, distribuição de equipes e acompanhamento de prazos</small></div>',unsafe_allow_html=True)
    st.caption(f'Fonte: aba NOTAS • {st.session_state.get("_excel_source", "Excel") } • {len(d):,} obras • consulta {datetime.now():%d/%m/%Y %H:%M}'.replace(',','.'))
    filtered=filters(d)
    st.sidebar.download_button('Exportar seleção CSV',filtered[FIELDS].to_csv(index=False,sep=';').encode('utf-8-sig'),file_name='obras_filtradas.csv',mime='text/csv')
    tabs=st.tabs(['🗺️ Visão Geral','👷 Levantadores','⏱️ Prazos e Prioridades','🔍 Qualidade da Base','📄 Relatórios'])
    with tabs[0]:
        metric_grid([('Obras totais',len(d)),('Obras filtradas',len(filtered)),('Com coordenadas',int(filtered.COORD_OK.sum())),('Sem coordenadas',int((~filtered.COORD_OK).sum())),('Sem NOME',int(filtered.SEM_NOME.sum()))])
        chart_counts(filtered,'STATUS LIST','Obras por STATUS LIST',20)
        map_view(filtered)
        st.dataframe(filtered[['ID SISCO','REGIONAL','MUNICIPIO','LEVANTADOR','STATUS LIST','STATUS SAP','STATUS SISCO','DATA DESPACHO']],hide_index=True,use_container_width=True)
    with tabs[1]:
        t=load_table(filtered)
        if t.empty:st.info('Nenhum levantador na carteira elegível com os filtros aplicados.')
        else:
            metric_grid([('Recarregar (<10)',int((t.TOTAL<10).sum())),('Faixa normal',int(t.TOTAL.between(10,50).sum())),('Acúmulo (>50)',int((t.TOTAL>50).sum())),('Obras sem levantador',int((filtered.ELEGIVEL&filtered.LEV_N.str.upper().eq('SEM LEVANTADOR')).sum()))])
            st.dataframe(t,hide_index=True,use_container_width=True)
            chart_counts(filtered[filtered.ELEGIVEL],'LEV_N','Carga por levantador',20)
            redistributions(filtered,t)
    with tabs[2]:
        metric_grid([('Carteira elegível',int(filtered.ELEGIVEL.sum())),('Atrasadas >10 dias',int(filtered.ATRASADA.sum())),('Prioritárias atrasadas',int((filtered.ATRASADA&filtered.PRIORITARIA).sum()))])
        aging(filtered)
        st.subheader('Top 10 das obras atrasadas')
        late=filtered[filtered.ATRASADA]
        c1,c2,c3=st.columns(3)
        with c1:chart_counts(late,'REGIONAL','Por regional',10)
        with c2:chart_counts(late,'MUNICIPIO','Por município',10)
        with c3:chart_counts(late,'LEV_N','Por levantador',10)
        st.subheader('Ranking das obras mais críticas');critical(filtered)
    with tabs[3]:
        quality(filtered)
        hist=snapshots()
        if len(hist)>1:
            previous=st.selectbox('Comparar base atual com versão anterior',hist.iloc[1:].apply(lambda x:f"{x['id']} — {x['imported_at']} ({x['filename']})",axis=1).tolist())
            old_id=previous.split(' — ')[0]
            new,gone,changes=comparison(d,prepared(get_snapshot(old_id)))
            metric_grid([('Novas notas',new),('Notas ausentes na versão atual',gone),('Mudanças de campos',len(changes))])
            if len(changes):st.dataframe(changes,hide_index=True,use_container_width=True)
            st.caption('O histórico reflete apenas diferenças entre as bases importadas, não eventos ocorridos entre as importações.')
        else:st.info('Importe uma segunda versão para habilitar a comparação e o histórico de mudanças.')
    with tabs[4]:
        st.write('Relatórios da seleção atual e carteira elegível. Dados são gerados sob demanda.')
        table=load_table(filtered)
        tables={'Resumo':pd.DataFrame([{'Data':date.today().isoformat(),'Obras totais':len(d),'Filtradas':len(filtered),'Elegíveis':int(filtered.ELEGIVEL.sum()),'Em atraso':int(filtered.ATRASADA.sum())}]), 'Carga levantadores':table if not table.empty else pd.DataFrame({'Informação':['Sem equipes']}),'Atrasadas':filtered.loc[filtered.ATRASADA,FIELDS+['DIAS']],'Base filtrada':filtered[FIELDS]}
        a,b=st.columns(2)
        if a.button('Preparar arquivo Excel'):
            st.session_state['xlsx_report']=xlsx_bytes(tables)
        if 'xlsx_report' in st.session_state:a.download_button('Baixar relatório Excel',st.session_state.xlsx_report,file_name='relatorio_nip.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        if b.button('Preparar relatório PDF'):
            st.session_state['pdf_report']=pdf_bytes({k:v for k,v in tables.items() if k!='Base filtrada'})
        if 'pdf_report' in st.session_state:b.download_button('Baixar relatório PDF',st.session_state.pdf_report,file_name='relatorio_nip.pdf',mime='application/pdf')
        st.caption('A geração de PDF resume as primeiras 20 linhas por seção; o Excel inclui a seleção completa.')

if __name__=='__main__': main()
