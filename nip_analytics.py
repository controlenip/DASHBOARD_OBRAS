"""Analises adicionais da BASE_LIST; nenhuma escrita na fonte oficial."""
import io,json,re,unicodedata
from collections import Counter
from datetime import datetime
import pandas as pd

def norm(v):
    return ''.join(c for c in unicodedata.normalize('NFKD',str(v or '')) if not unicodedata.combining(c)).upper().strip()

def parse_date(v):
    if pd.isna(v) or str(v).strip()=='':return ''
    try:
        dt=pd.to_datetime(v,dayfirst=True,errors='coerce')
        return dt.strftime('%Y-%m-%d') if pd.notna(dt) and 2000<=dt.year<=2100 else ''
    except (ValueError,TypeError):return ''

def number(v):
    if pd.isna(v) or str(v).strip()=='':return None
    try:
        n=float(str(v).strip().replace(',','.'))
        return int(n) if n.is_integer() and 0<=n<=100000 else None
    except (ValueError,TypeError):return None

def coordinate(v,limit):
    try:
        n=float(str(v).strip().replace(',','.'))
        return n if abs(n)<=limit else None
    except (ValueError,TypeError):return None

def records_from_blob(blob):
    df=pd.read_excel(io.BytesIO(blob),sheet_name=0,engine='openpyxl',dtype=object)
    if len(df.columns)<32:raise ValueError('BASE_LIST precisa ter ao menos 32 colunas.')
    expected={4:'STATUS DO PROJETO',7:'MUNICIPIO',14:'LEVANTADOR',16:'DATA DO LEVANTAMENTO',19:'ANALISTA',20:'DATA DA ANALISE',31:'OBSERVACOES'}
    bad=[str(df.columns[i]) for i,e in expected.items() if e not in norm(df.columns[i])]
    if bad:raise ValueError('Cabeçalhos inesperados na BASE_LIST: '+', '.join(bad))
    def txt(v):return '' if pd.isna(v) or str(v).strip().lower() in ('nan','none','null') else str(v).strip()
    out=[]
    for i,r in enumerate(df.itertuples(index=False,name=None)):
        row={'linha':i+2,'nota':txt(r[0]),'status':txt(r[4]),'regional':txt(r[6]),'municipio':txt(r[7]),'levantador':txt(r[14]),
             'analista':txt(r[19]),'q':parse_date(r[16]),'u':parse_date(r[20]),'existentes':number(r[17]),'novos':number(r[18]),
             'obs':txt(r[31]),'lat':coordinate(r[10],90),'lon':coordinate(r[11],180)}
        row['correcao']=norm(row['status']) in ('CORRECAO DE LEVANTAMENTO','EM CORRECAO DE LEVANTAMENTO')
        out.append(row)
    return out

def validate_unique(rows):
    c=Counter(r['nota'] for r in rows if r['nota']); duplicates={k for k,v in c.items() if v>1}
    for r in rows:r['duplicada']=bool(r['nota'] in duplicates)
    return {'linhas':len(rows),'notas_duplicadas':len(duplicates),'linhas_duplicadas':sum(r['duplicada'] for r in rows),
      'sem_nota':sum(not r['nota'] for r in rows),'sem_levantador':sum(not r['levantador'] for r in rows),
      'sem_q':sum(not r['q'] for r in rows),'sem_u':sum(not r['u'] for r in rows),
      'sem_postes_existentes':sum(r['existentes'] is None for r in rows),'sem_postes_novos':sum(r['novos'] is None for r in rows),
      'correcao_sem_af':sum(r['correcao'] and not r['obs'] for r in rows)}

def build_payload(blob):
    rows=records_from_blob(blob)
    warnings=validate_unique(rows)
    return json.dumps({'rows':rows,'validation':warnings},ensure_ascii=False,separators=(',',':'))
