"""NIP IA: consultas somente leitura e ações visuais, sem escrita nas bases."""
import json, math, re
from datetime import datetime, timezone
import pandas as pd

def _text(v):
    return '' if pd.isna(v) else str(v).strip()

def _km(a,b,c,d):
    rad=math.pi/180
    h=math.sin((c-a)*rad/2)**2+math.cos(a*rad)*math.cos(c*rad)*math.sin((d-b)*rad/2)**2
    return 12742*math.asin(min(1,math.sqrt(h)))

def query_data(df, people, op, args):
    """Return bounded facts and optional map action. All filtering happens server-side."""
    if op=='find_note':
        q=str(args.get('note','')).strip().casefold()
        if not q or len(q)>100:return {'error':'Informe um identificador válido.'}
        keys=['ID SISCO','PROTOCOLO','SGO','Nº SGO','N° SGO','NÚMERO SGO']
        match=df.loc[df.apply(lambda r:any(q==_text(r[k]).casefold() for k in keys if k in df),axis=1)]
        if match.empty:return {'count':0,'message':'Nota não localizada.'}
        cols=[c for c in ['ID SISCO','PROTOCOLO','REGIONAL','MUNICIPIO','LEVANTADOR','STATUS SAP','STATUS SISCO','STATUS LIST','DATA DESPACHO','LATITUDE','LONGITUDE'] if c in df]
        return {'count':len(match),'rows':json.loads(match[cols].head(5).to_json(orient='records',date_format='iso',force_ascii=False)),'action':{'type':'note','value':q}}
    if op=='summary':
        col=str(args.get('group_by','REGIONAL'))
        if col not in ('REGIONAL','MUNICIPIO','LEVANTADOR','STATUS SAP','STATUS SISCO','STATUS LIST','TIPO NOTA','CONTRATO'):return {'error':'Agrupamento não permitido.'}
        regional=str(args.get('regional','')).strip()
        subset=df[df['REGIONAL'].astype(str).str.casefold()==regional.casefold()] if regional and 'REGIONAL' in df else df
        counts=subset[col].fillna('(Vazio)').astype(str).value_counts().head(25)
        return {'total':len(subset),'group_by':col,'counts':counts.to_dict(),'regional':regional or 'Todas'}
    if op=='pending':
        subset=df.copy()
        if 'STATUS SAP' in subset:subset=subset[~subset['STATUS SAP'].astype(str).str.upper().isin(['FINL','CANC'])]
        if 'DATA DESPACHO' not in subset:return {'error':'Sem coluna DATA DESPACHO.'}
        days=pd.to_datetime(subset['DATA DESPACHO'],errors='coerce',utc=True)
        limit=min(365,max(0,int(args.get('days',10))))
        subset=subset.loc[days.notna() & ((pd.Timestamp.now(tz='UTC')-days).dt.days>limit)]
        cols=[c for c in ['ID SISCO','PROTOCOLO','MUNICIPIO','REGIONAL','LEVANTADOR','STATUS LIST','DATA DESPACHO'] if c in subset]
        return {'count':len(subset),'days':limit,'rows':json.loads(subset[cols].head(20).to_json(orient='records',date_format='iso',force_ascii=False)),'note':'Contagem indicativa por data de despacho; não equivale a dias sem movimentação.'}
    if op=='nearest':
        note=str(args.get('note','')).strip()
        found=query_data(df,people,'find_note',{'note':note})
        if not found.get('rows'):return found
        r=found['rows'][0]
        try:a,b=float(r['LATITUDE']),float(r['LONGITUDE'])
        except (TypeError,ValueError,KeyError):return {'error':'Obra sem coordenadas válidas.'}
        if not (-90<=a<=90 and -180<=b<=180 and a!=0 and b!=0):return {'error':'Coordenadas da obra inválidas.'}
        nearby=sorted([{'nome':p['nome'],'equipe':p['equipe'],'cidade':p.get('cidade',''),'km_linha_reta':round(_km(a,b,p['lat'],p['lon']),2)} for p in people],key=lambda p:p['km_linha_reta'])[:5]
        return {'note':note,'assigned_to':r.get('LEVANTADOR',''),'nearby_residences':nearby,'action':{'type':'note','value':note},'note':'Distância residencial em linha reta; não altera o responsável atribuído.'}
    if op=='map_filter':
        field=str(args.get('field','')).upper()
        if field not in ('REGIONAL','MUNICIPIO','LEVANTADOR','STATUS SAP','STATUS SISCO','STATUS LIST','TIPO NOTA','CONTRATO'):return {'error':'Filtro não permitido.'}
        value=str(args.get('value','')).strip()
        if not value or len(value)>120:return {'error':'Valor de filtro inválido.'}
        if field not in df:return {'error':'Coluna indisponível.'}
        count=int((df[field].fillna('').astype(str).str.casefold()==value.casefold()).sum())
        return {'matches':count,'action':{'type':'filter','field':field,'value':value} if count else None}
    if op=='clear_map':return {'action':{'type':'clear'},'message':'Limpeza dos filtros visuais solicitada.'}
    return {'error':'Operação desconhecida.'}

def ask(question, df, people, api_key, model='gpt-4.1-mini'):
    from openai import OpenAI
    client=OpenAI(api_key=api_key,timeout=35,max_retries=1)
    schema={'type':'object','additionalProperties':False,'properties':{'op':{'type':'string','enum':['find_note','summary','pending','nearest','map_filter','clear_map']},'args':{'type':'object','additionalProperties':{'type':['string','integer','number','boolean']}}},'required':['op','args']}
    planner=client.chat.completions.create(model=model,temperature=0,messages=[{'role':'system','content':'Você é o roteador de consultas NIP. Converta a pergunta em UMA operação e argumentos JSON. Consultar nota=find_note; contagens=summary; atraso desde despacho=pending; quem mora mais perto=nearest; filtrar mapa=map_filter; limpar mapa=clear_map. Não invente dados. Colunas aceitas: REGIONAL, MUNICIPIO, LEVANTADOR, STATUS SAP, STATUS SISCO, STATUS LIST, TIPO NOTA, CONTRATO. Para summary, group_by; para find_note e nearest, note; para pending, days; para map_filter, field e value.'},{'role':'user','content':question}],response_format={'type':'json_schema','json_schema':{'name':'nip_query','strict':True,'schema':schema}})
    plan=json.loads(planner.choices[0].message.content)
    result=query_data(df,people,plan['op'],plan['args'])
    answer=client.chat.completions.create(model=model,temperature=0,messages=[{'role':'system','content':'Você é o NIP IA, assistente operacional. Responda em português do Brasil com base APENAS no JSON da consulta. Não invente valores. Não confunda residência com atribuição de obra. Diga quando a distância é em linha reta. Se houver ação visual, explique que será aplicada apenas após confirmação do usuário. Seja breve e claro.'},{'role':'user','content':'Pergunta: '+question+'\nResultado da consulta: '+json.dumps(result,ensure_ascii=False,default=str)[:18000]}])
    return answer.choices[0].message.content,result.get('action')
