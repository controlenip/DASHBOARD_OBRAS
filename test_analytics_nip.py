"""Regressao de integracao sem precisar conectar ao GitHub/Streamlit Cloud."""
from pathlib import Path
import importlib.util,sys,types,json,re,subprocess,tempfile,collections
ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT))
from nip_analytics import build_payload,parse_date,number,validate_unique
src=ROOT/'BASE_LIST(1).xlsx'
assert src.exists(),'BASE_LIST de referencia nao encontrada'
data=json.loads(build_payload(src.read_bytes()))
rows=data['rows'];checks=data['validation']
assert len(rows)==7458,len(rows)
assert sum(r['correcao'] for r in rows)==187
assert sum(r['correcao'] and not r['obs'] for r in rows)==39
assert sum(bool(r['u']) for r in rows)==3520
assert checks['notas_duplicadas']>0 and checks['linhas_duplicadas']>checks['notas_duplicadas']
assert checks['sem_q']>0 and checks['sem_u']>0
assert parse_date('')=='' and number(None) is None and number('12')==12
safe=[r for r in rows if r['nota'] and not r['duplicada']]
assert len({r['nota'] for r in safe})==len(safe)
assert any(r['q'] and not r['u'] for r in safe)
assert any(r['u'] and not r['q'] for r in safe)
html=(ROOT/'nip_dashboard_template.html').read_text(encoding='utf8');app=(ROOT/'app_nip.py').read_text(encoding='utf8')
for marker in ('__NIP_DATA_PAYLOAD__','__NIP_LIST_PAYLOAD__','__NIP_LIST_STATUS__','__NIP_PRODUCTIVITY_PAYLOAD__','__NIP_PRODUCTIVITY_STATUS__','__NIP_ANALYTICS_PAYLOAD__'):
 assert html.count(marker)==1,(marker,html.count(marker))
for term in ('nip-tab-qualidade-list','nip-tab-produtividade','advQuality','advProduction','advQualityMap','advCalendar','advPareto','advMatrix','advValidCards','advAnalyst','advGoals','advEvolution'):
 assert term in html,term
assert app.count('build_payload(list_blob)')==1
assert 'BASE_LEVANTAMENTO_ATUALIZADA.xlsx' in app and 'BASE_LIST.xlsx' in app
assert "sheet_name='NOTAS'" in app
for i,block in enumerate(re.findall(r'<script(?:\s[^>]*)?>(.*?)</script>',html,re.S|re.I)):
 for key,value in [('__NIP_DATA_PAYLOAD__','[]'),('__NIP_LIST_PAYLOAD__','[]'),('__NIP_LIST_STATUS__','"ok"'),('__NIP_PRODUCTIVITY_PAYLOAD__','[]'),('__NIP_PRODUCTIVITY_STATUS__','"ok"'),('__NIP_ANALYTICS_PAYLOAD__','null')]:block=block.replace(key,value)
 with tempfile.NamedTemporaryFile(mode='w',suffix='.js') as f:
  f.write(block);f.flush();p=subprocess.run(['node','--check',f.name],capture_output=True,text=True)
  assert p.returncode==0,(i,p.stderr)
import py_compile
for name in ('app_nip.py','nip_analytics.py'):py_compile.compile(str(ROOT/name),doraise=True)
print('PASS:',len(rows),'linhas;',sum(r['correcao'] for r in rows),'em correcao;',checks['notas_duplicadas'],'notas duplicadas;',len(safe),'notas unicas elegiveis;',sum(bool(r['u']) for r in rows),'datas U;',len(re.findall(r'<script(?:\s[^>]*)?>',html,re.I)),'scripts validados.')
