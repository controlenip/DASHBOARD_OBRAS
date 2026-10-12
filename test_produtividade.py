import sys,types,importlib.util,pathlib,json,collections,re,subprocess,tempfile
ROOT=pathlib.Path(__file__).resolve().parent
class Fake(types.ModuleType):
 def __getattr__(self,k):
  if k=='cache_data':return lambda *a,**kw:lambda f:f
  if k=='secrets':return {}
  if k in ('markdown','set_page_config'):return lambda *a,**kw:None
  raise AttributeError(k)
st=Fake('streamlit');sys.modules['streamlit']=st
p=types.ModuleType('streamlit.components');v=types.ModuleType('streamlit.components.v1');sys.modules['streamlit.components']=p;sys.modules['streamlit.components.v1']=v;st.components=p;p.v1=v
spec=importlib.util.spec_from_file_location('nip_app',ROOT/'app_nip.py');mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)
blob=(ROOT/'BASE_LIST(1).xlsx').read_bytes();payload,total,completed=mod.parse_list_productivity(blob);records=json.loads(payload)
assert total==7458 and completed==len(records)>0
assert all(re.fullmatch(r'\d{4}-\d{2}-\d{2}',r['data']) for r in records)
assert all(isinstance(r['postes_novos'],int) and r['postes_novos']>=0 for r in records)
assert all(isinstance(r['postes_existentes'],int) and r['postes_existentes']>=0 for r in records)
assert all(r['levantador'] and r['municipio'] for r in records)
html=(ROOT/'nip_dashboard_template.html').read_text();py=(ROOT/'app_nip.py').read_text()
for marker in ('__NIP_DATA_PAYLOAD__','__NIP_LIST_PAYLOAD__','__NIP_LIST_STATUS__','__NIP_PRODUCTIVITY_PAYLOAD__','__NIP_PRODUCTIVITY_STATUS__'):assert html.count(marker)==1,marker
for part in ('nip-tab-produtividade','pdTop','pdCities','pdPredict','pdCsv','nip-tab-qualidade-list'):assert part in html
assert 'BASE_LEVANTAMENTO_ATUALIZADA.xlsx' in py and 'BASE_LIST.xlsx' in py
for i,b in enumerate(re.findall(r'<script(?:\s[^>]*)?>(.*?)</script>',html,re.S|re.I)):
 for key,value in [('__NIP_DATA_PAYLOAD__','[]'),('__NIP_LIST_PAYLOAD__','[]'),('__NIP_LIST_STATUS__','"ok"'),('__NIP_PRODUCTIVITY_PAYLOAD__','[]'),('__NIP_PRODUCTIVITY_STATUS__','"ok"')]:b=b.replace(key,value)
 with tempfile.NamedTemporaryFile(suffix='.js',mode='w') as f:
  f.write(b);f.flush();p=subprocess.run(['node','--check',f.name],capture_output=True,text=True);assert p.returncode==0,(i,p.stderr)
print('PASS produtividade',total,'linhas;',completed,'com data U;',sum(r['postes_novos'] for r in records),'postes novos;',len(set(r['levantador'] for r in records)),'levantadores')
print('Top 10 obras:',collections.Counter(r['levantador'] for r in records).most_common(10))
