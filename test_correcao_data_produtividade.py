import json, pathlib, sys, importlib.util, types, re, subprocess, tempfile
P=pathlib.Path(__file__).parent
class Fake(types.ModuleType):
 def __getattr__(self,k):
  if k=='cache_data':return lambda *a,**kw:lambda f:f
  if k=='secrets':return {}
  if k in ('markdown','set_page_config'):return lambda *a,**kw:None
  raise AttributeError(k)
st=Fake('streamlit');sys.modules['streamlit']=st
p=types.ModuleType('streamlit.components');v=types.ModuleType('streamlit.components.v1');sys.modules['streamlit.components']=p;sys.modules['streamlit.components.v1']=v;st.components=p;p.v1=v
spec=importlib.util.spec_from_file_location('nip_app',P/'app_nip.py');m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
rows=json.loads(m.parse_list_productivity((P/'BASE_LIST(1).xlsx').read_bytes())[0]); name='Raimundo Jhones Assunção da Luz'
q=[r for r in rows if r['levantador']==name and r['data_levantamento']=='2026-10-08']
assert len(q)==4,(len(q),q)
assert sum(r['postes_existentes'] for r in q)==13
assert sum(r['postes_novos'] for r in q)==2
assert sum(not r['data'] for r in q)==1
assert len([r for r in q if r['data']=='2026-10-09'])==3
assert len([r for r in rows if r['levantador']==name and r['data']=='2026-10-08'])==2
s=(P/'nip_dashboard_template.html').read_text();assert '<select id="pdBase">' in s and 'dateOf(r)' in s and "validRows()" in s
for i,b in enumerate(re.findall(r'<script(?:\s[^>]*)?>(.*?)</script>',s,re.S|re.I)):
 for key,val in [('__NIP_DATA_PAYLOAD__','[]'),('__NIP_LIST_PAYLOAD__','[]'),('__NIP_LIST_STATUS__','"ok"'),('__NIP_PRODUCTIVITY_PAYLOAD__','[]'),('__NIP_PRODUCTIVITY_STATUS__','"ok"'),('__NIP_ANALYTICS_PAYLOAD__','null')]: b=b.replace(key,val)
 with tempfile.NamedTemporaryFile(suffix='.js',mode='w') as f:
  f.write(b);f.flush();p=subprocess.run(['node','--check',f.name],capture_output=True,text=True);assert p.returncode==0,(i,p.stderr)
print('PASS: Raimundo 08/10 Q = 4 obras, R = 13, S = 2; 3 com U em 09/10 e 1 sem U; scripts JS válidos')
