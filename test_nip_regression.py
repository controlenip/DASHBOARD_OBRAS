import ast,io,re,json
from pathlib import Path
import pandas as pd
root=Path(__file__).parent
src=(root/'app_nip.py').read_text()
module=ast.parse(src)
fun=next(n for n in module.body if isinstance(n,ast.FunctionDef) and n.name=='parse_notas')
fun.decorator_list=[]
code=compile(ast.Module(body=[fun],type_ignores=[]),'<parse_notas>','exec')
ns={'pd':pd,'io':io,'REQUIRED':['ID SISCO','STATUS LIST','STATUS SAP','STATUS SISCO','REGIONAL','MUNICIPIO','LEVANTADOR','DATA DESPACHO']}
exec(code,ns)
blob=(root/'BASE_LEVANTAMENTO_ATUALIZADA.xlsx').read_bytes()
payload,n=ns['parse_notas'](blob)
rows=json.loads(payload)
assert n==len(rows)==11892,(n,len(rows))
assert all(k in rows[0] for k in ['PRIORIDADE','NOME','STATUS SAP','STATUS SISCO','STATUS LIST','CONTRATO','SGO'])
assert sum(1 for r in rows if r.get('LATITUDE') is not None and r.get('LONGITUDE') is not None)>0
from collections import Counter
cnt=Counter(str(r['STATUS LIST']).strip() for r in rows)
assert cnt['0']==7564,cnt
assert cnt['Em levantamento']==1887,cnt
html=(root/'nip_dashboard_template.html').read_text()
assert html.count('__NIP_DATA_PAYLOAD__')==1
assert html.count('id="nip-tab-')==0  # panels are dynamically generated, not hardcoded
for token in ['nipQuickBtns','nipLoadBody','nipCityBody','nipAdvancedBody','nipFullMap','nipEvolutionCards','nipQualityExtraStats','nipExecutivePrint','nipPreviousFile','nipXlsx','nipPdf']:
 assert token in html,token
print('PASS: aba NOTAS, 11892 linhas, campos antigos/novos, status original, cinco abas dinâmicas, controles de 10 melhorias')
