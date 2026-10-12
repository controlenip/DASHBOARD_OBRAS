import ast,io,re,sys,types,zipfile
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
ROOT=Path(__file__).parent
app=(ROOT/'app_nip.py').read_text()
html=(ROOT/'nip_dashboard_template.html').read_text()
assert ast.parse(app)
assert html.count('__NIP_DATA_PAYLOAD__')==1
assert len(re.findall(r'id="nipExt(?:Presets|Kml|Kmz|Pdf|Summary|ColumnChecks)"',html))==6
assert all(x in html for x in ['nipExtOperationalRefresh','nipReportRows','nipMarkerCache','nipOpsPreviewRefresh','nipExtSavePreset'])
assert 'comparação de bases' not in html.lower()
assert 'const nipTableColumns' in html
assert 'staffLoadStatus' in html
assert 'nipStatusColorLegend' in html
assert 'nipOpsLayerRefresh' in html
assert 'BASE_LEVANTAMENTO_ATUALIZADA.xlsx' in app
assert "sheet_name='NOTAS'" in app
# Execute only safe backend functions from AST without starting Streamlit.
tree=ast.parse(app)
funcs={n.name:n for n in tree.body if isinstance(n,ast.FunctionDef)}
selected=[funcs['validate_workbook'],funcs['parse_notas']]
for f in selected:f.decorator_list=[]
code=compile(ast.Module(body=selected,type_ignores=[]),'app_nip.py','exec')
ns={'pd':pd,'io':io,'zipfile':zipfile,'REQUIRED':['ID SISCO','STATUS LIST','STATUS SAP','STATUS SISCO','REGIONAL','MUNICIPIO','LEVANTADOR','DATA DESPACHO']}
exec(code,ns)
w=Workbook();ws=w.active;ws.title='NOTAS';ws.append(ns['REQUIRED']+['LATITUDE','LONGITUDE','ORÇAMENTO MODULAR (NÃO ENVIAR A CAMPO)']);ws.append([1,'0','ATIV','0','NORTE','SAO LUIS','ANA','2026-10-01',-2.5,-44.2,0]);ws.append([2,'Em levantamento','CANC','0','SUL','BALSAS','BIA','2026-10-02',-7.5,-46.0,1]);b=io.BytesIO();w.save(b);blob=b.getvalue()
v=ns['validate_workbook'](blob);assert not v['fatal'] and v['rows']==2,v
payload,n=ns['parse_notas'](blob);assert n==2 and '"STATUS SAP":"CANC"' in payload
w2=Workbook();w2.active.title='INVALIDA';b2=io.BytesIO();w2.save(b2)
assert ns['validate_workbook'](b2.getvalue())['fatal']
print('PASS: sintaxe backend, template, componentes, integridade, validação NOTAS e leitura de 2 registros.')
