"""Standalone tests for BASE_LIST quality integration, no Streamlit server needed."""
import sys, types, importlib.util, pathlib, json, re
ROOT=pathlib.Path(__file__).resolve().parent
class StreamlitFake(types.ModuleType):
    def __getattr__(self,k):
        if k=='cache_data': return lambda *a,**kw: (lambda fn:fn)
        if k=='secrets':return {}
        if k=='markdown':return lambda *a,**kw:None
        if k=='set_page_config':return lambda *a,**kw:None
        raise AttributeError(k)
st=StreamlitFake('streamlit'); sys.modules['streamlit']=st
parts=types.ModuleType('streamlit.components');v1=types.ModuleType('streamlit.components.v1')
sys.modules['streamlit.components']=parts;sys.modules['streamlit.components.v1']=v1
st.components=parts;parts.v1=v1
spec=importlib.util.spec_from_file_location('app_qualidade_test',ROOT/'app_nip.py')
module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
blob=(ROOT/'BASE_LIST(1).xlsx').read_bytes()
payload,total,correction=module.parse_list_quality(blob)
records=json.loads(payload)
assert correction==187,(total,correction)
assert len(records)==187
assert all(r['levantador'] and isinstance(r['observacoes'],str) for r in records)
assert sum(bool(r['observacoes']) for r in records)==148
assert len({r['nota'] for r in records})==187
assert all('analista' in r for r in records)
missing=[r for r in records if not r['observacoes']]
assert len(missing)==39
assert sum(r['analista']=='(Sem analista informado)' for r in missing)==10
from collections import Counter as QualityCounter
assert QualityCounter(r['analista'] for r in missing)['Jamille Wilena Brito Lago']==14
assert QualityCounter(r['analista'] for r in records)['Paulo Pacheco Moraes']==50
from collections import Counter
ranking=Counter(r['levantador'] for r in records)
assert ranking.most_common(1)[0][1]==46
assert len(ranking)>5
html=(ROOT/'nip_dashboard_template.html').read_text()
assert html.count('__NIP_LIST_PAYLOAD__')==1
assert html.count('__NIP_LIST_STATUS__')==1
assert 'Qualidade Levantamento' in html
assert 'nip-tab-qualidade-list' in html
assert 'qlAnalystRanking' in html and 'qlAnalystMissingRows' in html
assert "LIST_NAME = 'BASE_LIST.xlsx'" in (ROOT/'app_nip.py').read_text()
assert "EXCEL_NAME = 'BASE_LEVANTAMENTO_ATUALIZADA.xlsx'" in (ROOT/'app_nip.py').read_text()
assert 'Comparação entre bases' not in html
print('OK BASE_LIST:',total,'registros,',correction,'correcoes,',sum(bool(r['observacoes']) for r in records),'observacoes, top 5:',ranking.most_common(5))
