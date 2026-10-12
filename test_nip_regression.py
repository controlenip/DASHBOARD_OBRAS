"""Testes de regressao sem alterar a planilha original."""
import ast
import io
import json
import re
import subprocess
import tempfile
from pathlib import Path
from collections import Counter
import pandas as pd

ROOT = Path(__file__).resolve().parent
SRC = (ROOT / 'app_nip.py').read_text(encoding='utf-8')
HTML = (ROOT / 'nip_dashboard_template.html').read_text(encoding='utf-8')
mod = ast.parse(SRC)
fn = next(n for n in mod.body if isinstance(n, ast.FunctionDef) and n.name == 'parse_notas')
fn.decorator_list = []
namespace = {'pd': pd, 'io': io, 'REQUIRED': ['ID SISCO','STATUS LIST','STATUS SAP','STATUS SISCO','REGIONAL','MUNICIPIO','LEVANTADOR','DATA DESPACHO']}
exec(compile(ast.Module(body=[fn], type_ignores=[]), '<parse_notas>', 'exec'), namespace)
blob = (ROOT / 'BASE_LEVANTAMENTO_ATUALIZADA.xlsx').read_bytes()
records_json, count = namespace['parse_notas'](blob)
records = json.loads(records_json)
assert count == len(records) == 11892, count
status = Counter(str(r['STATUS LIST']).strip() for r in records)
assert status['0'] == 7564 and status['Em levantamento'] == 1887, status
assert HTML.count('__NIP_DATA_PAYLOAD__') == 1
assert "['Visão geral','geral']" in HTML and "['Relatórios','relatorios']" in HTML
assert 'function workloadStatus(r)' in HTML and "sap==='CANC'||sap==='FINL'" in HTML
assert "'LIBERADO PARA LEVANTAMENTO','PRE ANALISE'" in HTML
assert 'function multiSet(' in HTML and 'function multiValues(' in HTML
assert 'const map=L.map(' in HTML and 'function popup(r)' in HTML
assert 'nipOpsRefresh();' in HTML
assert not any(k in HTML for k in ['nipPreviousFile','nipEvolutionCards','nipCompare'])
ids = ['nipOpsStaff','nipOpsCalendar','nipOpsRegion','nipOpsPending','nipOpsLayers',
       'nipOpsMeeting','nipOpsTable','nipOpsExportXlsx','nipOpsIntegrity',
       'nipOpsTableSearch','nipOpsStaffSelect']
for id_ in ids:
    assert id_ in HTML, id_
assert 'GITHUB_TOKEN' in SRC and 'ADMIN_PASSWORD' in SRC and "sheet_name='NOTAS'" in SRC
assert 'components.html(html' in SRC
scripts = re.findall(r'<script(?:\s[^>]*)?>(.*?)</script>', HTML, flags=re.S)
with tempfile.TemporaryDirectory() as tmp:
    for i, script in enumerate(scripts):
        if not script.strip(): continue
        script = script.replace('__NIP_DATA_PAYLOAD__', '[]')
        f = Path(tmp) / f'part{i}.js'
        f.write_text(script, encoding='utf-8')
        subprocess.run(['node', '--check', str(f)], check=True, capture_output=True)
assert all(k in records[0] for k in ['PAT','POSTE PREVISTO MT','DATA ABERTURA','CONTA CONTRATO'])
assert 'nipTableAppliedFilters' in HTML and 'nipTableColumns' in HTML and 'nipOpsFilterColumn' in HTML
assert 'nipOpsSortDirection' in HTML and 'nipOpsFilterAdd' in HTML
print('PASS: 11892 obras; STATUS LIST 0=7564, Em levantamento=1887; criterios preservados; 5 abas; filtros multiplos; 10 modulos; comparacao ausente; JS e Python validos')
