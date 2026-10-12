from pathlib import Path
import ast, re, subprocess, tempfile
root=Path(__file__).resolve().parent
app=(root/'app_nip.py').read_text(encoding='utf-8')
html=(root/'nip_dashboard_template.html').read_text(encoding='utf-8')
ast.parse(app)
assert not re.search(r'github|reposit.rio',html,re.I)
assert "$('source').textContent=DATA.length.toLocaleString('pt-BR')+' registros disponíveis'" in html
assert 'Fonte: {source}' not in app
assert "source+' •'" not in app
assert 'Github.' not in app
assert 'list_status=list_source' not in app
assert 'prod_total:,} registros' in app
assert "st.caption(f'{count:,} registros disponíveis" in app
for i,script in enumerate(re.findall(r'<script(?:\s[^>]*)?>(.*?)</script>',html,re.S|re.I)):
    for key,value in [('__NIP_DATA_PAYLOAD__','[]'),('__NIP_LIST_PAYLOAD__','[]'),('__NIP_LIST_STATUS__','"ok"'),('__NIP_PRODUCTIVITY_PAYLOAD__','[]'),('__NIP_PRODUCTIVITY_STATUS__','"ok"'),('__NIP_ANALYTICS_PAYLOAD__','null')]:script=script.replace(key,value)
    with tempfile.NamedTemporaryFile(mode='w',suffix='.js') as f:
        f.write(script);f.flush();p=subprocess.run(['node','--check',f.name],capture_output=True,text=True)
        assert p.returncode==0,(i,p.stderr)
print('PASS: sem mencoes da origem na interface HTML; mensagens Streamlit neutralizadas; JavaScript valido.')
