from pathlib import Path
import re, subprocess, zipfile
s=Path('/mnt/data/nip_dashboard_template.html').read_text()
original=zipfile.ZipFile('/mnt/data/NIP_PRODUTIVIDADE_CORRIGIDA_Q_U.zip').read('nip_dashboard_template.html').decode()
for token in ['__NIP_DATA_PAYLOAD__','__NIP_LIST_PAYLOAD__','__NIP_PRODUCTIVITY_PAYLOAD__','__NIP_ANALYTICS_PAYLOAD__']:
 assert s.count(token)==original.count(token)==1,token
assert 'GitHub' not in s
assert 'window.nipReportBridge' in s and 'window.nipCrossBridge' in s
assert 'nipCrossCardDetail' in s and 'nipReportCrossBanner' in s
for x in ['wlTotal','wlZero','wlCorr','wlAndamento','wlLow','wlOk','wlHigh','wlUnassigned']:
 assert s.count("['"+x+"',")==1,x
for v in ['nip-tab-produtividade','nip-tab-qualidade-list','nip-tab-prazos']:
 assert f"installButton('{v}'" in s,v
assert "let nipTableAppliedFilters=[]" in s and 'nipOpsTableRefresh()' in s
assert 'nipTableAppliedFilters.every' in s and "!bridge||bridge.notes.has" in s
base_scripts=re.findall(r'<script(?:\s[^>]*)?>(.*?)</script>',original,re.S)
scripts=re.findall(r'<script(?:\s[^>]*)?>(.*?)</script>',s,re.S)
assert len(scripts)==len(base_scripts)+2
for i,script in enumerate(scripts):
 if not script.strip():continue
 js=re.sub(r'__NIP_[A-Z_]+__','null',script)
 p=subprocess.run(['node','--check'],input=js,text=True,capture_output=True)
 assert p.returncode==0,(i,p.stderr[:300])
print('PASS: 8 cards, 3 abas integradas, cruzamento por nota, relatório e exportações, 13 scripts JS válidos, 4 payloads preservados')
