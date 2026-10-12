from pathlib import Path
import re, subprocess
h=Path('/mnt/data/nip_dashboard_template.html').read_text()
assert h.count('id="nip-card-events-layout-fix"')==1
assert "lev.insertBefore(workload,lev.firstElementChild)" in h
assert 'card.addEventListener(\'click\'' in h and 'card.addEventListener(\'keydown\'' in h
assert "currentRows=findRows(conf[1])" in h and "CURRENT.filter(r=>!!staffLoadStatus(r))" in h
assert "window.nipCrossBridge.publish(currentLabel,currentRows,'notas')" in h
assert '#nipCardWorkDetail' in h and 'text-align:left!important' in h
for card in ['wlTotal','wlZero','wlCorr','wlAndamento','wlLow','wlOk','wlHigh','wlUnassigned']:
 assert re.search(r'\b'+card+r':\[',h),card
for i,js in enumerate(re.findall(r'<script(?:\s[^>]*)?>(.*?)</script>',h,re.S)):
 if not js.strip(): continue
 js=re.sub(r'__NIP_[A-Z_]+__','null',js)
 r=subprocess.run(['node','--check'],input=js,text=True,capture_output=True)
 assert r.returncode==0,(i,r.stderr)
assert 'GitHub' not in h
print('PASS: 8 cards vinculados diretamente; tabela no painel; atalhos para Relatorios; carga no topo; alinhamento; 14 scripts JS')
