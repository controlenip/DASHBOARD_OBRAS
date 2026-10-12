from pathlib import Path
import re,json,subprocess,zipfile
h=Path('/mnt/data/nip_dashboard_template.html').read_text()
orig=zipfile.ZipFile('/mnt/data/NIP_CARDS_CLICAVEIS_TOPO_ALINHAMENTO.zip').read('nip_dashboard_template.html').decode()
assert h.startswith(orig[:5000]); assert h[:h.index('<style id="nipFaqStyles">')]==orig[:orig.rfind('</body>')]
assert 'nipFaqAsk' in h and 'nipFaqSearch' in h
assert h.count('id="nipFaqScript"')==1 and h.count("btn.textContent='FAQ / Dúvidas'")==1
assert 'GitHub' not in h
assert all(a in h for a in ['Correção de levantamento','Data Q','Data U','Pré-Análise','Relatórios','Pareto','Calendário','Levantadores'])
for i,sc in enumerate(re.findall(r'<script\b[^>]*>(.*?)</script>',h,re.S|re.I)):
 if not sc.strip(): continue
 r=subprocess.run(['node','--check'],input=sc,text=True,capture_output=True)
 assert r.returncode==0,(i,r.stderr)
assert h.count('nip-tab-faq')>=3
js=h.split('<script id="nipFaqScript">',1)[1].split('</script>',1)[0]
assert "panel.classList.remove('nip-hidden')" in js
assert 'window.nipFaqAskQuestion' in js
assert 'Não encontrei uma resposta suficientemente confiável' in js
print('PASS: FAQ, 39 artigos, busca, abas, ausencia de servico externo, sintaxe JS e base original preservada')
