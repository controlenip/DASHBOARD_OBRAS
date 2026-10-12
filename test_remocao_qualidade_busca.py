from pathlib import Path
import json,re,subprocess,tempfile,os
s=Path('/mnt/data/nip_dashboard_template.html').read_text()
assert "x[1]==='qualidade'?'':" in s
assert 'if(i===3)i=4' in s
assert '#nip-tab-qualidade,#nipFind{display:none!important}' in s
assert 'background:#dc2626"></i>Prioritárias' in s
assert 'Buscar obras — nota, SGO, cliente, município ou endereço' in s
assert 'id="nipAdvancedQ"' in s and 'id="nipAdvancedGo"' in s
assert "'nip-tab-relatorios':4" in s
faq=json.loads(Path('/mnt/data/FAQ_NIP_CONTEUDO.json').read_text())
assert faq and not any('Qualidade da base' in json.dumps(d,ensure_ascii=False) for d in faq)
assert "'Qualidade da base','Relatórios'" not in s
for i,script in enumerate(re.findall(r'<script\b[^>]*>(.*?)</script>',s,re.S|re.I)):
 if not script.strip():continue
 with tempfile.NamedTemporaryFile('w',suffix='.js',delete=False) as t:t.write(script.replace('__NIP_DATA_PAYLOAD__','[]'));n=t.name
 p=subprocess.run(['node','--check',n],capture_output=True,text=True);os.unlink(n)
 assert p.returncode==0,(i,p.stderr)
print('PASS: aba antiga removida da navegação; FAQ sem referencias; busca unica visivel; prioridade na legenda; JavaScript valido')
