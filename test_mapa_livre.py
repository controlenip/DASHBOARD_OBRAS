from pathlib import Path
import re, subprocess, unittest
S=Path('/mnt/data/nip_dashboard_template.html').read_text(encoding='utf-8')
class MapTests(unittest.TestCase):
 def test_initial_center(self):
  self.assertIn('setView(MARANHAO_CENTER,7)',S)
 def test_no_max_bounds(self):
  self.assertNotIn('maxBounds:MARANHAO_BOUNDS',S)
  self.assertNotIn('maxBoundsViscosity:',S)
  self.assertNotIn('map.panInsideBounds(MARANHAO_BOUNDS',S)
 def test_popup_auto_pan_and_scroll(self):
  self.assertIn('autoPanPaddingBottomRight:',S)
  self.assertIn('maxHeight:Math.max(250,',S)
  self.assertIn('.leaflet-popup-scrolled',S)
 def test_map_center_on_filter_and_free_navigation(self):
  self.assertIn('map.fitBounds(L.latLngBounds(points)',S)
  self.assertIn('Navegação livre no mapa.',S)
 def test_other_sections_preserved(self):
  for term in ('QUALIDADE LEVANTAMENTO','PRODUTIVIDADE','RELATÓRIOS','FAQ / DÚVIDAS','function renderWorkload','function nipOpsRows','function popup(r)'):
   self.assertIn(term,S)
 def test_scripts_parse(self):
  scripts=re.findall(r'<script(?:\s[^>]*)?>(.*?)</script>',S,re.S|re.I)
  self.assertGreaterEqual(len(scripts),10)
  for i,sc in enumerate(scripts):
   if not sc.strip():continue
   p=Path('/mnt/data/_syntax_map_test.js');p.write_text(sc)
   r=subprocess.run(['node','--check',str(p)],capture_output=True,text=True)
   self.assertEqual(r.returncode,0,f'Script {i}: {r.stderr}')
if __name__=='__main__':unittest.main(verbosity=2)
