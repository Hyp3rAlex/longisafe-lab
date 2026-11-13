"""Report artifact checks separate from physics and monitor tests."""
import unittest,json,hashlib
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
try:
    from pypdf import PdfReader
    from PIL import Image
except ImportError: PdfReader=None

@unittest.skipIf(PdfReader is None,'Optional PDF/image inspection dependencies unavailable')
class ReportArtifact(unittest.TestCase):
    def test_pdf_one_page_and_traceable(self):
        p=ROOT/'reports/example_summary.pdf'
        if not p.exists():self.skipTest('Run scripts/render_demo.py to generate report')
        reader=PdfReader(p);self.assertEqual(len(reader.pages),1);text=reader.pages[0].extract_text()
        self.assertIn('UNVERIFIED',text)
        self.assertIn(hashlib.sha256((ROOT/'data/synthetic/S14.csv').read_bytes()).hexdigest(),text)
        self.assertIn('not BMW-specific',text)
    def test_pdf_numbers_match_observed_results(self):
        p=ROOT/'reports/example_summary.pdf'
        if not p.exists():self.skipTest('Report optional')
        text=PdfReader(p).pages[0].extract_text()
        for sid in ['S05','S14']:
            r=json.loads((ROOT/'reports'/sid/'analysis.json').read_text())
            self.assertIn(f"{r['metrics']['stop_distance_m']['value']:.2f} m",text)
        r=json.loads((ROOT/'reports/S14/analysis.json').read_text());e=next(x for x in r['rules'] if x['rule_id']=='MON-005' and x['status']=='FAIL')
        self.assertIn(f"{e['t_start_s']:.2f} to {e['t_end_s']:.2f} s",text)
    def test_chart_is_nonblank_image(self):
        p=ROOT/'reports/engineering_comparison.png'
        if not p.exists():self.skipTest('Report optional')
        image=Image.open(p).convert('RGB');self.assertGreater(image.width,1500)
        extrema=image.getextrema();self.assertTrue(all(lo<hi for lo,hi in extrema))
