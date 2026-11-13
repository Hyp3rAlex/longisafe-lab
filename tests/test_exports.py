import unittest,tempfile,csv,sys,json,hashlib
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from check_excel_exports import check
from longisafe.cli import run
from longisafe.generator import generate
from longisafe.config import MonitorConfig
from dataclasses import asdict

class ExportChecker(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.addCleanup(self.temp.cleanup)
        self.root=Path(self.temp.name);generate(self.root)
        self.input=self.root/'data/synthetic/S04.csv';self.folder=self.root/'export'
        run(self.input,self.folder)
        with (self.folder/'run_metadata.csv').open('w',newline='') as f:
            w=csv.writer(f);w.writerow(['key','value']);w.writerow(['sha256',hashlib.sha256(self.input.read_bytes()).hexdigest()]);w.writerow(['scenario_id','S04'])
        with (self.folder/'run_config.csv').open('w',newline='') as f:
            w=csv.writer(f);w.writerow(['parameter','value']);w.writerows(asdict(MonitorConfig()).items())
    def test_checker_accepts_reference_fixture(self):
        # Deliberately Python-produced fixture: validates comparison plumbing ONLY, not VBA.
        self.assertEqual(check(self.input,self.folder)['result'],'PASS')
    def test_checker_rejects_changed_distance(self):
        p=self.folder/'metrics.csv';rows=list(csv.DictReader(p.open()))
        for r in rows:
            if r['metric']=='stop_distance_m':r['value']='999'
        with p.open('w',newline='') as f:
            w=csv.DictWriter(f,fieldnames=list(rows[0]));w.writeheader();w.writerows(rows)
        r=check(self.input,self.folder)
        self.assertEqual(r['result'],'FAIL');self.assertIn('stop_distance_m: numerical mismatch',r['failures'])
    def test_checker_rejects_forged_rule_pass(self):
        p=self.folder/'rule_results.csv';rows=list(csv.DictReader(p.open()));rows[2]['status']='FAIL'
        with p.open('w',newline='') as f:
            w=csv.DictWriter(f,fieldnames=list(rows[0]));w.writeheader();w.writerows(rows)
        self.assertEqual(check(self.input,self.folder)['result'],'FAIL')
    def test_checker_rejects_missing_output(self):
        (self.folder/'metrics.csv').unlink()
        with self.assertRaises(FileNotFoundError):check(self.input,self.folder)
