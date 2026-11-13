"""unittest discovery executable here; also discoverable unchanged by pytest."""
import unittest,tempfile,math,copy,json,hashlib
from pathlib import Path
from longisafe.model import Vehicle,resistance,step,simulate,allocate,regen_capacity
from longisafe.generator import scenarios,generate
from longisafe.csv_schema import load_csv,write_csv,ContractError,REQUIRED,safe_text
from longisafe.analysis import analyze,evaluate,reference_metrics
from longisafe.config import MonitorConfig
from longisafe.cli import run

class Physics(unittest.TestCase):
    def test_constant_decel_analytic(self):
        cfg=dict(kind='brake',v0_mps=25.,decel_mps2=3.,brake_start_s=0.,duration_s=12.,dt_s=.07,accel_noise_std=0.)
        rows,truth=simulate(cfg,Vehicle(crr=0.,cda_m2=0.))
        self.assertAlmostEqual(truth['exact_stop']['time_s'],25/3,places=9)
        self.assertAlmostEqual(truth['exact_stop']['distance_m'],25**2/6,places=8)
        m=reference_metrics(rows)
        self.assertAlmostEqual(m['stop_distance_m']['value'],25**2/6,places=8)
        self.assertAlmostEqual(m['stop_time_s']['value'],25/3,places=8)
        self.assertTrue(all(r['vehicle_speed_mps']>=0 for r in rows))
    def test_invalid_scenario_rejected(self):
        for cfg in [dict(kind='unknown'),dict(kind='brake',dt_s=float('nan')),dict(kind='brake',decel_mps2=-1),dict(kind='brake',soc_pct=101),dict(kind='brake',fault='unknown'),dict(kind='brake',typo=1)]:
            with self.subTest(cfg=cfg):
                with self.assertRaises(ValueError):simulate(cfg)
    def test_grade_sign(self):
        p=Vehicle();self.assertLess(resistance(10,-.06,p),resistance(10,.06,p))
        self.assertLess(resistance(10,-.06,p),0)
    def test_stop_partial_step(self):
        vn,s,t=step(1.,-4.,1.);self.assertEqual(vn,0);self.assertEqual(t,.25);self.assertEqual(s,.125)
    def test_zero_request_no_brakes(self):
        self.assertEqual(allocate(20,0,0,-.06,60,25,Vehicle())[:2],(0,0))
    def test_capacity_limits(self):
        p=Vehicle();cap=regen_capacity(20,60,25,p)
        self.assertLess(regen_capacity(20,98,25,p),cap);self.assertLess(regen_capacity(20,60,-5,p),cap)
        self.assertEqual(regen_capacity(0,60,25,p),0)
    def test_compensation_restores_request(self):
        p=Vehicle();r,f,_=allocate(20,3,0,-.06,98,25,p,True,True)
        self.assertAlmostEqual((r+f+resistance(20,-.06,p))/p.mass_kg,3)
    def test_traction_cap(self):
        p=Vehicle(mu=.1);r,f,_=allocate(20,8,0,0,60,25,p)
        self.assertLessEqual(r+f,.1*p.mass_kg*p.g+1e-9)
    def test_delayed_braking_extends_stop(self):
        base=dict(kind='brake',duration_s=20,dt_s=.05,v0_mps=25,brake_start_s=1,soc_pct=100,accel_noise_std=0)
        a,_=simulate(base);b,_=simulate(dict(base,friction_delay_until_s=2))
        self.assertGreater(reference_metrics(b)['stop_distance_m']['value'],reference_metrics(a)['stop_distance_m']['value'])

class Contract(unittest.TestCase):
    def setUp(self):
        self.t=tempfile.TemporaryDirectory();self.addCleanup(self.t.cleanup);self.p=Path(self.t.name)/'x.csv'
        self.good='time_s,vehicle_speed_mps,accel_long_mps2,accelerator_pedal_pct,brake_pedal_pct\n0,1,-1,0,0\n0.05,.95,-1,0,0\n'
    def test_bom_and_reordered(self):
        self.p.write_text('\ufeff'+self.good);self.assertEqual(len(load_csv(self.p)),2)
        lines=[line.split(',') for line in self.good.splitlines()]
        self.p.write_text('\n'.join(','.join(reversed(r)) for r in lines)+'\n')
        self.assertEqual(load_csv(self.p)[1]['vehicle_speed_mps'],.95)
    def test_bad_csv_cases(self):
        for bad in [self.good.replace('.95','NaN'),self.good.replace('.95','3oops'),self.good.replace('.95','1e999'),self.good.replace('.95',''),self.good.replace('0.05','0'),self.good.replace('0.05','-1'),self.good.replace('.95','".95"'),self.good.replace('0.05,.95','0.05,-1'),self.good.replace('brake_pedal_pct','accelerator_pedal_pct'),self.good.replace('vehicle_speed_mps','v_kmh'),self.good+'\n',self.good.replace('0.05,.95,-1,0,0','0.05,.95,-1,0,101')]:
            with self.subTest(bad=bad):
                self.p.write_text(bad)
                with self.assertRaises(ContractError):load_csv(self.p)
    def test_formula_guard(self):
        self.assertEqual(safe_text('=1+1'),"'=1+1");self.assertEqual(safe_text('normal'),'normal')
    def test_config_rejected(self):
        for c in [MonitorConfig(evidence_s=0),MonitorConfig(nominal_dt_s=float('nan')),MonitorConfig(gap_factor=1)]:
            with self.assertRaises(ValueError):c.validate()

class Suite(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp=tempfile.TemporaryDirectory();cls.root=Path(cls.temp.name);generate(cls.root)
    @classmethod
    def tearDownClass(cls):cls.temp.cleanup()
    def rows(self,sid):return load_csv(self.root/'data/synthetic'/f'{sid}.csv')
    def statuses(self,sid):return {(r['rule_id'],r['status']) for r in evaluate(self.rows(sid),sid)}
    def test_nominal_no_false_fail(self):
        for sid in ('S01','S02','S04','S05'):
            self.assertFalse(any(s=='FAIL' for _,s in self.statuses(sid)),sid)
    def test_downhill_accelerates_without_fault(self):
        rows=self.rows('S02');self.assertGreater(rows[-1]['vehicle_speed_mps'],rows[0]['vehicle_speed_mps'])
        self.assertTrue(all(r['accelerator_pedal_pct']==0 for r in rows));self.assertIn(('MON-003','PASS'),self.statuses('S02'))
    def test_observable_faults(self):
        for sid,rid in [('S09','MON-003'),('S10','MON-003'),('S10','MON-004'),('S14','MON-005')]:
            self.assertIn((rid,'FAIL'),self.statuses(sid))
    def test_bad_timestamp_reject(self):
        with self.assertRaisesRegex(ContractError,'MON-001'):self.rows('S11')
    def test_missing_not_pass(self):self.assertIn(('MON-005','NOT_EVALUATED'),self.statuses('MISSING_CONTEXT'))
    def test_faulted_kpi_withheld(self):self.assertIn('analysis_withheld',analyze(self.rows('S09'))['metrics'])
    def test_performance_rows(self):self.assertEqual(len(self.rows('PERF_10001')),10001)
    def test_gap_warn_and_withhold(self):
        rows=self.rows('S04');rows=rows[:30]+rows[40:]
        r=analyze(rows);self.assertEqual(r['rules'][0]['status'],'WARN');self.assertIn('analysis_withheld',r['metrics'])
    def test_uncompensated_distance(self):
        s5=analyze(self.rows('S05'));s14=analyze(self.rows('S14'))
        self.assertGreater(s14['metrics']['stop_distance_m']['value'],s5['metrics']['stop_distance_m']['value'])
        self.assertIn(('MON-005','PASS'),self.statuses('S05'))
    def test_reproducible_bytes(self):
        before=(self.root/'data/synthetic/S04.csv').read_bytes();generate(self.root)
        self.assertEqual(before,(self.root/'data/synthetic/S04.csv').read_bytes())
    def test_config_input_controls_generation(self):
        dest=self.root/'custom'
        generate(dest,{'CUSTOM':dict(kind='steady',seed=5,dt_s=.1,duration_s=1.,v0_mps=15.,vehicle={'mass_kg':2100.})})
        self.assertEqual(len(load_csv(dest/'data/synthetic/CUSTOM.csv')),11)
        self.assertEqual(json.loads((dest/'data/ground_truth/CUSTOM.json').read_text())['parameters']['mass_kg'],2100.)
    def test_truth_not_input(self):
        text=(self.root/'data/synthetic/S09.csv').read_text();self.assertNotIn('fault',text);self.assertNotIn('original',text)
    def test_heldout_parameter_variants(self):
        # Regression robustness, not independent real-world validation. Defaults frozen before this run.
        for seed,v0,dt in [(173,22.,.04),(809,28.,.06),(992,24.,.025)]:
            cfg=dict(scenarios()['S05'],seed=seed,v0_mps=v0,dt_s=dt)
            rows,_=simulate(cfg)
            self.assertFalse(any(r['status']=='FAIL' for r in evaluate(rows)))
            rows,_=simulate(dict(cfg,compensate=False,soc_pct=60.))
            self.assertTrue(any(r['rule_id']=='MON-005' and r['status']=='FAIL' for r in evaluate(rows)))
    def test_short_record_not_evaluated(self):
        r=evaluate(self.rows('S04')[:3]);self.assertTrue(any(x['rule_id']=='MON-003' and x['status']=='NOT_EVALUATED' for x in r))
    def test_export_round_trip(self):
        out=self.root/'export';r=run(self.root/'data/synthetic/S04.csv',out)
        self.assertTrue((out/'rule_results.csv').exists());self.assertEqual(json.loads((out/'analysis.json').read_text())['data_sha256'],r['data_sha256'])
    def test_constant_accel_filtered_jerk(self):
        rows=self.rows('S01')
        for r in rows:r['accel_long_mps2']=2.
        self.assertAlmostEqual(reference_metrics(rows)['jerk_rms_mps3']['value'],0.)

if __name__=='__main__':unittest.main()
