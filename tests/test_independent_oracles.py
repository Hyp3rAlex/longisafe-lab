"""Small hand-solvable observations provide an oracle independent of plant generation."""
import unittest
from longisafe.analysis import reference_metrics,evaluate
from longisafe.csv_schema import REQUIRED
from longisafe.model import simulate,Vehicle

class IndependentOracles(unittest.TestCase):
    def test_manual_triangular_stop(self):
        # Exactly v=10-2*t, brake starts at t=0; distance=25m, duration=5s.
        rows=[]
        for i in range(7):
            rows.append(dict(zip(REQUIRED,(float(i),max(0.,10.-2*i),-2. if i<5 else 0.,0.,40. if i<5 else 0.))))
        m=reference_metrics(rows)
        self.assertEqual(m['stop_distance_m']['value'],25.)
        self.assertEqual(m['stop_time_s']['value'],5.)
    def test_manual_stop_inside_last_interval(self):
        rows=[dict(zip(REQUIRED,x)) for x in [(0.,1.,-4.,0.,40.),(1.,0.,0.,0.,0.)]]
        m=reference_metrics(rows)
        self.assertEqual(m['stop_time_s']['value'],.25);self.assertEqual(m['stop_distance_m']['value'],.125)
    def test_manual_brake_work_not_sample_fraction(self):
        # First 1s at average 3m/s regen=100N =>300J; second avg 1m/s friction=100N=>100J.
        rows=[]
        for x,reg,fric in [((0.,4.,-2.,0.,40.),100.,0.),((1.,2.,-2.,0.,40.),0.,100.),((2.,0.,0.,0.,0.),0.,0.)]:
            r=dict(zip(REQUIRED,x));r.update(regen_force_N=reg,friction_force_N=fric);rows.append(r)
        m=reference_metrics(rows);self.assertEqual(m['regen_work_J']['value'],300.);self.assertEqual(m['regen_work_fraction']['value'],.75)
    def test_brake_work_partial_stop(self):
        rows=[dict(zip(REQUIRED,x)) for x in [(0.,1.,-4.,0.,40.),(1.,0.,0.,0.,0.)]]
        for r in rows:r.update(regen_force_N=100.,friction_force_N=0.)
        self.assertEqual(reference_metrics(rows)['regen_work_J']['value'],12.5)
    def test_monitor_sustained_bias_without_generator(self):
        rows=[dict(zip(REQUIRED,(i*.05,10.,1.5,0.,0.))) for i in range(30)]
        r=evaluate(rows)
        self.assertTrue(any(x['rule_id']=='MON-003' and x['status']=='FAIL' for x in r))
        self.assertTrue(any(x['rule_id']=='MON-004' and x['status']=='FAIL' for x in r))
    def test_flat_no_traction_stays_stopped(self):
        rows,_=simulate(dict(kind='coast',v0_mps=0,dt_s=.05,duration_s=1,accel_noise_std=0))
        self.assertTrue(all(x['vehicle_speed_mps']==0 and x['accel_long_mps2']==0 for x in rows))
    def test_convergence_with_drag(self):
        distances=[]
        for dt in [.1,.05,.025]:
            _,truth=simulate(dict(kind='brake',v0_mps=25,dt_s=dt,duration_s=35,brake_start_s=0,loss_time_s=3,compensate=False,accel_noise_std=0))
            distances.append(truth['exact_stop']['distance_m'])
        self.assertLess(abs(distances[1]-distances[2]),abs(distances[0]-distances[1]))
