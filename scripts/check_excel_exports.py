"""Compare real Excel exports with independent Python calculations; never fabricates exports."""
import argparse,csv,json,math,sys,hashlib
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
from longisafe.csv_schema import load_csv
from longisafe.analysis import analyze
from longisafe.config import MonitorConfig

def check(input_csv,folder):
    folder=Path(folder)
    with (folder/'run_config.csv').open(encoding='utf-8-sig',newline='') as f:
        cfg=MonitorConfig(**{r['parameter']:float(r['value']) for r in csv.DictReader(f)}).validate()
    expected=analyze(load_csv(input_csv),Path(input_csv).stem,cfg);failures=[]
    metadata=folder/'run_metadata.csv'
    if metadata.exists():
        with metadata.open(encoding='utf-8-sig',newline='') as f:meta={r['key']:r['value'] for r in csv.DictReader(f)}
        if meta.get('sha256')!=hashlib.sha256(Path(input_csv).read_bytes()).hexdigest():failures.append('Input SHA-256 mismatch')
        if meta.get('scenario_id')!=Path(input_csv).stem:failures.append('Metadata scenario mismatch')
    else:failures.append('Missing run_metadata.csv (provenance unverified)')
    with (folder/'rule_results.csv').open(encoding='utf-8-sig',newline='') as f:actual_rules=list(csv.DictReader(f))
    for rid in [f'MON-{i:03}' for i in range(1,11)]:
        ex=[r for r in expected['rules'] if r['rule_id']==rid];ac=[r for r in actual_rules if r['rule_id']==rid]
        if sorted(r['status'] for r in ex)!=sorted(r['status'] for r in ac):
            failures.append(f'{rid}: status/event-count mismatch');continue
        for e,a in zip(ex,ac):
            if a['scenario_id']!=e['scenario_id']:failures.append(f'{rid}: scenario mismatch')
            for field in ['t_start_s','t_end_s','detection_time_s','start_csv_row','end_csv_row']:
                if e[field]=='':
                    if a[field]!='':failures.append(f'{rid} {field}: expected blank')
                elif a[field]=='' or not math.isclose(float(a[field]),float(e[field]),abs_tol=1e-6,rel_tol=1e-9):failures.append(f'{rid} {field}: mismatch')
            # Values and thresholds numeric when defined; monitor 002 threshold is descriptive text in VBA.
            for field in ['measured_value','threshold']:
                if isinstance(e[field],(float,int)):
                    if rid=='MON-002' and field=='threshold':continue
                    try:good=math.isclose(float(a[field]),float(e[field]),abs_tol=1e-6,rel_tol=1e-6)
                    except (ValueError,TypeError):good=False
                    if not good:failures.append(f'{rid} {field}: mismatch')
    with (folder/'metrics.csv').open(encoding='utf-8-sig',newline='') as f:actual={r['metric']:r for r in csv.DictReader(f)}
    if set(actual)!=set(expected['metrics']):failures.append('Metric inventory mismatch')
    # Algorithms match definitions but arithmetic implementations are independent.
    tolerances={'stop_time_s':.003,'stop_distance_m':.02,'regen_work_J':.1,'regen_work_fraction':1e-6,'jerk_peak_mps3':1e-6,'jerk_rms_mps3':1e-6}
    for name,e in expected['metrics'].items():
        if name not in actual:continue
        a=actual[name]
        if a['status']!=e['status'] or a['units']!=e['units']:failures.append(f'{name}: status/unit mismatch')
        if e['value'] is None:
            if a['value']!='':failures.append(f'{name}: expected blank, never zero')
        else:
            try:good=math.isclose(float(a['value']),e['value'],rel_tol=1e-6,abs_tol=tolerances.get(name,1e-6))
            except ValueError:good=False
            if not good:failures.append(f'{name}: numerical mismatch')
    result=dict(result='FAIL' if failures else 'PASS',input_sha256=hashlib.sha256(Path(input_csv).read_bytes()).hexdigest(),config_sha256=cfg.sha256,failures=failures,caveat='Compares provided exports; this script cannot attest that exports were produced by desktop Excel.')
    (folder/'crosscheck.json').write_text(json.dumps(result,indent=2));return result
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('csv');p.add_argument('export_folder');a=p.parse_args()
    r=check(a.csv,a.export_folder);print(json.dumps(r,indent=2));sys.exit(bool(r['failures']))
