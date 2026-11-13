"""Reproducible observed-input analysis and truth-separated regression evaluation."""
from pathlib import Path
import sys,json,csv,time
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
from longisafe.cli import run
from longisafe.csv_schema import ContractError
from longisafe.config import MonitorConfig
from dataclasses import asdict
root=Path(__file__).resolve().parents[1]
rows=[]
observed={}
expected={'S09':'MON-003','S10':'MON-004','S14':'MON-005'}
for path in sorted((root/'data/synthetic').glob('*.csv')):
    started=time.perf_counter()
    try:
        r=run(path,root/'reports'/path.stem)
        observed[path.stem]=r
        fails=[x for x in r['rules'] if x['status']=='FAIL']
        first=next((x for x in fails if x['rule_id']==expected.get(path.stem)),None)
        rows.append(dict(scenario_id=path.stem,status='ANALYZED',fail_rules=';'.join(x['rule_id'] for x in fails),detection_time_s=first['detection_time_s'] if first else '',latency_from_injection_s=round(first['detection_time_s']-3.,8) if first else '',stop_distance_m=r['metrics'].get('stop_distance_m',{}).get('value'),elapsed_s=time.perf_counter()-started))
    except ContractError as e:
        if path.stem!='S11':raise
        rows.append(dict(scenario_id=path.stem,status='REJECTED MON-001',fail_rules=str(e),detection_time_s='',latency_from_injection_s='',stop_distance_m='',elapsed_s=time.perf_counter()-started))
with (root/'reports/suite_results.csv').open('w',newline='') as f:
    w=csv.DictWriter(f,fieldnames=list(rows[0]));w.writeheader();w.writerows(rows)
(root/'scenarios/monitor_config.json').write_text(json.dumps(asdict(MonitorConfig()),indent=2))
fault_checks=[]
for sid,rid in expected.items():
    events=[r for r in observed[sid]['rules'] if r['rule_id']==rid and r['status']=='FAIL']
    fault_checks.append(dict(scenario_id=sid,expected_rule=rid,detected=bool(events)))
fault_checks.append(dict(scenario_id='S11',expected_rule='MON-001 REJECT',detected=any(r['scenario_id']=='S11' and r['status'].startswith('REJECTED') for r in rows)))
healthy=['S01','S02','S04','S05']
summary=dict(observable_faults=len(fault_checks),detected=sum(r['detected'] for r in fault_checks),fault_checks=fault_checks,healthy_cases=healthy,spurious_fail_findings=sum(sum(r['status']=='FAIL' for r in observed[sid]['rules']) for sid in healthy),scope='Small predeclared synthetic regression suite; no population-level accuracy or safety claim')
(root/'reports/evaluation_summary.json').write_text(json.dumps(summary,indent=2))
expected_rows=[]
for sid,result in observed.items():expected_rows.extend(result['rules'])
with (root/'tests/expected_vba_results.csv').open('w',newline='') as f:
    w=csv.DictWriter(f,fieldnames=list(expected_rows[0]));w.writeheader();w.writerows(expected_rows)
print(json.dumps(rows,indent=2));print(json.dumps(summary,indent=2))
