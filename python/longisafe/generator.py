import argparse,copy,hashlib,json,csv,re
from pathlib import Path
from .model import simulate
from .csv_schema import write_csv,REQUIRED
from . import __version__

def scenarios():
    base=dict(kind='brake',seed=41,dt_s=.05,duration_s=35.,v0_mps=25.,brake_start_s=1.,decel_mps2=3.,accel_noise_std=.015)
    return {
      'S01':dict(base,kind='steady',duration_s=8.),
      'S02':dict(base,kind='coast',grade_rad=-.06,duration_s=8.),
      'S04':dict(base),
      'S05':dict(base,soc_pct=98.,loss_time_s=3.,compensate=True),
      'S09':dict(base,fault='bias'),
      'S10':dict(base,fault='stuck_at'),
      'S11':dict(base,fault='timing_duplicate'),
      'S14':dict(base,loss_time_s=3.,compensate=False),
      'MISSING_CONTEXT':dict(base,core_only=True),
      'PERF_10001':dict(base,kind='steady',duration_s=500.,accel_noise_std=0.),
    }

def generate(root,scenario_configs=None):
    root=Path(root);manifest=[]
    for sid,cfg in (scenarios() if scenario_configs is None else scenario_configs).items():
        if not re.fullmatch(r'[A-Za-z0-9_-]{1,64}',sid):raise ValueError('Unsafe scenario identifier')
        rows,truth=simulate(cfg);original=copy.deepcopy(rows);injections=[]
        if cfg.get('fault') and cfg.get('duration_s',35.)<5.:raise ValueError('M1 signal fault window requires duration >=5 s')
        frozen_speed=next((r['vehicle_speed_mps'] for r in original if r['time_s']>=3.-1e-8),None)
        for i,r in enumerate(rows):
            if 3.-1e-8<=r['time_s']<=5.+1e-8 and cfg.get('fault') in ('bias','stuck_at'):
                field='accel_long_mps2' if cfg['fault']=='bias' else 'vehicle_speed_mps'
                old=r[field];r[field]=old+1.5 if cfg['fault']=='bias' else frozen_speed
                injections.append(dict(csv_row=i+2,time_s=r['time_s'],field=field,original=old,observed=r[field]))
        if cfg.get('fault')=='timing_duplicate':
            index=next(i for i,r in enumerate(rows) if r['time_s']>=4.-1e-8)
            old=rows[index]['time_s'];rows[index]['time_s']=rows[index-1]['time_s']
            injections.append(dict(csv_row=index+2,field='time_s',original=old,observed=rows[index]['time_s']))
        if sid in ('S05','S14'):injections.append(dict(kind='regen_capacity_loss',time_s=3.,compensated=cfg['compensate']))
        if cfg.get('core_only'):rows=[{k:r[k] for k in REQUIRED} for r in rows]
        data=root/'data/synthetic'/f'{sid}.csv';write_csv(data,rows)
        truth.update(scenario_id=sid,seed=cfg['seed'],generator_version=__version__,config=cfg,injections=injections,data_sha256=hashlib.sha256(data.read_bytes()).hexdigest())
        tp=root/'data/ground_truth'/f'{sid}.json';tp.parent.mkdir(parents=True,exist_ok=True);tp.write_text(json.dumps(truth,indent=2))
        cp=root/'scenarios/configs'/f'{sid}.json';cp.parent.mkdir(parents=True,exist_ok=True);cp.write_text(json.dumps(cfg,indent=2))
        expected={'S09':'MON-003 FAIL','S10':'MON-003/MON-004 FAIL','S11':'MON-001 REJECT','S14':'MON-005 FAIL','MISSING_CONTEXT':'MON-005 NOT_EVALUATED'}.get(sid,'No FAIL')
        if sid not in scenarios():expected='Custom configuration: review observed output; no predeclared oracle'
        manifest.append(dict(scenario_id=sid,seed=cfg['seed'],dt_s=cfg['dt_s'],model_version=__version__,data_path=str(data.relative_to(root)),truth_path=str(tp.relative_to(root)),expected_behavior=expected))
    write_csv(root/'scenarios/manifest.csv',manifest)
    truth_summary=[]
    for entry in manifest:
        truth=json.loads((root/entry['truth_path']).read_text())
        truth_summary.append(dict(scenario_id=entry['scenario_id'],seed=truth['seed'],generator_version=truth['generator_version'],data_sha256=truth['data_sha256'],injected_records=len(truth['injections']),truth_file=entry['truth_path']))
    write_csv(root/'data/ground_truth/injection_manifest.csv',truth_summary)
    print(f'Generated {len(manifest)} deterministic datasets; truth is separate')

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('--root',default='.');ap.add_argument('--config');ap.add_argument('--scenario-id',default='CUSTOM');args=ap.parse_args()
    configs={args.scenario_id:json.loads(Path(args.config).read_text())} if args.config else None
    generate(args.root,configs)
