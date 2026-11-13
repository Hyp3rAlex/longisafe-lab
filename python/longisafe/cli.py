import argparse,csv,hashlib,json
from pathlib import Path
from datetime import datetime,timezone
from .csv_schema import load_csv,ContractError,safe_text
from .analysis import analyze,RULE_FIELDS
from .config import MonitorConfig
from . import __version__

def export_result(path,result):
    path=Path(path);path.mkdir(parents=True,exist_ok=True)
    (path/'analysis.json').write_text(json.dumps(result,indent=2))
    with (path/'rule_results.csv').open('w',newline='',encoding='utf-8') as f:
        w=csv.DictWriter(f,fieldnames=RULE_FIELDS,lineterminator='\n');w.writeheader()
        for r in result['rules']:w.writerow({k:safe_text(v) if isinstance(v,str) else v for k,v in r.items()})
    with (path/'metrics.csv').open('w',newline='',encoding='utf-8') as f:
        w=csv.writer(f,lineterminator='\n');w.writerow(['metric','value','units','status','definition'])
        for k,v in result['metrics'].items():w.writerow([k,v['value'],v['units'],v['status'],v['definition']])

def run(path,out,config=None):
    rows=load_csv(path);r=analyze(rows,Path(path).stem,config)
    r.update(data_sha256=hashlib.sha256(Path(path).read_bytes()).hexdigest(),software_version=__version__,generated_at_utc=datetime.now(timezone.utc).isoformat())
    export_result(out,r);return r

def main():
    ap=argparse.ArgumentParser();ap.add_argument('csv');ap.add_argument('--out',required=True);ap.add_argument('--config');args=ap.parse_args()
    cfg=MonitorConfig(**json.loads(Path(args.config).read_text())) if args.config else None
    try:r=run(args.csv,args.out,cfg)
    except (ContractError,ValueError) as e:ap.exit(2,f'REJECTED: {e}\n')
    print(args.out,[(x['rule_id'],x['status']) for x in r['rules']])
if __name__=='__main__':main()
