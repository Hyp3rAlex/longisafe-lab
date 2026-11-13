from pathlib import Path
import argparse,sys,json
from check_excel_exports import check
p=argparse.ArgumentParser();p.add_argument('smoke_exports');a=p.parse_args()
root=Path(__file__).resolve().parents[1];exports=Path(a.smoke_exports);results={}
for sid in ['S01','S02','S04','S05','S09','S10','S14','MISSING_CONTEXT','PERF_10001']:
    results[sid]=check(root/'data/synthetic'/f'{sid}.csv',exports/sid)
print(json.dumps(results,indent=2));sys.exit(any(v['failures'] for v in results.values()))
