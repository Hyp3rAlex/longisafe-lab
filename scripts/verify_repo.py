"""Static integrity check. Not a VBA compiler and not a replacement for Excel."""
from pathlib import Path
from vba_source_sync import workbook_modules
import hashlib,sys,json
root=Path(__file__).resolve().parents[1]
failed=[]
wb=root/'excel/LongiSafeLab.xlsm'
if not wb.exists():failed.append('Missing validated XLSM')
else:
    modules=workbook_modules(wb)
    files=list((root/'excel/vba').glob('*.bas'))
    if len(modules)!=9 or len(files)!=9:failed.append('Expected nine VBA modules')
    for name,b in modules.items():
        p=root/'excel/vba'/f'{name}.bas'
        if not p.exists() or p.read_bytes()!=b:failed.append('Workbook/source mismatch: '+name)
for scenario in ('S01','S02','S04','S05','S09','S10','S11','S14','MISSING_CONTEXT','PERF_10001'):
    if not (root/'data/synthetic'/f'{scenario}.csv').exists():failed.append('Missing dataset '+scenario)
for scenario in ('S05','S14'):
    if not (root/'evidence/excel_smoke'/scenario/'summary.pdf').exists():failed.append('Missing Excel report '+scenario)
for fn in ('docs/TEST_RESULTS.md','docs/engineering_assumptions.md','docs/signal_dictionary.md','docs/safety_scope_and_limitations.md','docs/VBA_SOURCE_SYNC.md'):
    if not (root/fn).exists():failed.append('Missing documentation '+fn)
print('Release structural/source verification:', 'PASS' if not failed else 'FAIL')
for item in failed:print(' -',item)
sys.exit(bool(failed))
