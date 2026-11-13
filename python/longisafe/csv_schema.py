"""Strict numeric CSV subset shared with the VBA parser; no quoted fields."""
import csv, math, re
from pathlib import Path
REQUIRED = ('time_s','vehicle_speed_mps','accel_long_mps2','accelerator_pedal_pct','brake_pedal_pct')
OPTIONAL = ('road_grade_rad','drive_force_N','regen_force_N','friction_force_N','requested_decel_mps2',
            'regen_limit_N','soc_pct','battery_temp_C','estimated_mu','adas_brake_request_mps2')
# vehicle_mode is deliberately unsupported text in M1. Fail rather than coerce it.
NUMBER = re.compile(r'^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$')

class ContractError(ValueError): pass

def load_csv(path):
    text=Path(path).read_text(encoding='utf-8-sig')
    if '"' in text or '\x00' in text:
        raise ContractError('Quotes/NUL unsupported in M1 numeric CSV subset')
    lines=text.splitlines()
    if len(lines)<3: raise ContractError('At least two data rows required')
    header=lines[0].split(',')
    if len(set(header))!=len(header) or any(h not in REQUIRED+OPTIONAL for h in header):
        raise ContractError('Duplicate or unknown header; exact SI headers required')
    if not set(REQUIRED)<=set(header): raise ContractError('Missing required SI columns')
    rows=[]
    for rownum,line in enumerate(lines[1:],2):
        parts=line.split(',')
        if len(parts)!=len(header): raise ContractError(f'Row {rownum}: width mismatch')
        d={}
        for h,s in zip(header,parts):
            if not NUMBER.fullmatch(s): raise ContractError(f'Row {rownum}: invalid number in {h}')
            v=float(s)
            if not math.isfinite(v) or abs(v)>1e100: raise ContractError(f'Row {rownum}: nonfinite/oversized {h}')
            d[h]=v
        if d['time_s']<0: raise ContractError(f'Row {rownum}: negative time')
        if rows and d['time_s']<=rows[-1]['time_s']: raise ContractError(f'MON-001 row {rownum}: time not strictly increasing')
        if d['vehicle_speed_mps']<0 or any(not 0<=d[h]<=100 for h in ('accelerator_pedal_pct','brake_pedal_pct')):
            raise ContractError(f'MON-002 row {rownum}: speed/pedal range')
        if any(d.get(h,0)<0 for h in ('drive_force_N','regen_force_N','friction_force_N','requested_decel_mps2','regen_limit_N','adas_brake_request_mps2')):
            raise ContractError(f'MON-002 row {rownum}: negative force/request')
        if 'soc_pct' in d and not 0<=d['soc_pct']<=100: raise ContractError('SOC range')
        if abs(d.get('road_grade_rad',0))>=1.5707963267948966: raise ContractError('Grade outside forward-road domain')
        if 'estimated_mu' in d and not 0<=d['estimated_mu']<=2: raise ContractError('mu range')
        rows.append(d)
    return rows

def write_csv(path, rows):
    p=Path(path);p.parent.mkdir(parents=True,exist_ok=True)
    with p.open('w',newline='',encoding='utf-8') as f:
        w=csv.DictWriter(f,fieldnames=list(rows[0]),lineterminator='\n');w.writeheader();w.writerows(rows)

def safe_text(s):
    s=str(s)
    return "'"+s if s.lstrip().startswith(('=','+','-','@')) else s
