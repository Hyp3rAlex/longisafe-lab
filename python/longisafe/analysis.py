"""Observed-input analysis. This module never reads fault truth."""
import math
from dataclasses import asdict
from .config import MonitorConfig

RULE_FIELDS=('rule_id','scenario_id','status','t_start_s','t_end_s','detection_time_s','start_csv_row','end_csv_row','measured_value','threshold','units','prerequisite_fields','rationale','caveat')

def evaluate(rows,sid='input',config=None):
    c=(config or MonitorConfig()).validate();n=len(rows);out=[]
    t=[r['time_s'] for r in rows];v=[r['vehicle_speed_mps'] for r in rows];a=[r['accel_long_mps2'] for r in rows]
    def result(rid,status,reason,threshold='',value='',unit='',prereq='',start=None,end=None,det=None):
        out.append(dict(zip(RULE_FIELDS,(rid,sid,status,'' if start is None else t[start],'' if end is None else t[end],'' if det is None else t[det],'' if start is None else start+2,'' if end is None else end+2,value,threshold,unit,prereq,reason,'Illustrative condition, not a vehicle safety conclusion'))))
    gaps=[i for i in range(1,n) if t[i]-t[i-1]>c.gap_factor*c.nominal_dt_s+1e-9]
    result('MON-001','WARN' if gaps else 'PASS','Large timing gap; downstream analysis withheld' if gaps else 'Strict timing accepted; no large gaps',c.gap_factor*c.nominal_dt_s,max(t[i]-t[i-1] for i in range(1,n)),'s','time_s',gaps[0]-1 if gaps else None,gaps[-1] if gaps else None)
    bad=[i for i in range(n) if v[i]>c.max_speed_mps or abs(a[i])>c.max_abs_accel_mps2]
    result('MON-002','FAIL' if bad else 'PASS','Configured speed/acceleration range exceeded' if bad else 'Parser/range checks evaluated',f'speed<={c.max_speed_mps}; |accel|<={c.max_abs_accel_mps2}',max(abs(x) for x in a),'SI','core fields',bad[0] if bad else None,bad[-1] if bad else None)
    # Complete elapsed-time trailing window; acceleration integrated on outgoing intervals.
    prefix=[0.]*n
    for i in range(1,n):prefix[i]=prefix[i-1]+a[i-1]*(t[i]-t[i-1])
    residual=[None]*n; stale=[None]*n;k=0
    for i in range(1,n):
        while k+1<i and t[i]-t[k+1]>=c.filter_window_s-1e-9:k+=1
        span=t[i]-t[k]
        if span>=c.filter_window_s-1e-9:
            av=(prefix[i]-prefix[k])/span
            residual[i]=abs((v[i]-v[k])/span-av)
            stale[i]=abs(v[i]-v[k])<=c.stale_speed_mps and abs(av)>c.motion_accel_mps2
    def sustained(rid,values,limit,reason,prereq,unit,eligible=None):
        if gaps or bad:
            result(rid,'NOT_EVALUATED','Invalid timing/range prevents evaluation',limit,unit=unit,prereq=prereq);return
        available=[i for i,x in enumerate(values) if x is not None and (eligible is None or eligible[i])]
        # A single sample cannot justify PASS for a sustained monitor.
        if not available or t[available[-1]]-t[available[0]]<c.evidence_s-1e-9:
            result(rid,'NOT_EVALUATED','Prerequisites or usable duration absent',limit,unit=unit,prereq=prereq);return
        events=[];start=None;det=None
        for i in range(n+1):
            hit=i<n and values[i] is not None and (eligible is None or eligible[i]) and values[i]>limit
            if hit:
                if start is None:start=i
                if det is None and t[i]-t[start]>=c.evidence_s-1e-9:det=i
            elif start is not None:
                if det is not None:events.append((start,i-1,det))
                start=None;det=None
        if events:
            for s,e,d in events:result(rid,'FAIL',reason,limit,max(values[s:e+1]),unit,prereq,s,e,d)
        else:result(rid,'PASS','Evaluated; no sustained threshold breach',limit,max(values[i] for i in available),unit,prereq)
    sustained('MON-003',residual,c.coherence_mps2,'Speed change contradicts integrated measured acceleration','time_s; vehicle_speed_mps; accel_long_mps2','m/s^2')
    sustained('MON-004',[None if x is None else float(x) for x in stale],.5,'Near-stationary speed signal with independent measured acceleration','time_s; vehicle_speed_mps; accel_long_mps2','boolean')
    request='requested_decel_mps2' in rows[0]
    if request:
        error=[];eligible=[];onset=None
        for i,r in enumerate(rows):
            req=r['requested_decel_mps2']
            if req>c.min_request_mps2 and v[i]>c.min_speed_mps:
                if onset is None:onset=t[i]
                eligible.append(t[i]-onset>=c.response_delay_s-1e-9)
                error.append(max(0.,req+a[i]))
            else:onset=None;eligible.append(False);error.append(None)
        sustained('MON-005',error,c.decel_error_mps2,'Requested net deceleration undershoot; cause not diagnosed','requested_decel_mps2; accel_long_mps2; vehicle_speed_mps','m/s^2',eligible)
    else:result('MON-005','NOT_EVALUATED','Missing requested net deceleration; pedal is not a force request',c.decel_error_mps2,unit='m/s^2',prereq='requested_decel_mps2')
    for rid,reason in [('MON-006','No calibrated pedal-to-force mapping; zero accelerator is not a fault criterion'),('MON-007','Force residual monitor reserved for M2'),('MON-008','Transition screening monitor reserved for M2; work fraction computed separately'),('MON-009','Terminal comfort screening reserved for M2; full-record filtered jerk is available')]:
        result(rid,'NOT_EVALUATED',reason)
    missing=[x for x in ('requested_decel_mps2','road_grade_rad','drive_force_N','regen_force_N','friction_force_N') if x not in rows[0]]
    result('MON-010','NOT_EVALUATED' if missing else 'PASS','Missing optional context: '+','.join(missing) if missing else 'M1 context channels present; M2 monitors still unimplemented',prereq='optional context inventory')
    return out

def reference_metrics(rows,config=None):
    """Independent of plant: trapezoids and observed-event stopping calculation."""
    c=(config or MonitorConfig()).validate();t=[r['time_s'] for r in rows];v=[r['vehicle_speed_mps'] for r in rows];a=[r['accel_long_mps2'] for r in rows];n=len(t)
    metrics={}
    def put(name,value,unit,status='EVALUATED',note=''):
        metrics[name]=dict(value=value,units=unit,status=status,definition=note)
    put('record_duration_s',t[-1]-t[0],'s')
    peak=min(range(n),key=lambda i:a[i]);put('peak_decel_mps2',max(0.,-a[peak]),'m/s^2',note='Raw observed acceleration, full record')
    put('peak_decel_time_s',t[peak],'s')
    # time-aware trailing arithmetic acceleration mean, then da_filtered/dt.
    filtered=[];k=0
    for i in range(n):
        while k<i and t[i]-t[k]>c.filter_window_s+1e-9:k+=1
        filtered.append(sum(a[k:i+1])/(i-k+1))
    jerk=[(filtered[i]-filtered[i-1])/(t[i]-t[i-1]) for i in range(1,n)]
    put('jerk_peak_mps3',max(map(abs,jerk)),'m/s^3',note=f'Trailing {c.filter_window_s:g} s arithmetic mean, full record including stop')
    put('jerk_rms_mps3',math.sqrt(sum(j*j*(t[i+1]-t[i]) for i,j in enumerate(jerk))/(t[-1]-t[0])),'m/s^3',note='Time-weighted RMS, same full-record window')
    start=next((i for i,r in enumerate(rows) if r['brake_pedal_pct']>0 and v[i]>c.min_speed_mps),None)
    stop=next((i for i in range((start or 0)+1,n) if v[i]<=1e-9),None) if start is not None else None
    if start is None or stop is None:
        put('stop_time_s',None,'s','NOT_EVALUATED','No complete brake-to-zero event');put('stop_distance_m',None,'m','NOT_EVALUATED','No complete brake-to-zero event')
    else:
        distance=sum(.5*(v[i]+v[i+1])*(t[i+1]-t[i]) for i in range(start,stop-1))
        i=stop-1;h=t[stop]-t[i]
        # A zero endpoint may include a standstill tail: infer crossing from incoming accel.
        tau=min(h,v[i]/-a[i]) if a[i]<0 else h
        distance+=.5*v[i]*tau
        put('stop_time_s',t[i]+tau-t[start],'s',note='First brake pedal >0 to zero; last interval constant-decel interpolation')
        put('stop_distance_m',distance,'m',note='Observed speed trapezoids, partial final stop interval')
    if all(h in rows[0] for h in ('regen_force_N','friction_force_N')):
        # Outgoing zero-order-held forces; linear speed per integration interval.
        wr=wf=0.
        for i in range(n-1):
            h=t[i+1]-t[i]
            if v[i+1]<=1e-9 and v[i]>0 and a[i]<0:h=min(h,v[i]/-a[i])
            ds=.5*(v[i]+v[i+1])*h
            wr+=rows[i]['regen_force_N']*ds;wf+=rows[i]['friction_force_N']*ds
        put('regen_work_J',wr,'J',note='Mechanical retarding work, not battery energy')
        put('regen_work_fraction',wr/(wr+wf) if wr+wf>0 else None,'1','EVALUATED' if wr+wf>0 else 'NOT_EVALUATED')
    else:
        put('regen_work_J',None,'J','NOT_EVALUATED');put('regen_work_fraction',None,'1','NOT_EVALUATED')
    return metrics

def analyze(rows,sid='input',config=None):
    c=(config or MonitorConfig()).validate();rules=evaluate(rows,sid,c)
    invalid=any(x['rule_id'] in ('MON-001','MON-002','MON-003','MON-004') and x['status'] in ('WARN','FAIL') for x in rules)
    metrics=reference_metrics(rows,c) if not invalid else {'analysis_withheld':dict(value=None,units='',status='NOT_EVALUATED',definition='Timing/range/coherence concern: inspect input, do not trust KPIs')}
    return dict(scenario_id=sid,config=asdict(c),config_sha256=c.sha256,rules=rules,metrics=metrics)
