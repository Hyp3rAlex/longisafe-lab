"""Forward-only point-mass plant; instantaneous outputs, frozen-force time steps."""
from dataclasses import dataclass, asdict
import math, random

@dataclass(frozen=True)
class Vehicle:
    mass_kg: float=1800.
    rho_kgpm3: float=1.225
    cda_m2: float=.65
    crr: float=.012
    mu: float=.8
    regen_max_N: float=4500.
    charge_power_W: float=60000.
    fade_speed_mps: float=2.
    g: float=9.81

    def validate(self):
        for k,v in asdict(self).items():
            if not math.isfinite(v) or v<0: raise ValueError(k)
        if min(self.mass_kg,self.g,self.fade_speed_mps)<=0: raise ValueError('mass/g/fade must be positive')
        return self

def resistance(v,grade,p):
    return .5*p.rho_kgpm3*p.cda_m2*v*v + p.crr*p.mass_kg*p.g*math.cos(grade) + p.mass_kg*p.g*math.sin(grade)

def regen_capacity(v,soc,temp,p):
    # Illustrative piecewise factors, not battery/motor characterization.
    soc_factor=max(0.,min(1.,(100.-soc)/10.))
    temp_factor=max(0.,min(1.,(temp+10.)/30.))
    return min(p.regen_max_N,p.charge_power_W/max(v,1.))*soc_factor*temp_factor*min(1.,v/p.fade_speed_mps)

def allocate(v,request,drive,grade,soc,temp,p,loss=False,compensate=True,friction_scale=1.):
    envelope=max(0.,p.mu*p.mass_kg*p.g*math.cos(grade))
    needed=max(0.,p.mass_kg*request+drive-resistance(v,grade,p)) if request>0 else 0.
    baseline=min(needed,regen_capacity(v,soc,temp,p),envelope)
    regen=0. if loss else baseline
    target=needed-regen if compensate else needed-baseline
    friction=min(max(0.,target)*friction_scale,max(0.,envelope-regen))
    return regen,friction,regen_capacity(v,soc,temp,p)

def step(v,a,dt):
    """Return v_next, distance, elapsed-to-stop or None. No reverse motion."""
    if v<=0 and a<=0: return 0.,0.,None
    if a<0 and v+a*dt<=0:
        tau=v/-a
        return 0.,v*tau+.5*a*tau*tau,tau
    return max(0.,v+a*dt),v*dt+.5*a*dt*dt,None

def simulate(cfg,p=None):
    allowed={'kind','seed','dt_s','duration_s','v0_mps','brake_start_s','decel_mps2','accel_noise_std','grade_rad','soc_pct','battery_temp_C','loss_time_s','compensate','friction_delay_until_s','fault','core_only','vehicle'}
    if set(cfg)-allowed:raise ValueError('Unknown scenario parameter: '+str(set(cfg)-allowed))
    if cfg.get('kind') not in ('brake','steady','coast'):raise ValueError('Unknown scenario kind')
    for key,value in cfg.items():
        if key not in ('kind','fault','compensate','core_only','vehicle'):
            if isinstance(value,bool) or not isinstance(value,(float,int)) or not math.isfinite(value):raise ValueError('Invalid scenario numeric parameter: '+key)
    for key in ('compensate','core_only'):
        if key in cfg and not isinstance(cfg[key],bool):raise ValueError(key+' must be Boolean')
    if cfg.get('fault') not in (None,'bias','stuck_at','timing_duplicate'):raise ValueError('Unsupported fault class in M1')
    if not isinstance(cfg.get('seed',1),int):raise ValueError('Seed must be integer')
    if cfg.get('decel_mps2',3)<0 or cfg.get('accel_noise_std',.015)<0:raise ValueError('Negative demand/noise')
    if not 0<=cfg.get('soc_pct',60)<=100 or abs(cfg.get('grade_rad',0))>=math.pi/2:raise ValueError('SOC/grade domain')
    if any(cfg.get(k,0)<0 for k in ('brake_start_s','loss_time_s','friction_delay_until_s')):raise ValueError('Negative event time')
    p=(p or Vehicle(**cfg.get('vehicle',{}))).validate()
    dt=cfg.get('dt_s',.05);duration=cfg.get('duration_s',35.);v=cfg.get('v0_mps',25.)
    if dt<=0 or duration<=0 or v<0: raise ValueError('Invalid simulation timing/speed')
    rng=random.Random(cfg.get('seed',1)); t=0.; distance=0.;rows=[]; exact_stop=None
    while t<=duration+1e-9:
        braking=cfg['kind']=='brake' and t>=cfg.get('brake_start_s',1.)-1e-9
        request=cfg.get('decel_mps2',3.) if braking and v>0 else 0.
        grade=cfg.get('grade_rad',0.);drive=0.
        if cfg['kind']=='steady': drive=resistance(v,grade,p)
        loss=braking and t>=cfg.get('loss_time_s',1e9)-1e-9
        scale=0. if braking and t<cfg.get('friction_delay_until_s',-1.) else 1.
        regen,friction,limit=allocate(v,request,drive,grade,cfg.get('soc_pct',60.),cfg.get('battery_temp_C',25.),p,loss,cfg.get('compensate',True),scale)
        a=(drive-regen-friction-resistance(v,grade,p))/p.mass_kg
        if v==0 and a<0: a=0.
        rows.append(dict(time_s=round(t,10),vehicle_speed_mps=v,accel_long_mps2=a+rng.gauss(0,cfg.get('accel_noise_std',.015)),accelerator_pedal_pct=20. if drive else 0.,brake_pedal_pct=40. if request else 0.,road_grade_rad=grade,drive_force_N=drive,regen_force_N=regen,friction_force_N=friction,requested_decel_mps2=request,regen_limit_N=limit,soc_pct=cfg.get('soc_pct',60.),battery_temp_C=cfg.get('battery_temp_C',25.),estimated_mu=p.mu))
        if t>=duration-1e-9: break
        h=min(dt,duration-t)
        # Do not straddle known discontinuities with an old actuator state.
        for boundary in (cfg.get('brake_start_s',1.),cfg.get('loss_time_s',1e9),cfg.get('friction_delay_until_s',-1.)):
            if t+1e-9<boundary<t+h-1e-9: h=boundary-t
        vn,ds,tau=step(v,a,h)
        if tau is not None and tau<h-1e-9:
            h=tau
            if exact_stop is None and braking:exact_stop={'time_s':t+h,'distance_m':distance+ds}
        elif tau is not None and exact_stop is None and braking:exact_stop={'time_s':t+h,'distance_m':distance+ds}
        distance+=ds;v=vn;t+=h
    return rows,dict(exact_stop=exact_stop,total_distance_m=distance,parameters=asdict(p))
