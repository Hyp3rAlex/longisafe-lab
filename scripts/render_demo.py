"""Generate genuine Python charts and a one-page synthetic analysis report."""
from pathlib import Path
import sys,json,hashlib
from datetime import datetime,timezone
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from longisafe.csv_schema import load_csv
from longisafe.analysis import analyze
from longisafe import __version__
ROOT=Path(__file__).resolve().parents[1]
plt.rcParams.update({'font.size':10,'axes.spines.top':False,'axes.spines.right':False,'axes.grid':True,'grid.alpha':.2,'font.family':'DejaVu Sans'})
colors={'S04':'#546578','S05':'#167e8a','S14':'#bd493d'}
fig,axes=plt.subplots(2,2,figsize=(12.5,7.4),layout='constrained')
results={};data={}
for sid in ['S04','S05','S14']:
    rows=load_csv(ROOT/'data/synthetic'/f'{sid}.csv');data[sid]=rows;results[sid]=analyze(rows,sid)
    t=[r['time_s'] for r in rows];v=[r['vehicle_speed_mps'] for r in rows]
    axes[0,0].plot(t,v,label=sid,color=colors[sid],lw=1.8)
    if sid!='S04':axes[0,1].plot(t,[-r['accel_long_mps2'] for r in rows],label=sid,color=colors[sid],lw=1.1)
axes[0,1].plot([r['time_s'] for r in data['S14']],[r['requested_decel_mps2'] for r in data['S14']],color='#202a36',ls='--',lw=1.2,label='Net request (S14)')
for sid,style in [('S05','-'),('S14','--')]:
    r=data[sid];axes[1,0].plot([x['time_s'] for x in r],[x['regen_force_N'] for x in r],color=colors[sid],ls=style,label=f'{sid} regen')
    axes[1,0].plot([x['time_s'] for x in r],[x['friction_force_N'] for x in r],color=colors[sid],ls=':' if sid=='S05' else '-.',label=f'{sid} friction')
r=load_csv(ROOT/'data/synthetic/S02.csv')
axes[1,1].plot([x['time_s'] for x in r],[x['vehicle_speed_mps'] for x in r],color='#167e8a',lw=2)
axes[1,1].text(.04,.82,'Accelerator = 0% | Grade = -0.06 rad\nMON-003 PASS; no zero-pedal fault heuristic',transform=axes[1,1].transAxes,fontsize=9)
for ax,title,ylabel in zip(axes.flat,['Compensation controls stopping distance','Uncompensated loss creates net-decel undershoot','Synthetic brake allocation (not measured telemetry)','S02: downhill coasting is physically consistent'],['Speed [m/s]','Deceleration [m/s²]','Retarding force [N]','Speed [m/s]']):
    ax.set(title=title,xlabel='Time [s]',ylabel=ylabel)
for ax in [axes[0,0],axes[0,1],axes[1,0]]:ax.legend(fontsize=8,loc='best')
axes[0,1].axvspan(3.0,3.3,color='#bd493d',alpha=.12,label='Evidence accumulation')
fig.suptitle('LongiSafe Lab | SYNTHETIC DATA / EDUCATIONAL TOOL',fontsize=15,fontweight='bold')
image=ROOT/'reports/engineering_comparison.png';fig.savefig(image,dpi=165);plt.close(fig)
# An actual exported PDF, clearly Python, not an invented Excel screenshot.
pdf=ROOT/'reports/example_summary.pdf';c=canvas.Canvas(str(pdf),pagesize=(842,595));c.setTitle('LongiSafe Lab - synthetic M1 analysis')
c.setFillColor(HexColor('#142b40'));c.rect(0,532,842,63,fill=1,stroke=0)
c.setFillColor(HexColor('#ffffff'));c.setFont('Helvetica-Bold',18);c.drawString(28,565,'LongiSafe Lab | synthetic M1 comparison')
c.setFont('Helvetica',10);c.drawString(28,547,'Python-generated analysis evidence. Desktop Excel / VBA execution remains UNVERIFIED.')
c.drawImage(str(image),27,183,width=560,height=332,preserveAspectRatio=True)
c.setFillColor(HexColor('#172e40'));c.setFont('Helvetica-Bold',10);c.drawString(608,505,'Observed results')
y=483
for sid in ['S04','S05','S14']:
    m=results[sid]['metrics'];distance=m['stop_distance_m']['value'];stop=m['stop_time_s']['value']
    c.setFont('Helvetica-Bold',11);c.drawString(608,y,sid);c.setFont('Helvetica',10)
    c.drawString(608,y-17,f'Stop distance: {distance:.2f} m');c.drawString(608,y-31,f'Stop time: {stop:.2f} s');y-=69
c.setFont('Helvetica',9)
event=next(x for x in results['S14']['rules'] if x['rule_id']=='MON-005' and x['status']=='FAIL')
for i,line in enumerate(['MON-005: S05 PASS; S14 FAIL',f"S14 evidence: {event['t_start_s']:.2f} to {event['t_end_s']:.2f} s",f"First detection: {event['detection_time_s']:.2f} s",'Raw source has exact evidence rows.','MON-006..009: NOT_EVALUATED','S02 zero-pedal coast: no FAIL']):c.drawString(608,265-i*14,line)
c.setFont('Helvetica',8)
lines=['Model: forward-only point mass, m=1800 kg, CdA=0.65 m2, Crr=0.012, rho=1.225 kg/m3, mu=0.8; grade positive uphill.',
       'All parameters/thresholds are illustrative, not BMW-specific. Frozen-force steps <=0.05 s with an inserted exact stop node.',
       'S05: high SOC plus regen loss, friction substitutes. S14: regen loss, no extra friction. These have different initial regen splits.',
       'No real telemetry, vehicle control, compliance assessment, ASIL allocation or population-level accuracy claim.',
       'Limitations: point-mass model, no ABS/ESC, sensor bias can mimic decel undershoot, full-record jerk includes the stop transient.',
       f'Version {__version__} | UTC {datetime.now(timezone.utc).isoformat(timespec="seconds")} | Config SHA-256 {results["S14"]["config_sha256"]}',
       'S14 CSV SHA-256 '+hashlib.sha256((ROOT/'data/synthetic/S14.csv').read_bytes()).hexdigest(),
       'Source: supplied authoritative specification; this report uses only generated CSVs. Rule evidence: reports/S14/rule_results.csv.']
for i,line in enumerate(lines):c.drawString(28,155-i*16,line)
c.save();print(image);print(pdf)
