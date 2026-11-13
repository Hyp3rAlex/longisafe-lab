# Requirements and validation evidence

| Requirement | Implementation | Scenario/test | Evidence/status |
|---|---|---|---|
| REQ-DATA-001 strict time | csv_schema.load_csv / LoadCsv, MON-001 gap | S11; duplicate/out-of-order/gap tests | Python execution tested; VBA runtime pending |
| REQ-SCHEMA-001 strict SI input | numeric grammar / Num | Contract malformed-input subtests | No coercion/zero filling |
| REQ-PLAUS-002 independent coherence | integrated-a vs delta-v, MON-003/004 | S01/S09/S10; hand-solvable observation oracle | reports/S09 and S10 rule_results.csv |
| REQ-BRAKE-003 net demand tracking | MON-005 | S04/S05/S14/MISSING_CONTEXT | S05 PASS, S14 FAIL, missing NOT_EVALUATED |
| REQ-CONTEXT-004 avoid naive zero-pedal diagnosis | no zero-pedal heuristic; MON-006 deferred | S02, positive downhill acceleration | No FAIL on this input; not full MON-006 implementation |
| REQ-STOP-001 complete event / interpolated stop | reference_metrics / ComputeMetrics | closed-form v0/a, v0²/(2a), manual partial step | Independent numerical tests |
| REQ-REGEN-001 mechanical work fraction | force*distance integration | manual 300J/100J oracle | 75% fraction, not 50% sample share |
| REQ-EVID-005 trace every anomaly | rule ID, csv row, time, measured, limit, rationale, hashes | all reports | Original input/truth separate |
| REQ-EXPORT-001 independent comparison | check_excel_exports.py | altered metric/rule fixtures rejected | Checker plumbing tested on Python fixtures; actual VBA comparison pending |
| REQ-REPORT-001 readable traceable evidence | render_demo.py / MakeReport | Python PDF/plot rendered and inspected | Excel PDF requires user inspection |
| REQ-PERF-001 10000+ rows | in-memory arrays | PERF_10001 (10001 rows) | Python observed timing in suite log; Excel timing pending |

## Scenario status (full specification preserved in AUTHORITATIVE_SPECIFICATION.md)

| ID | Status | Expected diagnostic behavior |
|---|---|---|
| S01 | Generated/tested Python | Steady flat drive, no FAIL |
| S02 | Generated/tested Python | Downhill, no pedal, speed rises consistently, no FAIL |
| S03 | Not delivered | Uphill coast; signed grade verified by unit test only |
| S04 | Generated/tested Python | Nominal braking/stop/regen work, no FAIL |
| S05 | Generated/tested Python | High SOC then compensated regen loss; MON-005 PASS |
| S06 | Not delivered | Cold/low-speed compensated transition; capacity has unit tests |
| S07 | Not delivered | Traction limitation distinguished from corrupted input; force cap unit-tested |
| S08 | Not delivered | Smooth jerk-limited stop; current profiles are NOT jerk-limited |
| S09 | Generated/tested Python | Bias => MON-003 FAIL; MON-005 may also flag sensor-contaminated undershoot |
| S10 | Generated/tested Python | Stuck speed => MON-003 and MON-004 FAIL |
| S11 | Generated/tested Python | Duplicate timestamp rejected, no KPIs |
| S12 | Not delivered | Spike detector/fault class deferred |
| S13 | Not separately delivered | S05 demonstrates compensated loss; no separate regen abnormality monitor |
| S14 | Generated/tested Python | Uncompensated loss => MON-005 FAIL, longer complete stop |
| S15 | Not delivered | Friction delay has physical unit test, no scenario/report yet |
| S16 | Not delivered | Small accel noise exists; no full timestep-jitter/noise study |

MISSING_CONTEXT tests core-only observability; PERF_10001 tests scale. Neither is a replacement for any S01–S16 case. Three alternate parameter/seed pairs are regression robustness checks after defaults were set, not an independent dataset or population-level validation.

## Illustrative FMEA-inspired table (not formal HARA/FMEA compliance)

| Failure mode | Potential effect | Observable signals | Rule | Prerequisite / limitation | Case |
|---|---|---|---|---|---|
| Speed freezes during braking | Misleading motion estimate | v,a,t | MON-003/004 | Independent a needed; steady-state freeze may be invisible | S10 |
| Accel bias | Misleading apparent response | v,a,t | MON-003 | Signal inconsistency, not isolated root cause; can mimic MON-005 | S09 |
| Regen lost, friction substitutes | Braking maintained | request,a,regen,friction | MON-005 + work split | No lost-braking claim; no dedicated regen-health monitor | S05 |
| Regen lost, no extra friction | Decel undershoot/longer stop | request,a | MON-005 | Request and independent sensors needed; low mu can also undershoot | S14 |
| Duplicate timestamp | Invalid derivatives/KPIs | t | Import/MON-001 | Reject before calculation | S11 |
