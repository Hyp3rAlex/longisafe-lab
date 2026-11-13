# Engineering assumptions and numerical definitions

All data are SYNTHETIC. The model is an educational forward-only 1D point mass, not a BMW vehicle model. No real control or actuation is provided.

## Coordinates and forces

Speed v is nonnegative forward [m/s]; acceleration a is positive forward [m/s²]; road angle θ [rad] is positive uphill. Regen R and friction B are nonnegative retarding magnitudes [N].

`m*a = drive - R - B - 0.5*rho*CdA*v² - Crr*m*g*cos(theta) - m*g*sin(theta)`.

The requested channel is a *net deceleration target*, not pedal position or raw actuator force. `brake_needed=max(0, m*d_request + drive - resistance)` when request>0. Uphill gravity reduces braking demand; downhill gravity increases it. For zero request the allocator produces no brakes. Aero/rolling resistance remain present while moving. At v=0 negative acceleration is clamped to zero; reverse motion is outside scope. A downhill coast can start moving from rest. Hill-hold control is not modeled.

Illustrative defaults: m=1800 kg, rho=1.225 kg/m³, CdA=0.65 m², Crr=0.012, mu=0.8, g=9.81 m/s², nominal dt=0.05 s. Edit Vehicle parameters and scenario configs to explore them; JSON config files preserve every generated run. No parameter represents a BMW specification.

## Braking allocation

Traction envelope = `mu*m*g*cos(theta)`, simplified total braking limit, not a tire/ABS model. Regen capacity = `min(4500 N, 60000 W/max(v,1 m/s))*soc_factor*temperature_factor*speed_factor`, then clipped to demand/envelope. SOC factor = clipped `(100-SOC)/10`; temperature factor = clipped `(T+10)/30`; speed factor = clipped `v/2`. These shapes are assumptions, not a validated battery/motor acceptance map. Friction takes remaining demand within the total envelope.

S05 uses SOC 98%, already reducing initial regen; complete regen loss at 3s is compensated by friction. S14 uses SOC 60%, complete regen loss at 3s, and **no additional friction substitution for the lost capability**. Its scheduled nominal friction can still rise as the nominal low-speed regen limit fades. These scenarios have different initial brake splits. Both deliver the same requested net deceleration before loss. To isolate only the compensation flag use otherwise identical configs (covered by robustness tests). Low friction can also cause undershoot: MON-005 indicates request tracking mismatch, not a diagnosis of malfunction.

## Integration and observation alignment

Each CSV row is a state at its timestamp, with acceleration/forces applying to the outgoing interval. Forces are recomputed each step and held constant until the next state. Integration uses `v_next=v+a*h`, `ds=v*h+0.5*a*h²`. Known request/loss boundaries are inserted if not on the nominal grid. If a step would cross zero, use τ=v/(-a), insert an exact stop node at t+τ, integrate only the triangular area, then hold zero under nonpositive acceleration. Negative speeds are never integrated. Final hold records may be irregularly spaced because of the exact stop node.

Measurement acceleration has seeded Gaussian noise σ=0.015 m/s² by default. Speed is noiseless in demo fixtures so observed stop timing remains unambiguous. No inference about robustness to real sensor errors follows. S09 adds +1.5 m/s² at 3–5s; S10 freezes measured speed at its 3s value over 3–5s; S11 duplicates CSV row 82's timestamp. All injected original samples and seed/config are separate from input CSVs.

## Monitor definitions and evidence

- Strict importer checks exact required headers, numeric finite values, percentages, forward speed, optional domains, and strictly increasing time. Rejected files cannot leave previous usable results active.
- MON-001 flags gaps exceeding 3*declared nominal dt; warning withholds subsequent metrics/monitors rather than bridging missing evidence.
- MON-002 also screens speed>100m/s or |a|>20m/s² (Config). These are pedagogical domain limits.
- MON-003 uses the most recent **complete** >=0.2s window. Compare `(v_i-v_k)/(t_i-t_k)` with time-weighted outgoing acceleration over the same interval. Residual>0.7m/s² must persist for >=0.3s. A complete window adds up to one sample of latency on irregular grids; persistence is timestamp based.
- MON-004 uses the same window: near-constant net speed difference <=0.005m/s while mean independent measured |a|>0.5m/s². Persist >=0.3s. A steady-state frozen signal is unobservable; oscillating speed with zero net change can also confound this simplified test.
- MON-005 uses `max(0, requested_decel + a_measured)` (undershoot only), threshold 0.6m/s² sustained >=0.3s. Active at request>0.2m/s² and speed>0.2m/s, after 0.2s request onset grace. Net request already includes grade, so no invented slope correction. A sensor bias can trigger both MON-003 and MON-005; that is correlated evidence, not two independent faults.
- MON-006..009 explicitly NOT_EVALUATED in M1. There is **no** zero-accelerator fault heuristic. S02 is checked for kinematic consistency, without claiming proof of all braking-demand behavior.
- MON-010 inventories optional context; missing context -> NOT_EVALUATED. All M2 rules stay NOT_EVALUATED even when context exists.

Evidence rows are CSV rows (header row=1), inclusive start/end times; detection time is first time duration criterion is met. Separate events are separate result rows. PASS means no sustained configured breach, never verified vehicle safety. Optional missing/too short/invalid inputs return NOT_EVALUATED.

## Metrics, filters and limitations

First brake_pedal_pct>0 with speed>0.2m/s starts the reported event; first observed zero ends it. No complete event -> NOT_EVALUATED. Stop distance uses speed trapezoids. Final zero endpoint may include hold, so τ=min(dt, v_prev/(-a_prev)) when incoming acceleration is negative, and last distance=0.5*v_prev*τ. Noise on a_prev can slightly shorten the inferred event; tolerances are declared in crosschecker. Demo exact-stop nodes constrain that error. This is not a general noisy-zero-speed estimator or 100–0 benchmark.

Peak deceleration is raw measured maximum (-a), full record, with its timestamp. Jerk uses a trailing arithmetic sample mean of acceleration over <=0.2s, then differences on actual timestamps; RMS is time weighted. Full-record peak/RMS includes initial filter transients and terminal stop. Do not present it as Soft Stop quality. MON-003's time-weighted acceleration integral and the jerk sample mean serve different purposes and are documented separately.

Mechanical regen work uses outgoing force * observed trapezoidal distance for every interval (same partial-stop correction for a zero endpoint with a hold tail); regen fraction = W_regen/(W_regen+W_friction), not sample share and not battery energy. Exported definition windows are explicit. M1 has no transition/terminal comfort window detector, no brake-delay KPI, no ADAS request arbitration, no formal sensor fault isolation.
