from dataclasses import dataclass, asdict
import hashlib, json, math

@dataclass(frozen=True)
class MonitorConfig:
    nominal_dt_s: float = .05
    gap_factor: float = 3.
    filter_window_s: float = .2
    evidence_s: float = .3
    coherence_mps2: float = .7
    stale_speed_mps: float = .005
    motion_accel_mps2: float = .5
    decel_error_mps2: float = .6
    response_delay_s: float = .2
    min_request_mps2: float = .2
    min_speed_mps: float = .2
    max_speed_mps: float = 100.
    max_abs_accel_mps2: float = 20.

    def validate(self):
        for k,v in asdict(self).items():
            if isinstance(v,bool) or not isinstance(v,(int,float)) or not math.isfinite(v) or v <= 0:
                raise ValueError(f'Invalid config: {k}')
        if self.gap_factor <= 1 or self.max_speed_mps <= self.min_speed_mps:
            raise ValueError('Invalid config ordering')
        return self

    @property
    def sha256(self):
        return hashlib.sha256(json.dumps(asdict(self),sort_keys=True).encode()).hexdigest()
