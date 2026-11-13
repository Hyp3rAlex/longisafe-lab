# M1 CSV contract

One scenario per file. Input is UTF-8 with optional BOM, comma-delimited, dot decimals, LF or CRLF. M1 intentionally supports only ASCII **unquoted numeric fields**, including decimal/exponent notation. Quotes, text fields, empty cells/lines, duplicate/unknown headers, malformed widths, NaN/Inf/overflow, negative/nonincreasing time and invalid core domains are rejected. No implicit unit conversion. Header order may vary. Magnitudes >1e100 rejected. Numeric syntax is checked before VBA Val. `vehicle_mode` text is deferred: rejected, never converted to zero.

| Field | Unit | Availability/domain |
|---|---|---|
| time_s | s | Required, >=0, strictly increasing |
| vehicle_speed_mps | m/s | Required, >=0 |
| accel_long_mps2 | m/s² | Required, signed forward acceleration |
| accelerator_pedal_pct | % | Required, 0–100 |
| brake_pedal_pct | % | Required, 0–100; no force calibration |
| road_grade_rad | rad | Optional, positive uphill, abs<pi/2 |
| drive_force_N | N | Optional, nonnegative synthetic propulsion |
| regen_force_N | N | Optional, nonnegative retarding magnitude |
| friction_force_N | N | Optional, nonnegative retarding magnitude |
| requested_decel_mps2 | m/s² | Optional, >=0 desired *net* deceleration |
| regen_limit_N | N | Optional, >=0 illustrative capability |
| soc_pct | % | Optional, 0–100 |
| battery_temp_C | °C | Optional, finite |
| estimated_mu | 1 | Optional, 0–2 |
| adas_brake_request_mps2 | m/s² | Optional >=0; accepted but not analyzed in M1 |

Absent optional columns remain absent/Empty; never synthesize them from pedal input. Invalid optional numeric data are rejected. Legacy brake_bar is pressure, not brake_pedal_pct: legacy CSVs are intentionally incompatible, preserved as-is in the original ZIP.

CSV report outputs use standard escaping/quoted fields, unlike restricted numeric inputs. Untrusted strings beginning with =,+,-,@ (after whitespace) are prefixed with apostrophe before writing/exporting. Numeric negatives remain numeric. Native VBA output is ASCII content in Windows text encoding; non-ASCII local file paths are informational and may not survive text export in every locale. Analysis uses numeric input ASCII subset. VBA source files use ASCII with CRLF for portable VBE import.

Source SHA-256 is over exact input bytes, BOM/newlines included. Python config SHA-256 is over sorted JSON with default separators. VBA config SHA-256 is over ordered ASCII key=value LF using invariant Str numeric spelling; these intentionally have different encodings, so compare numeric config exports, not hash strings to each other. VBA hashes use native Windows CNG, pending Windows runtime verification.
