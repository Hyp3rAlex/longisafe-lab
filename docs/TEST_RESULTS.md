# Actual M1 validation evidence

## Dates and provenance

- **12 November 2026:** project owner compiled the VBA project and ran `SmokeTestM1` in desktop Microsoft Excel for Windows after installing the corrected clipboard-free reporting module. The 10-case `SmokeResults` sheet reported ten expected outcomes matched, including the deliberate S11 timestamp rejection. The genuine workbook and exported files were uploaded for review.
- **13 November 2026:** Python checker rerun on the archived genuine Excel exports: all nine valid scenario exports **PASS**, with zero reported failures. Run-specific SHA-256 values, rule evidence, times and metrics were compared according to script tolerances.
- **13 November 2026:** Python regression suite rerun on the original M1 Python source: **43 tests passed, 18 subtests passed**. The supported tests run in this environment; this does not constitute safety certification or independent full requirements coverage.
- **13 November 2026:** recovered nine module sources from the actual saved XLSM's `vbaProject.bin` and byte-matched them to `excel/vba/`. The workbook has *not* been edited and rebuilt since the successful Windows test.

## Commands

```bash
python -m pytest -q
python scripts/check_all_excel_exports.py evidence/excel_smoke
python scripts/vba_source_sync.py
python scripts/verify_repo.py
```

Python version in the 2026-11-13 runner: 3.x with pytest available; full original Work environment details under `reports/test_environment.json`. Windows Excel build/version and 32/64-bit specificity were not independently archived; they should be added on the next Windows run.

## Results from user-side Excel

- **10/10 smoke cases:** S01, S02, S04, S05, S09, S10, S11, S14, MISSING_CONTEXT, PERF_10001 returned `MATCH` / the expected explicit rejection. Source `evidence/excel_smoke/smoke_results.csv`.
- **9/9 independent comparisons:** S01, S02, S04, S05, S09, S10, S14, MISSING_CONTEXT and PERF_10001; all checks returned `PASS` and empty failure lists. JSON: `evidence/validation/crosscheck_2026-11-13.json`. S11 is a rejected input and has no valid numerical report to cross-check.
- **S05:** stopping distance approximately **102.0555915 m**, stop time **8.2480587 s**, MON-005 **PASS**, with low regen-work share due to scenario's illustrative limitations and friction compensation.
- **S14:** stopping distance approximately **318.5889225 m**, stop time approximately **32.50 s**, MON-005 **FAIL**; demand-tracking evidence **3.00–33.30 s** and first detection **3.30 s**.
- **S02:** zero accelerator on downhill coast does not spuriously cause a `FAIL`.
- **PERF_10001:** 10,001-row smoke elapsed about **7.17 seconds** in the shown Windows run. This is scenario-level wall time including workbook operations, not a pure analysis performance benchmark.
- **Reports:** all nine real Excel scenario PDF files and four CSV exports per scenario were created; S05/S14 reports were visually checked. The PDF configuration hash label may visually truncate; the full hash is in `run_metadata.csv`.

## What is and is not proved

These results demonstrate operational Excel analysis and agreement with an independent Python reference **on the supplied fixed synthetic fixtures**. They do not prove real-vehicle safety, statistical detection performance, controller behavior, complete fault observability, ISO compliance, or a deployed production safety system. Monitor MON-006..009, eight additional numbered scenarios, transition/terminal comfort screening, force-residual checks and M2/M3 studies are not implemented. The test suite cannot attest to Excel execution on other machines; an independent Windows rerun remains recommended before using the workbook in a live interview.

## Reproducible artifacts

- `excel/LongiSafeLab.xlsm` — preserved real desktop-Excel workbook
- `excel/vba/*.bas` — byte-matched nine source modules
- `evidence/excel_smoke/` — actual user-side smoke CSVs and PDFs
- `evidence/validation/crosscheck_2026-11-13.json` — independent checker results
- `reports/engineering_comparison.png` and `reports/example_summary.pdf` — separately generated Python plots/report
- `docs/workbook_source_manifest.json` — workbook and source SHA-256 manifest
