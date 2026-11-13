# Workbook / VBA source synchronization

**Source of truth:** user-uploaded, successfully smoke-tested `LongiSafeLab.xlsm` (12 Nov 2026). The nine standard VBA modules in `excel/vba/` were recovered **byte for byte** from the workbook's embedded `xl/vbaProject.bin` and checked against the workbook. These are not the obsolete sources in the early M1 candidate ZIP.

```bash
python scripts/vba_source_sync.py
python scripts/verify_repo.py
```

The extractor reads OLE compound streams and decompresses the VBA source; it does not execute macros. See `workbook_source_manifest.json` for the SHA-256 hashes. The release includes the Windows-tested workbook **unchanged**; no unsupported claim is made that Excel recompiled an edited copy in this environment.

**Intentional decision:** `modMain` and `modSmoke` still contain temporary `Debug.Print` diagnostics used to locate the initial smoke-test failures. These statements have no analytical/calculation effect and were present in the tested workbook. Keeping them allows an exact source/workbook match. They may be removed in a later patch only after the workbook is recompiled, rerun, and the updated `.bas` files re-exported.

**Fixed before validation:** `modReporting` uses clipboard-free chart generation (avoids intermittent `ChartObject.Copy` failures), preserves `Empty` Excel cells as blank CSV fields, and guards against one-cell worksheets. These changes were tested through a 10-fixture Excel smoke run and independent comparison of nine export datasets.

**Why a flat chart can appear:** the saved workbook's last smoke scenario is `PERF_10001`, a constant-speed dataset without braking demand. Open S05/S14 explicitly to see braking curves.

**Version relationship:** `excel/LongiSafeLab.xlsm`, its nine `.bas` sources, and the Excel `evidence/excel_smoke/` outputs are evidence from the same M1 branch. They are not signed releases, certified analysis or real-car data.
