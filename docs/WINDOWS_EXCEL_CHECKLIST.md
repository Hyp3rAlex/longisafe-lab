# Windows Excel — reproduce the validated M1 smoke test

**Validated reference:** `excel/LongiSafeLab.xlsm` is the byte-for-byte workbook tested on Windows by the project owner. Its nine source modules are under `excel/vba/` and match the workbook source exactly. The workbook's last loaded dataset is `PERF_10001`, so its initial report may show a legitimate zero-deceleration flat line.

## Use the included workbook

1. Extract/download the repository to a local Windows folder. Open **desktop Excel** (VBA7); Excel Online and macOS are not supported by the Windows CNG hash module.
2. Open `excel/LongiSafeLab.xlsm`. Enable macros only after reviewing and trusting the project; do not disable Office macro security globally. Windows may block downloaded macro files. Use a trusted local location according to your organization's security policy.
3. Press `Alt+F8`, run `SmokeTestM1`, and choose `data/synthetic/S05.csv` to locate the test dataset directory.
4. The macro runs ten fixtures and creates/updates a `smoke_exports/` folder **next to the workbook**. Preserve any preexisting run evidence before rerunning.
5. Inspect `SmokeResults`; nine valid cases should show `MATCH` and S11 should show an expected timestamp rejection. Open `smoke_exports/S05/summary.pdf` and `smoke_exports/S14/summary.pdf`.
6. From the repository root, run `python scripts/check_all_excel_exports.py excel/smoke_exports` (adjust path to the actual generated directory). Expect no failures.
7. Run `python scripts/vba_source_sync.py` to verify the exported VBA sources match the included workbook.

## Regenerate the workbook from source (optional)

Create a blank Macro-Enabled Workbook (`.xlsm`), then import the **nine** `.bas` modules from `excel/vba` using the Visual Basic Editor (Alt+F11, File > Import File). Compile via Debug > Compile VBAProject and save. Use `SetupLongiSafe` to build the user interface, then `SmokeTestM1` as above. This builds a new workbook; it will have a different binary hash from the included validated reference until separately tested.

## Normal workflow

`Alt+F8 > ImportOnly` selects one CSV. `Alt+F8 > AnalyzeActive` performs monitors, metrics, charts and report. `Alt+F8 > ExportActive` exports a PDF and four numerical CSVs. `CompareActive` appends a labeled metrics snapshot. `ResetActive` clears outputs. The spreadsheet is a snapshot interface: editing `RawData` cells does not modify in-memory data; edit the source CSV and re-import.

**Known limitations:** M2 monitors MON-006..009 are not implemented; Excel smoke `MATCH` is source-level expected behavior, independently corroborated by the checker on the archived nine real outputs. Run-specific timestamps, path strings and performance times change between Windows environments.
