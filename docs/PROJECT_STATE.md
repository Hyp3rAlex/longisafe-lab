# M1 release state — 13 November 2026

**VALIDATED CORE / READY FOR PUBLIC REPOSITORY REVIEW:** Offline educational Python/Excel/VBA longitudinal-dynamics analysis, fixed synthetic dataset suite, actual user-side Excel compilation/smoke test, independent Excel-vs-Python numerical checks, exported reports and source synchronization. The workbook is the unchanged successfully tested binary; it and its nine source modules match exactly.

**IMPLEMENTED:** Python plant/scenario generation, separate fault truth, CSV validation, MON-001..005/010, stop and jerk/braking-work metrics, native Excel charts and PDF/CSV exports, input SHA-256 and config digests, traceability and test case suite.

**TESTED:** 43 Python tests +18 subtests, 10/10 Windows Excel smoke expected outcomes, 9/9 cross-language comparisons, 10001-row scenario; no real-world telemetry.

**OPEN:** Excel UI/report PDF formatting polish, complete independent rerun on a second machine, removal of incidental Debug.Print logging after revalidation, M2/M3 features, choice of public license.

**EXTERNAL ACTIONS:** Public GitHub creation and push have NOT happened automatically; user must create/push repository or connect an authorized GitHub integration. No BMW application has been submitted.
