# Verification Report — Models 09–12

**Date:** 2026-11-09 · **Status:** Four synthetic models completed. The repository was published to GitHub on 2026-11-12. Baseline checks and isolated negative tests pass in the spreadsheet calculation engine (independently re-verified in LibreOffice: 4/4 baselines PASS, 4/4 negative tests PASS → FAIL). **Desktop Excel/Power Query refresh is unverified and requires setup.**

## Scope and authority

The source of record for original issue definitions is `Fix_Verification_Report_Round12_50_Issues.md`. The Round 13 report is a contextual check (0 certified new fixes), **not** a replacement issue register. In particular **M13 is Model 07 target GL-to-QoE evidence**, not a tax change; the original Model 03 uses a **25% tax assumption**, not 21%. No edit is made to either.

The original eight models and Supporting Schedules template were **not rebuilt or modified**. They are packaged as byte-for-byte source copies. Round 12 remains **38/50 fixed in synthetic scope, 12 not fixed**, with all seven original Critical issues previously fixed. None of those 12 original issues is newly certified here.

## New model design / financial semantics

| Model | Input and measure grain | Version and aggregation decision | Output and important assumption |
|---|---|---|---|
| **09** | Three manager-owned BLs × four accounts × six months, one fact per business line/cost centre/account/month/Budget/Base | Positive revenue, direct cost, OPEX and capital budget values; monthly business-line rollup via `SUMIFS` of canonical Staging; CAPEX excluded from EBITDA | Total-company revenue, COGS, OPEX, EBITDA, CAPEX. A live source extract is not needed to review the synthetic example. |
| **10** | Three BLs × three P&L accounts × 12 months; separate Budget, Forecast and six closed-month Actual facts | `Version` is Actual/Budget/Forecast; `SnapshotID=FY27_Q2`. `Inputs!B12=2` means 2027 Q1–Q2 actual and Q3–Q4 selected forecast; Budget remains the comparator | Four quarters plus full-year chosen rolling revenue, direct cost, OPEX, EBITDA and revenue budget variance |
| **11** | Four products for March 2027, each with Budget and Actual quantity, unit price and sales amount | Exact additive decomposition: volume = Δtotal units × budget weighted-average unit price; mix = Σ[(actual product units – budget product units × actual total/budget total) × budget price]; price = Σ[actual units × (actual price − budget price)] | Budget 466,300 → volume +38,401.18 + mix +10,048.82 + price +16,010 = actual 530,760 CAD; stacked-column waterfall visualization |
| **12** | 13 Monday-start weeks from 2027-01-04; bank opening snapshot and transaction-level AR, AP, inventory and payroll due dates | Open-item due dates drive weekly receipts/disbursements. No double count: opening bank balance is a control snapshot, not a weekly receipt. Positive AP/inventory/payroll facts are subtracted in the engine. | Opening 180,000; AR receipts 182,000; outflows 198,800; 13-week closing 163,200 CAD; min cash 152,750 CAD vs 75,000 minimum |

## Shared data contract and dynamic staging

All four **newly built from blank workbooks** have the exact same staging header list in `Staging!A6:U6` (21 fields):

`RecordID, SourceID, BusinessLineID, CostCenterID, AccountID, ProductID, PeriodKey, Version, Scenario, SnapshotID, AmountCAD, Quantity, UnitPriceCAD, EventDate, DueDate, CounterpartyID, DocumentID, SourceRef, SourceFormat, QualityStatus, QualityMessage`.

The shared account dictionary has `REV_CORE, COGS_CORE, OPEX_SGNA, CAPEX_TOTAL, A_CASH, A_AR, A_AP, A_INVENTORY, A_PAYROLL`. Shared business lines: `BL01`, `BL02`, `BL03`, `BL_CORP`; cost centres: `CC10`, `CC20`, `CC30`, `CC90`. `PeriodKey` always uses `YYYY-MM` while weekly placement uses typed `DueDate`. The canonical source type is a **versioned fact**, not a physical database table. Exchanging facts with Models 03/05/08 still requires an explicit bridge/mapping; no cross-workbook reference is embedded.

**Parameters:** `Inputs!B7` RootFolder; `B8` ActiveSourceID; `B9` LoadMode `SINGLE` or `ALL`; `B10` Scenario; `B11` SnapshotID; `B12` closed quarter (Model 10). `Source Register!A:Q` identifies source, CSV/XLSX format, relative path, object name, wide/narrow grain, provided/missing dimensions, defaults/derivations and enabled status.

**M-code:** Full query text is in `Power Query` (qCanonicalStaging and qStagingErrors) in every workbook, plus eight identical-name companion `.pq` files, one pair per model. **Per-source** ten-step extraction plans live in `Extract Steps` in every workbook. Source extraction/normalization pipeline:

1. Read workbook `tblSourceParams`, find active source or enabled-source union and combine RootFolder + RelativePath (no code modification for new CSV/XLSX paths).
2. Read CSV with `Csv.Document`, or Excel `Sheet`/`Table` with `Excel.Workbook`; promote headers.
3. Apply case-insensitive source-scoped `tblHeaderMap` aliases; ignore unexpected extra fields. Wide files retain only period columns recognized in `tblPeriodMap` and unpivot month columns.
4. Select all canonical raw input columns using `MissingField.UseNull`; defaults are applied **only for missing raw values** as specified in Source Register. Invalid populated codes never silently default.
5. Merge equivalent source values through `tblDimensionMap` (business line, cost centre, account, product). Convert CAD, quantities/prices, dates, period aliases and version/scenario to canonical types and keys.
6. Create 21-column typed canonical Staging; flag invalid unmapped dimensions, invalid version/scenario/period, nonnumeric/missing amount and duplicated generated record ID in `QualityMessage`; `QualityStatus=FAIL` rows remain auditable but are excluded from calculations; `qStagingErrors` loads rejected records into Quality Log.

**Source formats covered:** 09 ERP-like wide CSV + manager XLSX + GL flat CSV; 10 ERP wide CSV + forecast XLSX + GL flat CSV; 11 product GL CSV + operational Excel; 12 bank CSV/XLSX + GL cash CSV + AR/AP subledger + inventory / purchase commitments + payroll. The file reader supports CSV and XLSX. It does not support proprietary ERP binaries, live SQL, OFX or APIs.

## Spreadsheet-engine baselines and OOXML archive audit

| Model | Checks!B3 | Control statuses B7:B12 | Formula cells | Cached errors | External-link OOXML parts | Integrity |
|---|---|---|---:|---:|---:|---|
| 09 | PASS | six PASS | 238 | 0 | 0 | ZIP CRC PASS |
| 10 | PASS | six PASS | 94 | 0 | 0 | ZIP CRC PASS |
| 11 | PASS | six PASS | 84 | 0 | 0 | ZIP CRC PASS |
| 12 | PASS | six PASS | 141 | 0 | 0 | ZIP CRC PASS |

These are cached values calculated by the spreadsheet engine, **not** independently recalculated by Microsoft Excel. No `#REF!`, `#VALUE!`, `#N/A`, `#DIV/0!` or other cached Excel error values were present in the four exported XLSX worksheet XML parts. Model 11 includes one exported chart XML part for its stacked-column bridge waterfall.

## Isolated negative tests — exact source and target cells

Each test was performed in an in-memory new workbook immediately before final export. The input was changed once, its targeted control and the overall `Checks!B3` were inspected, then the original value was restored and both controls were inspected again. The **saved** workbook contains only the restored baseline values.

| Model | Source cell | Mutation and restoration | Target | Target status | Overall Checks!B3 |
|---|---|---|---|---|---|
| 09 | `Budget Input!D8` | `148800` → `-1` → restored | `Checks!B9` | PASS → FAIL → PASS | PASS → FAIL → PASS |
| 10 | `Inputs!B12` | `2` → `5` → restored | `Checks!B9` | PASS → FAIL → PASS | PASS → FAIL → PASS |
| 11 | `Staging!L7` | `1000` → `-1` → restored | `Checks!B9` | PASS → FAIL → PASS | PASS → FAIL → PASS |
| 12 | `Staging!O8` | `2027-01-05 00:00:00+00:00` → `None` → restored | `Checks!B9` | PASS → FAIL → PASS | PASS → FAIL → PASS |

**Test coverage (exact formulas/logic):** Model 09 `Checks!B9` rejects negative `Staging!K7:K78` amounts (the edited `Budget Input!D8` flows through `Staging!K7`); Model 10 `Checks!B9` requires integer `0 ≤ Inputs!B12 ≤ 4`; Model 11 `Checks!B9` rejects negative `Staging!L7:L14` quantities; Model 12 `Checks!B9` checks nonblank due-date counts for all AR/AP/inventory/payroll facts in `Staging!O7:O40`. The control set also checks version/scenario hygiene, financial reconciliation, coverage and source lineage. The full exact mutation JSON is in `docs/mutation_tests.json`.

## Two-format switching proof — reference normalization only

For each model, **a real CSV fixture and a real XLSX fixture** were read independently using a local deterministic contract-checking script. These have deliberately different source headers/codes and include unknown extra fields. They resolve to the same canonical financial keys and amounts **excluding provenance IDs and source format**, with the same 21-column output contract. Change `Inputs!B8` with `Inputs!B9=SINGLE` and refreshes the installed query to select each source; for Model 12, `Inputs!B9=ALL` unions enabled cash-related sources.

| Model | Parameter change `Inputs!B8` | Physical formats | Normalized fact rows each | Result |
|---|---|---|---:|---|
| 09 | `ERP_BUDGET` → `MANAGER_XLSX` | CSV → XLSX | 6 | Schema + mapped financial values equal in reference normalization |
| 10 | `ERP_FCST` → `FCST_XLSX` | CSV → XLSX | 3 | Schema + mapped financial values equal in reference normalization |
| 11 | `GL_PVM` → `PVM_XLSX` | CSV → XLSX | 8 | Schema + mapped financial values equal in reference normalization |
| 12 | `BANK_CSV` → `BANK_XLSX` | CSV → XLSX | 1 | Schema + mapped financial values equal in reference normalization |

**Important scope:** The CSV/XLSX normalization above **did not execute Power Query M in Microsoft Excel** and does not certify the pasted M code, refresh credentials, connector behavior, Load To placement, or `Excel.CurrentWorkbook` naming on the user's desktop. The query code must be copied into desktop Excel and installed. It is not an installed refresh connection, so dynamic refresh does not work out of the box. The workbook's `Staging` remains a static/formula-driven synthetic sample until that installation step. A native Power Query trial must switch each parameter, refresh, compare the loaded header order and key measures, and inspect `qStagingErrors`.

## Relation to Round 12's still-open issues

**Newly certified original fixes: ZERO.** This work adds related functionality. It does not repair or recertify any original workbook. In particular, Model 10's controlled demo versions **do not** close Model 03's **M06** approved monthly fact/historical-source-tie issue; new source maps **do not** close Model 08's **H18** immutable/effective-dated dynamic append issue; the provided M-code **does not** close **M20** immutable workflow/automated production refresh; adding new documentation does not clear **L01/L02/L05** for the original baseline portfolio.

The 12 outstanding authoritative Round 12 IDs remain:

- H05 — Model 03 integrated lease/tax/WC and full statement rollforwards
- H18 — Model 08 dynamic effective-dated source map / immutable batch lineage
- M02 — Model 01 automatic bid-event expansion beyond bounded engine
- M05 — Model 03 deferred tax, lease, IC and financing integration
- M06 — Model 03 approved monthly Actual/Budget/Forecast facts and historical ties
- M11 — Model 06 globally optimal allocation across constraints
- M14 — Model 07 SPA classification and post-close integration
- M20 — Model 08 immutable source audit / scheduled refresh
- L01 — Complete provenance catalogue of nested schedules
- L02 — Native Excel per-sheet visual/print/protection review
- L03 — Secure external edit/approval workflow
- L05 — Native Excel recalc / protection / owner signoff

## Explicit limitations / release blockers

- **Native Excel Power Query not run / installed.** The source-changing M code is in worksheets and `.pq` files; it is **not embedded as an executable query connection**. One-time desktop Excel creation and Load To setup are needed before clicking Refresh All. This is not an already-working one-click replacement with arbitrary files.
- **Native Excel recalculation, print rendering, protection and chart UI not verified.** Cached PASS results and archive checks do not replace independent testing in Microsoft Excel for Windows or Mac, or approval by a finance owner.
- **Data-source constraints.** The models support multiple CSV/XLSX source schemas through mapping. They do not support unregistered header changes or proprietary formats. An unseen new column containing a new period needs a Period Map alias. Missing optional fields are filled, while missing required dimensions/amounts are logged and exclude those rows.
- **Source grain and duplicate treatment.** The generated record key includes the source and input row index. It detects conflicting generated keys, but it is not a forensic audit for duplicate source records. A production source should provide immutable document/line IDs and independent source-total controls.
- **Bank balance assumption.** Model 12's bank demo is an **opening snapshot**, not an end-to-end bank statement ledger or automated bank reconciliation. For a transaction history, first determine the opening balance and adjust the Source Register / account classification; do not treat sum of a full transaction history as a balance.
- **No original 8-model modifications or new issue certification.** The historical 38/50 synthetic-scope result is carried forward from the prior report. It was not independently reaudited as part of this expansion.

**Files in the repository:** Four new XLSX workbooks, source fixtures, all 12 workbooks and the Supporting Schedules template, the Start Here guide, the repository structure and README, and the evidence JSON files. **Published to GitHub 2026-11-12.**
