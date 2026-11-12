# Proposed GitHub repository structure (design only)

This tree is the published repository contents.

```text
finance-models-showcase/
├── README.md                            # hiring manager entry point
├── models/
│   ├── Model_01_Universal_Bid_Pricing_Fixed.xlsx
│   ├── Model_02_Manufacturing_Costing_Fixed.xlsx
│   ├── Model_03_Three_Statement_FPA_Fixed.xlsx
│   ├── Model_04_Capital_Investment_Fixed.xlsx
│   ├── Model_05_Working_Capital_Cash_Forecast_Fixed.xlsx
│   ├── Model_06_Pricing_Product_Mix_Fixed.xlsx
│   ├── Model_07_Acquisition_Diligence_Fixed.xlsx
│   ├── Model_08_Management_Reporting_Forecast_Accuracy_Fixed.xlsx
│   ├── Model_09_PL_Budget.xlsx
│   ├── Model_10_Rolling_Forecast.xlsx
│   ├── Model_11_Variance_PVM_Bridge.xlsx
│   └── Model_12_13_Week_Cash_Flow.xlsx
├── supporting/
│   └── Supporting_Schedules_Template_Fixed.xlsx
├── docs/
│   ├── START_HERE.md                    # visitor path and canonical data contract
│   ├── REPO_STRUCTURE.md                # this design specification
│   ├── VERIFICATION_REPORT.md           # positive/negative tests and limitations
│   ├── mutation_tests.json              # cell-level mutation evidence
│   ├── source_switch_reference_tests.json
│   └── integrity_manifest.json
├── power-query/
│   ├── model09_qCanonicalStaging.pq      # code mirrors in-workbook Power Query sheet
│   ├── model09_qStagingErrors.pq
│   ├── ... same two files for models 10, 11, 12
└── examples/
    └── sources/
        ├── m09_erp_budget_wide.csv / m09_manager_budget.xlsx / m09_gl_budget.csv
        ├── m10_erp_forecast_wide.csv / m10_forecast_rows.xlsx / m10_gl_fcst.csv
        ├── m11_gl_pvm.csv / m11_pvm_export.xlsx
        └── m12_bank.csv / m12_bank_export.xlsx / m12_gl_cash.csv
            m12_ar.csv / m12_ap.xlsx / m12_inventory.csv / m12_payroll.csv
```

## Navigation rationale

Keep the root README short. Offer four discipline subsections with direct workbook links and one dedicated Start Here link. Separate verification and technical source code from hiring-manager narrative. Keep existing eight models untouched in `models/`; avoid internal workbook hyperlinks to other XLSX files. The staged `.pq` files make M source readable and reviewable without opening Excel.

## Publication restrictions

Published 2026-11-09. Confidentiality reviewed: all data synthetic, no employer data. Native Microsoft Excel recalculation and Power Query refresh tests are still owed — see Verification Report. Source paths in the sample workbook are placeholders, not credentials.
