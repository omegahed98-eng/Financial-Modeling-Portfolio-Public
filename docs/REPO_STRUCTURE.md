# GitHub repository structure

The repository contains the following files and folders.

```text
Financial-Modeling-Portfolio-Public/
├── README.md                            # start here
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

## Repository layout

The root README is the entry point. It groups the workbooks by discipline and links to Start Here. Verification results and source code sit in `docs/` and `power-query/`, separate from the portfolio overview. The original eight models remain in `models/`. Workbooks do not link to other XLSX files. The `.pq` files let readers inspect the Power Query M code without opening Excel.

## Publication restrictions

Published 2026-11-12. Confidentiality reviewed: all data synthetic, no employer data. Microsoft Excel recalculation and Power Query refresh have not been tested. See Verification Report. Source paths in the sample workbook are placeholders, not credentials.
