# Financial Modelling Portfolio — FP&A, Operations, Commercial Finance

**12 Excel models · synthetic data only**

A hands-on finance portfolio showing how business-line assumptions, source-system extracts, margin drivers and cash conversion flow into financial decisions. Each `.xlsx` is self-contained. No confidential employer data or cross-workbook Excel formulas are used.

> **Start with the [portfolio guide](docs/START_HERE.md)** for a five-minute reading path, the source-data contract and workbook operation.

## Models by discipline

### FP&A Core

- **03 — Integrated three-statement FP&A:** [Workbook](models/Model_03_Three_Statement_FPA_Fixed.xlsx)
- **05 — Working capital and cash forecast:** [Workbook](models/Model_05_Working_Capital_Cash_Forecast_Fixed.xlsx)
- **09 — P&L budget — manager inputs to company:** [Workbook](models/Model_09_PL_Budget.xlsx)
- **10 — Quarterly rolling forecast:** [Workbook](models/Model_10_Rolling_Forecast.xlsx)
- **11 — Variance / price–volume–mix waterfall:** [Workbook](models/Model_11_Variance_PVM_Bridge.xlsx)
- **12 — 13-week cash flow:** [Workbook](models/Model_12_13_Week_Cash_Flow.xlsx)

### Costing & Operations

- **02 — Manufacturing costing:** [Workbook](models/Model_02_Manufacturing_Costing_Fixed.xlsx)
- **04 — Capital investment and dated returns:** [Workbook](models/Model_04_Capital_Investment_Fixed.xlsx)

### Deals & Pricing

- **01 — Universal bid pricing:** [Workbook](models/Model_01_Universal_Bid_Pricing_Fixed.xlsx)
- **06 — Pricing, mix and margin:** [Workbook](models/Model_06_Pricing_Product_Mix_Fixed.xlsx)
- **07 — Acquisition diligence / QoE:** [Workbook](models/Model_07_Acquisition_Diligence_Fixed.xlsx)

### Reporting

- **08 — Management reporting and forecast accuracy:** [Workbook](models/Model_08_Management_Reporting_Forecast_Accuracy_Fixed.xlsx)

## How to explore

Open **09 → 10 → 11 → 12** and check `Cover`, `Outputs`, `Checks`. Then follow their fit with 03 / 05 / 08. New Models 09–12 have shared keys, source registers, header/dimension maps, a 21-column canonical staging schema, per-source ETL steps, and Power Query M text. Sample CSV/Excel input formats are included under [`examples/sources`](examples/sources/).

**Power Query setup:** One-time manual query installation is needed; M is provided in each workbook and under [`power-query/`](power-query/). Afterwards, `Inputs!B8` selects the source and `Inputs!B9` sets `SINGLE` or `ALL`, while Source Register controls paths and enabled sources. Please review [setup limitations](docs/VERIFICATION_REPORT.md) before assuming Refresh All is preconfigured.

## Evidence & boundaries

- [Start Here](docs/START_HERE.md) · [Verification and negative tests](docs/VERIFICATION_REPORT.md) · [Repository design](docs/REPO_STRUCTURE.md)
- Models 09–12: **baseline PASS**, one **PASS → FAIL → PASS** test each, zero cached formula errors, zero external workbook-link parts.
- Model 11's native chart is a **stacked-column waterfall visualization**, not a signed-off Microsoft Excel native Waterfall chart implementation.
- Models 01–08: unchanged Round 12 synthetic-scope baselines (**38/50 certified in prior audit, 12 remain outstanding**); not reworked or recertified here.
- **No production data, credentialed connectors, audited source documents, native Power Query execution or professional sign-off.** Published to GitHub as a hiring-manager showcase. Native Excel recalculation and Power Query refresh are still unverified — see Verification Report.

`supporting/Supporting_Schedules_Template_Fixed.xlsx` is a reference workbook, **not** a live dependency.
