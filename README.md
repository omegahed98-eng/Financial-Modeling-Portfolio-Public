# Financial Modelling Portfolio — FP&A, Operations, Commercial Finance

**12 Excel models · synthetic data only**

The models cover business-line planning, source-system extracts, margin analysis and cash forecasting. Each `.xlsx` is self-contained. No confidential employer data or cross-workbook Excel formulas are used.

> **Start with the [portfolio guide](docs/START_HERE.md)** for a quick read, the source-data contract and instructions for using the workbooks.

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

## Where to start

Start with **09 → 10 → 11 → 12**. Review `Cover`, `Outputs` and `Checks` in each workbook, then compare them with Models 03, 05 and 08. Models 09–12 use shared keys, source registers, header and dimension maps, a 21-column canonical staging schema, per-source ETL steps and Power Query M code. Sample CSV/Excel input formats are included under [`examples/sources`](examples/sources/).

**Power Query setup:** The queries must be installed manually once in desktop Excel; M is provided in each workbook and under [`power-query/`](power-query/). Afterwards, `Inputs!B8` selects the source and `Inputs!B9` sets `SINGLE` or `ALL`, while Source Register controls paths and enabled sources. See [setup limitations](docs/VERIFICATION_REPORT.md). Refresh All is not preconfigured.

## Test results and limitations

- [Start Here](docs/START_HERE.md) · [Verification and negative tests](docs/VERIFICATION_REPORT.md) · [Repository design](docs/REPO_STRUCTURE.md)
- Models 09–12: **baseline PASS**, one **PASS → FAIL → PASS** test each, zero cached formula errors, zero external workbook-link parts.
- Model 11's native chart is a **stacked-column waterfall visualization**, not a signed-off Microsoft Excel native Waterfall chart implementation.
- Models 01–08: unchanged Round 12 synthetic-scope baselines (**38/50 certified in prior audit, 12 remain outstanding**); not reworked or recertified here.
- **No production data, credentialed connectors, audited source documents, native Power Query execution or professional sign-off.** The repository is published on GitHub. Native Excel recalculation and Power Query refresh have not been verified. See the Verification Report.

`supporting/Supporting_Schedules_Template_Fixed.xlsx` is a reference workbook, **not** a live dependency.
