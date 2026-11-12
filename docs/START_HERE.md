# Start Here — Synthetic FP&A & Operational Finance Portfolio

**Twelve self-contained Excel models**

The models cover planning, operational costing, capital allocation, commercial decisions, diligence, cash management and management reporting. Every number and entity in the models is synthetic. No GMED/EGMED source data, confidential commercial inputs or live ERP records are included.

**To review the models:** Open Model 09 → `Outputs` to see the total-company P&L and manager inputs; Model 10 → `Outputs` for quarterly rolling forecast; Model 11 → `Outputs` for the price/volume/mix waterfall; Model 12 → `Outputs` for 13-week liquidity. Review `Checks!B3` in each new workbook. Then review Model 03, 05 or 08 for the related integrated models.

## Portfolio navigation by functional area

| Group | Models | Contents |
|---|---|---|
| **FP&A Core** | 03, 05, 09, 10, 11, 12 | Integrated statements; working capital; departmental budgeting; versioned rolling forecast; price/volume/mix performance explanation; near-term liquidity |
| **Costing & Operations** | 02, 04 | Manufacturing cost and inventory; capex investment cases and dated returns |
| **Deals & Pricing** | 01, 06, 07 | Bid contract economics; pricing/product mix; acquisition diligence and QoE |
| **Reporting** | 08 | Management reporting, variance reporting and forecast accuracy |

**Open the models in this order:** 09 → 10 → 11 → 08 → 12 → 05 → 03 → 02 → 04 → 06 → 01 → 07.

## New models and their neighboring baselines

| New workbook | What you interact with | Compatible portfolio context |
|---|---|---|
| **Model 09: P&L Budget** | 3 business-line manager input blocks → monthly revenue / COGS / OPEX / EBITDA / CAPEX and company total | Models 03, 05 |
| **Model 10: Quarterly Rolling Forecast** | Monthly Budget / Actual / Forecast facts; Q2 closed actuals + Q3–Q4 selected forecast snapshot | Models 03, 08 |
| **Model 11: Variance / Price–Volume–Mix Bridge** | Product-level actual vs budget revenues; exact volume / mix / price bridge and waterfall visual | Models 08, 10 |
| **Model 12: 13-Week Cash Flow** | Monday-start weeks; bank opening balance, AR receipts, AP/inventory/payroll disbursements | Models 05 |

### Shared dimension and account contract — Models 09–12

All four new files use the **identical 21-column `Staging!A6:U6` canonical fact contract**. The shared dimensions are `BusinessLineID`, `CostCenterID`, `AccountID`, `ProductID`, `PeriodKey`, `Version`, `Scenario`, `SnapshotID`, and `SourceID`; all staging rows carry lineage and quality flags. Grain: **one native input fact per source + document/record + month + version + scenario + mapped business line/cost centre/account/product**. Cash rows additionally use `EventDate` and `DueDate` for exact weekly placement.

| Dimension | Canonical values / policy |
|---|---|
| Business line | `BL01` Urban Services; `BL02` Digital Platforms; `BL03` Field Operations; `BL_CORP` Corporate / Treasury |
| Cost centre | `CC10` Service Delivery; `CC20` Software Operations; `CC30` Regional Operations; `CC90` Central Treasury |
| Revenue and costs | `REV_CORE`, `COGS_CORE`, `OPEX_SGNA`, `CAPEX_TOTAL` |
| Balance sheet / cash | `A_CASH`, `A_AR`, `A_AP`, `A_INVENTORY`, `A_PAYROLL` |
| Product | `PR01`–`PR04` for bridge detail, `ALL` for non-product facts |
| Period | `YYYY-MM` for all four; Model 10 derives quarter from calendar month; Model 12 places receipts/payments into Monday-start weeks using `DueDate` |
| Version | `Actual`, `Budget`, or `Forecast`; budget is the plan, actual is recognized closed data, forecast is the `SnapshotID`-selected view |
| Scenario | `Base`, `Upside`, `Downside`; baseline is `Base` |
| Currency and sign | CAD, `AmountCAD`; P&L/capex/invoice amounts are positive magnitudes; bank signed movements are credit less debit. Cash outflows are subtracted by the cash model, not pre-negated in open-item facts. |

**Staging columns, in exact order:**

`RecordID | SourceID | BusinessLineID | CostCenterID | AccountID | ProductID | PeriodKey | Version | Scenario | SnapshotID | AmountCAD | Quantity | UnitPriceCAD | EventDate | DueDate | CounterpartyID | DocumentID | SourceRef | SourceFormat | QualityStatus | QualityMessage`

`Dimensions` contains source-specific code crosswalks, account classes and stable canonical IDs. Different source labels (for example ERP `10` and Excel `Urban`) map to `BL01`. Do not join facts on raw labels. The existing Models 01–08 retain their original field structures; interoperability with them is **via a deliberate export/import mapping**, not a claim that their existing sheets already share this schema.

## Using one of the new workbooks

1. Open **Cover**, then **Outputs**, then **Checks**. The default example has `Checks!B3 = PASS` and no confidential data.
2. For Model 09, edit the yellow `AmountCAD` cells on the three business-line manager sheets. These feed the demo staging data and the total-company model through internal formulas.
3. For a real extract, read **Source Register** (grain, dimensions present/missing and defaults), **Header Map** (raw header aliases), **Dimensions** (cross-source aliases) and **Period Map** (month labels). Each source has a ten-step workflow on **Extract Steps**.
4. Set `Inputs!B7` to the root folder containing your extracts, `Inputs!B8` to the chosen source ID and `Inputs!B9` to `SINGLE`, or set `ALL` to combine `Enabled=Y` source rows. Adjust the source file `RelativePath` in Source Register column D. No changes to the M code are needed for another supported CSV/XLSX file schema once the map is configured.
5. **One-time setup required:** Install `qCanonicalStaging` and `qStagingErrors` from the **Power Query** worksheet (or `power-query/*.pq` files) in desktop Excel through Data → Get Data → Blank Query → Advanced Editor. Load the first query to **Staging** and the second to **Quality Log**, replacing the preloaded demo staging contents. Then **Refresh All** responds to parameter changes.
6. Inspect `QualityStatus=FAIL` and `QualityMessage` before relying on an output. Missing optional columns are filled with null/declared defaults; renamed headers use aliases; extra columns are ignored; unknown populated dimension codes and invalid amounts are rejected. The calculation engine includes only `PASS` rows.
7. For an integration to another portfolio workbook, export an approved canonical fact extract and map its keys to that workbook's input expectations. **No workbook contains formula links to another workbook.**

**Source formats and examples:** Models 09–11 include ERP-like exports, GL/CSV data and Excel alternatives. Model 12 includes bank CSV/Excel alternatives, a GL cash extract, AR/AP subledger records, inventory commitments and payroll. The bank illustration is an **opening balance snapshot**, not a full transaction-to-bank reconciliation. Source file formats demonstrated are CSV and XLSX; database APIs, OFX, encrypted banks and unrestricted schema inference are not implemented.

## Verification and limitations

- Models 09–12: baseline PASS and a documented isolated PASS → FAIL → PASS mutation each. Exported workbooks contain zero cached Excel error cells and zero external workbook-link parts.
- CSV ↔ XLSX source switches were independently normalized on provided fixtures to verify matching canonical semantics, but **Microsoft Power Query was not executed in this environment**. Native Excel installation, workbook refresh, chart UI, protection and owner signoff still require desktop validation.
- Existing Models 01–08 and the Supporting Schedules reference are **preserved byte-for-byte** in the repository, not modified or rebuilt.
- The authoritative Round 12 issue status remains **38/50 synthetic-scoped fixed; 12 outstanding**. No new original-issue fix is claimed by this expansion.
- The workbooks are not a production financial system, audited models or live operational data integration. The repository is published on GitHub.

For mutation coordinates, structural tests, source-switch evidence and remaining blockers, see [`VERIFICATION_REPORT.md`](VERIFICATION_REPORT.md).
