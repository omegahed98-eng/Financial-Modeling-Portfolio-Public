# Quickstart

This guide covers two ways to use the workbooks. The first uses the sample data already in the file. The second connects your own CSV or XLSX extracts to Models 09–12.

## 1. Try the demo in 2 minutes

1. Open the [models folder](../models/) and download one workbook (`.xlsx`). You do not need to download all 12.
2. Open the workbook in Excel. The sample data is already loaded in the workbook's `Staging` sheet where applicable. Models 01–08 also contain their own synthetic inputs and schedules.
3. Open `Cover`, then `Outputs`, then `Checks`. Start with the output sheet to see the model result. Use `Checks` to review the workbook's controls. For Models 09–12, the overall status is shown in `Checks!B3`.
4. That is all you need for the demo. Do not install Power Query or download the example source files just to review the preloaded model. In Model 09, you can also change the yellow `AmountCAD` cells on the three business-line manager sheets (`BL01`/`BL02`/`BL03 Manager`) to see the totals move.

**Models 01–08 need no input files.** They are self-contained and use built-in synthetic data. The Supporting Schedules workbook is not required to run any model.

## 2. Plug in your own data (Models 09–12 only)

Power Query setup is a one-time manual step in desktop Excel. It is not preconfigured, and the existing `Staging` data remains the sample data until you install and load the queries.

1. **Choose a model and prepare your extract.** Use this process only with Models 09, 10, 11 or 12. Prepare a supported CSV or XLSX file using the source layout you intend to map.
2. **Put the extract in a folder.** Set up a root folder on your computer and place the source file or files inside it. The folder path is the value for `Inputs!B7`.
3. **Check the source setup in the workbook.** In `Source Register`, find the row for your source ID. Check its file format, object name, grain, dimensions, defaults and enabled status. Set the file's `RelativePath` in column D to the path from the root folder to the file. Check `Header Map` for raw column aliases, `Dimensions` for source-code mappings, and `Period Map` for month labels. The `Extract Steps` sheet lists the source-specific extraction steps.
4. **Set the three input cells.**
   - `Inputs!B7`: the root folder containing your extracts.
   - `Inputs!B8`: the source ID from `Source Register`.
   - `Inputs!B9`: `SINGLE` to load the selected source, or `ALL` to combine the rows marked `Enabled=Y` in `Source Register`.
5. **Install the two Power Query queries once in desktop Excel.** In the workbook, open the `Power Query` sheet and copy the M code for `qCanonicalStaging` and `qStagingErrors`. The same code is supplied in `power-query/*.pq` if you prefer to copy it from separate files. In desktop Excel, create each query using **Data → Get Data → Blank Query → Advanced Editor**, paste the corresponding M code, and name the queries exactly `qCanonicalStaging` and `qStagingErrors`.
6. **Set the query load destinations.** Load `qCanonicalStaging` to the workbook's `Staging` sheet and `qStagingErrors` to `Quality Log`. Loading the first query replaces the preloaded demo staging contents.
7. **Run Refresh All.** Refresh after the queries are installed and the input cells and source paths are set. Refresh All is not preconfigured before this setup.
8. **Review the results and rejected rows.** Check `Checks!B3` for the overall status. Review `Quality Log`, especially `QualityStatus` and `QualityMessage`, before relying on the outputs. Rows marked `FAIL` are logged and excluded from the calculations. Fix source paths or mappings and refresh again if needed.

The queries support the CSV and XLSX source layouts represented by the source maps and examples. They do not provide unrestricted schema inference or support for proprietary ERP binaries, live SQL, OFX or APIs. A new source layout may require changes to the source register and mapping tables. Native Excel recalculation and Power Query refresh have not been verified as part of the repository tests.

## 3. What you need

Choose one workbook for the demo. For your own data, use one of Models 09–12 and install the two queries once. The other repository files are examples, documentation or reference material.

| What it is for | Do you need it? |
|---|---|
| **`models/Model_01_*.xlsx` through `models/Model_12_*.xlsx`**: The 12 financial models. Each workbook contains its own model and synthetic demo data. Download the model you want to review; you do not need all 12. | **Required** (one workbook for a demo; use Models 09–12 for the own-data process) |
| **`power-query/*.pq`**: Separate copies of the Power Query M code for `qCanonicalStaging` and `qStagingErrors` in Models 09–12. You can copy the code from the workbook's `Power Query` sheet instead. | **Optional** (use these files only if you prefer not to copy the code from the workbook) |
| **`examples/sources/*`**: Sample CSV and XLSX source files showing the input layouts used by Models 09–12. | **Optional** (use them to test the source setup; not needed for the demo) |
| **`supporting/Supporting_Schedules_Template_Fixed.xlsx`**: A separate reference template. Its schedules are not linked to the model workbooks. | **Reference only** |
| **`docs/*.md`**: Quickstart, portfolio guide, repository structure and verification details. | **Optional** (read the Quickstart if connecting your own data; not needed to open a demo) |
| **`docs/*.json`**: Machine-readable mutation tests, source-switch reference tests and integrity manifest. | **Reference only** |

## 4. Important limits

- Models 01–08 are unchanged Round 12 baselines: 38 of 50 original issues are certified fixed in synthetic scope; 12 remain outstanding. They have not been reworked or recertified in this expansion.
- Models 09–12 have preloaded synthetic data. Their Power Query code must be installed manually before Refresh All can load external files.
- Native Excel recalculation and Power Query refresh have not been verified. The spreadsheet-engine checks are not a substitute for a native Excel test.
- Model 11 uses a stacked-column chart for the waterfall visual. It is not a native Excel Waterfall chart.
- Model 12's bank demo is an opening balance snapshot, not a full bank reconciliation.
- `supporting/Supporting_Schedules_Template_Fixed.xlsx` is a reference workbook, not a live dependency.
- All model data is synthetic. No confidential employer data or live ERP records are included.

For the detailed test results and outstanding issues, see [VERIFICATION_REPORT.md](VERIFICATION_REPORT.md).
