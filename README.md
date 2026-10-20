# data_checker – Excel Compare

Desktop tool that compares two Excel/CSV files (headers + data) and writes an
Excel report with every issue highlighted in red.

## For users (no Python needed)
1. Get `ExcelCompare.zip` and unzip it anywhere.
2. Double-click `ExcelCompare\ExcelCompare.exe`. Keep the exe inside its folder, because it needs the `_internal` folder next to it.
3. If Windows SmartScreen warns about an unknown app, click **More info → Run anyway** (the exe is not code-signed).

## Build the exe
```bash
build_exe.bat
```
The output is `dist\ExcelCompare.zip`, which you can share.

## Setup (for developers)
```bash
python -m pip install -r requirements.txt
```

## Run
```bash
python app.py
```

1. Browse for **File 1** and **File 2** and pick a sheet for each.
2. (Optional) select **key column(s)** such as an ID. Rows are then matched by key,
   so it doesn't matter if the rows are in a different order. With no key, rows are compared by position.
3. Options: ignore columns, ignore case, trim spaces, number tolerance.
4. (Optional) tick **Compare ALL sheets**. Sheets are matched by name (ignoring case and spaces); sheets found in only one file are reported.
5. (Optional) choose a **mapping file** when the two files use different column names or File 2 holds calculated values (see below).
6. Click **Compare** and then **Open report**.

## Mapping file (different headers, formulas, plain English)
Click **Make template** to create a starter mapping from your two files, then edit it in Excel:

| Output Column | File 1 | File 2 | Key | Tolerance |
|---|---|---|---|---|
| Invoice No | Invoice Number | Inv No | Y | |
| Customer | Buyer Name | Customer Name | | |
| Invoice Date | | | | |
| Total Tax | [Cgst Tax] + [Sgst Tax] + [Igst Tax] | Total Tax | | |
| Avg Tax | ([Igst Tax] + [Cgst Tax]) / 2 | Avg IGST+CGST | | |
| Net Amount | ROUND([Invoice Amount] - [Total Tax Amount], 2) | Net | | 0.05 |
| Avg Tax (English) | add igst tax and cgst tax then divide by 2 | Avg IGST+CGST | | |

Each File 1 / File 2 cell can be:
- **a column name**, or **blank** for the same name as the Output Column
- **a formula**, with columns in `[brackets]`: `+ - * / ^ &`, `= <> < > <= >=`, `ROUND ABS SUM MIN MAX AVERAGE IF IFERROR AND OR NOT ISBLANK UPPER LOWER TRIM LEN LEFT RIGHT MID CONCAT VALUE`. Blank cells count as 0.
- **plain English**. Claude (`claude-opus-5-5`) turns it into a formula, which is checked against the real columns and shown on the report's **Mapping Used** sheet. This needs a Claude API key, entered in the app or set as `ANTHROPIC_API_KEY`, plus internet access.

Only the rows in the mapping are compared. **Key = Y** marks the column(s) used to match rows, and **Tolerance** sets the allowed number difference for that row.
Examples are in `test_data/`: `mapping_B2B.xlsx` compares `B2B_file1.xlsx` with `B2B_summary_file2.xlsx`.

## Report sheets
| Sheet | Contents |
|---|---|
| Summary | Counts and an overall MATCH / ISSUES FOUND result (one row per sheet in all-sheets mode) |
| Mapping Used | Only with a mapping file: what was used for each output column, including any AI-made formula |
| Header Issues | Missing, extra, renamed by case only, reordered or blank columns |
| Data Differences | Each different cell: key/row, column, File 1 value, File 2 value, number difference, Excel row numbers (plus a Sheet column in all-sheets mode) |
| Rows Only In File 1 / File 2 | Rows that have no match in the other file |
| Duplicate Keys | Keys that appear more than once (only shown if any) |
| File 2 Highlighted (or `F2 <sheet>`) | File 2 copy with differing cells, extra rows and extra columns in red |

Notes: numbers are compared as numbers (`1` = `1.0`), and all other values as text, so leading zeros are kept.
