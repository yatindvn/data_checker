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
4. Click **Compare** and then **Open report**.

## Report sheets
| Sheet | Contents |
|---|---|
| Summary | Counts and an overall MATCH / ISSUES FOUND result |
| Header Issues | Missing, extra, renamed by case only, reordered or blank columns |
| Data Differences | Each different cell: key/row, column, File 1 value, File 2 value, Excel row numbers |
| Rows Only In File 1 / File 2 | Rows that have no match in the other file |
| Duplicate Keys | Keys that appear more than once (only shown if any) |
| File 2 Highlighted | File 2 copy with differing cells, extra rows and extra columns in red |

Notes: numbers are compared as numbers (`1` = `1.0`), and all other values as text, so leading zeros are kept.
