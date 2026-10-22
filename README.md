# data_checker – Excel Compare

Desktop tool that compares two Excel/CSV files (headers + data) and writes an
Excel report with every issue highlighted in red.

## For users (no Python needed)
**[⬇ Download ExcelCompare.zip (latest release)](https://github.com/yatindvn/data_checker/releases/latest/download/ExcelCompare.zip)** · Windows only · [All releases](https://github.com/yatindvn/data_checker/releases)

1. Download `ExcelCompare.zip` using the link above and unzip it anywhere.
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

## Mapping file guide
A mapping file is an Excel sheet that tells the tool **what to compare**. Use one when File 1 and File 2 have different column names, or when File 2 holds calculated values.

### Steps
1. In the app, pick **File 1** and **File 2**. If you know your key column (e.g. Invoice Number), select it in the list now too.
2. Click **Make template** and save it (e.g. `mapping.xlsx`). The app lists every File 1 column and fills in the similarly named File 2 column where it finds one.
3. Open the template in Excel and edit it as described below. Save and close it.
4. Back in the app, the mapping file is already selected. Click **Compare**.

Next time you can skip steps 1–3: click **Browse...** next to the mapping box and pick the saved mapping file.

### The columns
| Column | What to put |
|---|---|
| **Output Column** | The name shown in the report. You can choose any name. |
| **File 1** | Where this value comes from in File 1 (see the options below). |
| **File 2** | Where this value comes from in File 2. |
| **Key** | `Y` on the column(s) that identify a row, e.g. Invoice No. At least one key is needed if the row order differs between files. |
| **Tolerance** | Optional. The allowed number difference for this row only, e.g. `0.05`. |
| **Note** | Template only. It flags rows where no File 2 match was found, and is ignored when comparing. |

### What a File 1 / File 2 cell can contain
| You write | Example | Meaning |
|---|---|---|
| A column name | `Invoice Number` | Take the value from that column |
| *(blank)* | | Same column name as the Output Column |
| A formula | `([Igst Tax] + [Cgst Tax]) / 2` | Calculate the value, then compare it |
| Plain English | `add igst tax and cgst tax then divide by 2` | Claude turns it into a formula (needs an API key) |

### Example mapping
| Output Column | File 1 | File 2 | Key | Tolerance |
|---|---|---|---|---|
| Invoice No | Invoice Number | Inv No | Y | |
| Customer | Buyer Name | Customer Name | | |
| Invoice Date | | | | |
| Total Tax | `[Cgst Tax] + [Sgst Tax] + [Igst Tax]` | Total Tax | | |
| Avg Tax | `([Igst Tax] + [Cgst Tax]) / 2` | Avg IGST+CGST | | |
| Net Amount | `ROUND([Invoice Amount] - [Total Tax Amount], 2)` | Net | | 0.05 |
| Avg Tax (English) | add igst tax and cgst tax then divide by 2 | Avg IGST+CGST | | |

### Writing formulas
- Put column names in **square brackets**: `[Igst Tax]`.
- **Maths:** `+  -  *  /  ^  ( )`
- **Join text:** `[City] & ", " & [State]`
- **Compare:** `=  <>  <  >  <=  >=` (text compares ignore upper/lower case)
- **Functions:**
  - Numbers: `ROUND(x, 2)`, `ABS`, `SUM`, `MIN`, `MAX`, `AVERAGE`
  - Logic: `IF(cond, a, b)`, `IFERROR(x, 0)`, `AND`, `OR`, `NOT`, `ISBLANK`
  - Text: `UPPER`, `LOWER`, `TRIM`, `LEN`, `LEFT`, `RIGHT`, `MID`, `CONCAT`
  - Convert: `VALUE`
- Blank cells count as **0**. Dividing by zero shows **#ERROR**.
- A leading `=` is optional.
- Examples:
  - `[Cgst Tax] + [Sgst Tax] + [Igst Tax]`
  - `ROUND([Invoice Amount] - [Total Tax Amount], 2)`
  - `IF([Igst Tax] = 0, "INTRA", "INTER")`

### Rules to remember
- **Only rows in the mapping are compared.** Delete the rows you don't care about.
- Each **Output Column** name must be unique.
- **Plain-English rows** need a Claude API key, entered in the app or set as `ANTHROPIC_API_KEY`, plus internet access. Claude (`claude-opus-5-5`) writes the formula, the tool checks it against the file's real columns, and the report's **Mapping Used** sheet shows it so you can confirm it.
- When a mapping file is selected, the app's key list and "Ignore columns" box are switched off, because the mapping file controls those.
- It also works with **Compare ALL sheets**: the same mapping is applied to every sheet.
- **Clear** next to the mapping box goes back to normal comparison.

### Try it
`test_data/mapping_B2B.xlsx` is a ready-made example. Use `B2B_file1.xlsx` as File 1 and `B2B_summary_file2.xlsx` as File 2. It should report 4 planted errors.

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
