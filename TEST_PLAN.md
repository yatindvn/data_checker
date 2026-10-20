# Test Plan – Excel Compare

Automated in [tests/test_compare.py](tests/test_compare.py). Run them with:

```bash
python -m pytest -v
```

## Test data
`python tests/make_test_data.py` creates two B2B-style files (same 89 columns as the
Amazon B2B report) in `test_data/`. File 1 has 300 invoices. File 2 is a copy with the planted changes below, and its rows are **shuffled**.

| # | Planted change in File 2 | Where |
|---|---|---|
| 1 | Number changed (Invoice Amount → 9999) | LDX1-5 |
| 2 | Upper/lower case only (Buyer Name → "KUSH SCIENTIFIC AGENCIES") | LDX1-12 |
| 3 | Text changed (Ship To City → GURUGRAM) | LDX1-20 |
| 4 | Date changed (Invoice Date → 1/1/2030) | LDX1-25 |
| 5 | Credit Note No changed to "CN-NEW-001" | LDX1-33 |
| 6 | Value → blank (Payment Method Code) | LDX1-50 |
| 7 | Extra spaces only (Ship To State) | LDX1-60 |
| 8 | Tiny number change +0.004 (Tax Exclusive Gross) | LDX1-70 |
| 9 | Rows deleted | LDX1-100, LDX1-101 |
| 10 | New row added | LDX1-NEW-1 |
| 11 | Duplicate invoice | LDX1-150 |
| 12 | Column removed | Irn Error Code |
| 13 | Column added | Remarks |
| 14 | Header case changed | Hsn/sac → HSN/SAC |

To try it in the app, compare the two files with key column **Invoice Number**.

## Positive test cases
| ID | Scenario | Expected result |
|---|---|---|
| P01 | Compare a file with itself | No issues, "FILES MATCH" |
| P02 | Column missing in File 2 | "Missing in File 2" header issue |
| P03 | Extra column in File 2 | "Extra in File 2" header issue |
| P04 | Header differs only by case | Reported, with an "Only case differs" note |
| P05 | Key match on shuffled rows | Exactly changes 1–6 and 8 found, nothing else |
| P06 | Rows deleted | Listed in "Rows Only In File 1" |
| P07 | Row added and duplicate row | Listed in "Rows Only In File 2" |
| P08 | Duplicate key | Listed in "Duplicate Keys" |
| P09 | Number tolerance 0.01 | Tiny change (8) hidden, big change (1) still found |
| P10 | Ignore case | Case-only change (2) hidden |
| P11 | Trim spaces on / off | Change 7 hidden when on, found when off |
| P12 | Ignore columns | Changes in an ignored column are not reported |
| P13 | `10` vs `10.00` | Treated as equal |
| P14 | Row-by-row mode (no key) | Changed cell and extra row found at the correct row numbers |
| P15 | CSV vs XLSX | Compares correctly |
| P16 | Report output | All sheets present; changed cells, extra rows and extra column headers are red |
| P17 | Two key columns (GSTIN + invoice) | Matches on the combined key |
| P18 | Performance: 100k rows × 20 columns | Correct result in under 60 s |
| P19 | UI loads the key-column list | Lists only columns in both files |

## Negative test cases
| ID | Scenario | Expected result |
|---|---|---|
| N01 | File does not exist | Error shown, no crash |
| N02 | Text file renamed to .xlsx | Error shown, no crash |
| N03 | Key column missing in one file | Clear "Key column(s) not present" error |
| N04 | Sheet name does not exist | Error shown |
| N05 | Headers only, no data rows | Runs fine, 0 rows, no issues |
| N06 | Completely blank sheet | Runs fine, columns reported as extra |
| N07 | Report path cannot be written (e.g. file open in Excel) | Error shown, no crash |
| N08 | UI: tolerance is "abc" | Warning shown, comparison not started |
| N09 | UI: Compare clicked with no files | Warning shown, comparison not started |
| N10 | Blank key values | Blanks reported as duplicate keys, still compared in order |
| N11 | No common columns | Only header issues, report still written |
| N12 | File 2 has headers but no rows | All File 1 rows listed as "only in File 1" |

---

# New features: all sheets, mapping file, formulas, plain English
Automated in [tests/test_features.py](tests/test_features.py).

## Extra test data (`python tests/make_test_data.py`)
| File | What it is |
|---|---|
| `MultiSheet_file1.xlsx` | Sheets **B2B**, **B2C**, **Credit Notes** |
| `MultiSheet_file2.xlsx` | Sheets **B2B** (same), **" b2c"** (different case, extra space, Invoice Amount of LDX1-7 changed), **Extra** |
| `B2B_summary_file2.xlsx` | Summary of `B2B_file1.xlsx` with **different headers** (Inv No, Customer Name, Net…) and **calculated columns** (Total Tax, Avg IGST+CGST, Tax Type, Ship To), shuffled rows |
| `mapping_B2B.xlsx` | Mapping between the two: column renames, blank (same name), 5 formulas, Key, a 0.05 tolerance on Net Amount |
| `mapping_B2B_plain_english.xlsx` | Same as above, plus one row written in plain English (needs a Claude API key) |

Planted errors in the summary file: Total Tax of LDX1-3, Avg Tax of LDX1-9, Customer of LDX1-15 and Tax Type of LDX1-21. Net Amount of LDX1-30 is +0.03, inside the 0.05 tolerance, so it should **not** be reported.

## Positive test cases
| ID | Scenario | Expected result |
|---|---|---|
| P20 | Compare all sheets | Sheets matched by name ignoring case and spaces; only the planted B2C change found; sheets in one file listed |
| P21 | All-sheets report | Summary has one row per sheet (MATCH / ISSUES FOUND / NOT COMPARED); tables have a Sheet column; one highlighted copy per sheet |
| P22 | Mapping with renamed columns and formulas | Exactly the 4 planted errors found; all 300 rows matched by key |
| P23 | Formula values | File 1 value equals the calculation, e.g. (Igst + Cgst) / 2 |
| P24 | Tolerance per row | The 0.03 difference is hidden with tolerance 0.05 and reported without it |
| P25 | Blank mapping cell | Uses the same column name in both files |
| P26 | Plain-English logic | Converted to a formula once; shown on the Mapping Used sheet with an "AI converted" note; finds the same error |
| P27 | Claude request (mocked) | Uses model `claude-opus-5-5`, sends the column list, result is cached |
| P28 | Make template | File 2 columns matched by similar name (`invoice_number` = `Invoice Number`); Key filled in; How-to sheet included |
| P29 | Formula engine (11 formulas) | Maths, Excel rounding, IF, joining text with &, text compare ignoring case, divide by zero shown as #ERROR, IFERROR, blank counts as 0, SUM/MAX, no float noise |
| P30 | Mapping + all sheets | Mapping applied to every sheet; each sheet shows only its own errors |
| P31 | UI controls | "All sheets" disables the sheet pickers; a mapping file disables the key list |

## Negative test cases
| ID | Scenario | Expected result |
|---|---|---|
| N13 | Mapping file missing its required columns | Clear error naming the required columns |
| N14 | Formula uses a column that doesn't exist | Error names the file, the output row and the column |
| N15 | Malicious formula (`__import__("os")...`) | Rejected; nothing runs |
| N16 | Plain English with no API key | Clear error: enter a key or write it as a formula |
| N17 | AI returns a bad formula, or says CANNOT | Rejected with a clear message |
| N18 | All sheets, one sheet fails (key missing) | Other sheets still compared; failure shown in the report |
| N19 | Same Output Column listed twice | Error |
| N20 | Tolerance is text ("abc") | Error |
| N21 | Formula syntax errors (6 kinds) | FormulaError, no crash |
| N22 | No sheet names in common | All sheets "missing"; report still written |
| N23 | UI given a bad mapping file | Warning shown; comparison not started |
