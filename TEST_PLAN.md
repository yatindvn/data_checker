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
