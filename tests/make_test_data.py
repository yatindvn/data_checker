"""Create B2B-style test files with known, documented differences.

    python tests/make_test_data.py

Writes test_data/B2B_file1.xlsx (the original) and test_data/B2B_file2.xlsx
(a modified copy). Every planted difference is listed in CHANGES so the tests
can check the tool finds exactly those and nothing else.
"""
from __future__ import annotations

import datetime as dt
import os

import numpy as np
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(os.path.dirname(HERE), "test_data")
N_ROWS = 300
KEY = "Invoice Number"

COLUMNS = [
    "Seller Gstin", "Invoice Number", "Invoice Date", "Transaction Type", "Order Id",
    "Shipment Id", "Shipment Date", "Order Date", "Shipment Item Id", "Quantity",
    "Item Description", "Asin", "Hsn/sac", "Sku", "Product Tax Code", "Bill From City",
    "Bill From State", "Bill From Country", "Bill From Postal Code", "Ship From City",
    "Ship From State", "Ship From Country", "Ship From Postal Code", "Ship To City",
    "Ship To State", "Ship To Country", "Ship To Postal Code", "Invoice Amount",
    "Tax Exclusive Gross", "Total Tax Amount", "Cgst Rate", "Sgst Rate", "Utgst Rate",
    "Igst Rate", "Compensatory Cess Rate", "Principal Amount", "Principal Amount Basis",
    "Cgst Tax", "Sgst Tax", "Utgst Tax", "Igst Tax", "Compensatory Cess Tax",
    "Shipping Amount", "Shipping Amount Basis", "Shipping Cgst Tax", "Shipping Sgst Tax",
    "Shipping Utgst Tax", "Shipping Igst Tax", "Shipping Cess Tax", "Gift Wrap Amount",
    "Gift Wrap Amount Basis", "Gift Wrap Cgst Tax", "Gift Wrap Sgst Tax",
    "Gift Wrap Utgst Tax", "Gift Wrap Igst Tax", "Gift Wrap Compensatory Cess Tax",
    "Item Promo Discount", "Item Promo Discount Basis", "Item Promo Tax",
    "Shipping Promo Discount", "Shipping Promo Discount Basis", "Shipping Promo Tax",
    "Gift Wrap Promo Discount", "Gift Wrap Promo Discount Basis", "Gift Wrap Promo Tax",
    "Tcs Cgst Rate", "Tcs Cgst Amount", "Tcs Sgst Rate", "Tcs Sgst Amount",
    "Tcs Utgst Rate", "Tcs Utgst Amount", "Tcs Igst Rate", "Tcs Igst Amount",
    "Warehouse Id", "Fulfillment Channel", "Payment Method Code", "Bill To City",
    "Bill To State", "Bill To Country", "Bill To Postalcode", "Customer Bill To Gstid",
    "Customer Ship To Gstid", "Buyer Name", "Credit Note No", "Credit Note Date",
    "Irn Number", "Irn Filing Status", "Irn Date", "Irn Error Code",
]

PRODUCTS = [
    ("GHADI DETERGENT POWDER 5 KG  SUPER SAVER PACK", "B0DSPY7L7T", 34029011, 500.0),
    ("Fem Fruity Squash Soft Safe Handwash 1500ml", "B0D314FYDY", 34013019, 200.0),
    ("Surf Excel Easy Wash 4 KG", "B07XYZ1234", 34022090, 650.0),
    ("Dettol Original Soap 125g x 8", "B08ABC5678", 34011190, 399.0),
]
CITIES = [("LUCKNOW", "UTTAR PRADESH", 226010), ("NEW DELHI", "DELHI", 110093),
          ("MUMBAI", "MAHARASHTRA", 400001), ("PUNE", "MAHARASHTRA", 411001)]
BUYERS = ["Kush Scientific Agencies", "PANKSHRI ENTERPRISRES", "Sharma Traders",
          "Gupta Wholesale"]


def make_base(n: int = N_ROWS, seed: int = 42) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    rows = []
    for i in range(1, n + 1):
        name, asin, hsn, price = PRODUCTS[rng.integers(len(PRODUCTS))]
        city, state, pin = CITIES[rng.integers(len(CITIES))]
        qty = int(rng.integers(1, 5))
        gross = round(price * qty / 1.18, 2)
        tax = round(price * qty - gross, 2)
        inv_date = dt.datetime(2025, 7, 1, 9, 0) + dt.timedelta(hours=int(rng.integers(0, 4000)))
        refund = rng.random() < 0.2
        r = dict.fromkeys(COLUMNS, 0)
        r.update({
            "Seller Gstin": rng.choice(["06AAECR0564M1Z7", "19AAECR0564M1Z0"]),
            KEY: f"LDX1-{i}",
            "Invoice Date": inv_date,
            "Transaction Type": "Refund" if refund else "Shipment",
            "Order Id": f"{rng.integers(100, 999)}-{rng.integers(1e6, 9e6)}-{rng.integers(1e6, 9e6)}",
            "Shipment Id": f"A0{rng.integers(1e9, 9e9)}X42LY",
            "Shipment Date": inv_date + dt.timedelta(hours=4),
            "Order Date": inv_date - dt.timedelta(days=1),
            "Shipment Item Id": int(rng.integers(5e11, 6e11)),
            "Quantity": qty,
            "Item Description": name,
            "Asin": asin, "Hsn/sac": hsn, "Sku": asin, "Product Tax Code": "A_GEN_STANDARD",
            "Bill From City": "Ahmedabad", "Bill From State": "GUJARAT",
            "Bill From Country": "IN", "Bill From Postal Code": 380009,
            "Ship From City": "RAJPURA", "Ship From State": "PUNJAB",
            "Ship From Country": "IN", "Ship From Postal Code": 140401,
            "Ship To City": city, "Ship To State": state, "Ship To Country": "IN",
            "Ship To Postal Code": pin,
            "Invoice Amount": price * qty, "Tax Exclusive Gross": gross,
            "Total Tax Amount": tax, "Igst Rate": 0.18, "Principal Amount": price * qty,
            "Principal Amount Basis": gross, "Igst Tax": tax,
            "Tcs Igst Rate": 0.005, "Tcs Igst Amount": round(-gross * 0.005, 2),
            "Warehouse Id": "LDX1", "Fulfillment Channel": "AFN",
            "Payment Method Code": rng.choice(["COD", "Prepaid"]),
            "Bill To City": city, "Bill To State": state, "Bill To Country": "IN",
            "Bill To Postalcode": pin,
            "Customer Bill To Gstid": f"09BPBPA{rng.integers(1000, 9999)}B1ZU",
            "Customer Ship To Gstid": "",
            "Buyer Name": BUYERS[rng.integers(len(BUYERS))],
            "Credit Note No": f"LDX1-C-{rng.integers(1e5, 9e5)}" if refund else "",
            "Credit Note Date": inv_date + dt.timedelta(days=5) if refund else "",
            "Irn Number": f"{rng.integers(1e15, 9e15):x}{rng.integers(1e15, 9e15):x}",
            "Irn Filing Status": "GOVT_ACCEPTED",
            "Irn Date": inv_date,
            "Irn Error Code": "",
        })
        rows.append(r)
    return pd.DataFrame(rows, columns=COLUMNS)


# Planted differences in file 2 (key, column, new value). Used by the tests.
CELL_CHANGES = [
    ("LDX1-5", "Invoice Amount", 9999.0),                      # number changed
    ("LDX1-12", "Buyer Name", "KUSH SCIENTIFIC AGENCIES"),     # case only (if buyer differs)
    ("LDX1-20", "Ship To City", "GURUGRAM"),                   # text changed
    ("LDX1-25", "Invoice Date", dt.datetime(2030, 1, 1, 10, 0)),  # date changed
    ("LDX1-33", "Credit Note No", "CN-NEW-001"),                # blank -> value (non refund row)
    ("LDX1-50", "Payment Method Code", ""),                    # value -> blank
]
CASE_ONLY_CHANGE = ("LDX1-12", "Buyer Name")
SPACE_ONLY_CHANGE = ("LDX1-60", "Ship To State")    # trailing spaces only -> not an issue (trim on)
TOLERANCE_CHANGE = ("LDX1-70", "Tax Exclusive Gross", 0.004)   # tiny diff, hidden by tolerance 0.01
DELETED_KEYS = ["LDX1-100", "LDX1-101"]
ADDED_KEY = "LDX1-NEW-1"
DUPLICATED_KEY = "LDX1-150"
REMOVED_COLUMN = "Irn Error Code"
ADDED_COLUMN = "Remarks"
RENAMED_CASE = ("Hsn/sac", "HSN/SAC")


def make_file2(base: pd.DataFrame) -> pd.DataFrame:
    df = base.copy().set_index(KEY, drop=False)
    for key, col, val in CELL_CHANGES:
        df[col] = df[col].astype(object)
        df.loc[key, col] = val
    k, c = SPACE_ONLY_CHANGE
    df.loc[k, c] = f"  {df.loc[k, c]}  "
    k, c, delta = TOLERANCE_CHANGE
    df.loc[k, c] = round(df.loc[k, c] + delta, 3)
    df = df.drop(index=DELETED_KEYS)
    new = base.iloc[[0]].copy()
    new[KEY] = ADDED_KEY
    dup = df.loc[[DUPLICATED_KEY]].copy()
    df = pd.concat([df.reset_index(drop=True), new, dup], ignore_index=True)
    df = df.drop(columns=[REMOVED_COLUMN])
    df[ADDED_COLUMN] = "checked"
    df = df.rename(columns={RENAMED_CASE[0]: RENAMED_CASE[1]})
    return df.sample(frac=1, random_state=7).reset_index(drop=True)   # shuffled order


def write_xlsx(df: pd.DataFrame, path: str) -> None:
    with pd.ExcelWriter(path, engine="xlsxwriter",
                        datetime_format="m/d/yyyy h:mm") as xw:
        df.to_excel(xw, index=False, sheet_name="Sheet1")


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    base = make_base()
    # make sure the case-only change really is case-only
    base.loc[base[KEY] == CASE_ONLY_CHANGE[0], CASE_ONLY_CHANGE[1]] = "Kush Scientific Agencies"
    f2 = make_file2(base)
    write_xlsx(base, os.path.join(OUT_DIR, "B2B_file1.xlsx"))
    write_xlsx(f2, os.path.join(OUT_DIR, "B2B_file2.xlsx"))
    print(f"Wrote {len(base)} / {len(f2)} rows to {OUT_DIR}")
    return base, f2


if __name__ == "__main__":
    main()
