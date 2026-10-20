"""Core logic: compare two Excel/CSV sheets (headers + data) and write a report.

All comparisons are vectorised with pandas/numpy so large files stay fast.
"""
from __future__ import annotations

import os
import time
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

try:
    import python_calamine  # noqa: F401  (much faster .xlsx reader)
    _ENGINE = "calamine"
except ImportError:
    _ENGINE = None

EXCEL_MAX_ROWS = 1_048_575
ROW_COL = "__excel_row__"


@dataclass
class Options:
    key_columns: list[str] = field(default_factory=list)  # empty -> compare row by row
    ignore_case: bool = False
    trim_spaces: bool = True
    numeric_tolerance: float = 0.0
    ignore_columns: list[str] = field(default_factory=list)


@dataclass
class Result:
    header_issues: pd.DataFrame
    cell_diffs: pd.DataFrame
    only_in_a: pd.DataFrame
    only_in_b: pd.DataFrame
    duplicate_keys: pd.DataFrame
    summary: dict
    df_b: pd.DataFrame          # file B as read (for the highlighted copy)
    diff_mask_b: dict           # {b_row_position: set(column names)}
    rows_only_b_pos: set
    extra_cols_b: list

    @property
    def has_issues(self) -> bool:
        return any(len(x) for x in (self.header_issues, self.cell_diffs,
                                    self.only_in_a, self.only_in_b, self.duplicate_keys))


# --------------------------------------------------------------------------- reading
def list_sheets(path: str) -> list[str]:
    if path.lower().endswith(".csv"):
        return ["(csv)"]
    return pd.ExcelFile(path, engine=_ENGINE).sheet_names


def read_headers(path: str, sheet: str) -> list[str]:
    return list(_read(path, sheet, nrows=0).columns)


def _read(path: str, sheet: str, nrows=None) -> pd.DataFrame:
    # Everything as text so "001" stays "001" and blanks stay blank.
    if path.lower().endswith(".csv"):
        df = pd.read_csv(path, dtype=str, keep_default_na=False, nrows=nrows)
    else:
        df = pd.read_excel(path, sheet_name=sheet, dtype=str, keep_default_na=False,
                           nrows=nrows, engine=_ENGINE)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _normalise(df: pd.DataFrame, opts: Options) -> pd.DataFrame:
    df = df.astype(str)
    if opts.trim_spaces:
        df = df.apply(lambda s: s.str.strip())
    if opts.ignore_case:
        df = df.apply(lambda s: s.str.lower())
    return df


# --------------------------------------------------------------------------- compare
def compare(path_a: str, sheet_a: str, path_b: str, sheet_b: str,
            opts: Options, progress=lambda msg: None) -> Result:
    t0 = time.perf_counter()
    progress("Reading file 1...")
    raw_a = _read(path_a, sheet_a)
    progress("Reading file 2...")
    raw_b = _read(path_b, sheet_b)

    progress("Comparing headers...")
    cols_a, cols_b = list(raw_a.columns), list(raw_b.columns)
    header_issues = _compare_headers(cols_a, cols_b)

    ignore = set(opts.ignore_columns)
    common = [c for c in cols_a if c in set(cols_b) and c not in ignore]
    missing_keys = [k for k in opts.key_columns if k not in common]
    if missing_keys:
        raise ValueError(f"Key column(s) not present in both files: {', '.join(missing_keys)}")

    a = _normalise(raw_a[common], opts)
    b = _normalise(raw_b[common], opts)
    a[ROW_COL] = np.arange(len(a)) + 2   # +2 = header row + 1-based
    b[ROW_COL] = np.arange(len(b)) + 2

    progress("Aligning rows...")
    duplicate_keys = pd.DataFrame(columns=["File", "Key", "Count"])
    if opts.key_columns:
        keys = opts.key_columns
        a.index = _make_key(a, keys)
        b.index = _make_key(b, keys)
        duplicate_keys = pd.concat([_dupes(a.index, "File 1"), _dupes(b.index, "File 2")],
                                   ignore_index=True)
        # make duplicate keys unique (#2, #3...) so they still align in order
        a.index = _dedupe(a.index)
        b.index = _dedupe(b.index)
        compare_cols = [c for c in common if c not in keys]
    else:
        a.index = pd.RangeIndex(len(a)).astype(str)
        b.index = pd.RangeIndex(len(b)).astype(str)
        compare_cols = common

    in_both = a.index.intersection(b.index, sort=False)
    only_a_idx = a.index.difference(b.index, sort=False)
    only_b_idx = b.index.difference(a.index, sort=False)

    progress(f"Comparing {len(in_both):,} rows x {len(compare_cols)} columns...")
    a2 = a.loc[in_both]
    b2 = b.loc[in_both]
    va = a2[compare_cols].to_numpy(dtype=object)
    vb = b2[compare_cols].to_numpy(dtype=object)
    diff = va != vb

    # numbers: "1" == "1.0" == "1.00", optional tolerance
    for j, c in enumerate(compare_cols):
        if not diff[:, j].any():
            continue
        na = pd.to_numeric(a2[c], errors="coerce").to_numpy()
        nb = pd.to_numeric(b2[c], errors="coerce").to_numpy()
        both = ~np.isnan(na) & ~np.isnan(nb)
        if both.any():
            diff[both, j] = np.abs(na[both] - nb[both]) > opts.numeric_tolerance

    rows, cols = np.nonzero(diff)
    key_label = "Key" if opts.key_columns else "Row #"
    raw_a_vals = raw_a[compare_cols].to_numpy(dtype=object) if compare_cols else None
    raw_b_vals = raw_b[compare_cols].to_numpy(dtype=object) if compare_cols else None
    ra = a2[ROW_COL].to_numpy()[rows] - 2
    rb = b2[ROW_COL].to_numpy()[rows] - 2
    cell_diffs = pd.DataFrame({
        key_label: in_both.to_numpy()[rows] if opts.key_columns else rb + 2,
        "Column": np.array(compare_cols, dtype=object)[cols],
        "File 1 Value": raw_a_vals[ra, cols] if len(rows) else [],
        "File 2 Value": raw_b_vals[rb, cols] if len(rows) else [],
        "File 1 Excel Row": ra + 2,
        "File 2 Excel Row": rb + 2,
    })

    diff_mask_b: dict[int, set] = {}
    for r, c in zip(rb.tolist(), cell_diffs["Column"].tolist()):
        diff_mask_b.setdefault(r, set()).add(c)

    only_in_a = raw_a.iloc[a.loc[only_a_idx, ROW_COL].to_numpy() - 2].copy()
    only_in_a.insert(0, "File 1 Excel Row", a.loc[only_a_idx, ROW_COL].to_numpy())
    only_b_pos = b.loc[only_b_idx, ROW_COL].to_numpy() - 2
    only_in_b = raw_b.iloc[only_b_pos].copy()
    only_in_b.insert(0, "File 2 Excel Row", only_b_pos + 2)

    summary = {
        "File 1": os.path.basename(path_a) + f"  [{sheet_a}]",
        "File 2": os.path.basename(path_b) + f"  [{sheet_b}]",
        "Match mode": f"By key: {', '.join(opts.key_columns)}" if opts.key_columns else "Row by row",
        "Rows in File 1": len(raw_a),
        "Rows in File 2": len(raw_b),
        "Columns in File 1": len(cols_a),
        "Columns in File 2": len(cols_b),
        "Header issues": len(header_issues),
        "Rows matched": len(in_both),
        "Rows with differences": int(diff.any(axis=1).sum()) if diff.size else 0,
        "Cell differences": len(cell_diffs),
        "Rows only in File 1": len(only_in_a),
        "Rows only in File 2": len(only_in_b),
        "Duplicate keys": len(duplicate_keys),
        "Compare time (s)": round(time.perf_counter() - t0, 2),
    }
    return Result(header_issues, cell_diffs, only_in_a, only_in_b, duplicate_keys, summary,
                  raw_b, diff_mask_b, set(only_b_pos.tolist()),
                  [c for c in cols_b if c not in set(cols_a)])


def _compare_headers(cols_a, cols_b) -> pd.DataFrame:
    out = []
    set_a, set_b = set(cols_a), set(cols_b)
    lower_b = {c.lower(): c for c in cols_b}
    for i, c in enumerate(cols_a):
        if c not in set_b:
            hint = lower_b.get(c.lower())
            note = f"Only case differs: '{hint}' in File 2" if hint else ""
            out.append((c, "Missing in File 2", i + 1, "", note))
    for i, c in enumerate(cols_b):
        if c not in set_a:
            out.append((c, "Extra in File 2", "", i + 1, ""))
    common_a = [c for c in cols_a if c in set_b]
    common_b = [c for c in cols_b if c in set_a]
    if common_a != common_b:
        for c in common_a:
            pa, pb = cols_a.index(c) + 1, cols_b.index(c) + 1
            if common_a.index(c) != common_b.index(c):
                out.append((c, "Column order changed", pa, pb, ""))
    for name, cols in (("File 1", cols_a), ("File 2", cols_b)):
        for c in pd.Index(cols)[pd.Index(cols).duplicated()].unique():
            out.append((c, f"Duplicate header in {name}", "", "", ""))
        for c in cols:
            if c == "" or c.startswith("Unnamed:"):
                out.append((c, f"Blank header in {name}", "", "", ""))
    return pd.DataFrame(out, columns=["Column", "Issue", "File 1 Position",
                                      "File 2 Position", "Note"])


def _make_key(df, keys) -> pd.Index:
    if len(keys) == 1:
        return pd.Index(df[keys[0]].to_numpy(dtype=object))
    return pd.Index(df[keys].agg(" | ".join, axis=1).to_numpy(dtype=object))


def _dupes(idx: pd.Index, name: str) -> pd.DataFrame:
    vc = idx.value_counts()
    vc = vc[vc > 1]
    return pd.DataFrame({"File": name, "Key": vc.index, "Count": vc.to_numpy()})


def _dedupe(idx: pd.Index) -> pd.Index:
    if idx.is_unique:
        return idx
    s = pd.Series(idx)
    n = s.groupby(s).cumcount()
    return pd.Index(np.where(n == 0, s, s + " #" + (n + 1).astype(str)))


# --------------------------------------------------------------------------- report
def write_report(res: Result, out_path: str, progress=lambda msg: None) -> None:
    progress("Writing report...")
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as xw:
        wb = xw.book
        hdr = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white",
                             "border": 1})
        red = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
        red_hdr = wb.add_format({"bold": True, "bg_color": "#C00000", "font_color": "white",
                                 "border": 1})
        bold = wb.add_format({"bold": True})
        ok = wb.add_format({"bold": True, "font_color": "#006100", "bg_color": "#C6EFCE"})

        # Summary
        ws = wb.add_worksheet("Summary")
        ws.write(0, 0, "Excel Comparison Report", wb.add_format({"bold": True, "font_size": 14}))
        for i, (k, v) in enumerate(res.summary.items(), start=2):
            ws.write(i, 0, k, bold)
            is_issue = k in ("Header issues", "Rows with differences", "Cell differences",
                             "Rows only in File 1", "Rows only in File 2", "Duplicate keys")
            ws.write(i, 1, v, red if is_issue and v else None)
        r = len(res.summary) + 3
        ws.write(r, 0, "Result", bold)
        ws.write(r, 1, "ISSUES FOUND" if res.has_issues else "FILES MATCH",
                 red if res.has_issues else ok)
        ws.set_column(0, 0, 26)
        ws.set_column(1, 1, 50)

        _table(xw, res.header_issues, "Header Issues", hdr, red)
        cd = res.cell_diffs
        if len(cd) > EXCEL_MAX_ROWS:
            progress(f"Too many differences; first {EXCEL_MAX_ROWS:,} written")
            cd = cd.head(EXCEL_MAX_ROWS)
        _table(xw, cd, "Data Differences", hdr, red)
        _table(xw, res.only_in_a, "Rows Only In File 1", hdr, red)
        _table(xw, res.only_in_b, "Rows Only In File 2", hdr, red)
        if len(res.duplicate_keys):
            _table(xw, res.duplicate_keys, "Duplicate Keys", hdr, red)

        progress("Highlighting File 2 copy...")
        _highlighted_copy(wb, res, hdr, red_hdr, red)


def _table(xw, df: pd.DataFrame, name: str, hdr, red) -> None:
    df = df.reset_index(drop=True)
    df.to_excel(xw, sheet_name=name, index=False, startrow=1, header=False)
    ws = xw.sheets[name]
    for j, c in enumerate(df.columns):
        ws.write(0, j, c, hdr)
        width = min(max(len(str(c)), *(df[c].astype(str).str.len().head(500).tolist() or [0])) + 2, 60)
        ws.set_column(j, j, width)
    if len(df):
        ws.conditional_format(1, 0, len(df), len(df.columns) - 1,
                              {"type": "no_blanks", "format": red})
        ws.conditional_format(1, 0, len(df), len(df.columns) - 1,
                              {"type": "blanks", "format": red})
        ws.autofilter(0, 0, len(df), len(df.columns) - 1)
    ws.freeze_panes(1, 0)


def _highlighted_copy(wb, res: Result, hdr, red_hdr, red) -> None:
    """File 2 as-is; differing cells, extra rows and extra columns in red."""
    ws = wb.add_worksheet("File 2 Highlighted")
    df = res.df_b
    cols = list(df.columns)
    extra = set(res.extra_cols_b)
    for j, c in enumerate(cols):
        ws.write_string(0, j, c, red_hdr if c in extra else hdr)
        ws.set_column(j, j, max(10, min(len(c) + 2, 40)))
    values = df.to_numpy(dtype=object)
    col_pos = {c: j for j, c in enumerate(cols)}
    for i in range(len(values)):
        row = values[i]
        if i in res.rows_only_b_pos:
            ws.write_row(i + 1, 0, row, red)
            continue
        ws.write_row(i + 1, 0, row)
        for c in res.diff_mask_b.get(i, ()):
            j = col_pos[c]
            ws.write(i + 1, j, row[j], red)
    for j, c in enumerate(cols):
        if c in extra and len(values):
            ws.conditional_format(1, j, len(values), j, {"type": "no_blanks", "format": red})
    ws.freeze_panes(1, 0)
    if cols:
        ws.autofilter(0, 0, len(values), len(cols) - 1)
