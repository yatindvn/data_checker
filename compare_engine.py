"""Core logic: compare two Excel/CSV sheets (headers + data) and write a report.

Three ways to compare:
  compare()             one sheet vs one sheet
  compare_all_sheets()  every sheet with the same name in both files
  mapping=...           a mapping file says which File 1 column / formula matches
                        which File 2 column / formula (works with both of the above)

All comparisons are vectorised with pandas/numpy so large files stay fast.
"""
from __future__ import annotations

import os
import re
import time
from dataclasses import dataclass, field, replace

import numpy as np
import pandas as pd

import formula

try:
    import python_calamine  # noqa: F401  (much faster .xlsx reader)
    _ENGINE = "calamine"
except ImportError:
    _ENGINE = None

EXCEL_MAX_ROWS = 1_048_575
ROW_COL = "__excel_row__"
FLOAT_EPS = 1e-9          # hides float noise such as 0.1 + 0.2 = 0.30000000000000004


@dataclass
class Options:
    key_columns: list[str] = field(default_factory=list)  # empty -> compare row by row
    ignore_case: bool = False
    trim_spaces: bool = True
    numeric_tolerance: float = 0.0
    ignore_columns: list[str] = field(default_factory=list)
    column_tolerance: dict = field(default_factory=dict)  # per column, from a mapping file
    leading_zeros_matter: bool = False  # True: "001" != "1" (codes, PINs); "10" still == "10.00"


@dataclass
class MapRow:
    output: str
    file1: str
    file2: str
    key: bool = False
    tolerance: float | None = None


@dataclass
class Result:
    header_issues: pd.DataFrame
    cell_diffs: pd.DataFrame
    only_in_a: pd.DataFrame
    only_in_b: pd.DataFrame
    duplicate_keys: pd.DataFrame
    summary: dict
    df_b: pd.DataFrame          # file B as read (or as mapped) for the highlighted copy
    diff_mask_b: dict           # {b_row_position: set(column names)}
    rows_only_b_pos: set
    extra_cols_b: list
    mapping_used: pd.DataFrame | None = None

    @property
    def has_issues(self) -> bool:
        return any(len(x) for x in (self.header_issues, self.cell_diffs,
                                    self.only_in_a, self.only_in_b, self.duplicate_keys))


@dataclass
class SheetResult:
    """One entry of a report: a compared sheet pair, or a sheet that could not be compared."""
    name: str
    result: Result | None = None
    error: str | None = None


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


# --------------------------------------------------------------------------- mapping file
_MAP_ALIASES = {
    "output": ("output column", "output", "column name", "name", "field"),
    "file1": ("file 1", "file1", "file 1 column", "file 1 column / logic", "source 1"),
    "file2": ("file 2", "file2", "file 2 column", "file 2 column / logic", "source 2"),
    "key": ("key", "is key", "key column"),
    "tolerance": ("tolerance", "number tolerance"),
}


def load_mapping(path: str) -> list[MapRow]:
    """Read the first sheet of a mapping file (Output Column | File 1 | File 2 | Key | Tolerance)."""
    if path.lower().endswith(".csv"):
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
    else:
        df = pd.read_excel(path, sheet_name=0, dtype=str, keep_default_na=False, engine=_ENGINE)
    cols = {str(c).strip().lower(): c for c in df.columns}
    pick = {k: next((cols[a] for a in aliases if a in cols), None)
            for k, aliases in _MAP_ALIASES.items()}
    if pick["output"] is None or (pick["file1"] is None and pick["file2"] is None):
        raise ValueError("Mapping file needs columns 'Output Column', 'File 1' and 'File 2' "
                         f"(found: {', '.join(map(str, df.columns))})")
    rows = []
    for _, r in df.iterrows():
        out = str(r[pick["output"]]).strip()
        if not out:
            continue
        get = lambda k: str(r[pick[k]]).strip() if pick[k] is not None else ""
        tol = get("tolerance")
        try:
            tol_v = float(tol) if tol else None
        except ValueError:
            raise ValueError(f"Mapping row '{out}': tolerance '{tol}' is not a number") from None
        rows.append(MapRow(out, get("file1"), get("file2"),
                           get("key").lower() in ("y", "yes", "true", "1", "x"), tol_v))
    if not rows:
        raise ValueError("Mapping file has no rows")
    dupes = pd.Series([r.output for r in rows]).duplicated()
    if dupes.any():
        raise ValueError(f"Mapping file lists '{rows[int(dupes.idxmax())].output}' more than once")
    return rows


def _resolve(text: str, output: str, df: pd.DataFrame, which: str, ai) -> tuple[pd.Series, str, str]:
    """One mapping cell -> (values, what was used, note)."""
    text = text.strip() or output          # blank = same name as the output column
    col = formula.find_column(text, df.columns)
    if col is not None and not formula.is_formula(text):
        return df[col], f"[{col}]", ""
    if formula.is_formula(text):
        try:
            return formula.evaluate(text, df), text.lstrip("="), ""
        except formula.FormulaError as e:
            raise ValueError(f"{which}, '{output}': {e}") from None
    if ai is None:
        raise ValueError(f"{which}, '{output}': '{text}' is not a column in {which} and not a "
                         "formula. Put column names in [brackets], e.g. [Igst Tax] + [Cgst Tax]")
    try:
        f = ai.to_formula(text, list(df.columns))
    except Exception as e:
        raise ValueError(f"{which}, '{output}': {e}") from None
    return formula.evaluate(f, df), f, f"AI converted: \"{text}\""


def apply_mapping(raw: pd.DataFrame, mapping: list[MapRow], which: str, ai=None):
    """Build the mapped table for one file. Returns (DataFrame, [(used, note), ...])."""
    out, used = {}, []
    for m in mapping:
        text = m.file1 if which == "File 1" else m.file2
        values, what, note = _resolve(text, m.output, raw, which, ai)
        out[m.output] = values.reset_index(drop=True).astype(str)
        used.append((what, note))
    return pd.DataFrame(out, columns=[m.output for m in mapping]), used


# --------------------------------------------------------------------------- compare
def compare(path_a: str, sheet_a: str, path_b: str, sheet_b: str, opts: Options,
            progress=lambda msg: None, mapping: list[MapRow] | None = None, ai=None) -> Result:
    t0 = time.perf_counter()
    progress(f"Reading file 1 [{sheet_a}]...")
    raw_a = _read(path_a, sheet_a)
    progress(f"Reading file 2 [{sheet_b}]...")
    raw_b = _read(path_b, sheet_b)
    labels = (os.path.basename(path_a) + f"  [{sheet_a}]",
              os.path.basename(path_b) + f"  [{sheet_b}]")
    mapping_used = None
    if mapping:
        progress("Applying mapping / formulas...")
        raw_a, used_a = apply_mapping(raw_a, mapping, "File 1", ai)
        raw_b, used_b = apply_mapping(raw_b, mapping, "File 2", ai)
        opts = replace(opts, key_columns=[m.output for m in mapping if m.key],
                       column_tolerance={m.output: m.tolerance for m in mapping
                                         if m.tolerance is not None})
        mapping_used = pd.DataFrame({
            "Output Column": [m.output for m in mapping],
            "Key": ["Y" if m.key else "" for m in mapping],
            "File 1 (as written)": [m.file1 or m.output for m in mapping],
            "File 1 used": [u[0] for u in used_a],
            "File 2 (as written)": [m.file2 or m.output for m in mapping],
            "File 2 used": [u[0] for u in used_b],
            "Note": ["; ".join(x for x in (a[1], b[1]) if x) for a, b in zip(used_a, used_b)],
        })
    res = compare_frames(raw_a, raw_b, opts, labels, progress)
    res.mapping_used = mapping_used
    res.summary["Compare time (s)"] = round(time.perf_counter() - t0, 2)
    return res


def compare_all_sheets(path_a: str, path_b: str, opts: Options, progress=lambda msg: None,
                       mapping: list[MapRow] | None = None, ai=None) -> list[SheetResult]:
    """Compare every sheet whose name is in both files (ignoring case and spaces)."""
    sa, sb = list_sheets(path_a), list_sheets(path_b)
    norm = lambda s: s.strip().lower()
    b_by = {norm(s): s for s in sb}
    out = []
    for s in sa:
        if norm(s) not in b_by:
            out.append(SheetResult(s, error="Sheet missing in File 2"))
            continue
        try:
            out.append(SheetResult(s, compare(path_a, s, path_b, b_by[norm(s)], opts,
                                              progress, mapping, ai)))
        except Exception as e:   # one bad sheet must not stop the others
            out.append(SheetResult(s, error=str(e)))
            progress(f"  [{s}] skipped: {e}")
    a_norms = {norm(s) for s in sa}
    out += [SheetResult(s, error="Sheet missing in File 1") for s in sb if norm(s) not in a_norms]
    return out


def compare_frames(raw_a: pd.DataFrame, raw_b: pd.DataFrame, opts: Options,
                   labels=("File 1", "File 2"), progress=lambda msg: None) -> Result:
    t0 = time.perf_counter()
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
        if opts.leading_zeros_matter:
            # values like "001" stay text, so they only match the exact same text
            both &= ~(_leading_zero(a2[c]) | _leading_zero(b2[c]))
        if both.any():
            tol = max(opts.column_tolerance.get(c, opts.numeric_tolerance), FLOAT_EPS)
            diff[both, j] = np.abs(na[both] - nb[both]) > tol

    rows, cols = np.nonzero(diff)
    key_label = "Key" if opts.key_columns else "Row #"
    raw_a_vals = raw_a[compare_cols].to_numpy(dtype=object) if compare_cols else None
    raw_b_vals = raw_b[compare_cols].to_numpy(dtype=object) if compare_cols else None
    ra = a2[ROW_COL].to_numpy()[rows] - 2
    rb = b2[ROW_COL].to_numpy()[rows] - 2
    va_out = raw_a_vals[ra, cols] if len(rows) else np.array([], dtype=object)
    vb_out = raw_b_vals[rb, cols] if len(rows) else np.array([], dtype=object)
    cell_diffs = pd.DataFrame({
        key_label: in_both.to_numpy()[rows] if opts.key_columns else rb + 2,
        "Column": np.array(compare_cols, dtype=object)[cols],
        "File 1 Value": va_out,
        "File 2 Value": vb_out,
        "Difference": _num_diff(va_out, vb_out),
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
        "File 1": labels[0],
        "File 2": labels[1],
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


def _leading_zero(s: pd.Series) -> np.ndarray:
    """'001', '-007', '0123.5' -> True;  '0', '0.5', '10' -> False."""
    return s.str.match(r"^\s*[+-]?0\d").fillna(False).to_numpy(bool)


def _num_diff(va, vb) -> np.ndarray:
    """File 2 - File 1 where both are numbers, else blank."""
    na = pd.to_numeric(pd.Series(va, dtype=object), errors="coerce").to_numpy()
    nb = pd.to_numeric(pd.Series(vb, dtype=object), errors="coerce").to_numpy()
    d = np.round(nb - na, 10)
    return np.where(np.isnan(d), "", d.astype(object))


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


# --------------------------------------------------------------------------- mapping template
def write_mapping_template(path_a: str, sheet_a: str, path_b: str, sheet_b: str,
                           out_path: str, key_columns=()) -> None:
    """Starter mapping file: File 1 columns, matched to File 2 columns by similar name."""
    ha, hb = read_headers(path_a, sheet_a), read_headers(path_b, sheet_b)
    squash = lambda s: re.sub(r"[\s_\-./]", "", s.lower())
    b_by = {squash(c): c for c in hb}
    rows = []
    for c in ha:
        match = b_by.get(squash(c), "")
        rows.append((c, c, match,"Y" if c in key_columns else "", "",
                     "" if match else "No similar column in File 2 - fill in, or delete this row"))
    df = pd.DataFrame(rows, columns=["Output Column", "File 1", "File 2", "Key", "Tolerance", "Note"])
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as xw:
        df.to_excel(xw, sheet_name="Mapping", index=False)
        wb, ws = xw.book, xw.sheets["Mapping"]
        hdr = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white"})
        red = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
        for j, c in enumerate(df.columns):
            ws.write(0, j, c, hdr)
        ws.set_column(0, 2, 34)
        ws.set_column(3, 4, 10)
        ws.set_column(5, 5, 55)
        ws.conditional_format(1, 2, len(df), 2, {"type": "blanks", "format": red})
        ws.freeze_panes(1, 0)
        help_ws = wb.add_worksheet("How to use")
        lines = [
            "How to fill the Mapping sheet",
            "",
            "Output Column : the name shown in the report.",
            "File 1 / File 2 : for each file, one of:",
            "   - a column name, e.g.  Invoice Number",
            "   - a formula, e.g.  ([Igst Tax] + [Cgst Tax]) / 2",
            "   - plain English, e.g.  add igst amt and cgst amt then divide by 2",
            "     (needs a Claude API key; the formula it makes is shown in the report)",
            "   - blank = same name as Output Column",
            "Key : Y on the column(s) that identify a row (e.g. Invoice Number).",
            "Tolerance : optional, allowed number difference for this row (e.g. 0.01).",
            "Only rows in this sheet are compared. Delete rows you don't need.",
            "",
        ] + formula.HELP.splitlines()
        help_ws.set_column(0, 0, 100)
        for i, line in enumerate(lines):
            help_ws.write(i, 0, line, wb.add_format({"bold": True}) if i == 0 else None)


# --------------------------------------------------------------------------- report
def write_report(results, out_path: str, progress=lambda msg: None) -> None:
    """results: a Result, or a list of SheetResult (one per compared sheet)."""
    if isinstance(results, Result):
        results = [SheetResult("", results)]
    multi = len(results) > 1
    progress("Writing report...")
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as xw:
        wb = xw.book
        fmt = {
            "hdr": wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white",
                                  "border": 1}),
            "red": wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"}),
            "red_hdr": wb.add_format({"bold": True, "bg_color": "#C00000", "font_color": "white",
                                      "border": 1}),
            "bold": wb.add_format({"bold": True}),
            "ok": wb.add_format({"bold": True, "font_color": "#006100", "bg_color": "#C6EFCE"}),
            "title": wb.add_format({"bold": True, "font_size": 14}),
        }
        if multi:
            _summary_multi(wb, results, fmt)
        else:
            _summary_single(wb, results[0], fmt)

        ok = [r for r in results if r.result is not None]

        def combined(attr):
            frames = []
            for r in ok:
                df = getattr(r.result, attr).reset_index(drop=True)
                if multi:
                    df.insert(0, "Sheet", r.name)
                frames.append(df)
            return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

        mapped = next((r.result.mapping_used for r in ok if r.result.mapping_used is not None), None)
        if mapped is not None:
            _table(xw, mapped, "Mapping Used", fmt, red=False)
        _table(xw, combined("header_issues"), "Header Issues", fmt)
        cd = combined("cell_diffs")
        if len(cd) > EXCEL_MAX_ROWS:
            progress(f"Too many differences; first {EXCEL_MAX_ROWS:,} written")
            cd = cd.head(EXCEL_MAX_ROWS)
        _table(xw, cd, "Data Differences", fmt)
        _table(xw, combined("only_in_a"), "Rows Only In File 1", fmt)
        _table(xw, combined("only_in_b"), "Rows Only In File 2", fmt)
        dk = combined("duplicate_keys")
        if len(dk):
            _table(xw, dk, "Duplicate Keys", fmt)

        used = set(wb.sheetnames)
        for r in ok:
            progress(f"Highlighting File 2 copy {r.name}...".replace("  ", " "))
            name = "File 2 Highlighted" if not multi else _sheet_name(f"F2 {r.name}", used)
            used.add(name)
            _highlighted_copy(wb, r.result, fmt, name)


def _summary_single(wb, sr: SheetResult, fmt) -> None:
    res = sr.result
    ws = wb.add_worksheet("Summary")
    ws.write(0, 0, "Excel Comparison Report", fmt["title"])
    for i, (k, v) in enumerate(res.summary.items(), start=2):
        ws.write(i, 0, k, fmt["bold"])
        ws.write(i, 1, v, fmt["red"] if k in _ISSUE_KEYS and v else None)
    r = len(res.summary) + 3
    ws.write(r, 0, "Result", fmt["bold"])
    ws.write(r, 1, "ISSUES FOUND" if res.has_issues else "FILES MATCH",
             fmt["red"] if res.has_issues else fmt["ok"])
    ws.set_column(0, 0, 26)
    ws.set_column(1, 1, 50)


_ISSUE_KEYS = ("Header issues", "Rows with differences", "Cell differences",
               "Rows only in File 1", "Rows only in File 2", "Duplicate keys")


def _summary_multi(wb, results, fmt) -> None:
    ws = wb.add_worksheet("Summary")
    ws.write(0, 0, "Excel Comparison Report - all sheets", fmt["title"])
    cols = ["Sheet", "Result", "Rows in File 1", "Rows in File 2", *_ISSUE_KEYS, "Details"]
    for j, c in enumerate(cols):
        ws.write(2, j, c, fmt["hdr"])
    any_issue = False
    for i, sr in enumerate(results, start=3):
        ws.write(i, 0, sr.name, fmt["bold"])
        if sr.result is None:
            any_issue = True
            ws.write(i, 1, "NOT COMPARED", fmt["red"])
            ws.write(i, len(cols) - 1, sr.error, fmt["red"])
            continue
        s, bad = sr.result.summary, sr.result.has_issues
        any_issue |= bad
        ws.write(i, 1, "ISSUES FOUND" if bad else "MATCH", fmt["red"] if bad else fmt["ok"])
        ws.write(i, 2, s["Rows in File 1"])
        ws.write(i, 3, s["Rows in File 2"])
        for j, k in enumerate(_ISSUE_KEYS, start=4):
            ws.write(i, j, s[k], fmt["red"] if s[k] else None)
        ws.write(i, len(cols) - 1, s["Match mode"])
    r = len(results) + 4
    ws.write(r, 0, "Overall", fmt["bold"])
    ws.write(r, 1, "ISSUES FOUND" if any_issue else "FILES MATCH",
             fmt["red"] if any_issue else fmt["ok"])
    ws.set_column(0, 0, 24)
    ws.set_column(1, len(cols) - 2, 16)
    ws.set_column(len(cols) - 1, len(cols) - 1, 60)
    ws.freeze_panes(3, 1)


def _sheet_name(name: str, used: set) -> str:
    base = re.sub(r"[\[\]:*?/\\]", "_", name)[:31]
    cand, i = base, 2
    while cand in used:
        suffix = f" ({i})"
        cand, i = base[:31 - len(suffix)] + suffix, i + 1
    return cand


def _table(xw, df: pd.DataFrame, name: str, fmt, red: bool = True) -> None:
    df = df.reset_index(drop=True)
    df.to_excel(xw, sheet_name=name, index=False, startrow=1, header=False)
    ws = xw.sheets[name]
    for j, c in enumerate(df.columns):
        ws.write(0, j, c, fmt["hdr"])
        sample = df.iloc[:500, j].astype(str).str.len().tolist()
        ws.set_column(j, j, min(max([len(str(c)), *sample]) + 2, 60))
    if len(df):
        if red:
            ws.conditional_format(1, 0, len(df), len(df.columns) - 1,
                                  {"type": "formula", "criteria": "TRUE", "format": fmt["red"]})
        ws.autofilter(0, 0, len(df), len(df.columns) - 1)
    ws.freeze_panes(1, 0)


def _highlighted_copy(wb, res: Result, fmt, name: str) -> None:
    """File 2 as-is; differing cells, extra rows and extra columns in red."""
    ws = wb.add_worksheet(name)
    df = res.df_b
    cols = list(df.columns)
    extra = set(res.extra_cols_b)
    red = fmt["red"]
    for j, c in enumerate(cols):
        ws.write_string(0, j, c, fmt["red_hdr"] if c in extra else fmt["hdr"])
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
