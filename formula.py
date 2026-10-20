"""Excel-style formulas evaluated on whole columns at once (vectorised, safe).

    ([Igst Tax] + [Cgst Tax]) / 2
    ROUND([Invoice Amount] * 0.18, 2)
    IF([Igst Rate] = 0, [Cgst Tax] * 2, [Igst Tax])
    [First] & " " & [Last]

Columns go in [square brackets]. Operators: + - * / ^ & = <> < > <= >=
Functions: see FUNCTIONS below. Blank cells count as 0 in arithmetic.
The formula is parsed with Python's ast and only whitelisted nodes are
evaluated, so nothing in a mapping file can run arbitrary code.
"""
from __future__ import annotations

import ast
import re

import numpy as np
import pandas as pd

HELP = """Formula syntax
  Columns:    [Column Name]           e.g. [Igst Tax]
  Maths:      + - * / ^  and ( )      e.g. ([Igst Tax] + [Cgst Tax]) / 2
  Join text:  &                       e.g. [City] & "-" & [State]
  Compare:    = <> < > <= >=          e.g. IF([Igst Rate] = 0, "Intra", "Inter")
  Functions:  ROUND(x, n)  ABS(x)  SUM(a, b, ...)  MIN(...)  MAX(...)  AVERAGE(...)
              IF(cond, a, b)  AND(...)  OR(...)  NOT(x)  IFERROR(x, alt)  ISBLANK(x)
              UPPER(x)  LOWER(x)  TRIM(x)  LEN(x)  LEFT(x, n)  RIGHT(x, n)
              MID(x, start, n)  CONCAT(a, b, ...)  VALUE(x)
  Blank cells count as 0 in maths. A leading "=" is optional."""

ERROR = "#ERROR"


class FormulaError(ValueError):
    pass


def is_formula(text: str) -> bool:
    t = str(text).strip()
    return t.startswith("=") or ("[" in t and "]" in t)


def find_column(name: str, columns) -> str | None:
    """Exact match first, then ignoring case and surrounding spaces."""
    if name in columns:
        return name
    norm = name.strip().lower()
    for c in columns:
        if str(c).strip().lower() == norm:
            return c
    return None


# ------------------------------------------------------------------ translate
_TOKEN = re.compile(r'\s+|"(?:[^"]|"")*"|\[[^\]]*\]|<>|<=|>=|[A-Za-z_][A-Za-z0-9_.]*|\d+\.?\d*(?:[eE][-+]?\d+)?|\.\d+|.')


def _translate(expr: str, columns) -> tuple[str, dict]:
    """Excel-ish text -> Python expression text + {var: column}."""
    expr = str(expr).strip()
    if expr.startswith("="):
        expr = expr[1:]
    if not expr:
        raise FormulaError("Formula is empty")
    out, cols = [], {}
    for tok in _TOKEN.findall(expr):
        if tok.startswith('"'):
            out.append(repr(tok[1:-1].replace('""', '"')))
        elif tok.startswith("["):
            name = tok[1:-1]
            col = find_column(name, columns)
            if col is None:
                raise FormulaError(f"Column [{name}] not found")
            var = next((v for v, c in cols.items() if c == col), f"_c{len(cols)}")
            cols[var] = col
            out.append(var)
        elif tok == "<>":
            out.append("!=")
        elif tok == "=":
            out.append("==")
        elif tok == "^":
            out.append("**")
        elif tok[0].isalpha() or tok[0] == "_":
            up = tok.upper()
            if up in ("TRUE", "FALSE"):
                out.append(up.title())
            elif up in FUNCTIONS:
                out.append(up)
            else:
                raise FormulaError(f"Unknown name '{tok}' (put column names in [brackets])")
        else:
            out.append(tok)
    return "".join(out), cols


def validate(expr: str, columns) -> None:
    src, _ = _translate(expr, columns)
    try:
        tree = ast.parse(src, mode="eval")
    except SyntaxError:
        raise FormulaError(f"Cannot understand formula: {expr}") from None
    for node in ast.walk(tree):
        if not isinstance(node, _ALLOWED):
            raise FormulaError(f"Not allowed in formula: {type(node).__name__}")


# ------------------------------------------------------------------ evaluate
def evaluate(expr: str, df: pd.DataFrame) -> pd.Series:
    """Evaluate on every row of df (all text columns). Returns a text Series."""
    validate(expr, df.columns)
    src, cols = _translate(expr, df.columns)
    env = {v: df[c].reset_index(drop=True) for v, c in cols.items()}
    n = len(df)
    try:
        with np.errstate(all="ignore"):
            res = _Eval(env, n).visit(ast.parse(src, mode="eval").body)
    except FormulaError:
        raise
    except Exception as e:  # unexpected maths/type problems -> readable message
        raise FormulaError(f"Error in formula {expr}: {e}") from None
    return to_text(_series(res, n))


def to_text(s: pd.Series) -> pd.Series:
    if pd.api.types.is_bool_dtype(s):
        return s.map({True: "TRUE", False: "FALSE"})
    if pd.api.types.is_numeric_dtype(s):
        return s.map(_fmt_num)
    return s.map(_fmt_any)


def _fmt_any(v) -> str:
    if isinstance(v, (bool, np.bool_)):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float, np.integer, np.floating)):
        return _fmt_num(v)
    if v is None:
        return ERROR
    return str(v)


def _fmt_num(x) -> str:
    if x is None or not np.isfinite(x):
        return ERROR
    x = round(float(x), 10)
    if x == int(x) and abs(x) < 1e15:
        return str(int(x))
    return f"{x:.10f}".rstrip("0").rstrip(".")


def _series(x, n) -> pd.Series:
    if isinstance(x, pd.Series):
        return x.reset_index(drop=True)
    if isinstance(x, np.ndarray):
        return pd.Series(x)
    return pd.Series([x] * n)


def _num(x):
    if isinstance(x, pd.Series):
        if pd.api.types.is_bool_dtype(x):
            return x.astype(float)
        if pd.api.types.is_numeric_dtype(x):
            return x.astype(float)
        s = x.astype(str).str.strip()
        return pd.to_numeric(s.where(s != "", "0"), errors="coerce")
    if isinstance(x, bool):
        return float(x)
    if isinstance(x, (int, float)):
        return float(x)
    try:
        return float(str(x).strip() or 0)
    except ValueError:
        return np.nan


def _text(x):
    if isinstance(x, pd.Series):
        if pd.api.types.is_object_dtype(x):
            return x.map(lambda v: "" if v is None or (isinstance(v, float) and np.isnan(v)) else _fmt_any(v))
        return to_text(x)
    if isinstance(x, bool):
        return "TRUE" if x else "FALSE"
    if isinstance(x, float):
        return _fmt_num(x)
    return str(x)


def _norm(x):
    t = _text(x)
    return t.str.strip().str.lower() if isinstance(t, pd.Series) else t.strip().lower()


def _truthy(x):
    if isinstance(x, pd.Series) and pd.api.types.is_bool_dtype(x):
        return x
    if isinstance(x, (bool, np.bool_)):
        return bool(x)
    v = _num(x)
    return (v.fillna(0) != 0) if isinstance(v, pd.Series) else bool(v)


def _clean(v):
    """Division by zero etc. -> NaN (shown as #ERROR)."""
    if isinstance(v, pd.Series):
        return v.replace([np.inf, -np.inf], np.nan)
    return np.nan if isinstance(v, float) and not np.isfinite(v) else v


def _round(x, n=0):
    v, n = _num(x), int(_num(n))
    f = 10.0 ** n
    return np.sign(v) * np.floor(np.abs(v) * f + 0.5 + 1e-9) / f   # Excel: half away from zero


def _nary(func):
    def run(*args):
        vals = [_num(a) for a in args]
        if not vals:
            return 0.0
        if any(isinstance(v, pd.Series) for v in vals):
            n = max(len(v) for v in vals if isinstance(v, pd.Series))
            vals = [_series(v, n) for v in vals]
            return pd.Series(func(np.vstack([v.to_numpy(float) for v in vals]), axis=0))
        return float(func(np.array(vals, dtype=float)))
    return run


def _str_fn(f):
    return lambda x, *a: x.astype(str).map(lambda s: f(s, *a)) if isinstance(x, pd.Series) else f(str(x), *a)


def _if(cond, a, b=False):
    c = _truthy(cond)
    if not isinstance(c, pd.Series):
        return a if c else b
    n = len(c)
    sa, sb = _series(a, n), _series(b, n)
    out = sa.where(c.to_numpy(), sb)
    return out


def _iferror(x, alt):
    if isinstance(x, pd.Series):
        bad = x.isna() | (x.astype(str) == ERROR)
        return x.where(~bad, _series(alt, len(x)))
    return alt if x is None or (isinstance(x, float) and np.isnan(x)) else x


def _logic(func):
    def run(*args):
        vals = [_truthy(a) for a in args]
        if any(isinstance(v, pd.Series) for v in vals):
            n = max(len(v) for v in vals if isinstance(v, pd.Series))
            return pd.Series(func(np.vstack([_series(v, n).to_numpy(bool) for v in vals]), axis=0))
        return bool(func(np.array(vals, dtype=bool)))
    return run


def _isblank(x):
    if isinstance(x, pd.Series):
        return x.isna() | (x.astype(str).str.strip() == "")
    return str(x).strip() == ""


def _concat(*args):
    out = ""
    for a in args:
        out = _text(out) + _text(a) if isinstance(a, pd.Series) or isinstance(out, pd.Series) else out + _text(a)
    return out


def _not(x):
    t = _truthy(x)
    return ~t if isinstance(t, pd.Series) else (not t)


FUNCTIONS = {
    "ROUND": _round,
    "ABS": lambda x: np.abs(_num(x)),
    "SUM": _nary(np.nansum),
    "MIN": _nary(np.nanmin),
    "MAX": _nary(np.nanmax),
    "AVERAGE": _nary(np.nanmean),
    "IF": _if,
    "IFERROR": _iferror,
    "AND": _logic(np.all),
    "OR": _logic(np.any),
    "NOT": _not,
    "ISBLANK": _isblank,
    "UPPER": _str_fn(lambda s: s.upper()),
    "LOWER": _str_fn(lambda s: s.lower()),
    "TRIM": _str_fn(lambda s: " ".join(s.split())),
    "LEN": lambda x: x.astype(str).str.len().astype(float) if isinstance(x, pd.Series) else float(len(str(x))),
    "LEFT": lambda x, n=1: _str_fn(lambda s, k: s[:k])(x, int(_num(n))),
    "RIGHT": lambda x, n=1: _str_fn(lambda s, k: s[-k:] if k else "")(x, int(_num(n))),
    "MID": lambda x, st, n: _str_fn(lambda s, a, k: s[a - 1:a - 1 + k])(x, int(_num(st)), int(_num(n))),
    "CONCAT": _concat,
    "VALUE": _num,
}

_ALLOWED = (ast.Expression, ast.BinOp, ast.UnaryOp, ast.Compare, ast.Call, ast.Name,
            ast.Constant, ast.Load, ast.Add, ast.Sub, ast.Mult, ast.Div, ast.Pow, ast.Mod,
            ast.BitAnd, ast.USub, ast.UAdd, ast.Eq, ast.NotEq, ast.Lt, ast.LtE, ast.Gt,
            ast.GtE)

_MATH = {ast.Add: np.add, ast.Sub: np.subtract, ast.Mult: np.multiply,
         ast.Div: np.divide, ast.Pow: np.power, ast.Mod: np.mod}
_CMP = {ast.Eq: np.equal, ast.NotEq: np.not_equal, ast.Lt: np.less,
        ast.LtE: np.less_equal, ast.Gt: np.greater, ast.GtE: np.greater_equal}


class _Eval(ast.NodeVisitor):
    def __init__(self, env, n):
        self.env, self.n = env, n

    def visit_Constant(self, node):
        return node.value

    def visit_Name(self, node):
        if node.id in self.env:
            return self.env[node.id]
        if node.id in ("True", "False"):
            return node.id == "True"
        raise FormulaError(f"Unknown name {node.id}")

    def visit_BinOp(self, node):
        a, b = self.visit(node.left), self.visit(node.right)
        if isinstance(node.op, ast.BitAnd):
            return _concat(a, b)
        return _clean(_MATH[type(node.op)](_num(a), _num(b)))

    def visit_UnaryOp(self, node):
        v = _num(self.visit(node.operand))
        return -v if isinstance(node.op, ast.USub) else v

    def visit_Compare(self, node):
        if len(node.ops) != 1:
            raise FormulaError("Chained comparisons are not supported; use AND()")
        a, b = self.visit(node.left), self.visit(node.comparators[0])
        kind = type(node.ops[0])
        na, nb = _num(a), _num(b)
        num_res = _CMP[kind](na, nb)
        if kind not in (ast.Eq, ast.NotEq):
            return num_res
        # = and <> : numbers compare as numbers, text compares ignoring case (like Excel)
        ta, tb = _norm(a), _norm(b)
        txt_res = (ta == tb) if kind is ast.Eq else (ta != tb)
        both = pd.notna(na) & pd.notna(nb)
        if not any(isinstance(v, pd.Series) for v in (na, nb, ta, tb)):
            return bool(num_res if both else txt_res)
        n = self.n
        return pd.Series(np.where(_series(both, n).to_numpy(bool),
                                  _series(num_res, n).to_numpy(bool),
                                  _series(txt_res, n).to_numpy(bool)))

    def visit_Call(self, node):
        if not isinstance(node.func, ast.Name) or node.func.id not in FUNCTIONS:
            raise FormulaError("Unknown function")
        if node.keywords:
            raise FormulaError("Keyword arguments are not supported")
        args = [self.visit(a) for a in node.args]
        try:
            return _clean(FUNCTIONS[node.func.id](*args))
        except TypeError:
            raise FormulaError(f"Wrong number of arguments for {node.func.id}()") from None

    def generic_visit(self, node):
        raise FormulaError(f"Not allowed in formula: {type(node).__name__}")
