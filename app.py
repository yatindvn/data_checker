"""Excel Compare - desktop UI.

Pick two Excel/CSV files, choose sheets and (optionally) key columns, and get
an Excel report with every header/data issue highlighted in red.

Run:  python app.py
"""
from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import compare_engine as ce

FILETYPES = [("Excel / CSV", "*.xlsx *.xlsm *.xls *.csv"), ("All files", "*.*")]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Compare")
        self.geometry("760x640")
        self.minsize(680, 560)
        self.q: queue.Queue = queue.Queue()
        self.paths = [tk.StringVar(), tk.StringVar()]
        self.sheets = [tk.StringVar(), tk.StringVar()]
        self.sheet_boxes: list[ttk.Combobox] = []
        self.out_path = tk.StringVar()
        self.ignore_case = tk.BooleanVar(value=False)
        self.trim = tk.BooleanVar(value=True)
        self.tolerance = tk.StringVar(value="0")
        self.status = tk.StringVar(value="Select two files to compare.")
        self._build()
        self.after(100, self._poll)

    # ------------------------------------------------------------------ layout
    def _build(self):
        pad = {"padx": 8, "pady": 4}
        files = ttk.LabelFrame(self, text="Files")
        files.pack(fill="x", **pad)
        for i in range(2):
            ttk.Label(files, text=f"File {i + 1}:").grid(row=i, column=0, sticky="w", **pad)
            ttk.Entry(files, textvariable=self.paths[i]).grid(row=i, column=1, sticky="ew", **pad)
            ttk.Button(files, text="Browse...", command=lambda i=i: self._browse(i)).grid(
                row=i, column=2, **pad)
            ttk.Label(files, text="Sheet:").grid(row=i, column=3, **pad)
            cb = ttk.Combobox(files, textvariable=self.sheets[i], state="readonly", width=18)
            cb.grid(row=i, column=4, **pad)
            cb.bind("<<ComboboxSelected>>", lambda e: self._load_columns())
            self.sheet_boxes.append(cb)
        files.columnconfigure(1, weight=1)

        keys = ttk.LabelFrame(self, text="Match rows by key column(s)  "
                                         "(select none = compare row by row)")
        keys.pack(fill="both", expand=True, **pad)
        self.key_list = tk.Listbox(keys, selectmode="multiple", exportselection=False, height=8)
        sb = ttk.Scrollbar(keys, command=self.key_list.yview)
        self.key_list.config(yscrollcommand=sb.set)
        self.key_list.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=6)
        sb.pack(side="left", fill="y", pady=6)
        side = ttk.Frame(keys)
        side.pack(side="left", fill="y", padx=8)
        ttk.Label(side, text="Ignore columns (comma separated):").pack(anchor="w")
        self.ignore_cols = ttk.Entry(side, width=32)
        self.ignore_cols.pack(anchor="w", pady=(0, 8))
        ttk.Checkbutton(side, text="Ignore upper/lower case", variable=self.ignore_case).pack(anchor="w")
        ttk.Checkbutton(side, text="Trim leading/trailing spaces", variable=self.trim).pack(anchor="w")
        tol = ttk.Frame(side)
        tol.pack(anchor="w", pady=6)
        ttk.Label(tol, text="Number tolerance:").pack(side="left")
        ttk.Entry(tol, textvariable=self.tolerance, width=10).pack(side="left", padx=4)

        out = ttk.LabelFrame(self, text="Report")
        out.pack(fill="x", **pad)
        ttk.Entry(out, textvariable=self.out_path).pack(side="left", fill="x", expand=True, **pad)
        ttk.Button(out, text="Save as...", command=self._browse_out).pack(side="left", **pad)

        run = ttk.Frame(self)
        run.pack(fill="x", **pad)
        self.run_btn = ttk.Button(run, text="Compare", command=self._start)
        self.run_btn.pack(side="left", padx=8)
        self.open_btn = ttk.Button(run, text="Open report", state="disabled",
                                   command=lambda: os.startfile(self.out_path.get()))
        self.open_btn.pack(side="left")
        self.bar = ttk.Progressbar(run, mode="indeterminate")
        self.bar.pack(side="left", fill="x", expand=True, padx=8)

        self.log = tk.Text(self, height=10, state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True, **pad)
        ttk.Label(self, textvariable=self.status, relief="sunken", anchor="w").pack(fill="x")

    # ------------------------------------------------------------------ actions
    def _browse(self, i):
        path = filedialog.askopenfilename(title=f"Select file {i + 1}", filetypes=FILETYPES)
        if not path:
            return
        self.paths[i].set(path)
        try:
            names = ce.list_sheets(path)
        except Exception as e:
            messagebox.showerror("Cannot open file", str(e))
            return
        self.sheet_boxes[i]["values"] = names
        other = self.sheets[1 - i].get()
        self.sheets[i].set(other if other in names else names[0])
        if not self.out_path.get():
            base = os.path.splitext(path)[0]
            self.out_path.set(base + "_comparison_report.xlsx")
        self._load_columns()

    def _browse_out(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")],
                                            initialfile="comparison_report.xlsx")
        if path:
            self.out_path.set(path)

    def _load_columns(self):
        if not all(p.get() for p in self.paths):
            return
        try:
            ha = ce.read_headers(self.paths[0].get(), self.sheets[0].get())
            hb = set(ce.read_headers(self.paths[1].get(), self.sheets[1].get()))
        except Exception as e:
            messagebox.showerror("Cannot read headers", str(e))
            return
        self.key_list.delete(0, "end")
        for c in ha:
            if c in hb:
                self.key_list.insert("end", c)
        self.status.set(f"{self.key_list.size()} common columns. Pick key column(s) or leave empty.")

    def _start(self):
        pa, pb, out = self.paths[0].get(), self.paths[1].get(), self.out_path.get()
        if not (pa and pb):
            messagebox.showwarning("Missing file", "Please select both files.")
            return
        if not out:
            messagebox.showwarning("Missing output", "Please choose where to save the report.")
            return
        try:
            tol = float(self.tolerance.get() or 0)
        except ValueError:
            messagebox.showwarning("Invalid tolerance", "Number tolerance must be a number.")
            return
        opts = ce.Options(
            key_columns=[self.key_list.get(i) for i in self.key_list.curselection()],
            ignore_case=self.ignore_case.get(),
            trim_spaces=self.trim.get(),
            numeric_tolerance=tol,
            ignore_columns=[c.strip() for c in self.ignore_cols.get().split(",") if c.strip()],
        )
        self.run_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.bar.start(12)
        self._clear_log()
        args = (pa, self.sheets[0].get(), pb, self.sheets[1].get(), opts, out)
        threading.Thread(target=self._worker, args=args, daemon=True).start()

    def _worker(self, pa, sa, pb, sb, opts, out):
        say = lambda m: self.q.put(("log", m))
        try:
            res = ce.compare(pa, sa, pb, sb, opts, progress=say)
            ce.write_report(res, out, progress=say)
            self.q.put(("done", res))
        except PermissionError:
            self.q.put(("error", f"Cannot write '{out}'.\nIs the report open in Excel? Close it and retry."))
        except Exception as e:
            self.q.put(("error", f"{type(e).__name__}: {e}"))

    def _poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self._write_log(payload)
                    self.status.set(payload)
                elif kind == "done":
                    self._finish()
                    for k, v in payload.summary.items():
                        self._write_log(f"  {k}: {v}")
                    msg = "Issues found - see report." if payload.has_issues else "Files match!"
                    self._write_log(msg)
                    self.status.set(f"{msg}  Report: {self.out_path.get()}")
                    self.open_btn.config(state="normal")
                elif kind == "error":
                    self._finish()
                    self._write_log("ERROR: " + payload)
                    self.status.set("Failed.")
                    messagebox.showerror("Comparison failed", payload)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _finish(self):
        self.bar.stop()
        self.run_btn.config(state="normal")

    def _clear_log(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

    def _write_log(self, text):
        self.log.config(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.config(state="disabled")


if __name__ == "__main__":
    App().mainloop()
