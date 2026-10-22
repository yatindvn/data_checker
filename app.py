"""Excel Compare - desktop UI.

Pick two Excel/CSV files, choose sheets (or all sheets), optionally a key column
or a mapping file, and get an Excel report with every issue highlighted in red.

Run:  python app.py
"""
from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import ai_logic
import compare_engine as ce

FILETYPES = [("Excel / CSV", "*.xlsx *.xlsm *.xls *.csv"), ("All files", "*.*")]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Compare")
        self.geometry("820x700")
        self.minsize(720, 600)
        self.q: queue.Queue = queue.Queue()
        self.paths = [tk.StringVar(), tk.StringVar()]
        self.sheets = [tk.StringVar(), tk.StringVar()]
        self.sheet_boxes: list[ttk.Combobox] = []
        self.all_sheets = tk.BooleanVar(value=False)
        self.mapping_path = tk.StringVar()
        self.api_key = tk.StringVar(value=os.environ.get("ANTHROPIC_API_KEY", ""))
        self.out_path = tk.StringVar()
        self.ignore_case = tk.BooleanVar(value=False)
        self.trim = tk.BooleanVar(value=True)
        self.leading_zeros = tk.BooleanVar(value=False)
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
        ttk.Checkbutton(files, text="Compare ALL sheets (matched by sheet name)",
                        variable=self.all_sheets, command=self._toggle_all).grid(
            row=2, column=1, columnspan=4, sticky="w", **pad)
        files.columnconfigure(1, weight=1)

        mp = ttk.LabelFrame(self, text="Mapping file (optional): which column / formula in File 1 "
                                       "matches which in File 2")
        mp.pack(fill="x", **pad)
        ttk.Entry(mp, textvariable=self.mapping_path).grid(row=0, column=0, sticky="ew", **pad)
        ttk.Button(mp, text="Browse...", command=self._browse_mapping).grid(row=0, column=1, **pad)
        ttk.Button(mp, text="Make template", command=self._make_template).grid(row=0, column=2, **pad)
        ttk.Button(mp, text="Clear", command=lambda: self._set_mapping("")).grid(row=0, column=3, **pad)
        ttk.Label(mp, text="Claude API key (only for plain-English logic):").grid(
            row=1, column=0, sticky="e", **pad)
        ttk.Entry(mp, textvariable=self.api_key, show="*", width=30).grid(
            row=1, column=1, columnspan=3, sticky="w", **pad)
        mp.columnconfigure(0, weight=1)

        self.keys_frame = ttk.LabelFrame(self, text="Match rows by key column(s)  "
                                                    "(select none = compare row by row)")
        self.keys_frame.pack(fill="both", expand=True, **pad)
        self.key_list = tk.Listbox(self.keys_frame, selectmode="multiple",
                                   exportselection=False, height=5)
        sb = ttk.Scrollbar(self.keys_frame, command=self.key_list.yview)
        self.key_list.config(yscrollcommand=sb.set)
        self.key_list.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=6)
        sb.pack(side="left", fill="y", pady=6)
        side = ttk.Frame(self.keys_frame)
        side.pack(side="left", fill="y", padx=8)
        ttk.Label(side, text="Ignore columns (comma separated):").pack(anchor="w")
        self.ignore_cols = ttk.Entry(side, width=32)
        self.ignore_cols.pack(anchor="w", pady=(0, 8))
        ttk.Checkbutton(side, text="Ignore upper/lower case", variable=self.ignore_case).pack(anchor="w")
        ttk.Checkbutton(side, text="Trim leading/trailing spaces", variable=self.trim).pack(anchor="w")
        ttk.Checkbutton(side, text="Leading zeros matter (001 ≠ 1)",
                        variable=self.leading_zeros).pack(anchor="w")
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

        self.log = tk.Text(self, height=6, state="disabled", wrap="word")
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

    def _toggle_all(self):
        state = "disabled" if self.all_sheets.get() else "readonly"
        for cb in self.sheet_boxes:
            cb.config(state=state)

    def _browse_mapping(self):
        path = filedialog.askopenfilename(title="Select mapping file", filetypes=FILETYPES)
        if path:
            self._set_mapping(path)

    def _set_mapping(self, path):
        self.mapping_path.set(path)
        on = bool(path)
        self.key_list.config(state="disabled" if on else "normal")
        self.ignore_cols.config(state="disabled" if on else "normal")
        self.keys_frame.config(text="Key columns come from the mapping file (Key = Y)" if on else
                               "Match rows by key column(s)  (select none = compare row by row)")

    def _make_template(self):
        if not all(p.get() for p in self.paths):
            messagebox.showwarning("Missing file", "Select File 1 and File 2 first.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")],
                                            initialfile="mapping.xlsx")
        if not path:
            return
        keys = [self.key_list.get(i) for i in self.key_list.curselection()]
        try:
            ce.write_mapping_template(self.paths[0].get(), self.sheets[0].get(),
                                      self.paths[1].get(), self.sheets[1].get(), path, keys)
        except PermissionError:
            messagebox.showerror("Cannot save", f"Cannot write '{path}'. Is it open in Excel?")
            return
        except Exception as e:
            messagebox.showerror("Cannot make template", str(e))
            return
        self._set_mapping(path)
        self._write_log(f"Mapping template saved: {path}\nEdit it in Excel (see the "
                        "'How to use' sheet), save, then click Compare.")
        if messagebox.askyesno("Template ready", "Open the mapping template in Excel now?"):
            os.startfile(path)

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
        state = self.key_list.cget("state")
        self.key_list.config(state="normal")
        self.key_list.delete(0, "end")
        for c in ha:
            if c in hb:
                self.key_list.insert("end", c)
        self.key_list.config(state=state)
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
        mapping = None
        if self.mapping_path.get():
            try:
                mapping = ce.load_mapping(self.mapping_path.get())
            except Exception as e:
                messagebox.showwarning("Mapping file problem", str(e))
                return
        opts = ce.Options(
            key_columns=[] if mapping else
            [self.key_list.get(i) for i in self.key_list.curselection()],
            ignore_case=self.ignore_case.get(),
            trim_spaces=self.trim.get(),
            numeric_tolerance=tol,
            leading_zeros_matter=self.leading_zeros.get(),
            ignore_columns=[] if mapping else
            [c.strip() for c in self.ignore_cols.get().split(",") if c.strip()],
        )
        ai = ai_logic.Translator(self.api_key.get()) if mapping else None
        self.run_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.bar.start(12)
        self._clear_log()
        args = (pa, self.sheets[0].get(), pb, self.sheets[1].get(), opts, out,
                self.all_sheets.get(), mapping, ai)
        threading.Thread(target=self._worker, args=args, daemon=True).start()

    def _worker(self, pa, sa, pb, sb, opts, out, all_sheets, mapping, ai):
        say = lambda m: self.q.put(("log", m))
        try:
            if all_sheets:
                results = ce.compare_all_sheets(pa, pb, opts, say, mapping, ai)
            else:
                results = [ce.SheetResult(sa, ce.compare(pa, sa, pb, sb, opts, say, mapping, ai))]
            ce.write_report(results, out, progress=say)
            self.q.put(("done", results))
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
                    self._show_results(payload)
                    self.open_btn.config(state="normal")
                elif kind == "error":
                    self._finish()
                    self._write_log("ERROR: " + payload)
                    self.status.set("Failed.")
                    messagebox.showerror("Comparison failed", payload)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _show_results(self, results):
        any_issue = False
        for sr in results:
            if sr.result is None:
                any_issue = True
                self._write_log(f"[{sr.name}] NOT COMPARED: {sr.error}")
                continue
            res = sr.result
            any_issue |= res.has_issues
            if len(results) == 1:
                for k, v in res.summary.items():
                    self._write_log(f"  {k}: {v}")
            else:
                s = res.summary
                self._write_log(f"[{sr.name}] {'ISSUES' if res.has_issues else 'MATCH'}  - "
                                f"cells: {s['Cell differences']}, only F1: {s['Rows only in File 1']}, "
                                f"only F2: {s['Rows only in File 2']}, headers: {s['Header issues']}")
            if res.mapping_used is not None:
                for _, m in res.mapping_used[res.mapping_used["Note"] != ""].iterrows():
                    self._write_log(f"  {m['Output Column']}: {m['Note']} -> "
                                    f"File 1: {m['File 1 used']} | File 2: {m['File 2 used']}")
        msg = "Issues found - see report." if any_issue else "Files match!"
        self._write_log(msg)
        self.status.set(f"{msg}  Report: {self.out_path.get()}")

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
