#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Complete Data Toolkit with Enhanced Excel to CSV Converter
- Tab 1: CSV Merger
- Tab 2: Data Processor  
- Tab 3: Enhanced Excel to CSV Converter (with sheet selection and preview)

Author: Jester Miranda (Enhanced)
"""

import os
import threading
import warnings
import time
from datetime import datetime

import chardet
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext

from openpyxl import load_workbook

# Optional Windows DPI helpers
import platform
import ctypes

# Settings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# =============================================================================
# TAB 1: CSV MERGER
# =============================================================================
class SimpleCSVMerger:
    def __init__(self, parent):
        self.parent = parent
        self.df1 = None
        self.df2 = None
        self.all_cols_f1 = []
        self.all_cols_f2 = []
        self.pull_vars = {}
        self.checkbox_widgets = []
        self.preview_tree = None
        self.preview_vscroll = None
        self.preview_hscroll = None
        self._prog_lock = threading.Lock()
        self.setup_ui()

    def setup_ui(self):
        self.setup_style()
        main = ttk.Frame(self.parent, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        # File Selection
        ttk.Label(main, text="1. Select CSV Files", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W)
        f_frame = ttk.Frame(main)
        f_frame.pack(fill=tk.X, pady=6)

        self.file1_path = ttk.Entry(f_frame, font=('Segoe UI', 9))
        self.file1_path.pack(fill=tk.X, pady=2)
        ttk.Button(f_frame, text="Browse Primary File", command=lambda: self.browse(1)).pack(fill=tk.X, pady=2)

        self.file2_path = ttk.Entry(f_frame, font=('Segoe UI', 9))
        self.file2_path.pack(fill=tk.X, pady=(8, 2))
        ttk.Button(f_frame, text="Browse Source File", command=lambda: self.browse(2)).pack(fill=tk.X, pady=2)

        # Matching with Search
        ttk.Label(main, text="2. Join Keys (Searchable)", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W, pady=(10, 0))
        m_frame = ttk.Frame(main)
        m_frame.pack(fill=tk.X, pady=6)

        self.s1_var = tk.StringVar()
        self.s1_var.trace_add("write", lambda *a: self.filter_key_list(1))
        ttk.Entry(m_frame, textvariable=self.s1_var, font=('Segoe UI', 9, 'italic')).grid(row=0, column=0, sticky="ew")
        self.match_f1 = ttk.Combobox(m_frame, state="readonly", width=40)
        self.match_f1.grid(row=1, column=0, padx=5, sticky="ew")

        ttk.Label(m_frame, text="↔", font=('Segoe UI', 10)).grid(row=1, column=1, padx=6)

        self.s2_var = tk.StringVar()
        self.s2_var.trace_add("write", lambda *a: self.filter_key_list(2))
        ttk.Entry(m_frame, textvariable=self.s2_var, font=('Segoe UI', 9, 'italic')).grid(row=0, column=2, sticky="ew")
        self.match_f2 = ttk.Combobox(m_frame, state="readonly", width=40)
        self.match_f2.grid(row=1, column=2, padx=5, sticky="ew")

        m_frame.columnconfigure(0, weight=1)
        m_frame.columnconfigure(2, weight=1)

        # Columns to Pull
        ttk.Label(main, text="3. Columns to Pull", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W, pady=(10, 0))
        ctrl = ttk.Frame(main)
        ctrl.pack(fill=tk.X, pady=6)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.filter_checkboxes)
        ttk.Entry(ctrl, textvariable=self.search_var, font=('Segoe UI', 9)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        ttk.Button(ctrl, text="All", width=6, command=self.select_all).pack(side=tk.LEFT, padx=4)
        ttk.Button(ctrl, text="None", width=6, command=self.deselect_all).pack(side=tk.LEFT)

        canvas_frame = ttk.Frame(main)
        canvas_frame.pack(fill=tk.BOTH, expand=False, pady=6)
        self.check_canvas = tk.Canvas(canvas_frame, bg="white", height=140, highlightthickness=1)
        self.check_scroll = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.check_canvas.yview)
        self.check_frame = ttk.Frame(self.check_canvas)
        self.check_canvas.create_window((0, 0), window=self.check_frame, anchor="nw")
        self.check_canvas.configure(yscrollcommand=self.check_scroll.set)
        self.check_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.check_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.check_frame.bind("<Configure>", lambda e: self.check_canvas.configure(scrollregion=self.check_canvas.bbox("all")))

        # Actions & Preview
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=10)
        self.prev_btn = ttk.Button(btn_frame, text="PREVIEW DATA", command=self.show_preview, state=tk.DISABLED)
        self.prev_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        self.merge_btn = ttk.Button(btn_frame, text="SAVE MERGED FILE", command=self.process_merge, state=tk.DISABLED)
        self.merge_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

        preview_label = ttk.Label(main, text="Preview (top rows)", font=('Segoe UI', 10, 'bold'))
        preview_label.pack(anchor=tk.W, pady=(6, 0))

        self.preview_frame = ttk.Frame(main, relief=tk.SOLID)
        self.preview_frame.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        # Status & Progress
        self.prog = ttk.Progressbar(main, mode='determinate', maximum=100)
        self.prog.pack(fill=tk.X, pady=6)
        self.stat_var = tk.StringVar(value="Ready")
        self.stat_bar = ttk.Label(main, textvariable=self.stat_var, relief=tk.SUNKEN, anchor=tk.W)
        self.stat_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def setup_style(self):
        style = ttk.Style()
        try:
            if 'vista' in style.theme_names():
                style.theme_use('vista')
            elif 'xpnative' in style.theme_names():
                style.theme_use('xpnative')
            else:
                style.theme_use('clam')
        except Exception:
            style.theme_use('clam')

        default_font = ('Segoe UI', 9)
        style.configure('.', font=default_font)
        style.configure('TButton', padding=(6, 4))
        style.configure('TEntry', padding=(4, 4))
        style.configure('Treeview', font=('Segoe UI', 9), rowheight=22)
        style.configure('Treeview.Heading', font=('Segoe UI', 9, 'bold'))

        try:
            if platform.system() == 'Windows':
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    def detect_enc(self, path):
        try:
            with open(path, 'rb') as f:
                raw = f.read(100000)
                res = chardet.detect(raw)
                return res['encoding'] if res and res.get('encoding') else 'utf-8'
        except Exception:
            return 'utf-8'

    def browse(self, num):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if path:
            target = self.file1_path if num == 1 else self.file2_path
            target.delete(0, tk.END)
            target.insert(0, path)
            self.load_data(num)

    def _set_progress(self, value, text=None):
        def _update():
            try:
                with self._prog_lock:
                    self.prog['value'] = value
                    if text is not None:
                        self.stat_var.set(text)
                    try:
                        self.parent.update_idletasks()
                    except Exception:
                        pass
            except Exception:
                pass

        try:
            self.parent.after(0, _update)
        except Exception:
            _update()

    def load_data(self, num):
        path = self.file1_path.get() if num == 1 else self.file2_path.get()
        if not path:
            return
        self.stat_var.set(f"Loading file {num}...")
        self._set_progress(2, f"Loading file {num}...")

        def run():
            try:
                for p in (5, 10, 18):
                    self._set_progress(p)
                    time.sleep(0.06)

                enc = self.detect_enc(path)
                try:
                    df = pd.read_csv(path, encoding=enc, engine='python')
                except Exception:
                    df = pd.read_csv(path, encoding='latin-1', engine='python')

                df.columns = [str(c).strip() for c in df.columns]
                df = df.loc[:, ~df.columns.duplicated()]

                for p in (30, 45, 60):
                    self._set_progress(p)
                    time.sleep(0.04)

                self.parent.after(0, lambda: self.on_load_success(num, df))
                self._set_progress(100, "File loaded.")
                time.sleep(0.08)
                self._set_progress(0, "Ready")
            except Exception as e:
                self.parent.after(0, lambda: messagebox.showerror("File Error", str(e)))
                self._set_progress(0, "Error")

        threading.Thread(target=run, daemon=True).start()

    def on_load_success(self, num, df):
        if num == 1:
            self.df1 = df
            self.all_cols_f1 = list(df.columns)
            self.filter_key_list(1)
            self.stat_var.set("Primary file loaded.")
        else:
            self.df2 = df
            self.all_cols_f2 = list(df.columns)
            self.filter_key_list(2)
            self.pull_vars = {col: tk.BooleanVar(value=False) for col in self.all_cols_f2}
            self.filter_checkboxes()
            self.stat_var.set("Source file loaded.")

        if self.df1 is not None and self.df2 is not None:
            self.prev_btn.config(state=tk.NORMAL)
            self.merge_btn.config(state=tk.NORMAL)

    def filter_key_list(self, num):
        term = self.s1_var.get().lower() if num == 1 else self.s2_var.get().lower()
        full = self.all_cols_f1 if num == 1 else self.all_cols_f2
        filt = [c for c in full if term in c.lower()]
        target = self.match_f1 if num == 1 else self.match_f2
        target['values'] = filt
        try:
            if filt:
                target.current(0)
            else:
                target.set('')
        except Exception:
            target.set('')

    def filter_checkboxes(self, *args):
        for w in self.checkbox_widgets:
            try:
                w.destroy()
            except Exception:
                pass
        self.checkbox_widgets = []

        term = self.search_var.get().lower()
        if not self.pull_vars:
            return

        for col in self.all_cols_f2:
            if term in col.lower():
                var = self.pull_vars.get(col)
                if var is None:
                    var = tk.BooleanVar(value=False)
                    self.pull_vars[col] = var
                cb = ttk.Checkbutton(self.check_frame, text=col, variable=var)
                cb.pack(fill=tk.X, padx=5, pady=1)
                self.checkbox_widgets.append(cb)

        self.check_canvas.configure(scrollregion=self.check_canvas.bbox("all"))

    def select_all(self):
        for v in self.pull_vars.values():
            v.set(True)

    def deselect_all(self):
        for v in self.pull_vars.values():
            v.set(False)

    def _merge_worker(self, k1, k2, pull, callback):
        try:
            self._set_progress(10, "Preparing data for merge...")
            time.sleep(0.05)

            d1 = self.df1.copy()
            d2_cols = [c for c in ([k2] + pull) if c in self.df2.columns]
            if k2 not in d2_cols:
                d2_cols.insert(0, k2)
            d2 = self.df2[d2_cols].copy()

            self._set_progress(30, "Normalizing keys...")
            time.sleep(0.05)

            d1[k1] = d1[k1].astype(str).str.strip()
            d2[k2] = d2[k2].astype(str).str.strip()

            self._set_progress(55, "Performing merge...")
            res = d1.merge(d2, left_on=k1, right_on=k2, how='left', suffixes=('', '_src'))

            if k1 != k2 and k2 in res.columns:
                try:
                    res.drop(columns=[k2], inplace=True)
                except Exception:
                    pass

            self._set_progress(85, "Finalizing merge...")
            time.sleep(0.05)

            self._set_progress(100, "Merge complete.")
            time.sleep(0.06)
            self._set_progress(0, "Ready")
            callback(res, None)
        except Exception as e:
            callback(None, e)

    def perform_merge(self):
        k1 = self.match_f1.get().strip()
        k2 = self.match_f2.get().strip()
        pull = [c for c, v in self.pull_vars.items() if v.get()]

        if not k1 or not k2 or not pull:
            messagebox.showwarning("Input Missing", "Select both keys and at least one column to pull.")
            return None

        if self.df1 is None or self.df2 is None:
            messagebox.showwarning("Data Missing", "Both files must be loaded before merging.")
            return None

        if k1 not in self.df1.columns:
            messagebox.showerror("Key Error", f"Key '{k1}' not found in primary file.")
            return None
        if k2 not in self.df2.columns:
            messagebox.showerror("Key Error", f"Key '{k2}' not found in source file.")
            return None

        result_container = {'df': None, 'err': None}
        done_event = threading.Event()

        def cb(res, err):
            result_container['df'] = res
            result_container['err'] = err
            done_event.set()

        threading.Thread(target=self._merge_worker, args=(k1, k2, pull, cb), daemon=True).start()

        while not done_event.is_set():
            try:
                self.parent.update()
            except Exception:
                pass
            time.sleep(0.02)

        if result_container['err'] is not None:
            messagebox.showerror("Merge Error", f"Error during merge: {result_container['err']}")
            return None
        return result_container['df']

    def populate_treeview_from_df(self, df, max_rows=200):
        if self.preview_tree:
            try:
                self.preview_tree.destroy()
            except Exception:
                pass
        if self.preview_vscroll:
            try:
                self.preview_vscroll.destroy()
            except Exception:
                pass
        if self.preview_hscroll:
            try:
                self.preview_hscroll.destroy()
            except Exception:
                pass

        cols = list(df.columns)
        if not cols:
            return

        self.preview_tree = ttk.Treeview(self.preview_frame, columns=cols, show='headings')
        self.preview_vscroll = ttk.Scrollbar(self.preview_frame, orient='vertical', command=self.preview_tree.yview)
        self.preview_hscroll = ttk.Scrollbar(self.preview_frame, orient='horizontal', command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=self.preview_vscroll.set, xscrollcommand=self.preview_hscroll.set)

        self.preview_tree.grid(row=0, column=0, sticky='nsew')
        self.preview_vscroll.grid(row=0, column=1, sticky='ns')
        self.preview_hscroll.grid(row=1, column=0, sticky='ew')
        self.preview_frame.columnconfigure(0, weight=1)
        self.preview_frame.rowconfigure(0, weight=1)

        sample = df.head(max_rows)
        for c in cols:
            try:
                max_sample_len = int(sample[c].astype(str).map(len).max()) if not sample.empty else 0
            except Exception:
                max_sample_len = 0
            header_len = len(str(c))
            est = min(max(80, (max(header_len, max_sample_len) * 7)), 400)
            self.preview_tree.heading(c, text=c)
            self.preview_tree.column(c, width=est, anchor='w', stretch=True)

        for idx, row in sample.iterrows():
            values = [self._safe_display_value(row.get(c)) for c in cols]
            self.preview_tree.insert('', 'end', values=values)

        self.preview_tree.bind("<Double-1>", self._on_treeview_double_click)

    def _safe_display_value(self, v):
        if pd.isna(v):
            return ''
        return str(v)

    def _on_treeview_double_click(self, event):
        tree = self.preview_tree
        if tree is None:
            return
        item = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not item or not col:
            return
        try:
            col_index = int(col.replace('#', '')) - 1
            vals = tree.item(item, 'values')
            if col_index < len(vals):
                val = vals[col_index]
                root = self.parent.winfo_toplevel()
                root.clipboard_clear()
                root.clipboard_append(val)
                self.stat_var.set("Copied cell to clipboard.")
        except Exception:
            pass

    def show_preview(self):
        def run_preview():
            self._set_progress(5, "Starting preview...")
            res = self.perform_merge()
            if res is not None:
                preview_rows = 200
                try:
                    self._set_progress(60, "Preparing preview...")
                    time.sleep(0.05)
                    self.parent.after(0, lambda: self.populate_treeview_from_df(res.head(preview_rows), max_rows=preview_rows))
                    self._set_progress(100, f"Previewing top {min(len(res), preview_rows)} rows.")
                    time.sleep(0.06)
                    self._set_progress(0, "Ready")
                except Exception:
                    try:
                        for w in self.preview_frame.winfo_children():
                            w.destroy()
                    except Exception:
                        pass
                    fallback = tk.Text(self.preview_frame, height=12, font=("Courier New", 9), bg="#f8f8f8")
                    fallback.pack(fill=tk.BOTH, expand=True)
                    fallback.delete(1.0, tk.END)
                    fallback.insert(tk.END, repr(res.head(20)))
                    self._set_progress(0, "Preview fallback to text.")
            else:
                self._set_progress(0, "Preview failed.")

        threading.Thread(target=run_preview, daemon=True).start()

    def process_merge(self):
        def run_save():
            self._set_progress(5, "Starting merge and save...")
            res = self.perform_merge()
            if res is not None:
                path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="merged_output.csv",
                                                    filetypes=[("CSV files", "*.csv")])
                if path:
                    try:
                        self._set_progress(60, "Writing CSV...")
                        try:
                            res.to_csv(path, index=False, encoding='utf-8-sig')
                        except Exception:
                            res.to_csv(path, index=False)
                        self.parent.after(0, lambda: messagebox.showinfo("Success", f"File saved to:\n{path}"))
                        self._set_progress(100, "Saved successfully.")
                        time.sleep(0.06)
                        self._set_progress(0, "Ready")
                    except Exception as e:
                        self.parent.after(0, lambda: messagebox.showerror("Save Error", f"Failed to save file: {e}"))
                        self._set_progress(0, "Save failed.")
                else:
                    self._set_progress(0, "Save cancelled.")
            else:
                self._set_progress(0, "Merge failed; nothing saved.")

        threading.Thread(target=run_save, daemon=True).start()


# =============================================================================
# TAB 2: DATA PROCESSOR
# =============================================================================
class DataProcessorGUI:
    def __init__(self, parent):
        self.parent = parent
        self.df = None
        self.file_path = None
        self.map_full_df = None
        self.map_df = None
        self._prog_lock = threading.Lock()
        self.setup_ui()
        self.log("System Ready. Please load MAP.csv if not already present.")
        self.load_map_silent()

    def setup_ui(self):
        main_frame = ttk.Frame(self.parent, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="1. Load Data", font=('Helvetica', 10, 'bold')).pack(anchor=tk.W)
        file_btn = ttk.Button(main_frame, text="Select Data File (CSV/XLSX/JSON)", command=self.load_file)
        file_btn.pack(fill=tk.X, pady=5)

        self.file_label = ttk.Label(main_frame, text="No file selected", foreground="gray")
        self.file_label.pack(anchor=tk.W, pady=(0, 15))

        ttk.Label(main_frame, text="2. Data Operations", font=('Helvetica', 10, 'bold')).pack(anchor=tk.W)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)

        ttk.Button(btn_frame, text="Remove BSG", command=self.remove_bsg).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(btn_frame, text="Filter ACT", command=self.filter_act).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(btn_frame, text="Full Process", command=self.run_full_process).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

        ttk.Label(main_frame, text="3. Progress & Activity Log", font=('Helvetica', 10, 'bold')).pack(anchor=tk.W, pady=(15, 0))

        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, mode='determinate', maximum=100)
        self.progress.pack(fill=tk.X, pady=10)

        self.log_area = scrolledtext.ScrolledText(main_frame, height=12, font=('Consolas', 9), bg="#f8f9fa")
        self.log_area.pack(fill=tk.BOTH, expand=True)

    def _set_progress(self, value, text=None):
        def _update():
            try:
                with self._prog_lock:
                    self.progress['value'] = value
                    if text:
                        self.log(text)
                    try:
                        self.parent.update_idletasks()
                    except Exception:
                        pass
            except Exception:
                pass

        try:
            self.parent.after(0, _update)
        except Exception:
            _update()

    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_area.see(tk.END)

    def load_map_silent(self):
        map_path = 'MAP.csv'
        if os.path.exists(map_path):
            try:
                self.map_full_df = pd.read_csv(map_path)
                if self.map_full_df.shape[1] >= 3:
                    self.map_df = self.map_full_df.iloc[:, [0, 2]].copy()
                    self.map_df.columns = ['PROVINCENAME', 'REGION']
                    self.map_df.drop_duplicates(subset=['PROVINCENAME'], inplace=True)
                    self.log("✅ MAP.csv loaded and indexed.")
                else:
                    self.log("⚠️ MAP.csv found but doesn't have expected columns.")
            except Exception as e:
                self.log(f"❌ Error loading MAP.csv: {e}")
        else:
            self.log("⚠️ MAP.csv not found in folder. Some features will be disabled.")

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Data Files", "*.csv *.xlsx *.json *.jsonl")])
        if path:
            try:
                self.log(f"Opening: {os.path.basename(path)}...")
                ext = os.path.splitext(path)[1].lower()

                if ext == '.csv':
                    try:
                        self.df = pd.read_csv(path, encoding='utf-8')
                    except Exception:
                        self.log("UTF-8 failed, trying latin1...")
                        self.df = pd.read_csv(path, encoding='latin1')
                elif ext == '.xlsx':
                    self.df = pd.read_excel(path)
                else:
                    self.df = pd.read_json(path, lines=ext == '.jsonl')

                self.file_path = path
                self.file_label.config(text=f"Loaded: {os.path.basename(path)}", foreground="#007bff")
                self.log(f"✅ Success: Loaded {len(self.df)} rows.")
            except Exception as e:
                self.log(f"❌ Load Error: {e}")
                messagebox.showerror("Error", f"Failed to load: {e}")

    def save_df(self, suffix):
        if not self.file_path or self.df is None:
            self.log("❌ Save Error: No file loaded or no data to save.")
            return None

        output_folder = 'data'
        os.makedirs(output_folder, exist_ok=True)
        base = os.path.splitext(os.path.basename(self.file_path))[0]
        ext = os.path.splitext(self.file_path)[1].lower()
        out_path = os.path.join(output_folder, f"{base}_{suffix}{ext}")

        try:
            if ext == '.csv':
                self.df.to_csv(out_path, index=False)
            elif ext == '.xlsx':
                self.df.to_excel(out_path, index=False)
            else:
                self.df.to_json(out_path, orient='records', indent=4)
            self.log(f"💾 File Saved: {out_path}")
            return out_path
        except Exception as e:
            self.log(f"❌ Save Error: {e}")
            return None

    def remove_bsg(self):
        if self.df is not None and 'DIVISIONCODE' in self.df.columns:
            initial = len(self.df)
            self.df = self.df[self.df['DIVISIONCODE'] != 'BSG']
            removed = initial - len(self.df)
            self.log(f"Removed {removed} 'BSG' rows.")
            self.save_df("removedbsg")
        else:
            messagebox.showwarning("Warning", "Data or 'DIVISIONCODE' column missing.")

    def filter_act(self):
        if self.df is not None and 'SUBSCRIBERSTATUSCODE' in self.df.columns:
            self.df = self.df[self.df['SUBSCRIBERSTATUSCODE'].astype(str).str.contains('ACT', na=False)]
            self.log(f"Filtered to {len(self.df)} 'ACT' rows.")
            self.save_df("actfiltered")
        else:
            messagebox.showwarning("Warning", "Data or 'SUBSCRIBERSTATUSCODE' column missing.")

    def _full_process_worker(self):
        try:
            self._set_progress(2, "Starting Full Process...")
            time.sleep(0.05)

            today = pd.to_datetime('today').normalize()
            yesterday = today - pd.Timedelta(days=1)

            dy_yesterday = f"{yesterday.strftime('%b%d').upper()} (STATUS)"
            dy_today = f"{today.strftime('%b%d').upper()} (STATUS)"
            dy_hours = f"AGED (HOURS) - {today.strftime('%b%d').upper()}"
            dy_bucket = f"AGED BUCKET - {today.strftime('%b%d').upper()}"
            dy_group = f"AGED BUCKET GROUP - {today.strftime('%b%d').upper()}"

            new_headers = ['ALIGNED ACCT', 'ALIGNED JONO', 'ACCT+JONO', 'SEGMENT', 'PRODUCT',
                           'JOCRYEAR', 'DATE TODAY', 'JOTODAY', 'AGEING', 'AGEING (2)',
                           'AREA', 'MSP', dy_yesterday, dy_today, 'JIRA TICKET STATUS',
                           'ACTION TAKEN', 'FINAL STATUS', dy_hours, dy_bucket, dy_group]

            self._set_progress(8, "Ensuring headers...")
            for col in new_headers:
                if col not in self.df.columns:
                    self.df[col] = None
            time.sleep(0.04)

            self._set_progress(18, "Alignment & Date Calculations...")
            if 'ACCTNO' in self.df.columns:
                self.df['ALIGNED ACCT'] = self.df['ACCTNO'].astype(str).str.strip().str.zfill(13)
            else:
                self.df['ALIGNED ACCT'] = None

            if 'JONO' in self.df.columns:
                self.df['ALIGNED JONO'] = self.df['JONO'].astype(str).str.strip().str.zfill(8)
            else:
                self.df['ALIGNED JONO'] = None

            self.df['ACCT+JONO'] = self.df.apply(
                lambda r: f"{r['ALIGNED ACCT']}:{r['ALIGNED JONO']}" if r.get('ALIGNED JONO') else None, axis=1
            )

            if 'DATEJOCREATED' in self.df.columns:
                self.df['DATEJOCREATED'] = pd.to_datetime(self.df['DATEJOCREATED'], errors='coerce')
                self.df['JOCRYEAR'] = self.df['DATEJOCREATED'].dt.year
                self.df['DATE TODAY'] = today
                if 'DATEJOCLOSED' in self.df.columns:
                    self.df['DATEJOCLOSED'] = pd.to_datetime(self.df['DATEJOCLOSED'], errors='coerce')
                    end_dates = self.df['DATEJOCLOSED']
                else:
                    end_dates = today
                self.df['JOTODAY'] = (end_dates - self.df['DATEJOCREATED']).dt.days
            else:
                self.df['DATEJOCREATED'] = pd.NaT
                self.df['JOCRYEAR'] = None
                self.df['DATE TODAY'] = today
                self.df['JOTODAY'] = None

            self._set_progress(35, "Calculating segment & product...")
            if 'PACKAGENAME' in self.df.columns and 'PROVINCENAME' in self.df.columns:
                conditions = [
                    self.df['PACKAGENAME'].astype(str).str.contains('BIDA', case=False, na=False),
                    self.df['PACKAGENAME'].astype(str).str.contains('S2S', case=False, na=False),
                    self.df['PACKAGENAME'].astype(str).str.contains('SKY', case=False, na=False) & self.df['PROVINCENAME'].astype(str).str.contains('METRO MANILA', case=False, na=False),
                    self.df['PACKAGENAME'].astype(str).str.contains('SKY', case=False, na=False),
                    self.df['PACKAGENAME'].astype(str).str.contains('Biz', case=False, na=False),
                    self.df['PACKAGENAME'].astype(str).str.contains('Streamtech', case=False, na=False),
                    self.df['PACKAGENAME'].astype(str).str.contains(r'(AIR INTERNET|AIRONFIBER|FIBER X|FIBERX|GAME CHANGER|GAMECHANGER|HOME BASE|HOMEBASE|HYPERWIRE|BSS)', case=False, na=False)
                ]

                segment_choices = ["BIDA", "S2S", "SKYNCR", "SKY REGIONAL", "SME", "Streamtech", "RES"]
                self.df['SEGMENT'] = np.select(conditions, segment_choices, default=None)

                product_choices = ["BIDA", "S2S", "SKY", "SKY", "FIBER", "Streamtech", "FIBER"]
                self.df['PRODUCT'] = np.select(conditions, product_choices, default=None)

                self.df.loc[self.df['PACKAGENAME'].isna(), ['SEGMENT', 'PRODUCT']] = None
            else:
                self.log("⚠️ Warning: Missing PACKAGENAME/PROVINCENAME. Skipping Segment/Product logic.")

            self._set_progress(50, "Applying ageing buckets...")
            time.sleep(0.04)

            if 'JOTODAY' in self.df.columns and self.df['JOTODAY'].notna().any():
                self.df['JOTODAY'] = pd.to_numeric(self.df['JOTODAY'], errors='coerce')
                self.df['AGEING'] = np.select(
                    [self.df['JOTODAY'] <= 1, self.df['JOTODAY'] <= 3, self.df['JOTODAY'] <= 5, self.df['JOTODAY'] <= 15, self.df['JOTODAY'] <= 30, self.df['JOTODAY'] <= 60, self.df['JOTODAY'] > 60],
                    ["0-1 D", "2-3 D", "3-5 D", "5-15 D", "15-30 D", "30-60 D", "> 60 D"], default=None)
                self.df['AGEING (2)'] = np.select(
                    [self.df['JOTODAY'] <= 5, self.df['JOTODAY'] <= 15, self.df['JOTODAY'] <= 30, self.df['JOTODAY'] <= 60, self.df['JOTODAY'] > 60],
                    ["0-5 D", "5-15 D", "15-30 D", "30-60 D", "> 60 D"], default=None)
                self.df[dy_hours] = self.df['JOTODAY'] * 24
            else:
                self.df['AGEING'] = None
                self.df['AGEING (2)'] = None
                self.df[dy_hours] = None

            self.df[dy_bucket] = today.strftime('%d-%b')
            self.df[dy_group] = self.df['AGEING (2)']

            self._set_progress(65, "Mapping area...")
            time.sleep(0.04)

            if self.map_df is not None and 'PROVINCENAME' in self.df.columns:
                area_dict = self.map_df.set_index('PROVINCENAME')['REGION'].to_dict()
                self.df['AREA'] = self.df['PROVINCENAME'].astype(str).map(lambda x: area_dict.get(x, None))
            else:
                self.log("Area mapping skipped (MAP.csv missing or PROVINCENAME not in data).")

            self._set_progress(75, "Starting MSP lookups...")
            time.sleep(0.04)

            if self.map_full_df is not None and 'PROVINCENAME' in self.df.columns:
                self.log("Starting Complex MSP Lookups...")
                self.df['MSP'] = None
                p_mask = self.df['PROVINCENAME'].notna() & (self.df['PROVINCENAME'].astype(str).str.strip() != '')

                try:
                    if 'BARANGAYNAME' in self.df.columns and self.map_full_df.shape[1] > 8:
                        cond1 = (self.df['BARANGAYNAME'].astype(str).str.lower() == 'holy spirit') & \
                                (self.df['PROVINCENAME'].astype(str).str.lower().str.contains('metro manila', na=False))
                        map_f_i = self.map_full_df.iloc[:, [5, 8]].dropna().copy()
                        map_f_i.columns = ['k', 'v']
                        map_f_i['k'] = map_f_i['k'].astype(str).str.lower().str.strip()
                        h_dict = dict(zip(map_f_i['k'], map_f_i['v']))
                        self.df.loc[p_mask & cond1, 'MSP'] = self.df.loc[p_mask & cond1, 'BARANGAYNAME'].astype(str).str.lower().map(h_dict)
                except Exception as e:
                    self.log(f"MSP step 1 error: {e}")

                try:
                    m_mask = p_mask & self.df['MSP'].isna()
                    if 'MUNICIPALITYNAME' in self.df.columns and self.map_full_df.shape[1] > 8:
                        df_key = (self.df['PROVINCENAME'].astype(str).str.lower().str.strip() + '|' +
                                  self.df['MUNICIPALITYNAME'].astype(str).str.lower().str.strip())
                        map_e_g_i = self.map_full_df.iloc[:, [4, 6, 8]].dropna().copy()
                        map_e_g_i['key'] = map_e_g_i.iloc[:, 0].astype(str).str.lower().str.strip() + '|' + map_e_g_i.iloc[:, 1].astype(str).str.lower().str.strip()
                        m_dict = dict(zip(map_e_g_i['key'], map_e_g_i.iloc[:, 2]))
                        self.df.loc[m_mask, 'MSP'] = df_key.loc[m_mask].map(m_dict)
                except Exception as e:
                    self.log(f"MSP step 2 error: {e}")

                try:
                    m_mask = p_mask & self.df['MSP'].isna()
                    map_e_i = self.map_full_df.iloc[:, [4, 8]].dropna().copy()
                    map_e_i.columns = ['k', 'v']
                    map_e_i['k'] = map_e_i['k'].astype(str).str.lower().str.strip()
                    p_dict = dict(zip(map_e_i['k'], map_e_i['v']))
                    self.df.loc[m_mask, 'MSP'] = self.df.loc[m_mask, 'PROVINCENAME'].astype(str).str.lower().str.strip().map(p_dict)
                except Exception as e:
                    self.log(f"MSP step 3 error: {e}")
            else:
                self.log("MSP mapping skipped (MAP.csv missing or PROVINCENAME not in data).")

            if 'MSP' in self.df.columns:
                self.df['MSP'] = self.df['MSP'].fillna('')

            self._set_progress(90, "Saving processed file...")
            time.sleep(0.04)
            self.save_df("processed")
            self._set_progress(100, "ALL CALCULATIONS COMPLETE.")
            self.log("✅ ALL CALCULATIONS COMPLETE.")
            time.sleep(0.06)
            self._set_progress(0, "Ready")
            try:
                self.parent.after(0, lambda: messagebox.showinfo("Done", "Processing successful!"))
            except Exception:
                pass

        except Exception as e:
            self.log(f"❌ Error: {e}")
            try:
                self.parent.after(0, lambda: messagebox.showerror("Error", f"Processing failed: {e}"))
            except Exception:
                pass
            self._set_progress(0, "Error")

    def run_full_process(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Load data first!")
            return
        threading.Thread(target=self._full_process_worker, daemon=True).start()


# =============================================================================
# TAB 3: ENHANCED EXCEL TO CSV CONVERTER
# =============================================================================
class EnhancedExcelToCsvConverter:
    def __init__(self, parent):
        self.parent = parent
        self.file_path = None
        self.output_folder = None
        self.workbook = None
        self.sheet_vars = {}
        self.preview_tree = None
        self.preview_vscroll = None
        self.preview_hscroll = None
        self._prog_lock = threading.Lock()
        self.setup_ui()

    def setup_ui(self):
        self.setup_style()
        main = ttk.Frame(self.parent, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        # File Selection
        ttk.Label(main, text="1. Select Excel File", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W)
        file_frame = ttk.Frame(main)
        file_frame.pack(fill=tk.X, pady=6)
        
        self.entry_file = ttk.Entry(file_frame, font=('Segoe UI', 9))
        self.entry_file.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(file_frame, text="Browse Excel File", command=self.select_file).pack(side=tk.LEFT)

        # Sheet Selection
        ttk.Label(main, text="2. Select Sheets to Convert", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W, pady=(10, 0))
        
        ctrl_frame = ttk.Frame(main)
        ctrl_frame.pack(fill=tk.X, pady=6)
        ttk.Button(ctrl_frame, text="Select All", width=12, command=self.select_all_sheets).pack(side=tk.LEFT, padx=2)
        ttk.Button(ctrl_frame, text="Deselect All", width=12, command=self.deselect_all_sheets).pack(side=tk.LEFT, padx=2)
        ttk.Button(ctrl_frame, text="Preview Selected", command=self.preview_selected_sheet).pack(side=tk.LEFT, padx=10)
        
        sheet_canvas_frame = ttk.Frame(main)
        sheet_canvas_frame.pack(fill=tk.X, pady=6)
        
        self.sheet_canvas = tk.Canvas(sheet_canvas_frame, bg="white", height=120, highlightthickness=1, highlightbackground="#ccc")
        sheet_scroll = ttk.Scrollbar(sheet_canvas_frame, orient="vertical", command=self.sheet_canvas.yview)
        self.sheet_frame = ttk.Frame(self.sheet_canvas)
        
        self.sheet_canvas.create_window((0, 0), window=self.sheet_frame, anchor="nw")
        self.sheet_canvas.configure(yscrollcommand=sheet_scroll.set)
        
        self.sheet_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sheet_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.sheet_frame.bind("<Configure>", lambda e: self.sheet_canvas.configure(scrollregion=self.sheet_canvas.bbox("all")))

        # Preview Area
        ttk.Label(main, text="3. Preview (Double-click cell to copy)", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W, pady=(10, 0))
        
        self.preview_frame = ttk.Frame(main, relief=tk.SOLID, borderwidth=1)
        self.preview_frame.pack(fill=tk.BOTH, expand=True, pady=6)
        
        self.preview_label = ttk.Label(self.preview_frame, text="No preview available. Load an Excel file and select a sheet to preview.", 
                                      foreground="gray", font=('Segoe UI', 9, 'italic'))
        self.preview_label.pack(expand=True)

        # Output Options
        ttk.Label(main, text="4. Output Options", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W, pady=(10, 0))
        
        output_frame = ttk.Frame(main)
        output_frame.pack(fill=tk.X, pady=6)
        
        self.entry_output = ttk.Entry(output_frame, font=('Segoe UI', 9))
        self.entry_output.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(output_frame, text="Browse Output Folder", command=self.select_output).pack(side=tk.LEFT)

        option_frame = ttk.Frame(main)
        option_frame.pack(fill=tk.X, pady=6)
        
        self.combine_sheets_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(option_frame, text="Combine all selected sheets into one CSV", 
                       variable=self.combine_sheets_var).pack(side=tk.LEFT, padx=5)

        # Convert Button
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=10)
        
        self.convert_btn = ttk.Button(btn_frame, text="CONVERT TO CSV", command=self.convert, state=tk.DISABLED)
        self.convert_btn.pack(fill=tk.X, padx=4)

        # Progress & Status
        self.progress = ttk.Progressbar(main, mode='determinate', maximum=100)
        self.progress.pack(fill=tk.X, pady=6)
        
        self.status_var = tk.StringVar(value="Ready. Please select an Excel file.")
        self.status_bar = ttk.Label(main, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def setup_style(self):
        style = ttk.Style()
        try:
            if 'vista' in style.theme_names():
                style.theme_use('vista')
            elif 'xpnative' in style.theme_names():
                style.theme_use('xpnative')
            else:
                style.theme_use('clam')
        except Exception:
            style.theme_use('clam')

        default_font = ('Segoe UI', 9)
        style.configure('.', font=default_font)
        style.configure('TButton', padding=(6, 4))
        style.configure('TEntry', padding=(4, 4))
        style.configure('Treeview', font=('Segoe UI', 9), rowheight=22)
        style.configure('Treeview.Heading', font=('Segoe UI', 9, 'bold'))

    def _set_progress(self, value, text=None):
        def _update():
            try:
                with self._prog_lock:
                    self.progress['value'] = value
                    if text is not None:
                        self.status_var.set(text)
                    try:
                        self.parent.update_idletasks()
                    except Exception:
                        pass
            except Exception:
                pass

        try:
            self.parent.after(0, _update)
        except Exception:
            _update()

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self.file_path = path
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, path)
            self.load_workbook_sheets()

    def load_workbook_sheets(self):
        if not self.file_path:
            return
            
        self._set_progress(10, "Loading Excel file...")
        
        def _load():
            try:
                self.workbook = load_workbook(self.file_path, read_only=True, data_only=True)
                sheets = self.workbook.sheetnames
                
                self._set_progress(50, "Loading sheets...")
                
                self.parent.after(0, lambda: self._populate_sheet_list(sheets))
                
                self._set_progress(100, f"Loaded {len(sheets)} sheet(s) successfully.")
                self.parent.after(200, lambda: self._set_progress(0, "Ready"))
                
            except Exception as e:
                self.parent.after(0, lambda: messagebox.showerror("Error", f"Failed to load Excel file:\n{e}"))
                self._set_progress(0, "Error loading file.")
        
        threading.Thread(target=_load, daemon=True).start()

    def _populate_sheet_list(self, sheets):
        for widget in self.sheet_frame.winfo_children():
            widget.destroy()
        
        self.sheet_vars.clear()
        
        if not sheets:
            ttk.Label(self.sheet_frame, text="No sheets found in workbook", 
                     foreground="gray", font=('Segoe UI', 9, 'italic')).pack(pady=10)
            return
        
        for sheet_name in sheets:
            var = tk.BooleanVar(value=True)
            self.sheet_vars[sheet_name] = var
            
            frame = ttk.Frame(self.sheet_frame)
            frame.pack(fill=tk.X, padx=5, pady=2)
            
            cb = ttk.Checkbutton(frame, text=sheet_name, variable=var)
            cb.pack(side=tk.LEFT)
            
            preview_btn = ttk.Button(frame, text="Preview", width=8,
                                    command=lambda s=sheet_name: self.preview_sheet(s))
            preview_btn.pack(side=tk.RIGHT, padx=5)
        
        self.sheet_canvas.configure(scrollregion=self.sheet_canvas.bbox("all"))
        self.convert_btn.config(state=tk.NORMAL)

    def select_all_sheets(self):
        for var in self.sheet_vars.values():
            var.set(True)

    def deselect_all_sheets(self):
        for var in self.sheet_vars.values():
            var.set(False)

    def preview_selected_sheet(self):
        selected = [name for name, var in self.sheet_vars.items() if var.get()]
        if not selected:
            messagebox.showwarning("No Selection", "Please select at least one sheet to preview.")
            return
        self.preview_sheet(selected[0])

    def preview_sheet(self, sheet_name):
        if not self.workbook:
            messagebox.showwarning("No File", "Please load an Excel file first.")
            return
        
        self._set_progress(10, f"Loading preview of '{sheet_name}'...")
        
        def _load_preview():
            try:
                ws = self.workbook[sheet_name]
                rows = list(ws.values)
                
                if not rows:
                    self.parent.after(0, lambda: messagebox.showinfo("Empty Sheet", f"Sheet '{sheet_name}' is empty."))
                    self._set_progress(0, "Ready")
                    return
                
                self._set_progress(40, "Processing data...")
                
                header = [str(c) if c is not None else f"Column_{i}" for i, c in enumerate(rows[0])]
                data_rows = rows[1:] if len(rows) > 1 else []
                
                preview_rows = data_rows[:500]
                df = pd.DataFrame(preview_rows, columns=header)
                
                self._set_progress(70, "Rendering preview...")
                
                self.parent.after(0, lambda: self._populate_preview(df, sheet_name))
                
                self._set_progress(100, f"Preview loaded: {len(df)} rows")
                self.parent.after(200, lambda: self._set_progress(0, "Ready"))
                
            except Exception as e:
                self.parent.after(0, lambda: messagebox.showerror("Preview Error", f"Failed to preview sheet:\n{e}"))
                self._set_progress(0, "Preview failed.")
        
        threading.Thread(target=_load_preview, daemon=True).start()

    def _populate_preview(self, df, sheet_name):
        if self.preview_tree:
            self.preview_tree.destroy()
        if self.preview_vscroll:
            self.preview_vscroll.destroy()
        if self.preview_hscroll:
            self.preview_hscroll.destroy()
        if self.preview_label:
            self.preview_label.destroy()
            self.preview_label = None

        if df.empty:
            self.preview_label = ttk.Label(self.preview_frame, 
                                          text=f"Sheet '{sheet_name}' has no data to preview.",
                                          foreground="gray", font=('Segoe UI', 9, 'italic'))
            self.preview_label.pack(expand=True)
            return

        cols = list(df.columns)
        self.preview_tree = ttk.Treeview(self.preview_frame, columns=cols, show='headings', height=15)
        self.preview_vscroll = ttk.Scrollbar(self.preview_frame, orient='vertical', command=self.preview_tree.yview)
        self.preview_hscroll = ttk.Scrollbar(self.preview_frame, orient='horizontal', command=self.preview_tree.xview)
        
        self.preview_tree.configure(yscrollcommand=self.preview_vscroll.set, 
                                   xscrollcommand=self.preview_hscroll.set)

        self.preview_tree.grid(row=0, column=0, sticky='nsew')
        self.preview_vscroll.grid(row=0, column=1, sticky='ns')
        self.preview_hscroll.grid(row=1, column=0, sticky='ew')
        
        self.preview_frame.columnconfigure(0, weight=1)
        self.preview_frame.rowconfigure(0, weight=1)

        for col in cols:
            try:
                max_len = df[col].astype(str).str.len().max()
            except:
                max_len = 0
            
            header_len = len(str(col))
            est_width = min(max(100, max(header_len, max_len) * 8), 300)
            
            self.preview_tree.heading(col, text=col)
            self.preview_tree.column(col, width=est_width, anchor='w', stretch=True)

        for idx, row in df.iterrows():
            values = [self._safe_str(row.get(c)) for c in cols]
            self.preview_tree.insert('', 'end', values=values)

        self.preview_tree.bind("<Double-1>", self._on_cell_double_click)
        
        total_rows = len(df)
        self.status_var.set(f"Previewing sheet '{sheet_name}' - Showing {total_rows} rows")

    def _safe_str(self, value):
        if pd.isna(value) or value is None:
            return ''
        return str(value)

    def _on_cell_double_click(self, event):
        tree = self.preview_tree
        if not tree:
            return
        
        item = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        
        if not item or not col:
            return
        
        try:
            col_index = int(col.replace('#', '')) - 1
            values = tree.item(item, 'values')
            
            if col_index < len(values):
                value = values[col_index]
                root = self.parent.winfo_toplevel()
                root.clipboard_clear()
                root.clipboard_append(value)
                self.status_var.set(f"Copied to clipboard: {value[:50]}...")
        except Exception:
            pass

    def select_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder = folder
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, folder)

    def convert(self):
        if not self.file_path or not self.workbook:
            messagebox.showwarning("No File", "Please select an Excel file first.")
            return

        selected_sheets = [name for name, var in self.sheet_vars.items() if var.get()]
        
        if not selected_sheets:
            messagebox.showwarning("No Selection", "Please select at least one sheet to convert.")
            return

        out_folder = self.output_folder or os.path.join(os.path.dirname(self.file_path), "converted_csvs")
        os.makedirs(out_folder, exist_ok=True)

        combine = self.combine_sheets_var.get()
        
        def _convert_worker():
            try:
                exported = []
                total = len(selected_sheets)
                
                if combine:
                    self._set_progress(5, "Combining sheets...")
                    combined = []
                    
                    for idx, sheet_name in enumerate(selected_sheets):
                        progress = int(10 + (idx / total) * 70)
                        self._set_progress(progress, f"Reading sheet {idx + 1}/{total}: {sheet_name}")
                        
                        ws = self.workbook[sheet_name]
                        rows = list(ws.values)
                        
                        if not rows:
                            continue
                        
                        header = [str(c) if c is not None else f"Column_{i}" for i, c in enumerate(rows[0])]
                        data_rows = rows[1:] if len(rows) > 1 else []
                        
                        df = pd.DataFrame(data_rows, columns=header)
                        df['__SheetName__'] = sheet_name
                        combined.append(df)
                    
                    if combined:
                        self._set_progress(85, "Merging data...")
                        final_df = pd.concat(combined, ignore_index=True)
                        
                        base_name = os.path.splitext(os.path.basename(self.file_path))[0]
                        out_path = os.path.join(out_folder, f"{base_name}_combined.csv")
                        
                        self._set_progress(95, "Writing CSV file...")
                        final_df.to_csv(out_path, index=False, encoding='utf-8-sig')
                        exported.append(out_path)
                
                else:
                    for idx, sheet_name in enumerate(selected_sheets):
                        progress = int(10 + (idx / total) * 85)
                        self._set_progress(progress, f"Converting {idx + 1}/{total}: {sheet_name}")
                        
                        ws = self.workbook[sheet_name]
                        rows = list(ws.values)
                        
                        if not rows:
                            continue
                        
                        header = [str(c) if c is not None else f"Column_{i}" for i, c in enumerate(rows[0])]
                        data_rows = rows[1:] if len(rows) > 1 else []
                        
                        df = pd.DataFrame(data_rows, columns=header)
                        
                        safe_name = "".join(c if c.isalnum() or c in (' ', '_', '-') else '_' for c in sheet_name)
                        base_name = os.path.splitext(os.path.basename(self.file_path))[0]
                        out_path = os.path.join(out_folder, f"{base_name}_{safe_name}.csv")
                        
                        df.to_csv(out_path, index=False, encoding='utf-8-sig')
                        exported.append(out_path)
                
                self._set_progress(100, f"Conversion complete: {len(exported)} file(s) created")
                
                msg = f"Successfully converted {len(exported)} file(s) to:\n{out_folder}"
                self.parent.after(0, lambda: messagebox.showinfo("Success", msg))
                
                self.parent.after(500, lambda: self._set_progress(0, "Ready"))
                
            except Exception as e:
                error_msg = f"Conversion failed:\n{e}"
                self.parent.after(0, lambda: messagebox.showerror("Error", error_msg))
                self._set_progress(0, "Conversion failed.")
        
        threading.Thread(target=_convert_worker, daemon=True).start()


# =============================================================================
# MAIN APPLICATION
# =============================================================================
def main():
    root = tk.Tk()
    root.title("Data Toolkit - Jester Miranda (Enhanced)")
    root.geometry("1100x750")

    nb = ttk.Notebook(root)
    nb.pack(fill=tk.BOTH, expand=True)

    tab1 = ttk.Frame(nb)
    tab2 = ttk.Frame(nb)
    tab3 = ttk.Frame(nb)

    nb.add(tab1, text="CSV Merger")
    nb.add(tab2, text="Data Processor")
    nb.add(tab3, text="Excel → CSV")

    SimpleCSVMerger(tab1)
    DataProcessorGUI(tab2)
    EnhancedExcelToCsvConverter(tab3)

    root.mainloop()


if __name__ == "__main__":
    main()
