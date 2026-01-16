#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Updated CSV merger, data processor, and Excel-to-CSV converter with Windows 10/11 styling,
an Excel-like scrollable preview using ttk.Treeview, and live progress updates.

Drop this file into your project and run with Python 3.8+.
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
# TAB 1: CSV MERGER (Updated with Windows styling, Treeview preview, and live progress)
# =============================================================================
class SimpleCSVMerger:
    def __init__(self, parent):
        self.parent = parent
        # Dataframes and metadata
        self.df1 = None
        self.df2 = None
        self.all_cols_f1 = []
        self.all_cols_f2 = []
        self.pull_vars = {}
        self.checkbox_widgets = []

        # Preview widgets
        self.preview_tree = None
        self.preview_vscroll = None
        self.preview_hscroll = None

        # Progress control lock
        self._prog_lock = threading.Lock()

        self.setup_ui()

    # -------------------------
    # UI and styling
    # -------------------------
    def setup_ui(self):
        # Apply Windows-like style and fonts
        self.setup_style()

        main = ttk.Frame(self.parent, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        # --- 1. File Selection ---
        ttk.Label(main, text="1. Select CSV Files", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W)
        f_frame = ttk.Frame(main)
        f_frame.pack(fill=tk.X, pady=6)

        self.file1_path = ttk.Entry(f_frame, font=('Segoe UI', 9))
        self.file1_path.pack(fill=tk.X, pady=2)
        ttk.Button(f_frame, text="Browse Primary File", command=lambda: self.browse(1)).pack(fill=tk.X, pady=2)

        self.file2_path = ttk.Entry(f_frame, font=('Segoe UI', 9))
        self.file2_path.pack(fill=tk.X, pady=(8, 2))
        ttk.Button(f_frame, text="Browse Source File", command=lambda: self.browse(2)).pack(fill=tk.X, pady=2)

        # --- 2. Matching with Search ---
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

        # --- 3. Columns to Pull ---
        ttk.Label(main, text="3. Columns to Pull", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W, pady=(10, 0))
        ctrl = ttk.Frame(main)
        ctrl.pack(fill=tk.X, pady=6)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.filter_checkboxes)
        ttk.Entry(ctrl, textvariable=self.search_var, font=('Segoe UI', 9)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        ttk.Button(ctrl, text="All", width=6, command=self.select_all).pack(side=tk.LEFT, padx=4)
        ttk.Button(ctrl, text="None", width=6, command=self.deselect_all).pack(side=tk.LEFT)

        # Scrollable checkbox area
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

        # --- 4. Actions & Preview ---
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=10)
        self.prev_btn = ttk.Button(btn_frame, text="PREVIEW DATA", command=self.show_preview, state=tk.DISABLED)
        self.prev_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        self.merge_btn = ttk.Button(btn_frame, text="SAVE MERGED FILE", command=self.process_merge, state=tk.DISABLED)
        self.merge_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

        # Replace Text preview with Treeview (Excel-like)
        preview_label = ttk.Label(main, text="Preview (top rows)", font=('Segoe UI', 10, 'bold'))
        preview_label.pack(anchor=tk.W, pady=(6, 0))

        self.preview_frame = ttk.Frame(main, relief=tk.SOLID)
        self.preview_frame.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        # --- 5. Status & Progress ---
        self.prog = ttk.Progressbar(main, mode='determinate', maximum=100)
        self.prog.pack(fill=tk.X, pady=6)
        self.stat_var = tk.StringVar(value="Ready")
        self.stat_bar = ttk.Label(main, textvariable=self.stat_var, relief=tk.SUNKEN, anchor=tk.W)
        self.stat_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def setup_style(self):
        style = ttk.Style()
        # Prefer Windows-like theme on Windows
        try:
            if 'vista' in style.theme_names():
                style.theme_use('vista')
            elif 'xpnative' in style.theme_names():
                style.theme_use('xpnative')
            else:
                style.theme_use('clam')
        except Exception:
            style.theme_use('clam')

        # Global font and padding
        default_font = ('Segoe UI', 9)
        style.configure('.', font=default_font)
        style.configure('TButton', padding=(6, 4))
        style.configure('TEntry', padding=(4, 4))
        style.configure('Treeview', font=('Segoe UI', 9), rowheight=22)
        style.configure('Treeview.Heading', font=('Segoe UI', 9, 'bold'))

        # Attempt to set scaling for high DPI on Windows
        try:
            if platform.system() == 'Windows':
                # set a reasonable scaling for high DPI monitors
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    # -------------------------
    # File loading & detection
    # -------------------------
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
        # Thread-safe progress update
        def _update():
            try:
                with self._prog_lock:
                    self.prog['value'] = value
                    if text is not None:
                        self.stat_var.set(text)
                    # ensure UI refresh
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
                # small animated progress while detecting/reading
                for p in (5, 10, 18):
                    self._set_progress(p)
                    time.sleep(0.06)

                enc = self.detect_enc(path)
                try:
                    df = pd.read_csv(path, encoding=enc, engine='python')
                except Exception:
                    df = pd.read_csv(path, encoding='latin-1', engine='python')

                df.columns = [str(c).strip() for c in df.columns]
                # Ensure columns are unique
                df = df.loc[:, ~df.columns.duplicated()]

                # simulate parsing progress
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
            finally:
                pass

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
            # initialize pull_vars for df2 columns
            self.pull_vars = {col: tk.BooleanVar(value=False) for col in self.all_cols_f2}
            self.filter_checkboxes()
            self.stat_var.set("Source file loaded.")

        if self.df1 is not None and self.df2 is not None:
            self.prev_btn.config(state=tk.NORMAL)
            self.merge_btn.config(state=tk.NORMAL)

    # -------------------------
    # UI helpers
    # -------------------------
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
        # Clear existing
        for w in self.checkbox_widgets:
            try:
                w.destroy()
            except Exception:
                pass
        self.checkbox_widgets = []

        term = self.search_var.get().lower()
        # If pull_vars not initialized yet, nothing to show
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

        # update scrollregion (bound to configure already)
        self.check_canvas.configure(scrollregion=self.check_canvas.bbox("all"))

    def select_all(self):
        for v in self.pull_vars.values():
            v.set(True)

    def deselect_all(self):
        for v in self.pull_vars.values():
            v.set(False)

    # -------------------------
    # Merge logic (now threaded with live progress)
    # -------------------------
    def _merge_worker(self, k1, k2, pull, callback):
        """
        Worker that performs the merge and updates progress periodically.
        Calls callback(result_df_or_None, error_or_None) on completion.
        """
        try:
            self._set_progress(10, "Preparing data for merge...")
            time.sleep(0.05)

            d1 = self.df1.copy()
            # ensure requested pull columns exist in df2
            d2_cols = [c for c in ([k2] + pull) if c in self.df2.columns]
            if k2 not in d2_cols:
                d2_cols.insert(0, k2)
            d2 = self.df2[d2_cols].copy()

            self._set_progress(30, "Normalizing keys...")
            time.sleep(0.05)

            # Normalize join keys to strings and strip
            d1[k1] = d1[k1].astype(str).str.strip()
            d2[k2] = d2[k2].astype(str).str.strip()

            self._set_progress(55, "Performing merge...")
            # perform merge (this may be the heaviest step)
            res = d1.merge(d2, left_on=k1, right_on=k2, how='left', suffixes=('', '_src'))

            # If both keys exist and are different, drop the right key column to avoid duplication
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

        # We'll run merge in a background thread and block the caller until result is ready,
        # but keep UI responsive by using a small wait loop. This allows the progressbar to animate.
        result_container = {'df': None, 'err': None}
        done_event = threading.Event()

        def cb(res, err):
            result_container['df'] = res
            result_container['err'] = err
            done_event.set()

        threading.Thread(target=self._merge_worker, args=(k1, k2, pull, cb), daemon=True).start()

        # Wait for completion but keep UI responsive
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

    # -------------------------
    # Preview (Treeview)
    # -------------------------
    def populate_treeview_from_df(self, df, max_rows=200):
        # Clear existing tree if present
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
            # nothing to show
            return

        # Create treeview
        self.preview_tree = ttk.Treeview(self.preview_frame, columns=cols, show='headings')
        self.preview_vscroll = ttk.Scrollbar(self.preview_frame, orient='vertical', command=self.preview_tree.yview)
        self.preview_hscroll = ttk.Scrollbar(self.preview_frame, orient='horizontal', command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=self.preview_vscroll.set, xscrollcommand=self.preview_hscroll.set)

        # Grid layout
        self.preview_tree.grid(row=0, column=0, sticky='nsew')
        self.preview_vscroll.grid(row=0, column=1, sticky='ns')
        self.preview_hscroll.grid(row=1, column=0, sticky='ew')
        self.preview_frame.columnconfigure(0, weight=1)
        self.preview_frame.rowconfigure(0, weight=1)

        # Setup headings and column widths based on header and sample values
        sample = df.head(max_rows)
        for c in cols:
            # estimate width: header length and sample content
            try:
                max_sample_len = int(sample[c].astype(str).map(len).max()) if not sample.empty else 0
            except Exception:
                max_sample_len = 0
            header_len = len(str(c))
            est = min(max(80, (max(header_len, max_sample_len) * 7)), 400)  # px estimate
            self.preview_tree.heading(c, text=c)
            self.preview_tree.column(c, width=est, anchor='w', stretch=True)

        # Insert rows (limit to max_rows)
        for idx, row in sample.iterrows():
            values = [self._safe_display_value(row.get(c)) for c in cols]
            self.preview_tree.insert('', 'end', values=values)

        # Allow copying cell text on double-click
        self.preview_tree.bind("<Double-1>", self._on_treeview_double_click)

    def _safe_display_value(self, v):
        if pd.isna(v):
            return ''
        return str(v)

    def _on_treeview_double_click(self, event):
        # Copy the clicked cell value to clipboard
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
        # Run preview in a thread to keep UI responsive and show progress
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

    # -------------------------
    # Save merged result (threaded)
    # -------------------------
    def process_merge(self):
        # Run save in background thread to show progress
        def run_save():
            self._set_progress(5, "Starting merge and save...")
            res = self.perform_merge()
            if res is not None:
                path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="merged_output.csv",
                                                    filetypes=[("CSV files", "*.csv")])
                if path:
                    try:
                        self._set_progress(60, "Writing CSV...")
                        # write in chunks to allow progress updates for very large frames
                        try:
                            # attempt to write with a streaming approach if large
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
# TAB 2: DATA PROCESSOR (mostly unchanged but with live progress and threaded full process)
# =============================================================================
class DataProcessorGUI:
    def __init__(self, parent):
        self.parent = parent

        # State variables
        self.df = None
        self.file_path = None
        self.map_full_df = None
        self.map_df = None

        # Progress lock
        self._prog_lock = threading.Lock()

        self.setup_ui()
        self.log("System Ready. Please load MAP.csv if not already present.")
        self.load_map_silent()

    def setup_ui(self):
        main_frame = ttk.Frame(self.parent, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Section 1: File Loading ---
        ttk.Label(main_frame, text="1. Load Data", font=('Helvetica', 10, 'bold')).pack(anchor=tk.W)
        file_btn = ttk.Button(main_frame, text="Select Data File (CSV/XLSX/JSON)", command=self.load_file)
        file_btn.pack(fill=tk.X, pady=5)

        self.file_label = ttk.Label(main_frame, text="No file selected", foreground="gray")
        self.file_label.pack(anchor=tk.W, pady=(0, 15))

        # --- Section 2: Actions ---
        ttk.Label(main_frame, text="2. Data Operations", font=('Helvetica', 10, 'bold')).pack(anchor=tk.W)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)

        ttk.Button(btn_frame, text="Remove BSG", command=self.remove_bsg).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(btn_frame, text="Filter ACT", command=self.filter_act).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(btn_frame, text="Full Process", command=self.run_full_process).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

        # --- Section 3: Progress & Logs ---
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
                    # Use first and third columns as PROVINCENAME and REGION if present
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
                        # try utf-8 first
                        self.df = pd.read_csv(path, encoding='utf-8')
                    except Exception:
                        self.log("UTF-8 failed, trying latin1...")
                        self.df = pd.read_csv(path, encoding='latin1')

                elif ext == '.xlsx':
                    self.df = pd.read_excel(path)
                else:
                    # json or jsonl
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
                # default to json
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

            # Dynamic Headers
            dy_yesterday = f"{yesterday.strftime('%b%d').upper()} (STATUS)"
            dy_today = f"{today.strftime('%b%d').upper()} (STATUS)"
            dy_hours = f"AGED (HOURS) - {today.strftime('%b%d').upper()}"
            dy_bucket = f"AGED BUCKET - {today.strftime('%b%d').upper()}"
            dy_group = f"AGED BUCKET GROUP - {today.strftime('%b%d').upper()}"

            new_headers = ['ALIGNED ACCT', 'ALIGNED JONO', 'ACCT+JONO', 'SEGMENT', 'PRODUCT',
                           'JOCRYEAR', 'DATE TODAY', 'JOTODAY', 'AGEING', 'AGEING (2)',
                           'AREA', 'MSP', dy_yesterday, dy_today, 'JIRA TICKET STATUS',
                           'ACTION TAKEN', 'FINAL STATUS', dy_hours, dy_bucket, dy_group]

            # Step 1: ensure headers
            self._set_progress(8, "Ensuring headers...")
            for col in new_headers:
                if col not in self.df.columns:
                    self.df[col] = None
            time.sleep(0.04)

            # Step 2: Alignment & Date Calculations
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

            # DATEJOCREATED handling
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
            # --- SEGMENT & PRODUCT ---
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

            # Ageing
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

            # Area Mapping
            if self.map_df is not None and 'PROVINCENAME' in self.df.columns:
                area_dict = self.map_df.set_index('PROVINCENAME')['REGION'].to_dict()
                self.df['AREA'] = self.df['PROVINCENAME'].astype(str).map(lambda x: area_dict.get(x, None))
            else:
                self.log("Area mapping skipped (MAP.csv missing or PROVINCENAME not in data).")

            self._set_progress(75, "Starting MSP lookups...")
            time.sleep(0.04)

            # MSP Logic
            if self.map_full_df is not None and 'PROVINCENAME' in self.df.columns:
                self.log("Starting Complex MSP Lookups...")
                self.df['MSP'] = None
                p_mask = self.df['PROVINCENAME'].notna() & (self.df['PROVINCENAME'].astype(str).str.strip() != '')

                # 1. Holy Spirit (example)
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

                # 2. Province | Municipality
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

                # 3. Province Only
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

            # Fill MSP blanks with empty string for consistency
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

        # Run the full process in a background thread so UI remains responsive and progress updates are visible
        threading.Thread(target=self._full_process_worker, daemon=True).start()


# =============================================================================
# TAB 3: XLSX TO CSV CONVERTER (completed, unchanged except minor progress feedback)
# =============================================================================
class ExcelToCsvConverter:
    def __init__(self, parent):
        self.parent = parent
        self.file_path = None
        self.output_folder = None
        self.setup_ui()

    def setup_ui(self):
        main = ttk.Frame(self.parent, padding="20")
        main.pack(fill=tk.BOTH, expand=True)

        # File Selection
        ttk.Label(main, text="Excel File (.xlsx):").grid(row=0, column=0, sticky="w", pady=5)
        self.entry_file = ttk.Entry(main, width=50)
        self.entry_file.grid(row=0, column=1, pady=5, sticky="ew")
        ttk.Button(main, text="Browse", command=self.select_file).grid(row=0, column=2, padx=5)

        # Output Selection
        ttk.Label(main, text="Output Folder:").grid(row=1, column=0, sticky="w", pady=5)
        self.entry_output = ttk.Entry(main, width=50)
        self.entry_output.grid(row=1, column=1, pady=5, sticky="ew")
        ttk.Button(main, text="Browse", command=self.select_output).grid(row=1, column=2, padx=5)

        # Options
        self.split_sheets_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(main, text="Export each sheet to separate CSV", variable=self.split_sheets_var).grid(row=2, column=1, sticky="w", pady=6)

        # Convert button
        ttk.Button(main, text="Convert", command=self.convert).grid(row=3, column=1, pady=10)

        # Status / log
        self.status_label = ttk.Label(main, text="No file selected", foreground="gray")
        self.status_label.grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))

        main.columnconfigure(1, weight=1)

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.file_path = path
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, path)
            self.status_label.config(text=f"Selected: {os.path.basename(path)}", foreground="#007bff")

    def select_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder = folder
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, folder)

    def convert(self):
        if not self.file_path:
            messagebox.showwarning("No file", "Please select an Excel file to convert.")
            return
        out_folder = self.output_folder or os.path.join(os.path.dirname(self.file_path), "converted_csvs")
        os.makedirs(out_folder, exist_ok=True)

        def _convert_worker():
            try:
                wb = load_workbook(self.file_path, read_only=True, data_only=True)
                sheets = wb.sheetnames
                exported = []
                total = len(sheets) if sheets else 1
                count = 0
                if self.split_sheets_var.get():
                    for sheet in sheets:
                        count += 1
                        ws = wb[sheet]
                        rows = list(ws.values)
                        if not rows:
                            continue
                        # Use first row as header if it looks like headers
                        header = [str(c) if c is not None else "" for c in rows[0]]
                        data_rows = rows[1:] if len(rows) > 1 else []
                        df = pd.DataFrame(data_rows, columns=header)
                        out_name = os.path.join(out_folder, f"{os.path.splitext(os.path.basename(self.file_path))[0]}_{sheet}.csv")
                        df.to_csv(out_name, index=False, encoding='utf-8-sig')
                        exported.append(out_name)
                        # update status
                        try:
                            self.parent.after(0, lambda s=count, t=total: self.status_label.config(text=f"Exported {s}/{t} sheets...", foreground="#007bff"))
                        except Exception:
                            pass
                else:
                    # Combine all sheets vertically with sheet name column
                    combined = []
                    for sheet in sheets:
                        count += 1
                        ws = wb[sheet]
                        rows = list(ws.values)
                        if not rows:
                            continue
                        header = [str(c) if c is not None else "" for c in rows[0]]
                        data_rows = rows[1:] if len(rows) > 1 else []
                        df = pd.DataFrame(data_rows, columns=header)
                        df['__sheet__'] = sheet
                        combined.append(df)
                        try:
                            self.parent.after(0, lambda s=count, t=total: self.status_label.config(text=f"Reading {s}/{t} sheets...", foreground="#007bff"))
                        except Exception:
                            pass
                    if combined:
                        big = pd.concat(combined, ignore_index=True)
                        out_name = os.path.join(out_folder, f"{os.path.splitext(os.path.basename(self.file_path))[0]}.csv")
                        big.to_csv(out_name, index=False, encoding='utf-8-sig')
                        exported.append(out_name)

                if exported:
                    try:
                        self.parent.after(0, lambda: self.status_label.config(text=f"Exported {len(exported)} file(s).", foreground="green"))
                        self.parent.after(0, lambda: messagebox.showinfo("Done", f"Exported {len(exported)} CSV file(s) to:\n{out_folder}"))
                    except Exception:
                        pass
                else:
                    try:
                        self.parent.after(0, lambda: self.status_label.config(text="No data exported.", foreground="orange"))
                        self.parent.after(0, lambda: messagebox.showwarning("No data", "No sheets with data were found to export."))
                    except Exception:
                        pass
            except Exception as e:
                try:
                    self.parent.after(0, lambda: messagebox.showerror("Error", f"Conversion failed: {e}"))
                    self.parent.after(0, lambda: self.status_label.config(text="Conversion failed.", foreground="red"))
                except Exception:
                    pass

        threading.Thread(target=_convert_worker, daemon=True).start()


# =============================================================================
# Main application
# =============================================================================
def main():
    root = tk.Tk()
    root.title("Data Toolkit - Jester Miranda")
    root.geometry("1100x700")

    # Notebook with tabs
    nb = ttk.Notebook(root)
    nb.pack(fill=tk.BOTH, expand=True)

    # Tab frames
    tab1 = ttk.Frame(nb)
    tab2 = ttk.Frame(nb)
    tab3 = ttk.Frame(nb)

    nb.add(tab1, text="CSV Merger")
    nb.add(tab2, text="Data Processor")
    nb.add(tab3, text="Excel → CSV")

    # Instantiate tools
    SimpleCSVMerger(tab1)
    DataProcessorGUI(tab2)
    ExcelToCsvConverter(tab3)

    # Start
    root.mainloop()


if __name__ == "__main__":
    main()
