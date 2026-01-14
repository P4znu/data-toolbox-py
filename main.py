import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import chardet
import os
import threading
import numpy as np
import warnings
import csv
from datetime import datetime
from openpyxl import load_workbook

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
        self.setup_ui()

    def setup_ui(self):
        main = ttk.Frame(self.parent, padding="15")
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main, text="1. Select CSV Files", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W)
        f_frame = ttk.Frame(main)
        f_frame.pack(fill=tk.X, pady=5)
        
        self.file1_path = ttk.Entry(f_frame); self.file1_path.pack(fill=tk.X, pady=2)
        ttk.Button(f_frame, text="Browse Primary File", command=lambda: self.browse(1)).pack(fill=tk.X)
        
        self.file2_path = ttk.Entry(f_frame); self.file2_path.pack(fill=tk.X, pady=(10, 2))
        ttk.Button(f_frame, text="Browse Source File", command=lambda: self.browse(2)).pack(fill=tk.X)

        ttk.Label(main, text="2. Join Keys (Searchable)", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W, pady=(10, 0))
        m_frame = ttk.Frame(main)
        m_frame.pack(fill=tk.X, pady=5)
        
        self.s1_var = tk.StringVar(); self.s1_var.trace_add("write", lambda *a: self.filter_key_list(1))
        ttk.Entry(m_frame, textvariable=self.s1_var, font=('Arial', 8, 'italic')).grid(row=0, column=0, sticky="ew")
        self.match_f1 = ttk.Combobox(m_frame, state="readonly", width=35)
        self.match_f1.grid(row=1, column=0, padx=5)

        ttk.Label(m_frame, text="↔").grid(row=1, column=1)

        self.s2_var = tk.StringVar(); self.s2_var.trace_add("write", lambda *a: self.filter_key_list(2))
        ttk.Entry(m_frame, textvariable=self.s2_var, font=('Arial', 8, 'italic')).grid(row=0, column=2, sticky="ew")
        self.match_f2 = ttk.Combobox(m_frame, state="readonly", width=35)
        self.match_f2.grid(row=1, column=2, padx=5)

        ttk.Label(main, text="3. Columns to Pull", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W, pady=(10, 0))
        ctrl = ttk.Frame(main); ctrl.pack(fill=tk.X, pady=5)
        self.search_var = tk.StringVar(); self.search_var.trace_add("write", self.filter_checkboxes)
        ttk.Entry(ctrl, textvariable=self.search_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        ttk.Button(ctrl, text="All", width=5, command=self.select_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(ctrl, text="None", width=5, command=self.deselect_all).pack(side=tk.LEFT)

        self.check_canvas = tk.Canvas(main, bg="white", height=120, highlightthickness=1)
        self.check_scroll = ttk.Scrollbar(main, orient="vertical", command=self.check_canvas.yview)
        self.check_frame = ttk.Frame(self.check_canvas)
        self.check_canvas.create_window((0, 0), window=self.check_frame, anchor="nw")
        self.check_canvas.configure(yscrollcommand=self.check_scroll.set)
        self.check_canvas.pack(fill=tk.BOTH, expand=False)
        self.check_scroll.pack(side=tk.RIGHT, fill=tk.Y, before=self.check_canvas)

        btn_frame = ttk.Frame(main); btn_frame.pack(fill=tk.X, pady=10)
        self.prev_btn = ttk.Button(btn_frame, text="PREVIEW DATA", command=self.show_preview, state=tk.DISABLED)
        self.prev_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.merge_btn = ttk.Button(btn_frame, text="SAVE MERGED FILE", command=self.process_merge, state=tk.DISABLED)
        self.merge_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)

        self.preview_box = tk.Text(main, height=10, font=("Consolas", 9), bg="#f8f8f8")
        self.preview_box.pack(fill=tk.BOTH, expand=True)

        self.prog = ttk.Progressbar(main, mode='determinate'); self.prog.pack(fill=tk.X, pady=5)
        self.stat_var = tk.StringVar(value="Ready"); self.stat_bar = ttk.Label(main, textvariable=self.stat_var, relief=tk.SUNKEN, anchor=tk.W)
        self.stat_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def detect_enc(self, path):
        with open(path, 'rb') as f:
            raw = f.read(100000)
            res = chardet.detect(raw)
            return res['encoding'] if res['encoding'] else 'utf-8'

    def browse(self, num):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if path:
            target = self.file1_path if num == 1 else self.file2_path
            target.delete(0, tk.END); target.insert(0, path)
            self.load_data(num)

    def load_data(self, num):
        path = self.file1_path.get() if num == 1 else self.file2_path.get()
        def run():
            try:
                enc = self.detect_enc(path)
                df = pd.read_csv(path, encoding=enc, engine='python')
                df.columns = [str(c).strip() for c in df.columns]
                self.parent.after(0, lambda: self.on_load_success(num, df))
            except Exception as e:
                self.parent.after(0, lambda: messagebox.showerror("Error", str(e)))
        threading.Thread(target=run, daemon=True).start()

    def on_load_success(self, num, df):
        if num == 1:
            self.df1 = df; self.all_cols_f1 = list(df.columns)
            self.filter_key_list(1)
        else:
            self.df2 = df; self.all_cols_f2 = list(df.columns)
            self.filter_key_list(2)
            self.pull_vars = {col: tk.BooleanVar() for col in self.all_cols_f2}
            self.filter_checkboxes()
        if self.df1 is not None and self.df2 is not None:
            self.prev_btn.config(state=tk.NORMAL); self.merge_btn.config(state=tk.NORMAL)

    def filter_key_list(self, num):
        term = (self.s1_var.get() if num == 1 else self.s2_var.get()).lower()
        full = self.all_cols_f1 if num == 1 else self.all_cols_f2
        filt = [c for c in full if term in c.lower()]
        target = self.match_f1 if num == 1 else self.match_f2
        target['values'] = filt
        if filt: target.current(0)

    def filter_checkboxes(self, *args):
        for w in self.checkbox_widgets: w.destroy()
        self.checkbox_widgets = []
        term = self.search_var.get().lower()
        for col in self.all_cols_f2:
            if term in col.lower():
                cb = tk.Checkbutton(self.check_frame, text=col, variable=self.pull_vars[col], bg="white")
                cb.pack(fill=tk.X, padx=5); self.checkbox_widgets.append(cb)
        self.check_frame.update_idletasks()
        self.check_canvas.config(scrollregion=self.check_canvas.bbox("all"))

    def select_all(self):
        for v in self.pull_vars.values(): v.set(True)
    def deselect_all(self):
        for v in self.pull_vars.values(): v.set(False)

    def perform_merge(self):
        k1, k2 = self.match_f1.get(), self.match_f2.get()
        pull = [c for c, v in self.pull_vars.items() if v.get()]
        if not k1 or not k2 or not pull: return None
        try:
            d1 = self.df1.copy()
            d2 = self.df2[list(set([k2] + pull))].copy()
            d1[k1] = d1[k1].astype(str).str.strip()
            d2[k2] = d2[k2].astype(str).str.strip()
            res = d1.merge(d2, left_on=k1, right_on=k2, how='left')
            if k1 != k2: res.drop(columns=[k2], inplace=True)
            return res
        except Exception as e:
            messagebox.showerror("Merge Error", str(e)); return None

    def show_preview(self):
        res = self.perform_merge()
        if res is not None:
            self.preview_box.delete(1.0, tk.END)
            self.preview_box.insert(tk.END, res.head(20).to_string())

    def process_merge(self):
        res = self.perform_merge()
        if res is not None:
            path = filedialog.asksaveasfilename(defaultextension=".csv")
            if path: res.to_csv(path, index=False, encoding='utf-8-sig')

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
        self.setup_ui()
        self.load_map_silent()

    def setup_ui(self):
        main_frame = ttk.Frame(self.parent, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="1. Load Data", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W)
        ttk.Button(main_frame, text="Select Data File", command=self.load_file).pack(fill=tk.X, pady=5)
        self.file_label = ttk.Label(main_frame, text="No file selected", foreground="gray")
        self.file_label.pack(anchor=tk.W, pady=(0, 15))

        ttk.Label(main_frame, text="2. Data Operations", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W)
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame, text="Remove BSG", command=self.remove_bsg).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(btn_frame, text="Filter ACT", command=self.filter_act).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(btn_frame, text="Full Process", command=self.run_full_process_threaded).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

        ttk.Label(main_frame, text="3. Progress & Log", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W, pady=(15, 0))
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.pack(fill=tk.X, pady=10)
        self.log_area = scrolledtext.ScrolledText(main_frame, height=12, font=('Consolas', 9), bg="#f8f9fa")
        self.log_area.pack(fill=tk.BOTH, expand=True)

    def log(self, message):
        self.log_area.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_area.see(tk.END)

    def load_map_silent(self):
        if os.path.exists('MAP.csv'):
            try:
                self.map_full_df = pd.read_csv('MAP.csv')
                if self.map_full_df.shape[1] >= 3:
                    self.map_df = self.map_full_df.iloc[:, [0, 2]].copy()
                    self.map_df.columns = ['PROVINCENAME', 'REGION']
                    self.map_df.drop_duplicates(subset=['PROVINCENAME'], inplace=True)
                    self.log("✅ MAP.csv loaded.")
            except Exception as e: self.log(f"❌ Map Error: {e}")

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Data Files", "*.csv *.xlsx *.json")])
        if path:
            try:
                ext = os.path.splitext(path)[1].lower()
                if ext == '.csv': self.df = pd.read_csv(path, encoding='latin1')
                elif ext == '.xlsx': self.df = pd.read_excel(path)
                else: self.df = pd.read_json(path)
                self.file_path = path
                self.file_label.config(text=f"Loaded: {os.path.basename(path)}", foreground="#007bff")
                self.log(f"✅ Loaded {len(self.df)} rows.")
            except Exception as e: messagebox.showerror("Error", str(e))

    def save_df(self, suffix):
        os.makedirs('data', exist_ok=True)
        out = os.path.join('data', f"{os.path.splitext(os.path.basename(self.file_path))[0]}_{suffix}.csv")
        self.df.to_csv(out, index=False)
        self.log(f"💾 Saved: {out}")

    def remove_bsg(self):
        if self.df is not None and 'DIVISIONCODE' in self.df.columns:
            self.df = self.df[self.df['DIVISIONCODE'] != 'BSG']
            self.log("Removed BSG rows."); self.save_df("no_bsg")

    def filter_act(self):
        if self.df is not None and 'SUBSCRIBERSTATUSCODE' in self.df.columns:
            self.df = self.df[self.df['SUBSCRIBERSTATUSCODE'].astype(str).str.contains('ACT', na=False)]
            self.log("Filtered ACT."); self.save_df("act_only")

    def run_full_process_threaded(self):
        if self.df is None: return
        threading.Thread(target=self.run_full_process_logic, daemon=True).start()

    def run_full_process_logic(self):
        try:
            self.log("🚀 Processing Logic Started...")
            self.progress['value'] = 10
            today = pd.to_datetime('today').normalize()
            yesterday = today - pd.Timedelta(days=1)
            
            # Dynamic Columns
            dy_yest = f"{yesterday.strftime('%b%d').upper()} (STATUS)"
            dy_today = f"{today.strftime('%b%d').upper()} (STATUS)"
            dy_hours = f"AGED (HOURS) - {today.strftime('%b%d').upper()}"
            
            for col in ['ALIGNED ACCT', 'ALIGNED JONO', 'SEGMENT', 'PRODUCT', 'AREA', 'MSP', dy_yest, dy_today, dy_hours]:
                if col not in self.df.columns: self.df[col] = None

            self.df['ALIGNED ACCT'] = self.df['ACCTNO'].astype(str).str.strip().str.zfill(13)
            self.df['ALIGNED JONO'] = self.df['JONO'].astype(str).str.strip().str.zfill(8)
            self.df['DATEJOCREATED'] = pd.to_datetime(self.df['DATEJOCREATED'], errors='coerce')
            self.df['JOTODAY'] = (today - self.df['DATEJOCREATED']).dt.days

            # Segment Regex
            if 'PACKAGENAME' in self.df.columns:
                cond = [
                    self.df['PACKAGENAME'].str.contains('BIDA', case=False, na=False),
                    self.df['PACKAGENAME'].str.contains('S2S', case=False, na=False),
                    self.df['PACKAGENAME'].str.contains(r'(AIR INTERNET|FIBER X|GAME CHANGER|HOME BASE)', case=False, na=False)
                ]
                self.df['SEGMENT'] = np.select(cond, ["BIDA", "S2S", "RES"], default="OTHER")
            
            # MSP Lookup (Barangay -> City -> Province)
            if self.map_full_df is not None:
                self.log("Mapping MSP tiers...")
                # Tier 3 (Province) fallback
                p_map = dict(zip(self.map_full_df.iloc[:,4].str.lower(), self.map_full_df.iloc[:,8]))
                self.df['MSP'] = self.df['PROVINCENAME'].str.lower().map(p_map)

            self.progress['value'] = 90
            self.save_df("full_processed")
            self.progress['value'] = 100
            self.log("✅ Finished.")
        except Exception as e: self.log(f"❌ Error: {e}")

# =============================================================================
# TAB 3: EXCEL TO CSV
# =============================================================================
class ExcelToCsvConverter:
    def __init__(self, parent):
        self.parent = parent
        self.setup_ui()

    def setup_ui(self):
        main = ttk.Frame(self.parent, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main, text="Excel File:").grid(row=0, column=0, sticky="w")
        self.ent = ttk.Entry(main, width=40); self.ent.grid(row=0, column=1)
        ttk.Button(main, text="Browse", command=self.load).grid(row=0, column=2)
        self.sheets = tk.Listbox(main, height=5); self.sheets.grid(row=1, column=1, sticky="ew", pady=10)
        ttk.Button(main, text="Convert All", command=self.convert).grid(row=2, column=1)

    def load(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            self.ent.delete(0, tk.END); self.ent.insert(0, path)
            wb = load_workbook(path, read_only=True)
            self.sheets.delete(0, tk.END)
            for s in wb.sheetnames: self.sheets.insert(tk.END, s)

    def convert(self):
        path = self.ent.get()
        out_dir = filedialog.askdirectory()
        if path and out_dir:
            wb = load_workbook(path, data_only=True)
            for s in wb.sheetnames:
                ws = wb[s]
                with open(os.path.join(out_dir, f"{s}.csv"), 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    for r in ws.iter_rows(values_only=True): writer.writerow(r)
            messagebox.showinfo("Success", "Done!")

# =============================================================================
# MAIN
# =============================================================================
class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Toolbox Pro")
        self.root.geometry("800x700")
        nb = ttk.Notebook(root)
        nb.pack(fill=tk.BOTH, expand=True)
        t1, t2, t3 = ttk.Frame(nb), ttk.Frame(nb), ttk.Frame(nb)
        nb.add(t1, text=" CSV Merger "); nb.add(t2, text=" Data Processor "); nb.add(t3, text=" Excel to CSV ")
        SimpleCSVMerger(t1); DataProcessorGUI(t2); ExcelToCsvConverter(t3)

if __name__ == "__main__":
    root = tk.Tk(); app = MainApp(root); root.mainloop()
