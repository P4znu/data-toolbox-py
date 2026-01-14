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
        # self.parent is the Tab Frame, acts as root for this section
        
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

        # --- 1. File Selection ---
        ttk.Label(main, text="1. Select CSV Files", font=('Arial', 10, 'bold')).pack(anchor=tk.W)
        f_frame = ttk.Frame(main)
        f_frame.pack(fill=tk.X, pady=5)
        
        self.file1_path = ttk.Entry(f_frame); self.file1_path.pack(fill=tk.X, pady=2)
        ttk.Button(f_frame, text="Browse Primary File", command=lambda: self.browse(1)).pack(fill=tk.X)
        
        self.file2_path = ttk.Entry(f_frame); self.file2_path.pack(fill=tk.X, pady=(10, 2))
        ttk.Button(f_frame, text="Browse Source File", command=lambda: self.browse(2)).pack(fill=tk.X)

        # --- 2. Matching with Search ---
        ttk.Label(main, text="2. Join Keys (Searchable)", font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(10, 0))
        m_frame = ttk.Frame(main)
        m_frame.pack(fill=tk.X, pady=5)
        
        # Search Entry for File 1 Keys
        self.s1_var = tk.StringVar(); self.s1_var.trace_add("write", lambda *a: self.filter_key_list(1))
        ttk.Entry(m_frame, textvariable=self.s1_var, font=('Arial', 8, 'italic')).grid(row=0, column=0, sticky="ew")
        self.match_f1 = ttk.Combobox(m_frame, state="readonly", width=40)
        self.match_f1.grid(row=1, column=0, padx=5)

        ttk.Label(m_frame, text="↔").grid(row=1, column=1)

        # Search Entry for File 2 Keys
        self.s2_var = tk.StringVar(); self.s2_var.trace_add("write", lambda *a: self.filter_key_list(2))
        ttk.Entry(m_frame, textvariable=self.s2_var, font=('Arial', 8, 'italic')).grid(row=0, column=2, sticky="ew")
        self.match_f2 = ttk.Combobox(m_frame, state="readonly", width=40)
        self.match_f2.grid(row=1, column=2, padx=5)

        # --- 3. Columns to Pull ---
        ttk.Label(main, text="3. Columns to Pull", font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(10, 0))
        ctrl = ttk.Frame(main); ctrl.pack(fill=tk.X, pady=5)
        self.search_var = tk.StringVar(); self.search_var.trace_add("write", self.filter_checkboxes)
        ttk.Entry(ctrl, textvariable=self.search_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        ttk.Button(ctrl, text="All", width=5, command=self.select_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(ctrl, text="None", width=5, command=self.deselect_all).pack(side=tk.LEFT)

        self.check_canvas = tk.Canvas(main, bg="white", height=150, highlightthickness=1)
        self.check_scroll = ttk.Scrollbar(main, orient="vertical", command=self.check_canvas.yview)
        self.check_frame = ttk.Frame(self.check_canvas)
        self.check_canvas.create_window((0, 0), window=self.check_frame, anchor="nw")
        self.check_canvas.configure(yscrollcommand=self.check_scroll.set)
        self.check_canvas.pack(fill=tk.BOTH, expand=False)
        self.check_scroll.pack(side=tk.RIGHT, fill=tk.Y, before=self.check_canvas)

        # --- 4. Actions & Preview ---
        btn_frame = ttk.Frame(main); btn_frame.pack(fill=tk.X, pady=10)
        self.prev_btn = ttk.Button(btn_frame, text="PREVIEW DATA", command=self.show_preview, state=tk.DISABLED)
        self.prev_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.merge_btn = ttk.Button(btn_frame, text="SAVE MERGED FILE", command=self.process_merge, state=tk.DISABLED)
        self.merge_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)

        self.preview_box = tk.Text(main, height=12, font=("Courier New", 9), bg="#f8f8f8")
        self.preview_box.pack(fill=tk.BOTH, expand=True)

        # --- 5. Status & Progress ---
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
        self.stat_var.set(f"Loading file {num}...")
        self.prog.start()
        
        def run():
            try:
                enc = self.detect_enc(path)
                try:
                    df = pd.read_csv(path, encoding=enc, engine='python')
                except:
                    df = pd.read_csv(path, encoding='latin-1', engine='python')
                
                df.columns = [str(c).strip() for c in df.columns]
                
                self.parent.after(0, lambda: self.on_load_success(num, df))
            except Exception as e:
                self.parent.after(0, lambda: messagebox.showerror("File Error", str(e)))
            finally:
                self.parent.after(0, self.prog.stop)
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
        self.stat_var.set(f"File {num} ready.")

    def filter_key_list(self, num):
        term = self.s1_var.get().lower() if num == 1 else self.s2_var.get().lower()
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
        
        if not k1 or not k2 or not pull:
            messagebox.showwarning("Input Missing", "Select both keys and at least one column to pull."); return None

        try:
            self.prog['value'] = 20; self.stat_var.set("Processing merge...")
            d1 = self.df1.copy()
            d2_cols = list(set([k2] + pull)) 
            d2 = self.df2[d2_cols].copy()

            d1[k1] = d1[k1].astype(str).str.strip()
            d2[k2] = d2[k2].astype(str).str.strip()

            res = d1.merge(d2, left_on=k1, right_on=k2, how='left')
            if k1 != k2: res.drop(columns=[k2], inplace=True)
            
            self.prog['value'] = 80; self.stat_var.set("Merge complete.")
            return res
        except Exception as e:
            messagebox.showerror("Merge Error", f"Error during merge: {str(e)}")
            return None

    def show_preview(self):
        res = self.perform_merge()
        if res is not None:
            self.preview_box.delete(1.0, tk.END)
            self.preview_box.insert(tk.END, res.head(20).to_string())
            self.prog['value'] = 100

    def process_merge(self):
        res = self.perform_merge()
        if res is not None:
            path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="merged_output.csv")
            if path:
                res.to_csv(path, index=False, encoding='utf-8-sig')
                messagebox.showinfo("Success", f"File saved to:\n{path}")
                self.stat_var.set("Saved successfully.")
            self.prog['value'] = 0

# =============================================================================
# TAB 2: DATA PROCESSOR
# =============================================================================
class DataProcessorGUI:
    def __init__(self, parent):
        self.parent = parent
        
        # State variables
        self.df = None
        self.file_path = None
        self.map_full_df = None
        self.map_df = None
        
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
        
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, mode='determinate')
        self.progress.pack(fill=tk.X, pady=10)

        self.log_area = scrolledtext.ScrolledText(main_frame, height=12, font=('Consolas', 9), bg="#f8f9fa")
        self.log_area.pack(fill=tk.BOTH, expand=True)

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
                    except UnicodeDecodeError:
                        self.log("UTF-8 failed, trying latin1...")
                        self.df = pd.read_csv(path, encoding='latin1')
                        
                elif ext == '.xlsx':
                    self.df = pd.read_excel(path)
                else:
                    self.df = pd.read_json(path)
                
                self.file_path = path
                self.file_label.config(text=f"Loaded: {os.path.basename(path)}", foreground="#007bff")
                self.log(f"✅ Success: Loaded {len(self.df)} rows.")
            except Exception as e:
                self.log(f"❌ Load Error: {e}")
                messagebox.showerror("Error", f"Failed to load: {e}")

    def save_df(self, suffix):
        output_folder = 'data'
        os.makedirs(output_folder, exist_ok=True)
        base = os.path.splitext(os.path.basename(self.file_path))[0]
        ext = os.path.splitext(self.file_path)[1]
        out_path = os.path.join(output_folder, f"{base}_{suffix}{ext}")
        
        if ext == '.csv': self.df.to_csv(out_path, index=False)
        elif ext == '.xlsx': self.df.to_excel(out_path, index=False)
        else: self.df.to_json(out_path, orient='records', indent=4)
        
        self.log(f"💾 File Saved: {out_path}")
        return out_path

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

    def run_full_process(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Load data first!")
            return

        try:
            self.log("🚀 Starting Full Process...")
            self.progress['value'] = 5
            
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

            for col in new_headers:
                if col not in self.df.columns: self.df[col] = None

            self.log("Alignment & Date Calculations...")
            self.df['ALIGNED ACCT'] = self.df['ACCTNO'].astype(str).str.strip().str.zfill(13)
            self.df['ALIGNED JONO'] = self.df['JONO'].astype(str).str.strip().str.zfill(8)
            self.df['ACCT+JONO'] = self.df.apply(lambda r: f"{r['ALIGNED ACCT']}:{r['ALIGNED JONO']}" if r['ALIGNED JONO'] else None, axis=1)
            
            self.df['DATEJOCREATED'] = pd.to_datetime(self.df['DATEJOCREATED'], errors='coerce')
            self.df['JOCRYEAR'] = self.df['DATEJOCREATED'].dt.year
            self.df['DATE TODAY'] = today
            self.df['JOTODAY'] = (today - self.df['DATEJOCREATED']).dt.days

            # --- SEGMENT & PRODUCT ---
            if 'PACKAGENAME' in self.df.columns and 'PROVINCENAME' in self.df.columns:
                self.log("Calculating SEGMENT and PRODUCT...")
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
            
            self.progress['value'] = 40

            # Ageing
            self.log("Applying Ageing Buckets...")
            self.df['AGEING'] = np.select(
                [self.df['JOTODAY'] <= 1, self.df['JOTODAY'] <= 3, self.df['JOTODAY'] <= 5, self.df['JOTODAY'] <= 15, self.df['JOTODAY'] <= 30, self.df['JOTODAY'] <= 60, self.df['JOTODAY'] > 60],
                ["0-1 D", "2-3 D", "3-5 D", "5-15 D", "15-30 D", "30-60 D", "> 60 D"], default=None)
            
            self.df['AGEING (2)'] = np.select(
                [self.df['JOTODAY'] <= 5, self.df['JOTODAY'] <= 15, self.df['JOTODAY'] <= 30, self.df['JOTODAY'] <= 60, self.df['JOTODAY'] > 60],
                ["0-5 D", "5-15 D", "15-30 D", "30-60 D", "> 60 D"], default=None)

            self.df[dy_hours] = self.df['JOTODAY'] * 24
            self.df[dy_bucket] = today.strftime('%d-%b')
            self.df[dy_group] = self.df['AGEING (2)']

            # Area Mapping
            if self.map_df is not None:
                self.log("Mapping AREA...")
                area_dict = self.map_df.set_index('PROVINCENAME')['REGION'].to_dict()
                self.df['AREA'] = self.df['PROVINCENAME'].map(area_dict)

            # MSP Logic
            if self.map_full_df is not None:
                self.log("Starting Complex MSP Lookups...")
                self.df['MSP'] = None
                p_mask = self.df['PROVINCENAME'].notna() & (self.df['PROVINCENAME'].astype(str).str.strip() != '')

                # 1. Holy Spirit
                if 'BARANGAYNAME' in self.df.columns and self.map_full_df.shape[1] > 8:
                    cond1 = (self.df['BARANGAYNAME'].astype(str).str.lower() == 'holy spirit') & \
                            (self.df['PROVINCENAME'].astype(str).str.lower().str.contains('metro manila', na=False))
                    map_f_i = self.map_full_df.iloc[:, [5, 8]].dropna().copy()
                    map_f_i.columns = ['k', 'v']
                    map_f_i['k'] = map_f_i['k'].astype(str).str.lower().str.strip()
                    h_dict = dict(zip(map_f_i['k'], map_f_i['v']))
                    self.df.loc[p_mask & cond1, 'MSP'] = h_dict.get('holy spirit')

                # 2. Province | Municipality
                m_mask = p_mask & self.df['MSP'].isna()
                if 'MUNICIPALITYNAME' in self.df.columns and self.map_full_df.shape[1] > 8:
                    df_key = self.df['PROVINCENAME'].astype(str).str.lower().str.strip() + '|' + self.df['MUNICIPALITYNAME'].astype(str).str.lower().str.strip()
                    map_e_g_i = self.map_full_df.iloc[:, [4, 6, 8]].dropna().copy()
                    map_e_g_i['key'] = map_e_g_i.iloc[:, 0].astype(str).str.lower().str.strip() + '|' + map_e_g_i.iloc[:, 1].astype(str).str.lower().str.strip()
                    m_dict = dict(zip(map_e_g_i['key'], map_e_g_i.iloc[:, 2]))
                    self.df.loc[m_mask, 'MSP'] = df_key.loc[m_mask].map(m_dict)

                # 3. Province Only
                m_mask = p_mask & self.df['MSP'].isna()
                map_e_i = self.map_full_df.iloc[:, [4, 8]].dropna().copy()
                map_e_i.columns = ['k', 'v']
                map_e_i['k'] = map_e_i['k'].astype(str).str.lower().str.strip()
                p_dict = dict(zip(map_e_i['k'], map_e_i['v']))
                self.df.loc[m_mask, 'MSP'] = self.df['PROVINCENAME'].astype(str).str.lower().str.strip().map(p_dict)

            self.df['MSP'] = self.df['MSP'].fillna('')
            self.progress['value'] = 90
            self.save_df("processed")
            self.progress['value'] = 100
            self.log("✅ ALL CALCULATIONS COMPLETE.")
            messagebox.showinfo("Done", "Processing successful!")

        except Exception as e:
            self.log(f"❌ Error: {e}")
            messagebox.showerror("Error", f"Processing failed: {e}")

# =============================================================================
# TAB 3: XLSX TO CSV CONVERTER
# =============================================================================
class ExcelToCsvConverter:
    def __init__(self, parent):
        self.parent = parent
        self.setup_ui()

    def setup_ui(self):
        main = ttk.Frame(self.parent, padding="20")
        main.pack(fill=tk.BOTH, expand=True)

        # File Selection
        ttk.Label(main, text="Excel File (.xlsx):").grid(row=0, column=0, sticky="w", pady=5)
        self.entry_file = ttk.Entry(main, width=50)
        self.entry_file.grid(row=0, column=1, pady=5)
        ttk.Button(main, text="Browse", command=self.select_file).grid(row=0, column=2, padx=5)

        # Output Selection
        ttk.Label(main, text="Output Folder:").grid(row=1, column=0, sticky="w", pady=5)
        self.entry_output = ttk.Entry(main, width=50)
        self.entry_output.grid(row=1, column=1, pady=5)
        ttk.Button(main, text="Browse", command=self.select_output_dir).grid(row=1, column=2, padx=5)

        # Sheet List
        ttk.Label(main, text="Sheet Names:").grid(row=2, column=0, sticky="nw", pady=10)
        self.sheet_list = tk.Listbox(main, height=6, width=50, selectmode=tk.SINGLE)
        self.sheet_list.grid(row=2, column=1, pady=10)

        # Buttons
        btn_frame = ttk.Frame(main)
        btn_frame.grid(row=3, column=1, pady=5)
        ttk.Button(btn_frame, text="Preview Selected Sheet", command=self.preview_sheet).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Convert All / Selected", command=self.convert).pack(side=tk.LEFT, padx=5)

        # Preview Area
        self.preview_text = tk.Text(main, height=10, width=60, bg="#f9f9f9")
        self.preview_text.grid(row=4, column=0, columnspan=3, padx=10, pady=15)

    def select_file(self):
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, file_path)
            self.load_sheets(file_path)

    def select_output_dir(self):
        dir_path = filedialog.askdirectory(title="Select Output Folder")
        if dir_path:
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, dir_path)

    def load_sheets(self, file_path):
        try:
            wb = load_workbook(file_path, read_only=True, data_only=True)
            self.sheet_list.delete(0, tk.END)
            for sheet in wb.sheetnames:
                self.sheet_list.insert(tk.END, sheet)
            wb.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read sheets: {e}")

    def xlsx_to_csv_stream(self, xlsx_file, output_dir, sheet_name=None):
        try:
            wb = load_workbook(xlsx_file, read_only=True, data_only=True)
            sheets = [sheet_name] if sheet_name else wb.sheetnames
            for sheet in sheets:
                ws = wb[sheet]
                csv_file = os.path.join(output_dir, f"{sheet}.csv")
                with open(csv_file, "w", newline="", encoding="utf-8") as f:
                    writer = csv.writer(f)
                    for row in ws.iter_rows(values_only=True):
                        writer.writerow(row)
            wb.close()
            return True, f"Converted {len(sheets)} sheet(s) to CSV in {output_dir}"
        except Exception as e:
            return False, str(e)

    def preview_sheet(self):
        xlsx_file = self.entry_file.get()
        selected = self.sheet_list.curselection()
        if not xlsx_file or not selected:
            messagebox.showerror("Error", "Please select a file and sheet first.")
            return

        sheet_name = self.sheet_list.get(selected[0])
        try:
            wb = load_workbook(xlsx_file, read_only=True, data_only=True)
            ws = wb[sheet_name]
            self.preview_text.delete("1.0", tk.END)
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                self.preview_text.insert(tk.END, f"{row}\n")
                if i >= 4: break # Limit preview
            wb.close()
        except Exception as e:
            messagebox.showerror("Error", f"Preview failed: {e}")

    def convert(self):
        xlsx_file = self.entry_file.get()
        output_dir = self.entry_output.get()
        selected = self.sheet_list.curselection()
        sheet_name = self.sheet_list.get(selected[0]) if selected else None

        if not xlsx_file or not output_dir:
            messagebox.showerror("Error", "Please select both input file and output folder.")
            return

        success, msg = self.xlsx_to_csv_stream(xlsx_file, output_dir, sheet_name)
        if success: messagebox.showinfo("Success", msg)
        else: messagebox.showerror("Error", msg)

# =============================================================================
# MAIN APP LAUNCHER
# =============================================================================
class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Jester's Data Toolbox")
        self.root.geometry("800x750")

        # Create the Tab Container (Notebook)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Create Frames for each tab
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        self.tab3 = ttk.Frame(self.notebook)

        # Add tabs to notebook
        self.notebook.add(self.tab1, text=" CSV Merger ")
        self.notebook.add(self.tab2, text=" Data Processor ")
        self.notebook.add(self.tab3, text=" Excel to CSV ")

        # Initialize the tools inside their respective tabs
        SimpleCSVMerger(self.tab1)
        DataProcessorGUI(self.tab2)
        ExcelToCsvConverter(self.tab3)

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()
