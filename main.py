#!/usr/bin/env python3
"""
Tkinter GUI XLSX → CSV Converter
Shows sheet names and previews first rows before conversion.
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import csv
import os
from openpyxl import load_workbook

def xlsx_to_csv_stream(xlsx_file, output_dir, sheet_name=None):
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

def select_file():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx")]
    )
    entry_file.delete(0, tk.END)
    entry_file.insert(0, file_path)

    # Show sheet names
    if file_path:
        try:
            wb = load_workbook(file_path, read_only=True, data_only=True)
            sheet_list.delete(0, tk.END)  # clear old entries
            for sheet in wb.sheetnames:
                sheet_list.insert(tk.END, sheet)
            wb.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read sheets: {e}")

def select_output_dir():
    dir_path = filedialog.askdirectory(title="Select Output Folder")
    entry_output.delete(0, tk.END)
    entry_output.insert(0, dir_path)

def preview_sheet():
    xlsx_file = entry_file.get()
    selected = sheet_list.curselection()
    if not xlsx_file or not selected:
        messagebox.showerror("Error", "Please select a file and sheet first.")
        return

    sheet_name = sheet_list.get(selected[0])
    try:
        wb = load_workbook(xlsx_file, read_only=True, data_only=True)
        ws = wb[sheet_name]

        preview_text.delete("1.0", tk.END)
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            preview_text.insert(tk.END, f"{row}\n")
            if i >= 4:  # show first 5 rows only
                break

        wb.close()
    except Exception as e:
        messagebox.showerror("Error", f"Preview failed: {e}")

def convert():
    xlsx_file = entry_file.get()
    output_dir = entry_output.get()
    selected = sheet_list.curselection()
    sheet_name = sheet_list.get(selected[0]) if selected else None

    if not xlsx_file or not output_dir:
        messagebox.showerror("Error", "Please select both input file and output folder.")
        return

    success, msg = xlsx_to_csv_stream(xlsx_file, output_dir, sheet_name)
    if success:
        messagebox.showinfo("Success", msg)
    else:
        messagebox.showerror("Error", msg)

# --- GUI Setup ---
root = tk.Tk()
root.title("XLSX → CSV Converter")

tk.Label(root, text="Excel File (.xlsx):").grid(row=0, column=0, sticky="w")
entry_file = tk.Entry(root, width=50)
entry_file.grid(row=0, column=1)
tk.Button(root, text="Browse", command=select_file).grid(row=0, column=2)

tk.Label(root, text="Output Folder:").grid(row=1, column=0, sticky="w")
entry_output = tk.Entry(root, width=50)
entry_output.grid(row=1, column=1)
tk.Button(root, text="Browse", command=select_output_dir).grid(row=1, column=2)

tk.Label(root, text="Sheet Names:").grid(row=2, column=0, sticky="nw")
sheet_list = tk.Listbox(root, height=6, width=50, selectmode=tk.SINGLE)
sheet_list.grid(row=2, column=1, pady=5)

tk.Button(root, text="Preview", command=preview_sheet, bg="lightblue").grid(row=3, column=1, pady=5)

preview_text = tk.Text(root, height=8, width=60, bg="#f9f9f9")
preview_text.grid(row=4, column=0, columnspan=3, padx=10, pady=5)

tk.Button(root, text="Convert", command=convert, bg="lightgreen").grid(row=5, column=1, pady=10)

root.mainloop()
