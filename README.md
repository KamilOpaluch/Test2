import tkinter as tk
from tkinter import ttk
from tkcalendar import Calendar
import pandas as pd
import os

def col_letter(n):
    result = ""
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result

def process_sheet(df):
    indices = []
    for i in range(df.shape[0]):
        if str(df.iat[i, 0]).strip() == "LEVEL_4" and str(df.iat[i, 1]).strip() == "LEVEL_9":
            indices.append(i)
    if not indices:
        return None, "Header not found"
    header_index = indices[1] if len(indices) > 1 else indices[0]
    headers = df.iloc[header_index].tolist()
    new_headers = [f"{str(h).strip()} ({col_letter(idx+1)})" for idx, h in enumerate(headers)]
    data = df.iloc[header_index+1:].dropna(how="all")
    data.columns = new_headers[:len(data.columns)]
    return data, ""

def find_column(df, prefix):
    for col in df.columns:
        if col.startswith(prefix):
            return col
    return None

def open_calendar():
    def select_date():
        selected_date = cal.selection_get()
        date_entry.delete(0, tk.END)
        date_entry.insert(0, selected_date.strftime("%Y%m%d"))
        top.destroy()
    top = tk.Toplevel(root)
    cal = Calendar(top, date_pattern="yyyyMMdd")
    cal.pack(padx=10, pady=10)
    select_btn = ttk.Button(top, text="Select", command=select_date)
    select_btn.pack(pady=5)

def run_analysis():
    date_str = date_entry.get()
    file_path = os.path.join(folder_path, f"{date_str}_CGML_CGME_Backtesting_Capital.xlsx")
    try:
        sheets = ["CGME_10D_VaR_SVaR_Details_ECB", "CGME_10D_VaR_SVaR_Details_PRA"]
        results = {}
        for sheet in sheets:
            df_raw = pd.read_excel(file_path, sheet_name=sheet, header=None)
            table_df, err = process_sheet(df_raw)
            if table_df is None:
                results[sheet] = f"Sheet {sheet}: {err}"
                continue
            col_change_e = find_column(table_df, "Change,% E")
            col_change_f = find_column(table_df, "Change F")
            col_change_l = find_column(table_df, "Change,% L")
            col_change_m = find_column(table_df, "Change M")
            if None in (col_change_e, col_change_f, col_change_l, col_change_m):
                results[sheet] = f"Sheet {sheet}: One or more required columns not found."
                continue
            initial_count = table_df.shape[0]
            cond1 = (table_df[col_change_e] >= 10) & (table_df[col_change_f] >= 500000)
            cond2 = (table_df[col_change_l] >= 10) & (table_df[col_change_m] >= 500000)
            df_filtered = table_df[~(cond1 | cond2)]
            filtered_count = df_filtered.shape[0]
            results[sheet] = f"Sheet {sheet}: Rows before filtering: {initial_count}, after filtering: {filtered_count}"
        output_text.delete("1.0", tk.END)
        for res in results.values():
            output_text.insert(tk.END, res + "\n")
    except Exception as e:
        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, f"Error: {str(e)}")

root = tk.Tk()
root.title("Excel Analysis")
folder_path = "C:/Reports"
date_label = ttk.Label(root, text="Select Date (YYYYMMDD):")
date_label.grid(row=0, column=0, padx=5, pady=5)
date_entry = ttk.Entry(root, width=12)
date_entry.grid(row=0, column=1, padx=5, pady=5)
cal_button = ttk.Button(root, text="📅", command=open_calendar)
cal_button.grid(row=0, column=2, padx=5, pady=5)
run_button = ttk.Button(root, text="Run", command=run_analysis)
run_button.grid(row=0, column=3, padx=5, pady=5)
output_text = tk.Text(root, height=10, width=60)
output_text.grid(row=1, column=0, columnspan=4, padx=5, pady=5)
root.mainloop()
