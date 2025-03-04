import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import pandas as pd
import requests
from io import BytesIO

def col_letter(n):
    result = ""
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result

def process_sheet(df):
    header_index = None
    for i in range(df.shape[0]):
        if str(df.iat[i, 0]).strip() == "LEVEL_4" and str(df.iat[i, 1]).strip() == "LEVEL_9":
            header_index = i
            break
    if header_index is None:
        return None, "Header not found"
    headers = df.iloc[header_index].tolist()
    new_headers = []
    seen = {}
    for idx, h in enumerate(headers):
        h_str = str(h).strip()
        if h_str in seen:
            new_headers.append(f"{h_str} ({col_letter(idx+1)})")
        else:
            seen[h_str] = 1
            new_headers.append(h_str)
    data = df.iloc[header_index+1:].dropna(how="all")
    data.columns = new_headers[:len(data.columns)]
    return data, ""

def run_analysis():
    date_str = date_entry.get_date().strftime("%Y%m%d")
    url = sharepoint_url_template.format(date_str)
    try:
        r = requests.get(url)
        if r.status_code != 200:
            output_text.delete("1.0", tk.END)
            output_text.insert(tk.END, f"Failed to download file. Status code: {r.status_code}")
            return
        file_bytes = BytesIO(r.content)
        sheets = ["CGME_10D_VaR_SVaR_Details_ECB", "CGME_10D_VaR_SVaR_Details_PRA"]
        results = {}
        for sheet in sheets:
            df_raw = pd.read_excel(file_bytes, sheet_name=sheet, header=None)
            table_df, err = process_sheet(df_raw)
            if table_df is None:
                results[sheet] = f"Sheet {sheet}: {err}"
                continue
            initial_count = table_df.shape[0]
            cond1 = (table_df["Change,% E"] >= 10) & (table_df["Change F"] >= 500000)
            cond2 = (table_df["Change,% L"] >= 10) & (table_df["Change M"] >= 500000)
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
sharepoint_url_template = "https://yoursharepointsite.com/path/to/file/{}_CGML_CGME_Backtesting_Capital.xlsx"
date_label = ttk.Label(root, text="Select Date:")
date_label.grid(row=0, column=0, padx=5, pady=5)
date_entry = DateEntry(root, width=12, background="darkblue", foreground="white", borderwidth=2, date_pattern="yyyyMMdd")
date_entry.grid(row=0, column=1, padx=5, pady=5)
run_button = ttk.Button(root, text="Run", command=run_analysis)
run_button.grid(row=0, column=2, padx=5, pady=5)
output_text = tk.Text(root, height=10, width=50)
output_text.grid(row=1, column=0, columnspan=3, padx=5, pady=5)
root.mainloop()
