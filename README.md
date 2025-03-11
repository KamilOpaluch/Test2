```python
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

def find_column(df, target):
    for col in df.columns:
        if col == target:
            return col
    return None

def find_column_by_suffix(df, suffix):
    for col in df.columns:
        if col.endswith(suffix):
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

def insert_sentence(parts):
    for text, tag in parts:
        output_text.insert(tk.END, text, tag)
    output_text.insert(tk.END, "\n")

def run_analysis():
    date_str = date_entry.get()
    file_path = os.path.join(folder_path, f"{date_str}_CGML_CGME_Backtesting_Capital.xlsx")
    try:
        sheets = ["CGME_10D_VaR_SVaR_Details_ECB", "CGME_10D_VaR_SVaR_Details_PRA"]
        output_text.delete("1.0", tk.END)
        for sheet in sheets:
            df_raw = pd.read_excel(file_path, sheet_name=sheet, header=None)
            table_df, err = process_sheet(df_raw)
            if table_df is None:
                output_text.insert(tk.END, f"Sheet {sheet}: {err}\n")
                continue

            # Basic columns for both comments
            name_col = find_column_by_suffix(table_df, "(B)")
            target1_col = find_column_by_suffix(table_df, "(C)")   # VaR target
            target2_col = find_column_by_suffix(table_df, "(J)")   # SVaR target
            change_e_col = find_column(table_df, "Change, % (E)")   # VaR % change
            change_f_col = find_column(table_df, "Change F (F)")     # VaR change amount
            change_l_col = find_column(table_df, "Change, % (L)")   # SVaR % change
            change_m_col = find_column(table_df, "Change M (M)")     # SVaR change amount

            # Extended columns for CVaR and CSVaR
            cv_change_col = find_column(table_df, "Change (I)")      # CVaR change amount
            cv_target_col = find_column_by_suffix(table_df, "(G)")   # CVaR target value
            cs_change_col = find_column(table_df, "Change (P)")      # CSVaR change amount
            cs_target_col = find_column_by_suffix(table_df, "(O)")   # CSVaR target value

            if None in (name_col, target1_col, target2_col, change_e_col, change_f_col, 
                        change_l_col, change_m_col, cv_change_col, cv_target_col, cs_change_col, cs_target_col):
                output_text.insert(tk.END, f"Sheet {sheet}: One or more required columns not found.\n")
                continue

            # Conditions using threshold 0.1 for % and >=500000 or <=-500000 for change amounts
            cond1 = (table_df[change_e_col] >= 0.1) & ((table_df[change_f_col] >= 500000) | (table_df[change_f_col] <= -500000))
            cond2 = (table_df[change_l_col] >= 0.1) & ((table_df[change_m_col] >= 500000) | (table_df[change_m_col] <= -500000))
            df_filtered = table_df[cond1 | cond2]

            # Separate lists for VaR and SVaR comments
            var_comments = []
            svar_comments = []

            for idx, row in df_filtered.iterrows():
                # VaR comment with extended CVaR details
                if (row[change_e_col] >= 0.1) and ((row[change_f_col] >= 500000) or (row[change_f_col] <= -500000)):
                    direction = "increased" if row[change_f_col] >= 0 else "decreased"
                    change_val = abs(row[change_f_col]) / 1e6
                    target_val = abs(row[target1_col]) / 1e6
                    ext_direction = "increased" if row[cv_change_col] >= 0 else "decreased"
                    ext_change_val = abs(row[cv_change_col]) / 1e6
                    ext_target_val = abs(row[cv_target_col]) / 1e6
                    var_parts = [
                        (f"{row[name_col]} VaR {direction} by ", None),
                        (f"${change_val:.2f}mm", "bold"),
                        (" to ", None),
                        (f"${target_val:.2f}mm", "bold"),
                        (f" ({row[change_e_col]*100:.2f}%) while CVaR ", None),
                        (f"{ext_direction}", None),
                        (" by ", None),
                        (f"${ext_change_val:.2f}mm", "bold"),
                        (" to ", None),
                        (f"${ext_target_val:.2f}mm", "bold"),
                        (")", None)
                    ]
                    var_comments.append(var_parts)
                # SVaR comment with extended CSVaR details
                if (row[change_l_col] >= 0.1) and ((row[change_m_col] >= 500000) or (row[change_m_col] <= -500000)):
                    direction = "increased" if row[change_m_col] >= 0 else "decreased"
                    change_val = abs(row[change_m_col]) / 1e6
                    target_val = abs(row[target2_col]) / 1e6
                    ext_direction = "increased" if row[cs_change_col] >= 0 else "decreased"
                    ext_change_val = abs(row[cs_change_col]) / 1e6
                    ext_target_val = abs(row[cs_target_col]) / 1e6
                    svar_parts = [
                        (f"{row[name_col]} SVaR {direction} by ", None),
                        (f"${change_val:.2f}mm", "bold"),
                        (" to ", None),
                        (f"${target_val:.2f}mm", "bold"),
                        (f" ({row[change_l_col]*100:.2f}%) while CSVaR ", None),
                        (f"{ext_direction}", None),
                        (" by ", None),
                        (f"${ext_change_val:.2f}mm", "bold"),
                        (" to ", None),
                        (f"${ext_target_val:.2f}mm", "bold"),
                        (")", None)
                    ]
                    svar_comments.append(svar_parts)

            # Output for this sheet
            output_text.insert(tk.END, f"Sheet {sheet} comments:\n")
            if var_comments:
                output_text.insert(tk.END, "10d VaR change:\n", "bold")
                for parts in var_comments:
                    insert_sentence(parts)
            if svar_comments:
                output_text.insert(tk.END, "10d SVaR change:\n", "bold")
                for parts in svar_comments:
                    insert_sentence(parts)
            output_text.insert(tk.END, "\n")
    except Exception as e:
        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, f"Error: {str(e)}")

root = tk.Tk()
root.title("Excel Analysis")
folder_path = "C:/Reports"  # Update this path to your local folder
date_label = ttk.Label(root, text="Select Date (YYYYMMDD):")
date_label.grid(row=0, column=0, padx=5, pady=5)
date_entry = ttk.Entry(root, width=12)
date_entry.grid(row=0, column=1, padx=5, pady=5)
cal_button = ttk.Button(root, text="📅", command=open_calendar)
cal_button.grid(row=0, column=2, padx=5, pady=5)
run_button = ttk.Button(root, text="Run", command=run_analysis)
run_button.grid(row=0, column=3, padx=5, pady=5)
output_text = tk.Text(root, height=20, width=160)
output_text.grid(row=1, column=0, columnspan=4, padx=5, pady=5)
output_text.tag_config("bold", font=("Helvetica", 10, "bold"))
root.mainloop()
```
