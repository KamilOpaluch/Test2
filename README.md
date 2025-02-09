DS
import pandas as pd
import tkinter as tk
from tkinter import scrolledtext
import os

# Function to locate the table and rename columns
def locate_and_rename_table(file_path, sheet_name):
    # Load the Excel sheet
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

    # Search for the row containing the headers
    header_row = None
    for i, row in df.iterrows():
        if all(col in row.values for col in ["VaR", "SVaR", "CVaR", "CSVaR", "change", "change %", "Desk ID"]):
            header_row = i
            break

    if header_row is None:
        raise ValueError("Header row not found in the sheet.")

    # Set the header row and drop rows above it
    df.columns = df.iloc[header_row]
    df = df.drop(range(header_row + 1))

    # Rename duplicated columns
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols[cols == dup].index.values.tolist()] = [f"{dup}_{i+1}" for i in range(sum(cols == dup))]
    df.columns = cols

    return df

# Function to analyze the data and generate comments
def analyze_data(df):
    comments = []
    for _, row in df.iterrows():
        desk_id = row["Desk ID"]
        comment_parts = []

        # Check conditions for change_1 and change %_1
        if abs(row["change_1"]) > 500000 and abs(row["change %_1"]) >= 10:
            direction = "increased" if row["change_1"] > 0 else "decreased"
            comment_parts.append(
                f"{desk_id} VaR/SVaR {direction} by {abs(row['change_1']):,.0f} to {row['VaR_1']:,.0f} ({row['change %_1']:.1f}%)"
            )

        # Check conditions for change_2 and change %_2
        if abs(row["change_2"]) > 500000 and abs(row["change %_2"]) >= 10:
            direction = "increased" if row["change_2"] > 0 else "decreased"
            comment_parts.append(
                f"{desk_id} CVaR/CSVaR {direction} by {abs(row['change_2']):,.0f} to {row['CVaR_1']:,.0f} ({row['change %_2']:.1f}%)"
            )

        if comment_parts:
            comments.append(" ".join(comment_parts))

    return comments

# Function to fetch data and display comments
def fetch_and_display():
    file_path = "path_to_your_excel_file.xlsx"  # Replace with your file path
    sheet_name = "sheet1"

    try:
        df = locate_and_rename_table(file_path, sheet_name)
        comments = analyze_data(df)
        comment_text.delete(1.0, tk.END)  # Clear previous comments
        for comment in comments:
            comment_text.insert(tk.END, comment + "\n")
    except Exception as e:
        comment_text.delete(1.0, tk.END)
        comment_text.insert(tk.END, f"Error: {str(e)}")

# Tkinter UI
root = tk.Tk()
root.title("Excel Data Analyzer")

# Textbox to display comments
comment_text = scrolledtext.ScrolledText(root, width=80, height=20, wrap=tk.WORD)
comment_text.pack(padx=10, pady=10)

# Button to fetch data and generate comments
fetch_button = tk.Button(root, text="Fetch Data and Analyze", command=fetch_and_display)
fetch_button.pack(pady=10)

# Run the Tkinter event loop
root.mainloop()


GPT
import tkinter as tk
from tkinter import filedialog, scrolledtext
import pandas as pd

# Function to fetch data, analyze it, and generate comments
def fetch_and_analyze():
    # Get the Excel file path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_path:
        return

    # Read the entire sheet
    df = pd.read_excel(file_path, sheet_name="sheet1", header=None)

    # Identify the start of the desired table by locating the headers
    table_start_row = None
    for idx, row in df.iterrows():
        if (row == "VaR").sum() == 2 and (row == "SVaR").sum() == 2 and (row == "CVaR").sum() == 2 and (row == "CSVaR").sum() == 1 and (row == "change").sum() == 2 and (row == "change %").sum() == 2 and (row == "Desk ID").sum() == 1:
            table_start_row = idx
            break

    if table_start_row is None:
        result_text.insert(tk.END, "Table headers not found.\n")
        return

    # Read the table
    headers = df.iloc[table_start_row].tolist()
    data = df.iloc[table_start_row + 1:].reset_index(drop=True)
    data.columns = headers

    # Rename duplicated columns
    new_columns = []
    column_counts = {}
    for col in headers:
        if col in column_counts:
            column_counts[col] += 1
            new_columns.append(f"{col}_{column_counts[col]}")
        else:
            column_counts[col] = 1
            new_columns.append(col if column_counts[col] == 1 else f"{col}_1")

    data.columns = new_columns

    # Analyze and generate comments
    comments = []
    for _, row in data.iterrows():
        desk_id = row["Desk ID"]
        comment_parts = []

        # Check change_1 and change_%_1
        if row["change_1"] > 500000 and row["change %_1"] >= 10:
            comment_parts.append(f"VaR/SVaR increased/decreased by {row['change_1']:,} to {row['VaR']} ({row['change %_1']}%)")

        # Check change_2 and change_%_2
        if row["change_2"] > 500000 and row["change %_2"] >= 10:
            comment_parts.append(f"CVaR/CSVaR increased/decreased by {row['change_2']:,} to {row['CVaR']} ({row['change %_2']}%)")

        if comment_parts:
            comment = f"{desk_id} - " + " while ".join(comment_parts)
            comments.append(comment)

    # Display comments in the text box
    result_text.delete(1.0, tk.END)  # Clear previous text
    if comments:
        result_text.insert(tk.END, "\n".join(comments))
    else:
        result_text.insert(tk.END, "No significant changes detected.\n")

# Create the Tkinter GUI
root = tk.Tk()
root.title("Excel Table Analyzer")
root.geometry("600x400")

# Instruction label
label = tk.Label(root, text="Click the button to select an Excel file and analyze it.")
label.pack(pady=10)

# Button to trigger the analysis
analyze_button = tk.Button(root, text="Analyze Excel File", command=fetch_and_analyze)
analyze_button.pack(pady=5)

# Text box to display the comments
result_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=70, height=20)
result_text.pack(pady=10)

# Start the GUI event loop
root.mainloop()

