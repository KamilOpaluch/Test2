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
