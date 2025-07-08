import pandas as pd
import re

# Load Excel file and specific sheet
file_path = r'C:\Users\<YourUsername>\Documents\BRD_Queries_Appendix.xlsx'  # Update with your actual username if needed
sheet_name = 'CGML_Backtesting'

# Load the sheet into a DataFrame
df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

# Find the column with the most rows starting with "SELECT"
select_counts = df.apply(lambda col: col.dropna().astype(str).str.strip().str.upper().str.startswith('SELECT').sum())
target_column = select_counts.idxmax()

# Filter rows in that column that start with "SELECT"
query_series = df[target_column].dropna().astype(str)
query_series = query_series[query_series.str.strip().str.upper().str.startswith("SELECT")]

# Function to extract column names in brackets [] and tables after FROM
def extract_info(query):
    columns = re.findall(r'\[(.*?)\]', query)
    tables = re.findall(r'\bFROM\s+([\[\]A-Za-z0-9_\.]+)', query, re.IGNORECASE)
    return columns, tables

# Apply function and store results
results = query_series.apply(lambda q: pd.Series(extract_info(q), index=['Columns', 'Tables']))

# Combine original query and the parsed results
final_df = pd.concat([query_series.reset_index(drop=True), results], axis=1)
final_df.columns = ['Query', 'Columns', 'Tables']

# Show results
print(final_df)
